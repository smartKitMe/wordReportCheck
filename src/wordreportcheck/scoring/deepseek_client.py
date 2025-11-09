import json
import os
from typing import List, Any, Dict, Optional
from openai import OpenAI

from ..schemas import ReportItem


SYSTEM_PROMPT = (
    "你是一位严格的教学助教，负责批改学生实验报告。"
    "请对每个题目进行评分，满分100分，并给出简要反馈。"
    "只输出 JSON 数组，数组元素为对象：{id, score, feedback}。"
)


def _ensure_scored(obj: Any, item: ReportItem, raw: str = None) -> Dict[str, Any]:
    """确保评分结果包含 id、score、feedback，异常时提供保底值。"""
    out_id = item.id
    score: Any = None
    feedback: Any = None
    if isinstance(obj, dict):
        out_id = obj.get("id", out_id) or out_id
        score = obj.get("score")
        feedback = obj.get("feedback")

    if not isinstance(score, (int, float)):
        score = 60

    if feedback is None or (isinstance(feedback, str) and not feedback.strip()):
        feedback = "模型未返回规范JSON，已使用保底评分与占位反馈。"

    result: Dict[str, Any] = {
        "id": out_id,
        "score": int(score) if isinstance(score, float) else score,
        "feedback": str(feedback),
    }
    if raw is not None:
        result["raw"] = raw
    return result


def _safe_chat_create(client: OpenAI, **kwargs) -> Any:
    """安全调用 chat.completions.create，异常时返回 None。"""
    try:
        return client.chat.completions.create(**kwargs)
    except Exception:
        return None


def score_items(items: List[ReportItem], api_key: str, model: str = "deepseek-chat") -> Any:
    base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
    client = OpenAI(api_key=api_key, base_url=base_url)

    payload = [
        {
            "id": i.id,
            "question": i.question,
            "answer": i.answer,
        }
        for i in items
    ]

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                "以下是学生的题干与答案，请逐题评分并给出反馈。"
                "严格以 JSON 数组返回，每个元素形如 {id, score, feedback}。\n\n"
                + json.dumps(payload, ensure_ascii=False)
            ),
        },
    ]

    resp = _safe_chat_create(client, model=model, messages=messages, temperature=0)
    if resp is None:
        return [score_item(it, api_key=api_key, model=model) for it in items]
    content = resp.choices[0].message.content

    try:
        parsed = None
        try:
            parsed = json.loads(content)
        except Exception:
            block = _extract_json_block(content)
            if block:
                parsed = json.loads(block)
        if isinstance(parsed, list):
            ensured: List[Dict[str, Any]] = []
            by_id: Dict[str, Any] = {}
            for x in parsed:
                if isinstance(x, dict) and "id" in x:
                    by_id[str(x["id"])]= x
            for idx, it in enumerate(items):
                obj = by_id.get(it.id)
                if obj is None and idx < len(parsed):
                    obj = parsed[idx] if isinstance(parsed[idx], dict) else {}
                ensured.append(_ensure_scored(obj or {}, it))
            return ensured
        else:
            return [score_item(it, api_key=api_key, model=model) for it in items]
    except Exception:
        return [score_item(it, api_key=api_key, model=model) for it in items]


def score_item(item: ReportItem, api_key: str, model: str = "deepseek-chat") -> Any:
    base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
    client = OpenAI(api_key=api_key, base_url=base_url)

    payload = {
        "id": item.id,
        "question": item.question,
        "answer": item.answer,
    }

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                "以下是学生的单个题干与答案，请为该题评分并给出反馈。"
                "严格以 JSON 对象返回，形如 {id, score, feedback}。\n\n"
                + json.dumps(payload, ensure_ascii=False)
            ),
        },
    ]

    resp = _safe_chat_create(client, model=model, messages=messages, temperature=0)
    if resp is None:
        return _ensure_scored({}, item, raw="模型调用失败：可能超出上下文限制或服务端错误")
    content = resp.choices[0].message.content

    try:
        obj = None
        try:
            obj = json.loads(content)
        except Exception:
            block = _extract_json_block(content)
            if block:
                obj = json.loads(block)
        return _ensure_scored(obj if isinstance(obj, dict) else {}, item)
    except Exception:
        return _ensure_scored({}, item, raw=content)


def segment_items_from_content(content: str, expected_count: int, api_key: str, model: str = "deepseek-chat") -> List[ReportItem]:
    """使用 DeepSeek 将实验内容原文分割为指定数量的题目（带重试与分隔符协议）。"""
    base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
    client = OpenAI(api_key=api_key, base_url=base_url)
    # 使用完整原文进行分割，无截断
    content_trunc = content

    SENT_ITEM_BEGIN = "§§§WRC_ITEM_BEGIN§§§"
    SENT_ITEM_END = "§§§WRC_ITEM_END§§§"
    SENT_FIELD_BEGIN = "§§§WRC_FIELD:"
    SENT_FIELD_END = "§§§WRC_FIELD_END§§§"

    base_sys_prompt = (
        "你是一位教学助教，负责解析实验报告的‘实验内容’。"
        "原文示例：题目一：标题\n 题目一：题目要求 xxx \n 题目一：实验方法和步骤 xxx \n 题目1代码：xxx \n 题目1运行结果: xxxx（忽略运行结果）"
        "将整段原文切分为指定数量的题目。每题必须包含以下四个字段：题目名称、题目要求、实验方法和步骤、代码；并且在开头给出编号（id: 题目n）。"
        "严格只使用以下分隔符输出纯文本（不要使用 JSON、不要使用引号、不要使用任何代码块或Markdown）："
        f"每题使用一段：\n{SENT_ITEM_BEGIN}\nid: 题目n\n{SENT_FIELD_BEGIN}题目名称§§§\n(题目名称内容)\n{SENT_FIELD_END}\n{SENT_FIELD_BEGIN}题目要求§§§\n(题目要求内容)\n{SENT_FIELD_END}\n{SENT_FIELD_BEGIN}实验方法和步骤§§§\n(方法和步骤内容)\n{SENT_FIELD_END}\n{SENT_FIELD_BEGIN}代码§§§\n(代码内容)\n{SENT_FIELD_END}\n{SENT_ITEM_END}\n"
        "其中 n 为题号 1..N。严格输出正好 N 段，按 题目1...题目N 的顺序。不得添加任何其他文本。"
        "必须严格以原文中‘题目n’标题为边界进行切分；若原文不足以分出 N 段，可在保持标题边界的前提下将较大的任务合理拆分为多个子题，以满足数量。"
        "每一段都必须包含全部四个字段（题目名称/题目要求/实验方法和步骤/代码），允许空内容但必须保留对应分隔段。"
    )

    base_messages = [
        {"role": "system", "content": base_sys_prompt},
        {
            "role": "user",
            "content": (
                f"请将以下‘实验内容’分割为 {expected_count} 个题目。"
                f"严格按照分隔符协议输出纯文本，不要使用 JSON：\n\n原文：\n\n" + content_trunc
            ),
        },
    ]

    try:
        attempts = int(os.getenv("WORDREPORTCHECK_SEGMENT_RETRY", "3"))
    except Exception:
        attempts = 3
    attempts = max(1, min(attempts, 5))

    last_raw: Optional[str] = None
    last_parsed: Any = None
    last_len: Optional[int] = None

    for attempt in range(attempts):
        if attempt == 0:
            messages = base_messages
        else:
            reinforce_msg = (
                f"上次返回的题目段数为 {last_len}，与期望 {expected_count} 不符。"
                f"请严格输出正好 {expected_count} 段，按 题目1..题目{expected_count} 顺序，每段都包含三部分，"
                "必须严格以原文中‘题目n’标题为边界进行切分，id 写成对应标题（题目1…题目N），"
                "不要输出任何 JSON/引号/多余文本。若不足，请在保持标题边界的前提下合理拆分为多个子题以满足数量。"
            )
            messages = [
                {"role": "system", "content": base_sys_prompt},
                {"role": "user", "content": reinforce_msg},
                base_messages[1],
            ]

        resp = _safe_chat_create(client, model=model, messages=messages, temperature=0, max_tokens=1600)
        if resp is None:
            last_raw = None
            last_parsed = None
            last_len = None
            continue

        raw = resp.choices[0].message.content
        last_raw = raw
        try:
            print("[DeepSeek 输出]", raw)
        except Exception:
            pass

        def _parse_delimited(raw_text: str) -> List[Dict[str, Any]]:
            items: List[Dict[str, Any]] = []
            if not raw_text:
                return items
            begins: List[int] = []
            idx = 0
            while True:
                pos = raw_text.find(SENT_ITEM_BEGIN, idx)
                if pos == -1:
                    break
                begins.append(pos)
                idx = pos + len(SENT_ITEM_BEGIN)
            for i, b in enumerate(begins):
                start = b + len(SENT_ITEM_BEGIN)
                end_search_from = start
                epos = raw_text.find(SENT_ITEM_END, end_search_from)
                block = raw_text[start:] if epos == -1 else raw_text[start:epos]
                rid = None
                for line in block.splitlines():
                    t = line.strip()
                    if t.lower().startswith("id:"):
                        rid = t.split(":", 1)[1].strip()
                        break
                def _extract_field(name: str) -> str:
                    marker = f"{SENT_FIELD_BEGIN}{name}§§§"
                    if marker not in block:
                        return ""
                    after = block.split(marker, 1)[1]
                    if SENT_FIELD_END in after:
                        return after.split(SENT_FIELD_END, 1)[0].strip()
                    return after.strip()
                obj = {
                    "id": rid,
                    "题目名称": _extract_field("题目名称"),
                    "题目要求": _extract_field("题目要求"),
                    "实验方法和步骤": _extract_field("实验方法和步骤"),
                    "代码": _extract_field("代码"),
                }
                items.append(obj)
            return items

        result_items: Optional[List[ReportItem]] = None
        parsed: Any = None
        try:
            delimited = _parse_delimited(raw)
            if delimited:
                parsed = delimited
            if parsed is None:
                try:
                    parsed = json.loads(raw)
                except Exception:
                    parsed = None
            if parsed is None:
                block = _extract_json_block(raw)
                if block:
                    try:
                        parsed = json.loads(block)
                    except Exception:
                        parsed = None
            if isinstance(parsed, dict):
                candidates = []
                for key in ("items", "segments", "data"):
                    if key in parsed:
                        candidates.append(parsed.get(key))
                arr = None
                for c in candidates:
                    if isinstance(c, list):
                        arr = c
                        break
                    if isinstance(c, str):
                        try:
                            loaded = json.loads(c)
                            if isinstance(loaded, list):
                                arr = loaded
                                break
                        except Exception:
                            pass
                    if isinstance(c, dict):
                        values = list(c.values())
                        try:
                            values.sort(key=lambda x: int(str(x.get("id", "").lstrip("Q题目")) or 0))
                        except Exception:
                            pass
                        arr = values
                        break
                if arr is not None:
                    parsed = arr
                else:
                    keys = set(parsed.keys())
                    if {"id", "题目名称", "题目要求", "实验方法和步骤", "代码"}.issubset(keys) and expected_count == 1:
                        parsed = [parsed]
                    else:
                        values = list(parsed.values())
                        if values and all(isinstance(v, dict) for v in values):
                            try:
                                def _key_order(k: str) -> int:
                                    try:
                                        return int(str(k).lstrip('Q题目'))
                                    except Exception:
                                        return 0
                                parsed = [parsed[k] for k in sorted(parsed.keys(), key=_key_order)]
                            except Exception:
                                parsed = values
            if not isinstance(parsed, list):
                block = _extract_json_block(raw)
                if block:
                    try:
                        tmp = json.loads(block)
                        if isinstance(tmp, list):
                            parsed = tmp
                    except Exception:
                        pass
            if not isinstance(parsed, list):
                last_parsed = parsed
                last_len = None
                continue
            if len(parsed) != expected_count:
                # 若使用了分隔符解析且数量不足，尝试本地启发式拆分最长段以填充到 N
                try:
                    def _length_of(x: Dict[str, Any]) -> int:
                        if not isinstance(x, dict):
                            return 0
                        ql = len(str(x.get("题目要求") or x.get("题目") or x.get("question") or ""))
                        ml = len(str(x.get("实验方法和步骤") or x.get("方法和步骤") or x.get("steps") or x.get("methods") or ""))
                        cl = len(str(x.get("代码") or x.get("code") or ""))
                        al = len(str(x.get("answer") or ""))
                        return ql + ml + cl + al

                    def _split_by_lines(text: str) -> List[str]:
                        lines = [l for l in str(text).splitlines()]
                        if len(lines) <= 1:
                            t = str(text)
                            mid = max(1, len(t) // 2)
                            return [t[:mid], t[mid:]]
                        mid = max(1, len(lines) // 2)
                        return ["\n".join(lines[:mid]).strip(), "\n".join(lines[mid:]).strip()]

                    deficit = expected_count - len(parsed)
                    # 仅当 parsed 为 list 且 deficit > 0 时进行启发式拆分
                    if isinstance(parsed, list) and deficit > 0:
                        work = list(parsed)
                        while deficit > 0 and work:
                            # 选择当前内容最长的一段进行拆分
                            idx_long = 0
                            max_len = -1
                            for i, seg in enumerate(work):
                                l = _length_of(seg if isinstance(seg, dict) else {})
                                if l > max_len:
                                    max_len = l
                                    idx_long = i
                            seg = work.pop(idx_long)
                            if not isinstance(seg, dict):
                                seg = {}
                            q0 = seg.get("题目要求") or seg.get("题目") or seg.get("question") or ""
                            m0 = seg.get("实验方法和步骤") or seg.get("方法和步骤") or seg.get("steps") or seg.get("methods") or ""
                            c0 = seg.get("代码") or seg.get("code") or ""
                            a0 = seg.get("answer") or ""
                            # 优先对步骤拆分，其次代码，最后直接把 answer 拆分
                            parts = None
                            if m0 and len(m0) > 50:
                                left, right = _split_by_lines(m0)
                                parts = [
                                    {"题目要求": q0, "实验方法和步骤": left, "代码": c0, "answer": a0},
                                    {"题目要求": q0, "实验方法和步骤": right, "代码": c0, "answer": a0},
                                ]
                            elif c0 and len(c0) > 80:
                                left, right = _split_by_lines(c0)
                                parts = [
                                    {"题目要求": q0, "实验方法和步骤": m0, "代码": left, "answer": a0},
                                    {"题目要求": q0, "实验方法和步骤": m0, "代码": right, "answer": a0},
                                ]
                            elif a0 and len(a0) > 50:
                                left, right = _split_by_lines(a0)
                                parts = [
                                    {"题目要求": q0, "实验方法和步骤": m0, "代码": c0, "answer": left},
                                    {"题目要求": q0, "实验方法和步骤": m0, "代码": c0, "answer": right},
                                ]
                            else:
                                # 无法有效拆分，则复制占位补足一题
                                parts = [seg, {"题目要求": q0, "实验方法和步骤": m0, "代码": c0, "answer": ""}]

                            # 将拆分出的两段加入工作集末尾，直到凑够数量
                            work.extend(parts)
                            if len(work) >= expected_count:
                                parsed = work[:expected_count]
                                break
                            else:
                                deficit = expected_count - len(work)
                                parsed = work

                    # 若仍不足，进入下一次重试；若已满足，继续构建结果
                    if not isinstance(parsed, list) or len(parsed) != expected_count:
                        last_parsed = parsed
                        try:
                            last_len = len(parsed) if isinstance(parsed, list) else None
                        except Exception:
                            last_len = None
                        continue
                except Exception:
                    last_parsed = parsed
                    try:
                        last_len = len(parsed) if isinstance(parsed, list) else None
                    except Exception:
                        last_len = None
                    continue
            items: List[ReportItem] = []
            for idx, x in enumerate(parsed):
                if not isinstance(x, dict):
                    x = {}
                rid = x.get("id") or f"题目{idx + 1}"
                title = x.get("题目名称") or ""
                q = x.get("题目要求") or x.get("题目") or x.get("question") or ""
                methods = x.get("实验方法和步骤") or x.get("方法和步骤") or x.get("steps") or x.get("methods") or ""
                code = x.get("代码") or x.get("code") or ""
                parts = []
                if methods:
                    parts.append(f"实验方法和步骤：{methods}")
                if code:
                    parts.append(f"代码：{code}")
                a = "\n".join(parts) or x.get("answer") or ""
                # 题目名称存入 ReportItem 的可选字段（在 schemas 中定义）
                try:
                    items.append(ReportItem(id=str(rid), question=str(q), answer=str(a), methods=str(methods) if methods else None, code=str(code) if code else None, title=str(title) if title else None))
                except TypeError:
                    # 若当前版本的 ReportItem 尚未包含 title/methods/code，可退化为旧构造以避免崩溃
                    items.append(ReportItem(id=str(rid), question=str(q), answer=str(a)))
            result_items = items
        except Exception:
            try:
                last_len = len(parsed) if isinstance(parsed, list) else None
            except Exception:
                last_len = None
            last_parsed = parsed
            continue

        if result_items is not None:
            return result_items

    if last_len is None:
        raise ValueError("模型返回非数组或无法解析题目分段")
    else:
        raise ValueError(f"题目数量不匹配：期望 {expected_count}，实际 {last_len}")


def _extract_json_block(text: str):
    if not text:
        return None
    for open_ch, close_ch in (('[', ']'), ('{', '}')):
        start = text.find(open_ch)
        if start == -1:
            continue
        depth = 0
        for i in range(start, len(text)):
            ch = text[i]
            if ch == open_ch:
                depth += 1
            elif ch == close_ch:
                depth -= 1
                if depth == 0:
                    return text[start : i + 1]
    return None

# 已移除截断功能：始终使用完整文本，避免任何长度裁剪