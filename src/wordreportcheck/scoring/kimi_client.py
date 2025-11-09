import json
import os
from typing import List, Any, Dict
from openai import OpenAI
from typing import Optional, Tuple
from pathlib import Path

from ..schemas import ReportItem


SYSTEM_PROMPT = (
    "你是一位严格的教学助教，负责批改学生实验报告。"
    "请对每个题目进行评分，满分100分，并给出简要反馈。"
    "只输出 JSON，严格遵守：不允许任何解释、前后缀、Markdown 或代码块。"
)


def _ensure_scored(obj: Any, item: ReportItem, raw: str = None) -> Dict[str, Any]:
    """Ensure the result contains id, numeric score, and feedback.
    If missing or unparsable, fill with fallback values and keep raw if available.
    """
    out_id = item.id
    score: Any = None
    feedback: Any = None
    if isinstance(obj, dict):
        out_id = obj.get("id", out_id) or out_id
        score = obj.get("score")
        feedback = obj.get("feedback")

    # Normalize score
    if not isinstance(score, (int, float)):
        # fallback score
        score = 60

    # Normalize feedback
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


def _extract_json_block(text: str) -> Optional[str]:
    """Extract a JSON array/object substring from free-form text by bracket matching.
    Returns the best-effort block or None if not found.
    """
    if not text:
        return None
    # Prefer array extraction first (we expect a list of items for batch scoring)
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

def _sanitize_json_like(text: str) -> str:
    """Best-effort sanitize of a JSON-like string:
    - Converts newlines within string literals to escaped "\\n" to avoid control-character errors
    - Removes trailing commas before closing brackets/brace
    """
    if not text:
        return ""
    out_chars = []
    in_string = False
    escaped = False
    for ch in text:
        if in_string:
            if escaped:
                # Previous char was a backslash; current char is escaped
                out_chars.append(ch)
                escaped = False
                continue
            if ch == "\\":
                out_chars.append(ch)
                escaped = True
                continue
            if ch == '"':
                out_chars.append(ch)
                in_string = False
                continue
            if ch == "\n" or ch == "\r":
                out_chars.append("\\n")
                continue
            out_chars.append(ch)
        else:
            if ch == '"':
                out_chars.append(ch)
                in_string = True
                escaped = False
            else:
                out_chars.append(ch)
    sanitized = "".join(out_chars)
    # Remove trailing commas before array/object close
    sanitized = sanitized.replace(",]", "]").replace(",}", "}")
    return sanitized


def _truncate(text: str, max_chars: int) -> str:
    try:
        t = (text or "")
        if len(t) <= max_chars:
            return t
        return t[:max_chars] + "\n[内容过长，已截断，仅评估上述片段]"
    except Exception:
        return (text or "")


def _safe_chat_create(client: OpenAI, **kwargs) -> Any:
    """安全调用 chat.completions.create，异常时返回 None。"""
    try:
        return client.chat.completions.create(**kwargs)
    except Exception:
        return None


def score_items(items: List[ReportItem], api_key: str, model: str = "moonshot-v1-128k") -> Any:
    # Kimi (Moonshot) 使用 OpenAI 兼容接口，设置 base_url 即可
    base_url = os.getenv("MOONSHOT_BASE_URL", "https://api.moonshot.cn/v1")
    client = OpenAI(api_key=api_key, base_url=base_url)

    # 输入截断以避免超过模型的上下文限制
    max_input_chars = int(os.getenv("WORDREPORTCHECK_MAX_INPUT_CHARS", "128000"))
    payload = [
        {
            "id": i.id,
            "question": _truncate(i.question, max_input_chars // 4),
            "answer": _truncate(i.answer, max_input_chars),
        }
        for i in items
    ]

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT + "仅返回 JSON 数组，每个元素为 {id, score, feedback}。"},
        {
            "role": "user",
            "content": (
                "以下是学生的题干与答案，请逐题评分并给出反馈。"
                "只返回 JSON 数组，不要任何解释或标记。\n\n"
                + json.dumps(payload, ensure_ascii=False)
            ),
        },
    ]

    # 尝试启用 JSON 响应模式；若不被服务端支持，则回退
    # 先尝试 JSON 响应模式，失败则回退普通文本模式；若两次都失败，则逐题回退
    resp = _safe_chat_create(
        client,
        model=model,
        messages=messages,
        temperature=0,
        response_format={"type": "json_object"},
        max_tokens=1024,
    )
    if resp is None:
        resp = _safe_chat_create(
            client,
            model=model,
            messages=messages,
            temperature=0,
            max_tokens=1024,
        )
    if resp is None:
        # Fallback: per-item calls
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
        # Expect a list; ensure each mapped item has score/feedback
        if isinstance(parsed, list):
            ensured: List[Dict[str, Any]] = []
            # try align by id; fallback by index
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
            # Fallback: per-item calls to guarantee outputs
            return [score_item(it, api_key=api_key, model=model) for it in items]
    except Exception:
        # Fallback: per-item calls
        return [score_item(it, api_key=api_key, model=model) for it in items]


def score_item(item: ReportItem, api_key: str, model: str = "moonshot-v1-128k") -> Any:
    base_url = os.getenv("MOONSHOT_BASE_URL", "https://api.moonshot.cn/v1")
    client = OpenAI(api_key=api_key, base_url=base_url)

    # 输入截断以避免超过模型的上下文限制
    max_input_chars = int(os.getenv("WORDREPORTCHECK_MAX_INPUT_CHARS", "128000"))
    payload = {
        "id": item.id,
        "question": _truncate(item.question, max_input_chars // 4),
        "answer": _truncate(item.answer, max_input_chars),
    }

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT + "仅返回 JSON 对象 {id, score, feedback}。"},
        {
            "role": "user",
            "content": (
                "以下是学生的单个题干与答案，请为该题评分并给出反馈。"
                "只返回 JSON 对象，不要任何解释或标记。\n\n"
                + json.dumps(payload, ensure_ascii=False)
            ),
        },
    ]

    resp = _safe_chat_create(
        client,
        model=model,
        messages=messages,
        temperature=0,
        response_format={"type": "json_object"},
        max_tokens=512,
    )
    if resp is None:
        resp = _safe_chat_create(
            client,
            model=model,
            messages=messages,
            temperature=0,
            max_tokens=512,
        )
    if resp is None:
        # 构造保底返回（未调用到模型）
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


def segment_items_from_content(content: str, expected_count: int, api_key: str, model: str = "moonshot-v1-128k") -> List[ReportItem]:
    """使用 Kimi/Moonshot 将实验内容原文分割为指定数量的题目。

    加强版：使用分隔符协议并增加重试。当返回题目数量不符时，追加更强约束提示重新生成，最多重试次数由
    环境变量 WORDREPORTCHECK_SEGMENT_RETRY 控制（默认 3 次）。
    """
    base_url = os.getenv("MOONSHOT_BASE_URL", "https://api.moonshot.cn/v1")
    client = OpenAI(api_key=api_key, base_url=base_url)

    max_input_chars = int(os.getenv("WORDREPORTCHECK_MAX_INPUT_CHARS", "128000"))
    content_trunc = _truncate(content, max_input_chars)

    # 采用稀有分隔符的纯文本协议，避免 JSON/引号带来的解析问题
    SENT_ITEM_BEGIN = "§§§WRC_ITEM_BEGIN§§§"
    SENT_ITEM_END = "§§§WRC_ITEM_END§§§"
    SENT_FIELD_BEGIN = "§§§WRC_FIELD:"
    SENT_FIELD_END = "§§§WRC_FIELD_END§§§"

    base_sys_prompt = (
        "你是一位教学助教，负责解析实验报告的‘实验内容’。"
        "将整段原文切分为指定数量的题目，每题包含三部分：题目要求、实验方法和步骤、代码。"
        "严格只使用以下分隔符输出纯文本（不要使用 JSON、不要使用引号、不要使用任何代码块或Markdown）："
        f"每题使用一段：\n{SENT_ITEM_BEGIN}\nid: Qn\n{SENT_FIELD_BEGIN}题目要求§§§\n(题目要求内容)\n{SENT_FIELD_END}\n{SENT_FIELD_BEGIN}实验方法和步骤§§§\n(方法和步骤内容)\n{SENT_FIELD_END}\n{SENT_FIELD_BEGIN}代码§§§\n(代码内容)\n{SENT_FIELD_END}\n{SENT_ITEM_END}\n"
        "其中 n 为题号 1..N。严格输出正好 N 段，按 Q1...QN 的顺序。不得添加任何其他文本。"
        "如果原文中题目数量不足以达到 N，请将较大的任务合理拆分为多个独立的子题，确保总数为 N。"
        "每一段都必须包含三个字段（题目要求/实验方法和步骤/代码），允许空内容但必须保留对应分隔段。"
    )

    # 基础消息（首轮）：
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

    # 重试次数（默认 3）
    try:
        attempts = int(os.getenv("WORDREPORTCHECK_SEGMENT_RETRY", "3"))
    except Exception:
        attempts = 3
    attempts = max(1, min(attempts, 5))

    last_raw = None
    last_parsed: Any = None
    last_len = None
    for attempt in range(attempts):
        # 为后续重试构造更强约束提示
        if attempt == 0:
            messages = base_messages
        else:
            reinforce_msg = (
                f"上次返回的题目段数为 {last_len}，与期望 {expected_count} 不符。"
                f"请严格输出正好 {expected_count} 段，按 Q1..Q{expected_count} 顺序，每段都包含三部分，"
                "只使用分隔符协议，不要输出任何 JSON/引号/多余文本。若不足，请将问题拆分为多个子题以满足数量。"
            )
            messages = [
                {"role": "system", "content": base_sys_prompt},
                {"role": "user", "content": reinforce_msg},
                base_messages[1],
            ]

        resp = _safe_chat_create(
            client,
            model=model,
            messages=messages,
            temperature=0,
            max_tokens=1600,
        )
        if resp is None:
            # 进入下一次重试
            last_raw = None
            last_parsed = None
            last_len = None
            continue

        raw = resp.choices[0].message.content
        last_raw = raw
        # 打印模型的原始输出，便于调试
        try:
            print("[Kimi/Moonshot 输出]", raw)
        except Exception:
            pass

        def _parse_delimited(raw_text: str) -> List[Dict[str, Any]]:
            items: List[Dict[str, Any]] = []
            if not raw_text:
                return items
            # 扫描每个 BEGIN，并截取到对应 END；若末尾缺失 END，则取到文本末尾
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
                if epos == -1:
                    block = raw_text[start:]
                else:
                    block = raw_text[start:epos]
                # 读取 id
                rid = None
                # 寻找以 id: 开头的行
                for line in block.splitlines():
                    t = line.strip()
                    if t.lower().startswith("id:"):
                        rid = t.split(":", 1)[1].strip()
                        break
                # 读取三个字段
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
                    "题目要求": _extract_field("题目要求"),
                    "实验方法和步骤": _extract_field("实验方法和步骤"),
                    "代码": _extract_field("代码"),
                }
                items.append(obj)
            return items
        result_items: Optional[List[ReportItem]] = None
        parsed: Any = None
        try:
            # 0) 优先尝试分隔符解析
            delimited = _parse_delimited(raw)
            if delimited:
                parsed = delimited
            # 1) 直接尝试解析完整响应
            if parsed is None:
                try:
                    parsed = json.loads(raw)
                except Exception:
                    parsed = None
            # 2) 若失败，尝试提取文本中的 JSON 片段再解析
            if parsed is None:
                block = _extract_json_block(raw)
                if block:
                    try:
                        parsed = json.loads(block)
                    except Exception:
                        parsed = None
            # 3) 接受对象形式并从 items/segments 中取数组（容忍字典或字符串形式）
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
                        # 可能返回 {"items": {"Q1": {...}, ...}}
                        values = list(c.values())
                        # 尝试按 Q 序号排序
                        try:
                            values.sort(key=lambda x: int(str(x.get("id", "").lstrip("Q")) or 0))
                        except Exception:
                            pass
                        arr = values
                        break
                if arr is not None:
                    parsed = arr
                else:
                    # 若字典本身像一个题目对象且期望数量为1，则包一层数组
                    keys = set(parsed.keys())
                    if {"id", "题目要求", "实验方法和步骤", "代码"}.issubset(keys) and expected_count == 1:
                        parsed = [parsed]
                    else:
                        # 顶层可能是 {"Q1": {...}, "Q2": {...}}
                        values = list(parsed.values())
                        if values and all(isinstance(v, dict) for v in values):
                            try:
                                # 尝试按键中的序号排序
                                def _key_order(k: str) -> int:
                                    try:
                                        return int(str(k).lstrip('Q'))
                                    except Exception:
                                        return 0
                                ordered = [parsed[k] for k in sorted(parsed.keys(), key=_key_order)]
                                parsed = ordered
                            except Exception:
                                parsed = values
            # 4) 不为数组则尝试更宽松的提取
            if not isinstance(parsed, list):
                block = _extract_json_block(raw)
                if block:
                    try:
                        tmp = json.loads(block)
                        if isinstance(tmp, list):
                            parsed = tmp
                    except Exception:
                        try:
                            tmp2 = json.loads(_sanitize_json_like(block))
                            if isinstance(tmp2, list):
                                parsed = tmp2
                        except Exception:
                            pass
                else:
                    try:
                        tmp3 = json.loads(_sanitize_json_like(raw))
                        if isinstance(tmp3, list):
                            parsed = tmp3
                    except Exception:
                        pass
            # 5) 校验并转换
            if not isinstance(parsed, list):
                raise ValueError("解析结果非数组")
            if len(parsed) != expected_count:
                raise ValueError(f"题目数量不匹配：期望 {expected_count}，实际 {len(parsed)}")
            items: List[ReportItem] = []
            for idx, x in enumerate(parsed):
                if not isinstance(x, dict):
                    x = {}
                rid = x.get("id") or f"Q{idx + 1}"
                q = x.get("题目要求") or x.get("题目") or x.get("question") or ""
                methods = x.get("实验方法和步骤") or x.get("方法和步骤") or x.get("steps") or x.get("methods") or ""
                code = x.get("代码") or x.get("code") or ""
                parts = []
                if methods:
                    parts.append(f"实验方法和步骤：{methods}")
                if code:
                    parts.append(f"代码：{code}")
                a = "\n".join(parts) or x.get("answer") or ""
                items.append(ReportItem(id=str(rid), question=str(q), answer=str(a), methods=str(methods) if methods is not None else None, code=str(code) if code is not None else None))
            result_items = items
        except Exception:
            # 解析失败，记录并准备重试
            try:
                last_len = len(parsed) if isinstance(parsed, list) else None
            except Exception:
                last_len = None
            last_parsed = parsed
            continue
        # 若成功，直接返回
        if result_items is not None:
            return result_items

    # 全部尝试失败，构造更清晰的错误信息
    if last_len is None:
        raise ValueError("模型返回非数组或无法解析题目分段")
    else:
        raise ValueError(f"题目数量不匹配：期望 {expected_count}，实际 {last_len}")
def _load_env_file() -> None:
    """Load .env from current working directory or project root without overriding existing envs."""
    try:
        candidates = [
            Path.cwd() / ".env",
            Path(__file__).resolve().parents[2] / ".env",
        ]
        for env_path in candidates:
            if env_path.exists():
                for raw in env_path.read_text(encoding="utf-8").splitlines():
                    line = raw.strip()
                    if not line or line.startswith("#"):
                        continue
                    if line.startswith("export "):
                        line = line[len("export ") :].strip()
                    if "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    if key and (key not in os.environ):
                        os.environ[key] = value
                break
    except Exception:
        pass

# 默认尝试加载 .env（不会覆盖已有环境变量）
_load_env_file()