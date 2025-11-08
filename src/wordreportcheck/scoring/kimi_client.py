import json
import os
from typing import List, Any, Dict
from openai import OpenAI
from typing import Optional, Tuple

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


def score_items(items: List[ReportItem], api_key: str, model: str = "moonshot-v1-8k") -> Any:
    # Kimi (Moonshot) 使用 OpenAI 兼容接口，设置 base_url 即可
    base_url = os.getenv("MOONSHOT_BASE_URL", "https://api.moonshot.cn/v1")
    client = OpenAI(api_key=api_key, base_url=base_url)

    # 输入截断以避免超过模型的上下文限制
    max_input_chars = int(os.getenv("WORDREPORTCHECK_MAX_INPUT_CHARS", "8000"))
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


def score_item(item: ReportItem, api_key: str, model: str = "moonshot-v1-8k") -> Any:
    base_url = os.getenv("MOONSHOT_BASE_URL", "https://api.moonshot.cn/v1")
    client = OpenAI(api_key=api_key, base_url=base_url)

    # 输入截断以避免超过模型的上下文限制
    max_input_chars = int(os.getenv("WORDREPORTCHECK_MAX_INPUT_CHARS", "8000"))
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