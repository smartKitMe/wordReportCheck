import json
from typing import List, Any
from openai import OpenAI

from ..schemas import ReportItem


SYSTEM_PROMPT = (
    "你是一位严格的教学助教，负责批改学生实验报告。"
    "请对每个题目进行评分，满分100分，并给出简要反馈。"
    "只输出 JSON 数组，数组元素为对象：{id, score, feedback}。"
)


def score_items(items: List[ReportItem], api_key: str, model: str = "deepseek-chat") -> Any:
    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

    payload = [
        {"id": i.id, "question": i.question, "answer": i.answer} for i in items
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

    resp = client.chat.completions.create(model=model, messages=messages, temperature=0)
    content = resp.choices[0].message.content

    try:
        return json.loads(content)
    except Exception:
        # 若模型未严格返回 JSON，则封装为兼容对象
        return {"raw": content}


def score_item(item: ReportItem, api_key: str, model: str = "deepseek-chat") -> Any:
    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

    payload = {"id": item.id, "question": item.question, "answer": item.answer}

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

    resp = client.chat.completions.create(model=model, messages=messages, temperature=0)
    content = resp.choices[0].message.content

    try:
        obj = json.loads(content)
        # 确保包含 id
        if isinstance(obj, dict) and "id" not in obj:
            obj["id"] = item.id
        return obj
    except Exception:
        return {"id": item.id, "raw": content}