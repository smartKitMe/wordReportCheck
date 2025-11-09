from dataclasses import dataclass, asdict
from typing import List, Optional, Dict, Any
import json


@dataclass
class ReportItem:
    id: str
    question: str
    answer: str
    # 新增：为更细粒度的 JSON 输出保留分割字段
    methods: Optional[str] = None
    code: Optional[str] = None
    # 新增：题目名称（与题目要求区分，用于展示与导出）
    title: Optional[str] = None


@dataclass
class ReportDocument:
    学院信息: Optional[str] = None
    专业信息: Optional[str] = None
    时间: Optional[str] = None
    姓名: Optional[str] = None
    学号: Optional[str] = None
    班级: Optional[str] = None
    指导老师: Optional[str] = None
    课程名称: Optional[str] = None
    周次: Optional[str] = None
    实验名称: Optional[str] = None
    实验环境: Optional[str] = None
    实验分析与体会: Optional[str] = None
    实验日期: Optional[str] = None
    备注: Optional[str] = None
    成绩: Optional[str] = None
    签名: Optional[str] = None
    日期: Optional[str] = None
    # 原始“实验内容”全文（不经分割，用于发送到 AI）
    实验内容原文: Optional[str] = None
    # 实验内容（题干与答案）
    content_items: List[ReportItem] = None  # type: ignore

    def to_json_obj(self) -> Dict[str, Any]:
        # 自定义 items 的序列化：同时输出英文与中文字段，满足评分与展示需求
        items_out: List[Dict[str, Any]] = []
        for i in (self.content_items or []):
            item_dict: Dict[str, Any] = {
                "id": i.id,
                # 兼容：保留英文键用于评分逻辑
                "question": i.question,
                "answer": i.answer,
                # 中文展示字段（满足用户要求的键名）
                "编号": i.id,
                "题目名称": i.title,
                "题目要求": i.question,
                "实验方法和步骤": i.methods,
                "代码": i.code,
            }
            items_out.append(item_dict)

        obj: Dict[str, Any] = {
            "学院信息": self.学院信息,
            "专业信息": self.专业信息,
            "时间": self.时间,
            "姓名": self.姓名,
            "学号": self.学号,
            "班级": self.班级,
            "指导老师": self.指导老师,
            "课程名称": self.课程名称,
            "周次": self.周次,
            "实验名称": self.实验名称,
            "实验环境": self.实验环境,
            # 同时输出原文与分割后的 items，便于后续 AI 分割或调试
            "实验内容": {
                "raw": self.实验内容原文,
                "items": items_out,
            },
            "实验分析与体会": self.实验分析与体会,
            "实验日期": self.实验日期,
            "备注": self.备注,
            "成绩": self.成绩,
            "签名": self.签名,
            "日期": self.日期,
        }
        return obj


def report_to_json(report: ReportDocument) -> str:
    return json.dumps(report.to_json_obj(), ensure_ascii=False, indent=2)


def report_from_json(json_str: str) -> ReportDocument:
    data = json.loads(json_str)
    items_data = []
    content_raw: Optional[str] = None
    # 兼容：可能直接提供顶层 items 或嵌套在 "实验内容" 下
    if isinstance(data.get("实验内容"), dict):
        content_raw = data["实验内容"].get("raw")
        items_data = data["实验内容"].get("items", [])
    elif isinstance(data.get("items"), list):
        items_data = data.get("items", [])

    items: List[ReportItem] = []
    for i in items_data:
        # 优先从中文键读取，兼容英文键
        title = i.get("题目名称") or i.get("题目")
        q_cn = i.get("题目要求") or i.get("题目")
        methods = i.get("实验方法和步骤") or i.get("方法和步骤") or i.get("methods")
        code = i.get("代码") or i.get("code")
        q = (q_cn or i.get("question", ""))
        # 若未提供英文 answer，则由 methods+code 拼接生成
        ans = i.get("answer")
        if not ans:
            parts = []
            if methods:
                parts.append(f"实验方法和步骤：{methods}")
            if code:
                parts.append(f"代码：{code}")
            ans = "\n".join(parts)
        items.append(ReportItem(id=i.get("id", ""), question=q or "", answer=ans or "", methods=methods, code=code, title=title))

    return ReportDocument(
        学院信息=data.get("学院信息"),
        专业信息=data.get("专业信息"),
        时间=data.get("时间"),
        姓名=data.get("姓名"),
        学号=data.get("学号"),
        班级=data.get("班级"),
        指导老师=data.get("指导老师"),
        课程名称=data.get("课程名称"),
        周次=data.get("周次"),
        实验名称=data.get("实验名称"),
        实验环境=data.get("实验环境"),
        实验分析与体会=data.get("实验分析与体会"),
        实验日期=data.get("实验日期"),
        备注=data.get("备注"),
        成绩=data.get("成绩"),
        签名=data.get("签名"),
        日期=data.get("日期"),
        实验内容原文=content_raw,
        content_items=items,
    )


def items_to_json(items: List[ReportItem]) -> str:
    return json.dumps([asdict(i) for i in items], ensure_ascii=False, indent=2)


def items_from_json(json_str: str) -> List[ReportItem]:
    raw = json.loads(json_str)
    items: List[ReportItem] = []
    for i in raw:
        items.append(ReportItem(id=i.get("id", ""), question=i.get("question", ""), answer=i.get("answer", "")))
    return items