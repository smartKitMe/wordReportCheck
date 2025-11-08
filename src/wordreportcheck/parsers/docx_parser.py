from pathlib import Path
from typing import List, Optional, Dict, Tuple
import re
from docx import Document

from ..schemas import ReportItem, ReportDocument


def _strip(text: Optional[str]) -> str:
    return (text or "").strip()


LABEL_MAP: Dict[str, str] = {
    "学院": "学院信息",
    "学院信息": "学院信息",
    "专业": "专业信息",
    "专业信息": "专业信息",
    "时间": "时间",
    "姓名": "姓名",
    "学号": "学号",
    "班级": "班级",
    "指导老师": "指导老师",
    "课程名称": "课程名称",
    "周次": "周次",
    "实验名称": "实验名称",
    "实验环境": "实验环境",
    "实验内容": "实验内容",
    "实验分析与体会": "实验分析与体会",
    "实验日期": "实验日期",
    "备注": "备注",
    "成绩": "成绩",
    "签名": "签名",
    "日期": "日期",
}


def _normalize_label(text: str) -> str:
    t = _strip(text)
    # 去掉末尾的冒号/全角冒号及空白
    t = re.sub(r"[：:]\s*$", "", t)
    # 替换常见空格/制表符
    t = re.sub(r"\s+", "", t)
    return LABEL_MAP.get(t, t)


def _extract_row_label_value(row) -> Optional[Dict[str, str]]:
    cells = row.cells
    if not cells:
        return None
    if len(cells) >= 2:
        label = _normalize_label(cells[0].text)
        value = _strip(cells[1].text)
        return {"label": label, "value": value}
    # 单列情况：尝试按冒号分割
    raw = _strip(cells[0].text)
    parts = re.split(r"[：:]", raw, maxsplit=1)
    if len(parts) == 2:
        label = _normalize_label(parts[0])
        value = _strip(parts[1])
        return {"label": label, "value": value}
    return None


def _parse_content_items(content_text: str) -> List[ReportItem]:
    text = content_text.replace("\r\n", "\n").replace("\r", "\n")
    # 识别“题目X”标题；允许前缀括号说明，标题行到下一个标题之间视为该题的回答内容
    pattern = re.compile(r"(^|\n)(?:（[^）]*）)?\s*题目\s*(\d+)[：:](.*?)(?=\n(?:（[^）]*）)?\s*题目\s*\d+[：:]|\Z)", re.S)
    items: List[ReportItem] = []
    for m in pattern.finditer(text):
        num = m.group(2)
        title_tail = _strip(m.group(3))
        # 将第一行标题尾部作为题干，其余作为回答
        # 题干：截取到首个换行（若有）；回答：其余文本
        if "\n" in title_tail:
            first_line, rest = title_tail.split("\n", 1)
            question = _strip(f"题目{num}：{first_line}")
            answer = _strip(rest)
        else:
            question = _strip(f"题目{num}：{title_tail}")
            # 标题后若无内容，则回答为空字符串
            answer = ""
        items.append(ReportItem(id=f"Q{num}", question=question, answer=answer))
    # 若未识别到题目结构，退化为整段作为一个“未知题目”项（避免空输出）
    if not items and content_text.strip():
        items.append(ReportItem(id="Q1", question="实验内容", answer=_strip(content_text)))
    return items


def _get_grid_span(cell) -> int:
    # 兼容 python-docx 的 grid_span 属性，不存在时按 1 处理
    return getattr(cell, "grid_span", 1)


def _parse_by_template(doc: Document) -> Optional[ReportDocument]:
    """按照用户提供的统一模板（行号 + grid_span 顺序）解析 18 项字段。

    该模板特点：
    - 顶部若干行是信息汇总行，单元格可能合并，需用 grid_span 累积列索引
    - “实验内容”从包含该关键词的行开始，直到出现“实验分析与体会”为止
    - “实验分析与体会”之后的若干行依次为：实验日期、备注、成绩、（可能有签名行）、日期
    """
    if not doc.tables:
        return None

    table = doc.tables[0]
    rows = table.rows
    if not rows:
        return None

    report = ReportDocument(content_items=[])

    def set_field(name: str, value: str):
        try:
            setattr(report, name, _strip(value))
        except Exception:
            pass

    content_started = False
    analys_row_idx: Optional[int] = None
    content_buffer: List[str] = []

    for row_idx in range(len(rows)):
        row = rows[row_idx]
        # 行文本聚合（用于关键词检测）
        row_text = " ".join([_strip(c.text) for c in row.cells])

        # 先处理顶部信息汇总区（参考提供脚本的行号与索引）
        grid_span_index = 0
        col_idx = 0
        while col_idx < len(row.cells):
            cell = row.cells[col_idx]
            text = _strip(cell.text)
            span = _get_grid_span(cell)

            if row_idx == 0:
                if grid_span_index == 0:
                    set_field("学院信息", text)
                elif grid_span_index == 1:
                    set_field("专业信息", text)
                elif grid_span_index == 2:
                    set_field("时间", text)
            elif row_idx == 1:
                if grid_span_index == 1:
                    set_field("姓名", text)
                elif grid_span_index == 3:
                    set_field("学号", text)
            elif row_idx == 2:
                if grid_span_index == 1:
                    set_field("班级", text)
                elif grid_span_index == 3:
                    set_field("指导老师", text)
            elif row_idx == 3:
                if grid_span_index == 1:
                    set_field("课程名称", text)
                elif grid_span_index == 3:
                    set_field("周次", text)
            elif row_idx == 4:
                if grid_span_index == 1:
                    set_field("实验名称", text)

            # “实验分析与体会”之后的区块（以 analys_row_idx 为基准）
            if analys_row_idx is not None:
                if row_idx == analys_row_idx + 1:
                    set_field("实验分析与体会", text)
                elif row_idx == analys_row_idx + 2:
                    set_field("实验日期", text)
                elif row_idx == analys_row_idx + 3 and grid_span_index == 1:
                    set_field("备注", text)
                elif row_idx == analys_row_idx + 4 and grid_span_index == 1:
                    set_field("成绩", text)
                elif row_idx == analys_row_idx + 6 and grid_span_index == 1:
                    set_field("日期", text)

            col_idx += span
            grid_span_index += 1

        # 识别“实验内容”起止范围
        if not content_started and ("实验内容" in row_text):
            content_started = True
            continue

        if content_started and analys_row_idx is None:
            # 如果本行包含“实验分析与体会”，标记结束，并记录分析起始行索引
            if "实验分析与体会" in row_text:
                analys_row_idx = row_idx
            else:
                # 在内容区，累积文本（整行作为一个段）
                if row_text:
                    content_buffer.append(row_text)

    # 将实验内容解析为题目 items
    if content_buffer:
        full_content = "\n\n".join(content_buffer)
        report.content_items = _parse_content_items(full_content)

    # 若至少有一个关键字段或内容则认为解析成功
    has_any = any([
        report.学院信息, report.专业信息, report.时间, report.姓名, report.学号,
        report.课程名称, report.实验名称, report.实验分析与体会, report.实验日期,
        len(report.content_items or []) > 0,
    ])
    return report if has_any else None


def parse_docx_to_report(docx_path: Path) -> ReportDocument:
    doc = Document(str(docx_path))
    # 优先按统一模板严格解析
    tmpl_report = _parse_by_template(doc)
    if tmpl_report:
        return tmpl_report

    # 模板不匹配时使用通用标签映射解析
    report = ReportDocument(content_items=[])
    content_accumulator: List[str] = []

    for table in doc.tables:
        for row in table.rows:
            pair = _extract_row_label_value(row)
            if not pair:
                continue
            label = pair["label"]
            value = pair["value"]
            if label == "实验内容":
                if value:
                    content_accumulator.append(value)
                continue
            # 仅当是我们识别的18个单元之一时赋值
            if label in LABEL_MAP.values():
                try:
                    setattr(report, label, value)
                except Exception:
                    pass

    # 解析实验内容为题干与答案
    full_content = "\n\n".join(content_accumulator).strip()
    report.content_items = _parse_content_items(full_content)

    # 回退策略：若未识别到题目，则尝试从全局表格按两列结构提取
    if not report.content_items:
        items: List[ReportItem] = []
        counter = 1
        for table in doc.tables:
            rows = table.rows
            if not rows:
                continue
            headers = [_strip(c.text) for c in rows[0].cells]

            def find_idx(*keywords: str) -> Optional[int]:
                for kw in keywords:
                    for idx, h in enumerate(headers):
                        if kw in h:
                            return idx
                return None

            q_idx = find_idx("题干", "题目", "问题")
            a_idx = find_idx("答案", "回答")

            if q_idx is not None and a_idx is not None and len(rows) > 1:
                data_rows = rows[1:]
                for r in data_rows:
                    q = _strip(r.cells[q_idx].text)
                    a = _strip(r.cells[a_idx].text)
                    if q or a:
                        items.append(ReportItem(id=f"Q{counter}", question=q, answer=a))
                        counter += 1
                continue

            # 回退策略：表格前两列分别视为题干与答案
            col_count = len(rows[0].cells)
            if col_count >= 2:
                for r in rows:
                    q_raw = r.cells[0].text
                    q_norm_label = _normalize_label(q_raw)
                    # 若第一列是已知的18个单元标签，则跳过，不将其纳入内容题目
                    if q_norm_label in LABEL_MAP.values():
                        continue
                    q = _strip(q_raw)
                    a = _strip(r.cells[1].text)
                    if q or a:
                        # 进一步过滤：优先包含包含“题目/问题/任务”等关键词的行
                        if re.search(r"(题目|问题|任务)", q):
                            items.append(ReportItem(id=f"Q{counter}", question=q, answer=a))
                            counter += 1

        report.content_items = items

    return report