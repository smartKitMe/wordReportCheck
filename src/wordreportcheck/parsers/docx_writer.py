from pathlib import Path
from typing import Optional
from datetime import datetime
from docx import Document


def _find_label_row_and_value_cell(table, label: str) -> Optional[tuple]:
    """在给定表格中查找包含指定标签的行，并返回(行对象, 值单元格index)。
    规则：
    - 优先选择同一行中非标签的下一个单元格作为值单元格；
    - 若整行只有一个单元格，则返回该单元格index=0（覆盖写入）。
    """
    try:
        for row in table.rows:
            cells = list(row.cells)
            if not cells:
                continue
            for idx, cell in enumerate(cells):
                if label in (cell.text or ""):
                    # 寻找同一行的值单元格（优先下一个）
                    if len(cells) >= 2:
                        # 如果标签在第一个单元格，优先写到第二个；否则写到第一个非标签单元格
                        if idx + 1 < len(cells):
                            return (row, idx + 1)
                        else:
                            # 标签在最后一个，则写到第一个单元格
                            return (row, 0)
                    else:
                        return (row, 0)
        return None
    except Exception:
        return None


def write_grade_and_date(doc_path: Path, grade: str, date_str: Optional[str] = None) -> bool:
    """将成绩与日期写回 docx：
    - 查找包含“成绩”的行，将其同一行的值单元格设置为 grade；
    - 查找包含“日期”的行，将其同一行的值单元格设置为 date_str（默认为当前日期 YYYY.MM.DD）。
    返回 True 表示至少写入了一个字段。
    """
    try:
        doc = Document(str(doc_path))
    except Exception:
        return False

    wrote_any = False
    date_val = date_str or datetime.now().strftime("%Y.%m.%d")

    try:
        for table in doc.tables:
            # 写入成绩
            loc = _find_label_row_and_value_cell(table, "成绩")
            if loc is not None:
                row, val_idx = loc
                try:
                    row.cells[val_idx].text = str(grade)
                    wrote_any = True
                except Exception:
                    pass
            # 写入日期
            loc2 = _find_label_row_and_value_cell(table, "日期")
            if loc2 is not None:
                row2, val_idx2 = loc2
                try:
                    row2.cells[val_idx2].text = str(date_val)
                    wrote_any = True
                except Exception:
                    pass
        if wrote_any:
            doc.save(str(doc_path))
    except Exception:
        return False

    return wrote_any