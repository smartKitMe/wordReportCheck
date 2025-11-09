from pathlib import Path
from typing import Optional
from datetime import datetime
from docx import Document


def _find_label_row_and_value_cell(table, label: str, prefer_prev_if_last: bool = False) -> Optional[tuple]:
    """在给定表格中查找包含指定标签的行，并返回(行对象, 值单元格index)。
    规则：
    - 仅处理“可见上”为两列的行（考虑水平合并后的两列）；
    - 优先选择同一行中非标签的右侧相邻可见单元格作为值单元格；
    - 如标签位于该行最后一列，且启用 prefer_prev_if_last，则写到前一列。
    """
    try:
        for row in table.rows:
            # 构建可见单元格序列（按首次出现顺序去重 _tc）
            raw_cells = list(row.cells)
            unique_order_indices = []  # 可见单元格对应的 row.cells 索引
            seen_tcs = set()
            for i, c in enumerate(raw_cells):
                tc = getattr(c, "_tc", None)
                # _tc 作为底层单元格标识，合并后多个 Cell 可能共享同一 _tc
                key = tc if tc is not None else id(c)
                if key not in seen_tcs:
                    seen_tcs.add(key)
                    unique_order_indices.append(i)

            # 仅处理可见为两列的行
            if len(unique_order_indices) != 2:
                continue

            # 遍历可见单元格以匹配标签
            for vis_idx, raw_idx in enumerate(unique_order_indices):
                cell = raw_cells[raw_idx]
                text = (cell.text or "").strip()
                # 兼容中英文冒号、去除所有空白字符
                normalized = text.replace("：", ":")
                normalized = "".join(normalized.split())
                lbl_norm = label.replace("：", ":")
                lbl_norm = "".join(lbl_norm.split())

                if (label in text) or (lbl_norm in normalized):
                    # 选择值单元格（优先右侧相邻的可见单元格）
                    target_vis_idx = None
                    if vis_idx + 1 < len(unique_order_indices):
                        target_vis_idx = vis_idx + 1
                    elif prefer_prev_if_last and vis_idx - 1 >= 0:
                        target_vis_idx = vis_idx - 1
                    else:
                        target_vis_idx = 0

                    target_raw_idx = unique_order_indices[target_vis_idx]
                    return (row, target_raw_idx)
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
            # 写入成绩（当标签在最后一列时，优先写到前一列）
            loc = _find_label_row_and_value_cell(table, "成绩", prefer_prev_if_last=True)
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