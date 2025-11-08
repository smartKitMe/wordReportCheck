import sys
from pathlib import Path
from docx import Document


def find_value_cell(row, label: str):
    cells = list(row.cells)
    for idx, cell in enumerate(cells):
        if label in (cell.text or ""):
            if len(cells) >= 2:
                if idx + 1 < len(cells):
                    return idx + 1
                else:
                    return 0
            else:
                return 0
    return None


def main(doc_path: str):
    p = Path(doc_path)
    doc = Document(str(p))
    for t_i, table in enumerate(doc.tables):
        for r_i, row in enumerate(table.rows):
            texts = [c.text for c in row.cells]
            joined = " | ".join(texts)
            if ("成绩" in joined) or ("日期" in joined):
                vi_grade = find_value_cell(row, "成绩")
                vi_date = find_value_cell(row, "日期")
                print(f"[table {t_i} row {r_i}] {joined}")
                if vi_grade is not None:
                    print(f"  成绩值单元格 -> '{row.cells[vi_grade].text}'")
                if vi_date is not None:
                    print(f"  日期值单元格 -> '{row.cells[vi_date].text}'")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python scripts/dump_labels.py <docx>")
        sys.exit(2)
    main(sys.argv[1])