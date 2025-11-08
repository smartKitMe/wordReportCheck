from pathlib import Path
from datetime import datetime
from docx import Document
import argparse


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


def write(doc_path: Path, grade: str, date_str: str):
    doc = Document(str(doc_path))
    wrote = False
    for table in doc.tables:
        for row in table.rows:
            vi_g = find_value_cell(row, "成绩")
            if vi_g is not None:
                row.cells[vi_g].text = grade
                wrote = True
            vi_d = find_value_cell(row, "日期")
            if vi_d is not None:
                row.cells[vi_d].text = date_str
                wrote = True
    if wrote:
        doc.save(str(doc_path))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Force write grade/date into DOCX next to labels.")
    parser.add_argument("--doc", required=True, help="Path to the target DOCX file")
    parser.add_argument("--grade", default="88.0", help="Grade value to write (default: 88.0)")
    parser.add_argument("--date", default=datetime.now().strftime("%Y.%m.%d"), help="Date string to write (default: today)")
    args = parser.parse_args()

    p = Path(args.doc)
    write(p, args.grade, args.date)
    print("done")