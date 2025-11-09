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
        # 写入成绩：按标签“成绩”匹配并写入其相邻单元格
        for row in table.rows:
            vi_g = find_value_cell(row, "成绩")
            if vi_g is not None:
                row.cells[vi_g].text = grade
                wrote = True

        # 写入日期：仅写入该表格的最后一行
        if len(table.rows) > 0:
            last_row = table.rows[-1]
            vi_d = find_value_cell(last_row, "日期")
            if vi_d is not None:
                last_row.cells[vi_d].text = date_str
                wrote = True
            else:
                # 回退：若最后一行没有“日期”标签，则写入第二个单元格（若存在），否则写入第一个单元格
                cells = list(last_row.cells)
                if len(cells) >= 2:
                    cells[1].text = date_str
                elif len(cells) >= 1:
                    cells[0].text = date_str
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