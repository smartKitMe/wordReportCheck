import json
import sys
from pathlib import Path
from typing import List, Dict

from wordreportcheck.parsers.docx_parser import parse_docx_to_report
from wordreportcheck.schemas import report_to_json


SAMPLES_DIR = Path(__file__).resolve().parent.parent / "samples"
OUTPUTS_DIR = Path(__file__).resolve().parent.parent / "outputs"

REQUIRED_TOP_KEYS: List[str] = [
    "学院信息",
    "专业信息",
    "时间",
    "姓名",
    "学号",
    "班级",
    "指导老师",
    "课程名称",
    "周次",
    "实验名称",
    "实验环境",
    "实验内容",
    "实验分析与体会",
    "实验日期",
    "备注",
    "成绩",
    "签名",
    "日期",
]


def validate_json_structure(obj: Dict) -> Dict[str, List[str]]:
    errors: List[str] = []
    warnings: List[str] = []

    # 顶层键存在性
    for k in REQUIRED_TOP_KEYS:
        if k not in obj:
            errors.append(f"缺少顶层键: {k}")

    # 实验内容结构
    content = obj.get("实验内容")
    if not isinstance(content, dict):
        errors.append("实验内容 不是对象")
    else:
        items = content.get("items")
        if not isinstance(items, list):
            errors.append("实验内容.items 不是列表")
        else:
            if len(items) == 0:
                warnings.append("实验内容.items 为空（可能模板未按题目X格式）")
            else:
                # 抽样检查首条结构
                sample = items[0]
                for subk in ("id", "question", "answer"):
                    if subk not in sample:
                        errors.append(f"items[0] 缺少字段: {subk}")

    return {"errors": errors, "warnings": warnings}


def _iter_docx_files(root: Path):
    for p in root.rglob("*.docx"):
        # 跳过临时文件（如 ~$ 开头）
        if p.name.startswith("~$"):
            continue
        yield p


def main() -> int:
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

    docx_files = sorted(_iter_docx_files(SAMPLES_DIR))
    if not docx_files:
        print("❌ 未在 samples 目录发现 .docx 文件")
        return 2

    total = 0
    passed = 0
    failed = 0
    empty_items = 0
    results = []

    for docx_path in docx_files:
        total += 1
        try:
            report = parse_docx_to_report(docx_path)
            json_str = report_to_json(report)
            # 在 outputs 下镜像 samples 的相对路径结构
            rel = docx_path.relative_to(SAMPLES_DIR)
            out_path = OUTPUTS_DIR / rel.parent / (rel.stem + ".json")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(json_str, encoding="utf-8")

            obj = json.loads(json_str)
            v = validate_json_structure(obj)
            errs = v["errors"]
            warns = v["warnings"]

            status = "passed"
            if errs:
                status = "failed"
                failed += 1
            else:
                passed += 1
            if any("items 为空" in w for w in warns):
                empty_items += 1

            results.append({
                "file": docx_path.name,
                "status": status,
                "errors": errs,
                "warnings": warns,
                "items_count": len(obj.get("实验内容", {}).get("items", []) if isinstance(obj.get("实验内容"), dict) else [])
            })
        except Exception as e:
            failed += 1
            results.append({
                "file": docx_path.name,
                "status": "error",
                "errors": [str(e)],
                "warnings": [],
                "items_count": 0,
            })

    # 汇总输出
    print("=== 验证汇总 ===")
    print(f"总文件数: {total}")
    print(f"结构通过: {passed}")
    print(f"结构失败: {failed}")
    print(f"content_items 为空: {empty_items}")
    print("")

    # 列出失败与警告样例
    for r in results:
        if r["status"] != "passed" or r["warnings"]:
            print(f"-- {r['file']} | {r['status']} | items: {r['items_count']}")
            if r["errors"]:
                for e in r["errors"]:
                    print(f"  error: {e}")
            if r["warnings"]:
                for w in r["warnings"]:
                    print(f"  warn: {w}")

    return 0


if __name__ == "__main__":
    sys.exit(main())