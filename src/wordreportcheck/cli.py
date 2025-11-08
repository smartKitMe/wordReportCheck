import argparse
import json
import os
import sys
from pathlib import Path

from . import __version__
from .parsers.docx_parser import parse_docx_to_report
from .schemas import (
    report_to_json,
    report_from_json,
    ReportDocument,
)
from .scoring.deepseek_client import score_items


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="wordreportcheck",
        description="WordReportCheck: 批改学生实验报告的命令行工具"
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"wordreportcheck {__version__}",
        help="显示版本信息"
    )

    subparsers = parser.add_subparsers(dest="command")

    parse_parser = subparsers.add_parser("parse", help="解析 docx 文档为包含18个单元的 JSON")
    parse_parser.add_argument("--doc", required=True, help="docx 文档路径")
    parse_parser.add_argument("--out", required=False, help="输出 JSON 文件路径（默认打印到控制台）")

    score_parser = subparsers.add_parser("score", help="解析并提交到 DeepSeek 进行评分（仅评分实验内容中的题目）")
    score_parser.add_argument("--doc", required=False, help="docx 文档路径（与 --json 二选一）")
    score_parser.add_argument("--json", required=False, help="已解析的 JSON 文件路径（与 --doc 二选一）")
    score_parser.add_argument("--api-key", required=False, help="DeepSeek API Key（默认读取环境变量 DEEPSEEK_API_KEY）")
    score_parser.add_argument("--model", required=False, default="deepseek-chat", help="评分模型，默认 deepseek-chat")
    score_parser.add_argument("--per-item", action="store_true", help="逐题提交评分（每题单独请求）")
    score_parser.add_argument("--write-back", action="store_true", help="将平均分写入到 JSON 的“成绩”字段（仅在 --json 时生效）")

    args = parser.parse_args()

    if args.command == "parse":
        doc_path = Path(args.doc)
        report: ReportDocument = parse_docx_to_report(doc_path)
        json_str = report_to_json(report)
        if args.out:
            out_path = Path(args.out)
            out_path.write_text(json_str, encoding="utf-8")
            print(f"✅ 已写出 JSON 至: {out_path}")
        else:
            print(json_str)
        return 0

    if args.command == "score":
        report: ReportDocument = None  # type: ignore
        if args.doc:
            report = parse_docx_to_report(Path(args.doc))
        elif args.json:
            report = report_from_json(Path(args.json).read_text(encoding="utf-8"))
        else:
            print("❌ 请提供 --doc 或 --json 其中之一。")
            return 2

        items = report.content_items or []
        if not items:
            print("❌ 未在文档的“实验内容”中识别到题目，请检查模板是否包含“题目X：”格式。")
            return 2

        api_key = args.api_key or os.getenv("DEEPSEEK_API_KEY")
        if not api_key:
            print("❌ 未找到 DeepSeek API Key。请通过 --api-key 或设置环境变量 DEEPSEEK_API_KEY。")
            return 2

        if args.per_item:
            from .scoring.deepseek_client import score_item
            all_results = []
            for it in items:
                r = score_item(it, api_key=api_key, model=args.model)
                all_results.append(r)
            # 写回成绩（平均分）
            if args.write_back and args.json:
                # 统计有效分数
                scores = [x.get("score") for x in all_results if isinstance(x, dict) and isinstance(x.get("score"), (int, float))]
                avg = sum(scores) / len(scores) if scores else None
                if avg is not None:
                    from .schemas import report_from_json, report_to_json
                    p = Path(args.json)
                    rd = report_from_json(p.read_text(encoding="utf-8"))
                    rd.成绩 = f"{avg:.1f}"
                    p.write_text(report_to_json(rd), encoding="utf-8")
                    print(f"✅ 已写入成绩: {rd.成绩} 至: {p}")
                else:
                    print("⚠️ 未得到有效分数，未写入成绩。")
            # 打印详细评分结果
            print(json.dumps(all_results, ensure_ascii=False, indent=2))
        else:
            results = score_items(items, api_key=api_key, model=args.model)
            # 写回成绩（平均分）
            if args.write_back and args.json:
                # results 期望是数组；若为其他格式则不写回
                if isinstance(results, list):
                    scores = [x.get("score") for x in results if isinstance(x, dict) and isinstance(x.get("score"), (int, float))]
                    avg = sum(scores) / len(scores) if scores else None
                    if avg is not None:
                        from .schemas import report_from_json, report_to_json
                        p = Path(args.json)
                        rd = report_from_json(p.read_text(encoding="utf-8"))
                        rd.成绩 = f"{avg:.1f}"
                        p.write_text(report_to_json(rd), encoding="utf-8")
                        print(f"✅ 已写入成绩: {rd.成绩} 至: {p}")
                    else:
                        print("⚠️ 未得到有效分数，未写入成绩。")
                else:
                    print("⚠️ 返回结果不是评分数组，未写入成绩。")
            # 打印详细评分结果
            print(json.dumps(results, ensure_ascii=False, indent=2))
        return 0

    parser.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())