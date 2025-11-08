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


def _load_env_file() -> None:
    """Load .env from current working directory or project root.
    Values set here will NOT override existing environment variables.
    Supported syntax: KEY=VALUE or export KEY=VALUE; lines starting with '#' are comments.
    """
    try:
        candidates = [
            Path.cwd() / ".env",
            Path(__file__).resolve().parents[2] / ".env",
        ]
        for env_path in candidates:
            if env_path.exists():
                for raw in env_path.read_text(encoding="utf-8").splitlines():
                    line = raw.strip()
                    if not line or line.startswith("#"):
                        continue
                    if line.startswith("export "):
                        line = line[len("export "):].strip()
                    if "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    # do not override existing values
                    if key and (key not in os.environ):
                        os.environ[key] = value
                # only load first existing .env
                break
    except Exception:
        # fail silently; .env is optional
        pass


def _apply_env_defaults(args: argparse.Namespace) -> None:
    """Apply environment-provided defaults to CLI args if not explicitly set.
    Env keys:
      - WORDREPORTCHECK_PROVIDER
      - WORDREPORTCHECK_MODEL
      - WORDREPORTCHECK_API_KEY
      - WORDREPORTCHECK_PER_ITEM (true/false/1/0)
      - WORDREPORTCHECK_WRITE_BACK (true/false/1/0)
      - WORDREPORTCHECK_DOC
      - WORDREPORTCHECK_JSON
    """
    env = os.environ

    def truthy(s: str) -> bool:
        return s.strip().lower() in {"1", "true", "yes", "y", "on"}

    cmd = getattr(args, "command", None)

    # Common paths defaults
    if not getattr(args, "doc", None) and env.get("WORDREPORTCHECK_DOC"):
        setattr(args, "doc", env.get("WORDREPORTCHECK_DOC"))
    if not getattr(args, "json", None) and env.get("WORDREPORTCHECK_JSON"):
        setattr(args, "json", env.get("WORDREPORTCHECK_JSON"))

    # Only apply score-related defaults when scoring
    if cmd == "score":
        # Provider & model (env can override default values if CLI not set)
        if env.get("WORDREPORTCHECK_PROVIDER") and getattr(args, "provider", "deepseek") == "deepseek":
            args.provider = env.get("WORDREPORTCHECK_PROVIDER")
        if env.get("WORDREPORTCHECK_MODEL") and getattr(args, "model", "deepseek-chat") == "deepseek-chat":
            args.model = env.get("WORDREPORTCHECK_MODEL")

        # API key (generic fallback)
        if not getattr(args, "api_key", None) and env.get("WORDREPORTCHECK_API_KEY"):
            args.api_key = env.get("WORDREPORTCHECK_API_KEY")

        # Flags
        if not getattr(args, "per_item", False) and env.get("WORDREPORTCHECK_PER_ITEM"):
            args.per_item = truthy(env.get("WORDREPORTCHECK_PER_ITEM", ""))
        if not getattr(args, "write_back", False) and env.get("WORDREPORTCHECK_WRITE_BACK"):
            args.write_back = truthy(env.get("WORDREPORTCHECK_WRITE_BACK", ""))


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

    score_parser = subparsers.add_parser("score", help="解析并提交到评分服务进行评分（仅评分实验内容中的题目）")
    score_parser.add_argument("--doc", required=False, help="docx 文档路径（与 --json 二选一）")
    score_parser.add_argument("--json", required=False, help="已解析的 JSON 文件路径（与 --doc 二选一）")
    score_parser.add_argument("--api-key", required=False, help="DeepSeek API Key（默认读取环境变量 DEEPSEEK_API_KEY）")
    score_parser.add_argument("--model", required=False, default="deepseek-chat", help="评分模型，deepseek默认为 deepseek-chat，kimi默认为 moonshot-v1-8k")
    score_parser.add_argument("--per-item", action="store_true", help="逐题提交评分（每题单独请求）")
    score_parser.add_argument("--write-back", action="store_true", help="将平均分写入到 JSON 的“成绩”字段（仅在 --json 时生效）")
    score_parser.add_argument("--provider", choices=["deepseek", "kimi"], default="deepseek", help="选择评分服务提供者：deepseek 或 kimi")
    score_parser.add_argument("--write-docx", action="store_true", help="将平均分写回到 docx 文档的“成绩/日期”单元格（仅在 --doc 时生效）")

    # 新增：从 JSON 写回成绩到 DOCX（必须在 parse_args 之前注册）
    write_docx_parser = subparsers.add_parser("write-docx", help="将 JSON 中的‘成绩’写回到 DOCX 的‘成绩/日期’单元格")
    write_docx_parser.add_argument("--doc", required=False, help="docx 文档路径（必填）")
    write_docx_parser.add_argument("--json", required=False, help="JSON 文件路径（必填，用于读取成绩）")

    # 先加载 .env，使其中的变量对后续读取生效
    _load_env_file()
    args = parser.parse_args()
    # 使用环境变量为未显式提供的参数设置默认值
    _apply_env_defaults(args)

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

        # Provider与API Key选择
        provider = args.provider
        api_key = args.api_key
        if not api_key:
            if provider == "deepseek":
                api_key = os.getenv("DEEPSEEK_API_KEY")
            else:
                api_key = os.getenv("MOONSHOT_API_KEY")
        # 若仍无 API Key，尝试通用 KEY
        if not api_key:
            api_key = os.getenv("WORDREPORTCHECK_API_KEY")
        if not api_key:
            if provider == "deepseek":
                print("❌ 未找到 DeepSeek API Key。请通过 --api-key 或设置环境变量 DEEPSEEK_API_KEY。")
            else:
                print("❌ 未找到 Kimi/Moonshot API Key。请通过 --api-key 或设置环境变量 MOONSHOT_API_KEY。")
            return 2

        # 模型默认值按Provider校准
        model = args.model
        if provider == "kimi" and model == "deepseek-chat":
            model = "moonshot-v1-8k"

        if args.per_item:
            if provider == "deepseek":
                from .scoring.deepseek_client import score_item
            else:
                from .scoring.kimi_client import score_item
            all_results = []
            for it in items:
                r = score_item(it, api_key=api_key, model=model)
                all_results.append(r)
            # 写回成绩（平均分）
            if args.write_back and args.json:
                # 统计有效分数
                scores = [x.get("score") for x in all_results if isinstance(x, dict) and isinstance(x.get("score"), (int, float))]
                avg = sum(scores) / len(scores) if scores else None
                if avg is not None:
                    p = Path(args.json)
                    rd = report_from_json(p.read_text(encoding="utf-8"))
                    rd.成绩 = f"{avg:.1f}"
                    p.write_text(report_to_json(rd), encoding="utf-8")
                    print(f"✅ 已写入成绩: {rd.成绩} 至: {p}")
                else:
                    print("⚠️ 未得到有效分数，未写入成绩。")
            # 写回 docx（平均分）
            if args.write_docx and args.doc:
                scores = [x.get("score") for x in all_results if isinstance(x, dict) and isinstance(x.get("score"), (int, float))]
                avg = sum(scores) / len(scores) if scores else None
                if avg is not None:
                    try:
                        from .parsers.docx_writer import write_grade_and_date
                        ok = write_grade_and_date(Path(args.doc), f"{avg:.1f}")
                        if ok:
                            print(f"✅ 已写回成绩到 DOCX: {args.doc}")
                        else:
                            print("⚠️ 未找到可写入的‘成绩/日期’单元格，未写入 DOCX。")
                    except Exception as e:
                        print(f"⚠️ 写回 DOCX 失败: {e}")
            # 打印详细评分结果
            print(json.dumps(all_results, ensure_ascii=False, indent=2))
            # 额外：将详细评分结果写出到与原 JSON 同目录的 .scores.json
            if args.json:
                try:
                    src = Path(args.json)
                    out_scores = src.with_name(src.stem + ".scores.json")
                    out_scores.write_text(json.dumps(all_results, ensure_ascii=False, indent=2), encoding="utf-8")
                    print(f"✅ 已写出评分明细 JSON 至: {out_scores}")
                except Exception as e:
                    print(f"⚠️ 写出评分明细失败: {e}")
        else:
            if provider == "deepseek":
                from .scoring.deepseek_client import score_items
            else:
                from .scoring.kimi_client import score_items
            results = score_items(items, api_key=api_key, model=model)
            # 写回成绩（平均分）
            if args.write_back and args.json:
                # results 期望是数组；若为其他格式则不写回
                if isinstance(results, list):
                    scores = [x.get("score") for x in results if isinstance(x, dict) and isinstance(x.get("score"), (int, float))]
                    avg = sum(scores) / len(scores) if scores else None
                    if avg is not None:
                        p = Path(args.json)
                        rd = report_from_json(p.read_text(encoding="utf-8"))
                        rd.成绩 = f"{avg:.1f}"
                        p.write_text(report_to_json(rd), encoding="utf-8")
                        print(f"✅ 已写入成绩: {rd.成绩} 至: {p}")
                    else:
                        print("⚠️ 未得到有效分数，未写入成绩。")
                else:
                    print("⚠️ 返回结果不是评分数组，未写入成绩。")
            # 写回 docx（平均分）
            if args.write_docx and args.doc and isinstance(results, list):
                scores = [x.get("score") for x in results if isinstance(x, dict) and isinstance(x.get("score"), (int, float))]
                avg = sum(scores) / len(scores) if scores else None
                if avg is not None:
                    try:
                        from .parsers.docx_writer import write_grade_and_date
                        ok = write_grade_and_date(Path(args.doc), f"{avg:.1f}")
                        if ok:
                            print(f"✅ 已写回成绩到 DOCX: {args.doc}")
                        else:
                            print("⚠️ 未找到可写入的‘成绩/日期’单元格，未写入 DOCX。")
                    except Exception as e:
                        print(f"⚠️ 写回 DOCX 失败: {e}")
            # 打印详细评分结果
            print(json.dumps(results, ensure_ascii=False, indent=2))
            # 额外：将详细评分结果写出到与原 JSON 同目录的 .scores.json
            if args.json and isinstance(results, list):
                try:
                    src = Path(args.json)
                    out_scores = src.with_name(src.stem + ".scores.json")
                    out_scores.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
                    print(f"✅ 已写出评分明细 JSON 至: {out_scores}")
                except Exception as e:
                    print(f"⚠️ 写出评分明细失败: {e}")
        return 0

    if args.command == "write-docx":
        if not args.doc or not args.json:
            print("❌ 请同时提供 --doc 与 --json。")
            return 2
        p = Path(args.json)
        if not p.exists():
            print(f"❌ 未找到 JSON 文件: {p}")
            return 2
        from .parsers.docx_writer import write_grade_and_date
        rd = report_from_json(p.read_text(encoding="utf-8"))
        grade = (rd.成绩 or "").strip()
        if not grade:
            print("⚠️ JSON 顶层‘成绩’为空，无法写回 DOCX。请先评分并写入成绩。")
            return 2
        ok = write_grade_and_date(Path(args.doc), grade)
        if ok:
            print(f"✅ 已写回成绩到 DOCX: {args.doc}")
            return 0
        else:
            print("⚠️ 未找到可写入的‘成绩/日期’单元格，未写入 DOCX。")
            return 2

    parser.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())