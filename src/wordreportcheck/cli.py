import argparse
import csv
import json
import os
import sys
import shutil
from pathlib import Path

from . import __version__
from .parsers.docx_parser import parse_docx_to_report, _parse_content_items
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
    # 可选：将实验内容原文交给 AI 分割为指定题目数量
    parse_parser.add_argument("--segment-count", type=int, required=False, help="AI 分割题目数量（可选；必须与模型输出一致）")
    parse_parser.add_argument("--provider", choices=["deepseek", "kimi"], default="deepseek", help="选择分割服务提供者：deepseek 或 kimi")
    parse_parser.add_argument("--api-key", required=False, help="分割服务 API Key（可选，未提供则读取环境变量）")
    parse_parser.add_argument("--model", required=False, default="deepseek-chat", help="分割模型，deepseek默认为 deepseek-chat，kimi默认为 moonshot-v1-128k")

    score_parser = subparsers.add_parser("score", help="解析并提交到评分服务进行评分（仅评分实验内容中的题目）")
    score_parser.add_argument("--doc", required=False, help="docx 文档路径（与 --json 二选一）")
    score_parser.add_argument("--json", required=False, help="已解析的 JSON 文件路径（与 --doc 二选一）")
    score_parser.add_argument("--api-key", required=False, help="DeepSeek API Key（默认读取环境变量 DEEPSEEK_API_KEY）")
    score_parser.add_argument("--model", required=False, default="deepseek-chat", help="评分模型，deepseek默认为 deepseek-chat，kimi默认为 moonshot-v1-128k")
    score_parser.add_argument("--per-item", action="store_true", help="逐题提交评分（每题单独请求）")
    score_parser.add_argument("--write-back", action="store_true", help="将平均分写入到 JSON 的“成绩”字段（仅在 --json 时生效）")
    score_parser.add_argument("--provider", choices=["deepseek", "kimi"], default="deepseek", help="选择评分服务提供者：deepseek 或 kimi")
    score_parser.add_argument("--write-docx", action="store_true", help="将平均分写回到 docx 文档的“成绩/日期”单元格（仅在 --doc 时生效）")
    # 可选：若未识别题目，可先进行 AI 分割
    score_parser.add_argument("--segment-count", type=int, required=False, help="AI 分割题目数量（可选；必须与模型输出一致）")

    # 新增：从 JSON 写回成绩到 DOCX（必须在 parse_args 之前注册）
    write_docx_parser = subparsers.add_parser("write-docx", help="将 JSON 中的‘成绩’写回到 DOCX 的‘成绩/日期’单元格")
    write_docx_parser.add_argument("--doc", required=False, help="docx 文档路径（必填）")
    write_docx_parser.add_argument("--json", required=False, help="JSON 文件路径（必填，用于读取成绩）")

    # 新增：一键执行命令（解析 -> 评分 -> 写回）
    auto_parser = subparsers.add_parser(
        "auto",
        help="一键解析 DOCX、评分并写回成绩，同时输出 JSON 和评分明细"
    )
    auto_parser.add_argument("--doc", required=True, help="docx 文档路径（必填）")
    auto_parser.add_argument("--api-key", required=False, help="评分服务 API Key（可选，未提供则读取环境变量）")
    auto_parser.add_argument("--out-dir", required=False, help="输出目录（用于存放解析 JSON 与评分明细 JSON）")
    # 新增：当未识别题目时，允许在 auto 流程中进行 AI 分割
    auto_parser.add_argument("--segment-count", type=int, required=False, help="AI 分割题目数量（可选，默认 6）")
    auto_parser.add_argument("--provider", choices=["deepseek", "kimi"], required=False, help="选择服务提供者（覆盖环境变量）")

    # 批量执行：遍历目录下的所有 .docx 并逐个运行 auto
    auto_dir_parser = subparsers.add_parser(
        "auto-dir",
        help="批量执行 auto：遍历目录下的所有 .docx 并输出到指定目录"
    )
    auto_dir_parser.add_argument("--in-dir", required=True, help="输入目录（遍历其中的 .docx 文件）")
    auto_dir_parser.add_argument("--out-dir", required=True, help="输出目录（解析 JSON、评分明细及写回 DOCX 将写入此目录）")
    auto_dir_parser.add_argument("--api-key", required=False, help="评分服务 API Key（可选，未提供则读取环境变量）")
    auto_dir_parser.add_argument("--recursive", action="store_true", help="是否递归遍历子目录")
    # 新增：允许在批量模式下进行 AI 分割，指定题目数量
    auto_dir_parser.add_argument("--segment-count", type=int, required=False, help="AI 分割题目数量（可选）")
    auto_dir_parser.add_argument("--provider", choices=["deepseek", "kimi"], required=False, help="选择服务提供者（覆盖环境变量）")

    # 先加载 .env，使其中的变量对后续读取生效
    _load_env_file()
    args = parser.parse_args()
    # 使用环境变量为未显式提供的参数设置默认值
    _apply_env_defaults(args)

    if args.command == "parse":
        doc_path = Path(args.doc)
        report: ReportDocument = parse_docx_to_report(doc_path)
        # 若指定分割数量，则使用本地规则按题目进行分割
        if getattr(args, "segment_count", None):
            if not (report.实验内容原文 and report.实验内容原文.strip()):
                print("❌ 未在文档中提取到‘实验内容’原文，无法进行分割。")
                return 2
            items = _parse_content_items(report.实验内容原文 or "", int(args.segment_count))
            if len(items) != int(args.segment_count):
                print(f"⚠️ 分割结果数量与期望不一致：期望 {args.segment_count}，实际 {len(items)}。已按规则补齐/截断。")
            report.content_items = items
            print(f"✅ 已按文档规则分割为 {len(items)} 个题目。")
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
        # 若尚未有题目且指定了分割数量，先进行本地规则分割
        if (not items) and getattr(args, "segment_count", None):
            if not (report.实验内容原文 and report.实验内容原文.strip()):
                print("❌ 未在文档中提取到‘实验内容’原文，无法进行分割。")
                return 2
            items = _parse_content_items(report.实验内容原文 or "", int(args.segment_count))
            if len(items) != int(args.segment_count):
                print(f"⚠️ 分割结果数量与期望不一致：期望 {args.segment_count}，实际 {len(items)}。已按规则补齐/截断。")
            report.content_items = items
            print(f"✅ 已按文档规则分割为 {len(items)} 个题目。")
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
            model = "moonshot-v1-128k"

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

    if args.command == "auto":
        # 1) 解析 DOCX -> JSON 并写出到目标输出目录（默认与 DOCX 同目录）
        doc_path = Path(args.doc)
        if not doc_path.exists():
            print(f"❌ 未找到 DOCX 文件: {doc_path}")
            return 2
        out_dir = Path(getattr(args, "out_dir", "")) if getattr(args, "out_dir", None) else doc_path.parent
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"❌ 创建输出目录失败: {out_dir}，错误: {e}")
            return 2
        report: ReportDocument = parse_docx_to_report(doc_path)
        json_str = report_to_json(report)
        json_path = out_dir / (doc_path.stem + ".json")
        try:
            json_path.write_text(json_str, encoding="utf-8")
            print(f"✅ 已写出解析后的 JSON 至: {json_path}")
        except Exception as e:
            print(f"⚠️ 写出解析 JSON 失败: {e}")

        # 2) 若未识别到题目，则尝试 AI 分割（支持 --segment-count 或环境变量，默认 6）
        items = report.content_items or []
        if not items:
            env = os.environ
            seg_count = getattr(args, "segment_count", None)
            if not seg_count:
                # 从环境读取，默认 6
                try:
                    seg_count = int(env.get("WORDREPORTCHECK_SEGMENT_COUNT", "6"))
                except Exception:
                    seg_count = 6
            if not (report.实验内容原文 and (report.实验内容原文 or "").strip()):
                print("❌ 未在文档中提取到‘实验内容’原文，无法进行分割。")
                return 2
            items = _parse_content_items(report.实验内容原文 or "", int(seg_count))
            if len(items) != int(seg_count):
                print(f"⚠️ 分割结果数量与期望不一致：期望 {seg_count}，实际 {len(items)}。已按规则补齐/截断。")
            report.content_items = items
            json_str = report_to_json(report)
            try:
                json_path.write_text(json_str, encoding="utf-8")
                print(f"✅ 已通过本地规则分割得到 {len(items)} 个题目，并写出解析后的 JSON 至: {json_path}")
            except Exception as e:
                print(f"⚠️ 写出解析 JSON 失败: {e}")

        # 读取 Provider/Model/API Key（无需命令行参数，尽量简化）
        env = os.environ
        provider = (env.get("WORDREPORTCHECK_PROVIDER") or "deepseek").strip().lower()
        api_key_arg = getattr(args, "api_key", None)

        # 先根据当前 provider 读取对应的 key
        if provider == "deepseek":
            api_key = api_key_arg or env.get("DEEPSEEK_API_KEY") or env.get("WORDREPORTCHECK_API_KEY")
        else:
            api_key = api_key_arg or env.get("MOONSHOT_API_KEY") or env.get("WORDREPORTCHECK_API_KEY")

        # 若未找到 key，则尝试根据存在的密钥自动切换 provider
        if not api_key:
            if env.get("MOONSHOT_API_KEY"):
                provider = "kimi"
                api_key = env.get("MOONSHOT_API_KEY") or env.get("WORDREPORTCHECK_API_KEY")
            elif env.get("DEEPSEEK_API_KEY"):
                provider = "deepseek"
                api_key = env.get("DEEPSEEK_API_KEY") or env.get("WORDREPORTCHECK_API_KEY")

        # 模型选择在最终 provider 决定后进行
        model_env = env.get("WORDREPORTCHECK_MODEL")
        if model_env:
            model = model_env
        else:
            model = "deepseek-chat" if provider == "deepseek" else "moonshot-v1-128k"
        # 归一化：如果选择 kimi 但模型是 deepseek 默认值，则切换为 moonshot 默认
        if provider == "kimi" and model == "deepseek-chat":
            model = "moonshot-v1-128k"

        if not api_key:
            if provider == "deepseek":
                print("❌ 未找到 DeepSeek API Key。请设置环境变量 DEEPSEEK_API_KEY 或 WORDREPORTCHECK_API_KEY。")
            else:
                print("❌ 未找到 Kimi/Moonshot API Key。请设置环境变量 MOONSHOT_API_KEY 或 WORDREPORTCHECK_API_KEY。")
            return 2

        # 执行评分（批量方式）
        if provider == "deepseek":
            from .scoring.deepseek_client import score_items as _score_items
        else:
            from .scoring.kimi_client import score_items as _score_items
        results = _score_items(items, api_key=api_key, model=model)

        # 3) 写出评分明细至与解析 JSON 同目录的 .scores.json
        scores_path = out_dir / (doc_path.stem + ".scores.json")
        try:
            scores_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"✅ 已写出评分明细 JSON 至: {scores_path}")
        except Exception as e:
            print(f"⚠️ 写出评分明细失败: {e}")

        # 4) 计算平均分，写回解析 JSON 的‘成绩’字段
        avg = None
        if isinstance(results, list):
            valid_scores = [x.get("score") for x in results if isinstance(x, dict) and isinstance(x.get("score"), (int, float))]
            avg = (sum(valid_scores) / len(valid_scores)) if valid_scores else None
        if avg is not None:
            try:
                rd = report_from_json(json_path.read_text(encoding="utf-8"))
                rd.成绩 = f"{avg:.1f}"
                json_path.write_text(report_to_json(rd), encoding="utf-8")
                print(f"✅ 已在解析 JSON 写入成绩: {rd.成绩} -> {json_path}")
            except Exception as e:
                print(f"⚠️ 写入成绩到解析 JSON 失败: {e}")
        else:
            print("⚠️ 未得到有效分数，未在解析 JSON 写入‘成绩’。")

        # 5) 将平均分写回到 DOCX 的‘成绩/日期’单元格（不修改源文件，复制到输出目录后写回）
        if avg is not None:
            try:
                from .parsers.docx_writer import write_grade_and_date
                # 复制源 DOCX 到输出目录
                out_docx = out_dir / doc_path.name
                try:
                    shutil.copy2(doc_path, out_docx)
                except Exception as e:
                    print(f"⚠️ 复制 DOCX 到输出目录失败: {e}")
                    return 2
                # 在复制的 DOCX 上写回成绩
                ok = write_grade_and_date(out_docx, f"{avg:.1f}")
                if ok:
                    print(f"✅ 已写回成绩到输出 DOCX: {out_docx}")
                else:
                    print("⚠️ 未找到可写入的‘成绩/日期’单元格，未写入 DOCX。")
            except Exception as e:
                print(f"⚠️ 写回 DOCX 失败: {e}")

        # 6) 打印简要总结
        summary = {
            "doc": str(doc_path),
            "doc_out": str(out_dir / doc_path.name),
            "json": str(json_path),
            "scores_json": str(scores_path),
            "provider": provider,
            "model": model,
            "average_score": None if avg is None else float(f"{avg:.1f}")
        }
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 0

    if args.command == "auto-dir":
        in_dir = Path(args.in_dir)
        if not in_dir.exists() or not in_dir.is_dir():
            print(f"❌ 输入目录不存在或不可用: {in_dir}")
            return 2
        out_dir = Path(args.out_dir)
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"❌ 创建输出目录失败: {out_dir}，错误: {e}")
            return 2

        # 收集待处理的 DOCX 文件
        iterator = in_dir.rglob("*.docx") if getattr(args, "recursive", False) else in_dir.glob("*.docx")
        docs = [p for p in iterator if p.is_file() and not p.name.startswith("~$")]
        if not docs:
            print("⚠️ 目录中未找到 .docx 文件（已忽略临时锁文件 ~$ 开头）。")
            return 0

        total = 0
        failed = 0
        csv_rows = []
        for doc_path in docs:
            try:
                # 1) 解析 DOCX -> JSON
                report: ReportDocument = parse_docx_to_report(doc_path)
                json_path = out_dir / (doc_path.stem + ".json")

                # 解析后若未有题目且指定分割数量，尝试 AI 分割
                seg_count = getattr(args, "segment_count", None)
                if (not report.content_items) and seg_count:
                    if not (report.实验内容原文 and report.实验内容原文.strip()):
                        print(f"❌ [{doc_path.name}] 未在文档中提取到‘实验内容’原文，无法进行 AI 分割。")
                        failed += 1
                        continue
                    items = _parse_content_items(report.实验内容原文 or "", int(seg_count))
                    if len(items) != int(seg_count):
                        print(f"⚠️ [{doc_path.name}] 分割结果数量与期望不一致：期望 {seg_count}，实际 {len(items)}。已按规则补齐/截断。")
                    report.content_items = items
                    print(f"✅ [{doc_path.name}] 已通过本地规则分割得到 {len(items)} 个题目。")

                # 写出解析（或分割后）的 JSON
                json_str = report_to_json(report)
                try:
                    json_path.write_text(json_str, encoding="utf-8")
                    print(f"✅ [{doc_path.name}] 已写出解析后的 JSON 至: {json_path}")
                except Exception as e:
                    print(f"⚠️ [{doc_path.name}] 写出解析 JSON 失败: {e}")

                # 2) 评分
                items = report.content_items or []
                if not items:
                    print(f"❌ [{doc_path.name}] 未在文档的‘实验内容’中识别到题目，且未进行 AI 分割。")
                    csv_rows.append({
                        "文件名": doc_path.name,
                        "姓名": (report.姓名 or ""),
                        "学号": (report.学号 or ""),
                        "班级": (report.班级 or ""),
                        "课程名称": (report.课程名称 or ""),
                        "实验名称": (report.实验名称 or ""),
                        "成绩": ""
                    })
                    failed += 1
                    continue

                env = os.environ
                provider = (getattr(args, "provider", None) or env.get("WORDREPORTCHECK_PROVIDER") or "deepseek").strip().lower()
                api_key_arg = getattr(args, "api_key", None)
                if provider == "deepseek":
                    api_key = api_key_arg or env.get("DEEPSEEK_API_KEY") or env.get("WORDREPORTCHECK_API_KEY")
                else:
                    api_key = api_key_arg or env.get("MOONSHOT_API_KEY") or env.get("WORDREPORTCHECK_API_KEY")
                if not api_key:
                    if env.get("MOONSHOT_API_KEY"):
                        provider = "kimi"
                        api_key = env.get("MOONSHOT_API_KEY") or env.get("WORDREPORTCHECK_API_KEY")
                    elif env.get("DEEPSEEK_API_KEY"):
                        provider = "deepseek"
                        api_key = env.get("DEEPSEEK_API_KEY") or env.get("WORDREPORTCHECK_API_KEY")

                model_env = env.get("WORDREPORTCHECK_MODEL")
                if model_env:
                    model = model_env
                else:
                    model = "deepseek-chat" if provider == "deepseek" else "moonshot-v1-128k"
                if provider == "kimi" and model == "deepseek-chat":
                    model = "moonshot-v1-128k"

                if not api_key:
                    if provider == "deepseek":
                        print(f"❌ [{doc_path.name}] 未找到 DeepSeek API Key。请设置环境变量 DEEPSEEK_API_KEY 或 WORDREPORTCHECK_API_KEY。")
                    else:
                        print(f"❌ [{doc_path.name}] 未找到 Kimi/Moonshot API Key。请设置环境变量 MOONSHOT_API_KEY 或 WORDREPORTCHECK_API_KEY。")
                    csv_rows.append({
                        "文件名": doc_path.name,
                        "姓名": (report.姓名 or ""),
                        "学号": (report.学号 or ""),
                        "班级": (report.班级 or ""),
                        "课程名称": (report.课程名称 or ""),
                        "实验名称": (report.实验名称 or ""),
                        "成绩": ""
                    })
                    failed += 1
                    continue

                if provider == "deepseek":
                    from .scoring.deepseek_client import score_items as _score_items
                else:
                    from .scoring.kimi_client import score_items as _score_items
                results = _score_items(items, api_key=api_key, model=model)

                # 3) 写出评分明细
                scores_path = out_dir / (doc_path.stem + ".scores.json")
                try:
                    scores_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
                    print(f"✅ [{doc_path.name}] 已写出评分明细 JSON 至: {scores_path}")
                except Exception as e:
                    print(f"⚠️ [{doc_path.name}] 写出评分明细失败: {e}")

                # 4) 计算平均分并写回 JSON
                avg = None
                if isinstance(results, list):
                    valid_scores = [x.get("score") for x in results if isinstance(x, dict) and isinstance(x.get("score"), (int, float))]
                    avg = (sum(valid_scores) / len(valid_scores)) if valid_scores else None
                if avg is not None:
                    try:
                        rd = report_from_json(json_path.read_text(encoding="utf-8"))
                        rd.成绩 = f"{avg:.1f}"
                        json_path.write_text(report_to_json(rd), encoding="utf-8")
                        print(f"✅ [{doc_path.name}] 已在解析 JSON 写入成绩: {rd.成绩} -> {json_path}")
                    except Exception as e:
                        print(f"⚠️ [{doc_path.name}] 写入成绩到解析 JSON 失败: {e}")
                else:
                    print(f"⚠️ [{doc_path.name}] 未得到有效分数，未在解析 JSON 写入‘成绩’。")

                # 5) 将平均分写回到复制后的 DOCX（不修改源文件）
                if avg is not None:
                    try:
                        from .parsers.docx_writer import write_grade_and_date
                        out_docx = out_dir / doc_path.name
                        try:
                            shutil.copy2(doc_path, out_docx)
                        except Exception as e:
                            print(f"⚠️ [{doc_path.name}] 复制 DOCX 到输出目录失败: {e}")
                            failed += 1
                            continue
                        ok = write_grade_and_date(out_docx, f"{avg:.1f}")
                        if ok:
                            print(f"✅ [{doc_path.name}] 已写回成绩到输出 DOCX: {out_docx}")
                        else:
                            print(f"⚠️ [{doc_path.name}] 未找到可写入的‘成绩/日期’单元格，未写入 DOCX。")
                    except Exception as e:
                        print(f"⚠️ [{doc_path.name}] 写回 DOCX 失败: {e}")

                # 6) 打印简要总结（逐份）
                summary = {
                    "doc": str(doc_path),
                    "doc_out": str(out_dir / doc_path.name),
                    "json": str(json_path),
                    "scores_json": str(scores_path),
                    "provider": provider,
                    "model": model,
                    "average_score": None if avg is None else float(f"{avg:.1f}")
                }
                print(json.dumps(summary, ensure_ascii=False, indent=2))
                # 加入 CSV 汇总行
                csv_rows.append({
                    "姓名": (report.姓名 or ""),
                    "学号": (report.学号 or ""),
                    "班级": (report.班级 or ""),
                    "课程名称": (report.课程名称 or ""),
                    "实验名称": (report.实验名称 or ""),
                    "成绩": ("" if avg is None else f"{avg:.1f}")
                })
                total += 1
            except Exception as e:
                failed += 1
                print(f"❌ [{doc_path.name}] 处理失败: {e}")

        # 写出成绩汇总 CSV
        csv_path = out_dir / "grades.csv"
        try:
            headers = ["姓名", "学号", "班级", "课程名称", "实验名称", "成绩"]
            with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=headers)
                writer.writeheader()
                for row in csv_rows:
                    writer.writerow({k: row.get(k, "") for k in headers})
            print(f"✅ 已写出成绩汇总 CSV 至: {csv_path}")
        except Exception as e:
            print(f"⚠️ 写出成绩汇总 CSV 失败: {e}")

        final_summary = {
            "in_dir": str(in_dir),
            "out_dir": str(out_dir),
            "processed": total,
            "failed": failed,
            "csv": str(csv_path)
        }
        print(json.dumps(final_summary, ensure_ascii=False, indent=2))
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