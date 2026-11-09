"""命令行入口：供 Qt QProcess 调用，也可独立用于批处理/测试。"""

from __future__ import annotations

import argparse
import json
import sys
import uuid
from pathlib import Path
from typing import Any

from . import config
from .job import run_job


def _print_json(payload: dict[str, Any]) -> None:
    print(json.dumps(payload, ensure_ascii=False), flush=True)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="datatrans-core",
        description="DataTransTool Python 后端：物料总表 -> TC / Eplan",
    )
    parser.add_argument("--job", help="job.json 路径；提供后其余输入参数可省略")
    parser.add_argument("--input", action="append", default=[], help="输入 .xlsx，可重复")
    parser.add_argument("--out-dir", help="输出目录")
    parser.add_argument("--templates", help="模板目录（默认自动查找 templates/）")
    parser.add_argument("--sheet", help="指定工作表；默认活动工作表")
    parser.add_argument("--tc", dest="export_tc", action="store_true", default=None)
    parser.add_argument("--no-tc", dest="export_tc", action="store_false")
    parser.add_argument("--eplan", dest="export_eplan", action="store_true", default=None)
    parser.add_argument("--no-eplan", dest="export_eplan", action="store_false")
    parser.add_argument("--skip-invalid", action="store_true", help="跳过错误行继续导出")
    parser.add_argument("--abort-on-invalid", action="store_true", help="有错误时整批不导出")
    parser.add_argument("--progress", action="store_true", help="stdout 输出 JSONL 进度")
    parser.add_argument("--version", action="version", version=f"%(prog)s {config.VERSION}")
    return parser


def main(argv: list[str] | None = None) -> int:
    # Windows 控制台/管道默认可能不是 UTF-8；显式设置，便于 Qt 读取中文。
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = _build_parser()
    args = parser.parse_args(argv)

    try:
        payload: dict[str, Any]
        if args.job:
            job_path = Path(args.job)
            payload = json.loads(job_path.read_text(encoding="utf-8-sig"))
            if not isinstance(payload, dict):
                raise ValueError("job.json 顶层必须是对象")
        else:
            payload = {}

        if args.input:
            payload["inputs"] = args.input
        if args.out_dir:
            payload["output_dir"] = args.out_dir
        if args.templates:
            payload["templates_dir"] = args.templates
        if args.sheet:
            payload.setdefault("options", {})["sheet_name"] = args.sheet
        if args.export_tc is not None:
            payload.setdefault("options", {})["export_tc"] = args.export_tc
        if args.export_eplan is not None:
            payload.setdefault("options", {})["export_eplan"] = args.export_eplan
        if args.skip_invalid:
            payload.setdefault("options", {})["on_invalid"] = "skip"
        if args.abort_on_invalid:
            payload.setdefault("options", {})["on_invalid"] = "abort"
        payload.setdefault("job_id", uuid.uuid4().hex[:8])

        def progress(event: dict[str, Any]) -> None:
            if args.progress:
                _print_json(event)

        result = run_job(payload, progress)
        if not args.progress:
            print(result.message)
            if result.log_path:
                print(f"日志：{result.log_path}")
            for kind, path in result.outputs.items():
                if path:
                    print(f"{kind.upper()}：{path}")
        return 0 if result.ok else 2
    except KeyboardInterrupt:
        if args.progress:
            _print_json({"type": "error", "message": "已取消"})
        else:
            print("已取消", file=sys.stderr)
        return 130
    except Exception as exc:
        if args.progress:
            _print_json({"type": "error", "message": f"后端异常：{exc}"})
        else:
            print(f"后端异常：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
