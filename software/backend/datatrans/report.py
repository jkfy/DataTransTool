"""文本日志与 result.json 写出。"""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from .models import Issue, JobResult


def _safe(value: Any, limit: int = 200) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t")
    if len(text) > limit:
        text = text[:limit] + "..."
    return text


def format_issue(issue: Issue) -> str:
    parts = [f"[{issue.level.upper()}]", f"code={issue.code}"]
    if issue.file:
        parts.append(f"file={_safe(issue.file)}")
    if issue.sheet:
        parts.append(f"sheet={_safe(issue.sheet)}")
    if issue.row is not None:
        parts.append(f"row={issue.row}")
    if issue.field:
        parts.append(f"field={_safe(issue.field)}")
    if issue.value is not None:
        parts.append(f"value={_safe(issue.value, 80)}")
    parts.append(issue.message)
    return " ".join(parts)


class JobLogger:
    """一个任务一个文本日志；同时提供 info/warning/error 便捷方法。"""

    def __init__(self, output_dir: str | Path, job_id: str = "") -> None:
        directory = Path(output_dir)
        directory.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        suffix = f"_{job_id[:6]}" if job_id else ""
        self.log_path = directory / f"import_{stamp}{suffix}.log"
        self._stream = self.log_path.open("w", encoding="utf-8-sig", newline="")

    def _write(self, level: str, message: str, **fields: Any) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        parts = [timestamp, f"[{level.upper()}]"]
        for key, value in fields.items():
            if value is None or value == "":
                continue
            parts.append(f"{key}={_safe(value)}")
        parts.append(message)
        self._stream.write(" ".join(parts) + "\n")
        self._stream.flush()

    def info(self, message: str, **fields: Any) -> None:
        self._write("info", message, **fields)

    def warning(self, message: str, **fields: Any) -> None:
        self._write("warning", message, **fields)

    def error(self, message: str, **fields: Any) -> None:
        self._write("error", message, **fields)

    def issue(self, issue: Issue) -> None:
        self._write(
            issue.level,
            issue.message,
            code=issue.code,
            file=issue.file,
            sheet=issue.sheet,
            row=issue.row,
            field=issue.field,
            value=issue.value,
        )

    def close(self) -> None:
        if not self._stream.closed:
            self._stream.flush()
            self._stream.close()


def write_result_json(path: str | Path, result: JobResult) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(target.name + ".tmp")
    payload = json.dumps(result.to_dict(), ensure_ascii=False, indent=2)
    try:
        tmp.write_text(payload, encoding="utf-8")
        os.replace(tmp, target)
    except Exception:
        try:
            if tmp.exists():
                tmp.unlink()
        finally:
            raise
