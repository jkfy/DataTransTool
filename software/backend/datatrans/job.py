"""任务编排：读取 -> 校验 -> 写 TC / EPLAN XML -> 日志/结果。"""

from __future__ import annotations

import time
import uuid
from pathlib import Path
from typing import Any, Callable

from . import config
from .excel_com import read_with_excel, write_table_with_excel
from .mapper import map_tc
from .models import FileSummary, Issue, JobOptions, JobResult, MaterialRecord
from .reader import InputError, read_input_file
from .report import JobLogger, write_result_json
from .validator import parse_record
from .writer import OutputError, make_output_paths, write_table
from .xml_writer import write_parts_xml

ProgressCallback = Callable[[dict[str, Any]], None]


def _emit(callback: ProgressCallback | None, payload: dict[str, Any]) -> None:
    if callback is None:
        return
    try:
        callback(payload)
    except Exception:
        pass


def _as_bool(value: Any, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "on"}:
        return True
    if text in {"0", "false", "no", "n", "off"}:
        return False
    return default


def _issue(
    level: str,
    code: str,
    message: str,
    *,
    file: str = "",
    sheet: str = "",
    row: int | None = None,
    field_name: str = "",
    value: Any = None,
) -> Issue:
    return Issue(
        level=level,
        code=code,
        message=message,
        file=file,
        sheet=sheet,
        row=row,
        field=field_name,
        value=value,
    )


def run_job(payload: dict[str, Any], progress: ProgressCallback | None = None) -> JobResult:
    started = time.perf_counter()
    job_id = str(payload.get("job_id") or uuid.uuid4().hex[:8])
    input_values = payload.get("inputs") or []
    if isinstance(input_values, str):
        input_values = [input_values]
    inputs = [str(item) for item in input_values if str(item).strip()]

    options_raw = payload.get("options") or {}
    output_dir = str(payload.get("output_dir") or options_raw.get("output_dir") or "").strip()
    options = JobOptions(
        output_dir=output_dir,
        export_tc=_as_bool(options_raw.get("export_tc"), True),
        export_xml=_as_bool(options_raw.get("export_xml"), False),
        on_invalid=str(options_raw.get("on_invalid") or "abort").strip().lower(),
        engine=str(options_raw.get("engine") or "auto").strip().lower(),
        sheet_name=(str(options_raw.get("sheet_name")).strip() or None)
        if options_raw.get("sheet_name")
        else None,
        templates_dir=(str(payload.get("templates_dir")).strip() or None)
        if payload.get("templates_dir")
        else None,
    )
    if options.on_invalid not in {"abort", "skip"}:
        options.on_invalid = "abort"
    if options.engine not in {"auto", "openpyxl", "excel"}:
        options.engine = "auto"

    issues: list[Issue] = []
    files: list[FileSummary] = []
    records: list[MaterialRecord] = []
    written_rows = {"tc": 0, "xml": 0}
    used_excel = False
    logger: JobLogger | None = None
    result_path = Path(output_dir) / "result.json" if output_dir else Path()

    def finish(
        status: str,
        message: str,
        outputs: dict[str, str | None] | None = None,
        *,
        ok: bool | None = None,
    ) -> JobResult:
        duration_ms = int((time.perf_counter() - started) * 1000)
        error_count = sum(1 for issue in issues if issue.level == "error")
        warning_count = sum(1 for issue in issues if issue.level == "warning")
        stats = {
            "files": len(files),
            "rows_read": sum(item.rows_read for item in files),
            "rows_valid": sum(item.rows_valid for item in files),
            "rows_invalid": sum(item.rows_invalid for item in files),
            "tc_rows": written_rows["tc"],
            "xml_rows": written_rows["xml"],
            "errors": error_count,
            "warnings": warning_count,
            "duration_ms": duration_ms,
        }
        result = JobResult(
            job_id=job_id,
            status=status,
            ok=(status != "failed") if ok is None else ok,
            outputs=outputs or {"tc": None, "xml": None},
            stats=stats,
            issues=issues,
            log_path=str(logger.log_path) if logger else "",
            result_path=str(result_path) if output_dir else "",
            message=message,
            files=files,
        )
        if logger:
            logger.info("任务结束", status=status, result=message, duration_ms=duration_ms)
            logger.close()
        if output_dir:
            try:
                write_result_json(result_path, result)
            except Exception as exc:
                result.message = f"{result.message}（result.json 写出失败：{exc}）"
                result.status = "partial" if result.status != "failed" else "failed"
                result.ok = result.status != "failed"
        _emit(progress, result.to_dict())
        return result

    if not output_dir:
        issues.append(_issue("error", "OUTPUT_DIR_MISSING", "未指定输出目录"))
        return finish("failed", "未指定输出目录")
    if not inputs:
        issues.append(_issue("error", "INPUT_MISSING", "未选择任何输入文件"))
        return finish("failed", "未选择任何输入文件")
    if not options.export_tc and not options.export_xml:
        issues.append(_issue("error", "NO_OUTPUT_SELECTED", "至少选择生成一种输出文件"))
        return finish("failed", "至少选择生成一种输出文件")

    try:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        logger = JobLogger(output_dir, job_id)
    except Exception as exc:
        issues.append(
            _issue(
                "error",
                "OUTPUT_DIR_ERROR",
                f"无法创建/写入输出目录：{output_dir}（{exc}）",
            )
        )
        return finish("failed", "无法创建输出目录")

    logger.info(
        "任务开始",
        job_id=job_id,
        inputs=len(inputs),
        output_dir=output_dir,
        export_tc=options.export_tc,
        export_xml=options.export_xml,
        on_invalid=options.on_invalid,
        engine=options.engine,
    )

    _emit(
        progress,
        {
            "type": "progress",
            "stage": "start",
            "current": 0,
            "total": len(inputs),
            "message": f"开始处理 {len(inputs)} 个文件",
        },
    )

    template_paths = {
        "tc": config.template_path("tc", options.templates_dir) if options.export_tc else None,
        "xml": config.template_path("xml", options.templates_dir) if options.export_xml else None,
    }
    for kind, path in template_paths.items():
        if path is not None and not Path(path).exists():
            issues.append(
                _issue(
                    "error",
                    "TEMPLATE_MISSING",
                    f"输出模板不存在：{path}",
                    field_name=kind,
                )
            )
            logger.error("输出模板不存在", kind=kind, path=path)

    has_template_error = any(issue.code == "TEMPLATE_MISSING" for issue in issues)

    for index, input_text in enumerate(inputs, start=1):
        path = Path(input_text)
        summary = FileSummary(file=str(path))
        _emit(
            progress,
            {
                "type": "progress",
                "stage": "read",
                "current": index - 1,
                "total": len(inputs),
                "file": str(path),
                "message": f"读取 {path.name}",
            },
        )

        if path.suffix.lower() != config.INPUT_EXTENSION:
            message = f"只支持 {config.INPUT_EXTENSION} 格式"
            issues.append(
                _issue("error", "UNSUPPORTED_EXTENSION", message, file=str(path))
            )
            summary.error = message
            files.append(summary)
            logger.issue(issues[-1])
            continue
        if not path.exists():
            message = "文件不存在"
            issues.append(_issue("error", "FILE_NOT_FOUND", message, file=str(path)))
            summary.error = message
            files.append(summary)
            logger.issue(issues[-1])
            continue

        try:
            if options.engine == "excel":
                read_result = read_with_excel(path, options.sheet_name)
            elif options.engine == "openpyxl":
                read_result = read_input_file(path, options.sheet_name)
            else:
                try:
                    read_result = read_input_file(path, options.sheet_name)
                except Exception as openpyxl_exc:
                    logger.warning(
                        "openpyxl 读取失败，尝试用本机 Excel 读取（兼容加密文件）",
                        file=str(path),
                        error=str(openpyxl_exc),
                    )
                    try:
                        read_result = read_with_excel(path, options.sheet_name)
                    except Exception as excel_exc:
                        raise InputError(
                            f"openpyxl 读取失败：{openpyxl_exc}；"
                            f"Excel COM 也失败：{excel_exc}"
                        ) from excel_exc
                    logger.info(
                        "已切换到 Excel COM 读取",
                        file=str(path),
                        engine=read_result.engine,
                    )
            if read_result.engine == "excel":
                used_excel = True
        except InputError as exc:
            issues.append(_issue("error", "INPUT_ERROR", str(exc), file=str(path)))
            summary.error = str(exc)
            files.append(summary)
            logger.issue(issues[-1])
            continue
        except Exception as exc:
            issues.append(
                _issue("error", "INPUT_ERROR", f"读取失败：{exc}", file=str(path))
            )
            summary.error = str(exc)
            files.append(summary)
            logger.issue(issues[-1])
            continue

        summary.sheet = read_result.sheet
        summary.engine = read_result.engine
        summary.rows_read = len(read_result.rows)
        for issue in read_result.issues:
            issues.append(issue)
            logger.issue(issue)

        logger.info(
            "读取完成",
            file=str(path),
            sheet=read_result.sheet,
            engine=read_result.engine,
            rows=len(read_result.rows),
            header_row=read_result.header.row if read_result.header else None,
        )

        for raw in read_result.rows:
            row_issues: list[Issue] = []
            record = parse_record(raw, row_issues)
            for issue in row_issues:
                issues.append(issue)
                logger.issue(issue)
            if record.valid:
                records.append(record)
                summary.rows_valid += 1
            else:
                summary.rows_invalid += 1

        files.append(summary)
        _emit(
            progress,
            {
                "type": "progress",
                "stage": "read",
                "current": index,
                "total": len(inputs),
                "file": str(path),
                "message": f"已读取 {path.name}：{summary.rows_read} 行",
            },
        )

    seen: dict[tuple[str, str], MaterialRecord] = {}
    for record in records:
        key = (record.code, record.part_no)
        if key in seen:
            issues.append(
                _issue(
                    "warning",
                    "DUPLICATE_ROW",
                    f"类型编号/部件编号重复：{record.code} / {record.part_no}",
                    file=record.source_file,
                    sheet=record.source_sheet,
                    row=record.source_row,
                )
            )
            logger.issue(issues[-1])
        else:
            seen[key] = record

    error_count = sum(1 for issue in issues if issue.level == "error")

    if has_template_error:
        return finish("failed", "输出模板缺失，未生成文件")

    if error_count and not options.allow_partial:
        logger.error("校验未通过，按默认策略不生成输出文件", errors=error_count)
        return finish(
            "failed",
            f"校验发现 {error_count} 个错误，未生成输出文件；请查看日志或勾选“跳过错误行继续导出”。",
        )

    valid_records = [record for record in records if record.valid]
    tc_records = list(valid_records) if options.export_tc else []
    xml_records = list(valid_records) if options.export_xml else []

    if not tc_records and not xml_records:
        logger.warning("没有可导出的数据行")
        issues.append(_issue("warning", "NO_EXPORT_ROWS", "没有可导出的数据行"))

    output_paths = make_output_paths(output_dir, job_id=job_id)
    outputs: dict[str, str | None] = {"tc": None, "xml": None}
    write_failed = False
    write_total = int(options.export_tc) + int(options.export_xml)
    write_index = 0

    if options.export_tc:
        _emit(
            progress,
            {
                "type": "progress",
                "stage": "write",
                "current": write_index,
                "total": write_total,
                "message": f"写出 TC 表（{len(tc_records)} 行）",
            },
        )
        try:
            rows = [map_tc(record) for record in tc_records]
            use_excel_writer = options.engine == "excel" or (
                options.engine == "auto" and used_excel
            )
            if use_excel_writer:
                unknown = write_table_with_excel(
                    template_path=template_paths["tc"],
                    output_path=output_paths["tc"],
                    rows=rows,
                )
            else:
                try:
                    unknown = write_table(
                        template_path=template_paths["tc"],
                        output_path=output_paths["tc"],
                        rows=rows,
                    )
                except OutputError as exc:
                    if options.engine != "auto":
                        raise
                    logger.warning(
                        "openpyxl 写出失败，尝试用本机 Excel 写出",
                        path=str(output_paths["tc"]),
                        error=str(exc),
                    )
                    unknown = write_table_with_excel(
                        template_path=template_paths["tc"],
                        output_path=output_paths["tc"],
                        rows=rows,
                    )
            outputs["tc"] = str(output_paths["tc"])
            written_rows["tc"] = len(tc_records)
            logger.info("TC 表写出完成", path=str(output_paths["tc"]), rows=len(tc_records))
            if unknown:
                issues.append(
                    _issue(
                        "warning",
                        "UNKNOWN_OUTPUT_HEADER",
                        "TC 模板中存在未识别的表头，已保持空白：" + "、".join(unknown),
                    )
                )
                logger.issue(issues[-1])
        except OutputError as exc:
            write_failed = True
            issues.append(_issue("error", "OUTPUT_ERROR", str(exc), field_name="tc"))
            logger.issue(issues[-1])
        write_index += 1

    if options.export_xml:
        _emit(
            progress,
            {
                "type": "progress",
                "stage": "write",
                "current": write_index,
                "total": write_total,
                "message": f"写出 EPLAN XML（{len(xml_records)} 行）",
            },
        )
        try:
            write_parts_xml(
                template_paths["xml"],
                output_paths["xml"],
                xml_records,
            )
            outputs["xml"] = str(output_paths["xml"])
            written_rows["xml"] = len(xml_records)
            logger.info("XML 写出完成", path=str(output_paths["xml"]), rows=len(xml_records))
        except OutputError as exc:
            write_failed = True
            issues.append(_issue("error", "OUTPUT_ERROR", str(exc), field_name="xml"))
            logger.issue(issues[-1])
        write_index += 1

    if write_failed:
        successful = [path for path in outputs.values() if path]
        status = "partial" if successful else "failed"
        message = "部分输出失败，请查看日志" if successful else "输出文件写出失败"
        return finish(status, message, outputs)

    status = "success" if error_count == 0 and (tc_records or xml_records) else "partial"
    if status == "success":
        message = f"完成：{len(files)} 个文件，TC {len(tc_records)} 行，XML {len(xml_records)} 行"
    elif error_count:
        message = (
            f"完成（跳过 {sum(item.rows_invalid for item in files)} 行错误数据）："
            f"TC {len(tc_records)} 行，XML {len(xml_records)} 行"
        )
    else:
        message = "完成，但没有可导出的数据行（已生成空表）"
    _emit(
        progress,
        {
            "type": "progress",
            "stage": "write",
            "current": write_total,
            "total": write_total,
            "message": "输出完成",
        },
    )
    return finish(status, message, outputs)
