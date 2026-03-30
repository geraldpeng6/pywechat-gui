from __future__ import annotations

import json
from collections import Counter
from dataclasses import asdict

from .models import ChatBatchExportRequest, ChatExportRequest, ExecutionRecord, ExecutionRowResult, ExportHistoryRecord, ResourceExportKind, ResourceExportRequest, TaskTemplate, TaskType


def template_type_label(task_type: TaskType) -> str:
    return "批量消息" if task_type is TaskType.MESSAGE else "批量文件"


def filter_templates(templates: list[TaskTemplate], query: str) -> list[TaskTemplate]:
    normalized = query.strip().lower()
    if not normalized:
        return list(templates)
    filtered: list[TaskTemplate] = []
    for template in templates:
        haystack = " ".join([template.name, template_type_label(template.task_type), template.updated_at or ""]).lower()
        if normalized in haystack:
            filtered.append(template)
    return filtered


def template_metrics(templates: list[TaskTemplate]) -> dict[str, int]:
    return {
        "total": len(templates),
        "message": sum(1 for template in templates if template.task_type is TaskType.MESSAGE),
        "file": sum(1 for template in templates if template.task_type is TaskType.FILE),
    }


def filter_executions(executions: list[ExecutionRecord], query: str, failed_only: bool) -> list[ExecutionRecord]:
    normalized = query.strip().lower()
    filtered: list[ExecutionRecord] = []
    for execution in executions:
        if failed_only and execution.failure_count == 0:
            continue
        haystack = " ".join(
            [
                str(execution.id or ""),
                template_type_label(execution.task_type),
                execution.started_at,
                execution.status,
                str(execution.success_count),
                str(execution.failure_count),
            ]
        ).lower()
        if normalized and normalized not in haystack:
            continue
        filtered.append(execution)
    return filtered


def execution_metrics(executions: list[ExecutionRecord]) -> dict[str, int]:
    return {
        "total": len(executions),
        "success": sum(1 for execution in executions if execution.failure_count == 0 and execution.status == "completed"),
        "failed": sum(1 for execution in executions if execution.failure_count > 0),
    }


def summarize_failures(rows: list[ExecutionRowResult]) -> str:
    counter: Counter[str] = Counter()
    for row in rows:
        if row.success:
            continue
        reason = row.error_code or row.error_message or "未知原因"
        counter[reason] += 1
    if not counter:
        return "这次执行没有失败项，全部处理成功。"
    ordered = sorted(counter.items(), key=lambda item: (-item[1], item[0]))
    summary = "失败原因摘要：" + "；".join(f"{reason} x{count}" for reason, count in ordered[:4])
    if len(ordered) > 4:
        summary += "；..."
    return summary


def parse_export_detail(record: ExportHistoryRecord) -> dict:
    try:
        detail = json.loads(record.detail_json or "{}")
    except json.JSONDecodeError:
        detail = {}
    return detail if isinstance(detail, dict) else {}


def format_export_history_detail(record: ExportHistoryRecord) -> str:
    detail = parse_export_detail(record)
    result = detail.get("result", detail)
    lines = [
        f"导出类型：{record.export_kind}",
        f"标题：{record.title}",
        f"导出目录：{record.export_folder}",
        f"导出数量：{record.exported_count}",
        f"摘要文件：{record.summary_path or '无'}",
    ]
    request = detail.get("request")
    if isinstance(request, dict):
        lines.append("")
        lines.append("本次执行参数：")
        for label, value in _request_lines(record.export_kind, request):
            lines.append(f"- {label}：{value}")
    lines.append("")
    lines.append("结果摘要：")
    for label, value in _result_lines(record.export_kind, result):
        lines.append(f"- {label}：{value}")
    if "request" not in detail:
        lines.append("")
        lines.append("提示：这条记录来自较早版本，缺少原始执行参数，可能无法直接重新执行。")
    return "\n".join(lines)


def export_history_can_rerun(record: ExportHistoryRecord) -> bool:
    detail = parse_export_detail(record)
    return isinstance(detail.get("request"), dict)


def export_history_failed_sessions(record: ExportHistoryRecord) -> list[str]:
    detail = parse_export_detail(record)
    result = detail.get("result", detail)
    failed = result.get("failed_sessions", []) if isinstance(result, dict) else []
    names: list[str] = []
    for item in failed:
        if not isinstance(item, dict):
            continue
        name = str(item.get("session_name", "")).strip()
        if name:
            names.append(name)
    return names


def export_history_can_retry_failed(record: ExportHistoryRecord) -> bool:
    return record.export_kind == "chat_batch" and bool(export_history_failed_sessions(record))


def rebuild_export_request(record: ExportHistoryRecord, failed_only: bool = False) -> tuple[str, object] | None:
    detail = parse_export_detail(record)
    request = detail.get("request")
    if not isinstance(request, dict):
        return None
    if record.export_kind == "chat":
        return "chat", ChatExportRequest(**request)
    if record.export_kind == "chat_batch":
        if failed_only:
            failed_sessions = export_history_failed_sessions(record)
            if not failed_sessions:
                return None
            request = {**request, "session_names": failed_sessions}
        return "chat_batch", ChatBatchExportRequest(**request)
    if record.export_kind in {kind.value for kind in ResourceExportKind}:
        return "resource", ResourceExportRequest(
            export_kind=ResourceExportKind(request["export_kind"]),
            target_folder=str(request.get("target_folder", "")),
            year=str(request.get("year", "")),
            month=str(request.get("month", "")),
        )
    return None


def serialize_export_detail(export_kind: str, request: object, result: dict) -> str:
    return json.dumps(
        {
            "type": export_kind,
            "request": asdict(request) if hasattr(request, "__dataclass_fields__") else request,
            "result": result,
        },
        ensure_ascii=False,
    )


def _request_lines(export_kind: str, request: dict) -> list[tuple[str, str]]:
    if export_kind == "chat":
        return [
            ("会话名称", str(request.get("session_name", ""))),
            ("导出目录", str(request.get("target_folder", ""))),
            ("导出文本消息", "是" if request.get("export_messages") else "否"),
            ("导出聊天文件", "是" if request.get("export_files") else "否"),
            ("尝试导出图片/视频", "是" if request.get("export_images") else "否"),
            ("消息条数", str(request.get("message_limit", ""))),
            ("文件数量", str(request.get("file_limit", ""))),
        ]
    if export_kind == "chat_batch":
        session_names = request.get("session_names", [])
        count = len(session_names) if isinstance(session_names, list) else 0
        return [
            ("批量会话数", str(count)),
            ("导出目录", str(request.get("target_folder", ""))),
            ("导出文本消息", "是" if request.get("export_messages") else "否"),
            ("导出聊天文件", "是" if request.get("export_files") else "否"),
            ("尝试导出图片/视频", "是" if request.get("export_images") else "否"),
            ("消息条数", str(request.get("message_limit", ""))),
            ("文件数量", str(request.get("file_limit", ""))),
        ]
    return [
        ("导出类型", _resource_kind_label(str(request.get("export_kind", "")))),
        ("导出目录", str(request.get("target_folder", ""))),
        ("年份", str(request.get("year", ""))),
        ("月份", str(request.get("month", "") or "全部")),
    ]


def _result_lines(export_kind: str, result: dict) -> list[tuple[str, str]]:
    if export_kind == "chat":
        warnings = result.get("warnings", [])
        warning_text = "；".join(str(item) for item in warnings[:3]) if isinstance(warnings, list) and warnings else "无"
        return [
            ("消息数量", str(result.get("message_count", 0))),
            ("文件数量", str(result.get("file_count", 0))),
            ("注意事项", warning_text),
        ]
    if export_kind == "chat_batch":
        failed_names = export_history_failed_sessions(
            ExportHistoryRecord(
                export_kind=export_kind,
                title="",
                export_folder="",
                exported_count=0,
                detail_json=json.dumps({"result": result}, ensure_ascii=False),
            )
        )
        return [
            ("成功会话", str(result.get("success_count", 0))),
            ("失败会话", str(result.get("failure_count", 0))),
            ("失败名单预览", "、".join(failed_names[:5]) if failed_names else "无"),
        ]
    exported_paths = result.get("exported_paths", [])
    preview = "；".join(str(item) for item in exported_paths[:3]) if isinstance(exported_paths, list) and exported_paths else "无"
    return [
        ("导出类型", _resource_kind_label(str(result.get("export_kind", export_kind)))),
        ("导出数量", str(result.get("exported_count", 0))),
        ("导出结果预览", preview),
    ]


def _resource_kind_label(kind: str) -> str:
    mapping = {
        "recent_files": "最近聊天文件",
        "wxfiles": "微信聊天文件",
        "videos": "微信聊天视频",
    }
    return mapping.get(kind, kind or "未知")
