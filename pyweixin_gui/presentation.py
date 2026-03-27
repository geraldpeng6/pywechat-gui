from __future__ import annotations

from collections import Counter

from .models import ExecutionRecord, ExecutionRowResult, TaskTemplate, TaskType


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
