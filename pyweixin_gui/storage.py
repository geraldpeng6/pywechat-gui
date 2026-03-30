from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path

from .models import ExecutionRecord, ExecutionRowResult, ExportHistoryRecord, TaskTemplate, TaskType


class AppStorage:
    def __init__(self, database_path: Path):
        self.database_path = database_path
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def connect(self):
        connection = sqlite3.connect(self.database_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        try:
            yield connection
            connection.commit()
        finally:
            connection.close()

    def _init_db(self) -> None:
        with self.connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    task_type TEXT NOT NULL,
                    rows_json TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS executions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_type TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    finished_at TEXT,
                    status TEXT NOT NULL,
                    row_count INTEGER NOT NULL,
                    success_count INTEGER NOT NULL,
                    failure_count INTEGER NOT NULL,
                    source_template_id INTEGER,
                    source_execution_id INTEGER
                );
                CREATE TABLE IF NOT EXISTS execution_rows (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    execution_id INTEGER NOT NULL,
                    row_index INTEGER NOT NULL,
                    session_name TEXT NOT NULL,
                    success INTEGER NOT NULL,
                    error_code TEXT,
                    error_message TEXT,
                    raw_error TEXT,
                    row_payload_json TEXT NOT NULL,
                    FOREIGN KEY(execution_id) REFERENCES executions(id) ON DELETE CASCADE
                );
                CREATE TABLE IF NOT EXISTS export_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    export_kind TEXT NOT NULL,
                    title TEXT NOT NULL,
                    export_folder TEXT NOT NULL,
                    exported_count INTEGER NOT NULL,
                    summary_path TEXT,
                    detail_json TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """
            )

    def list_templates(self, task_type: TaskType | None = None) -> list[TaskTemplate]:
        query = "SELECT * FROM templates"
        params: tuple[str, ...] = ()
        if task_type is not None:
            query += " WHERE task_type=?"
            params = (task_type.value,)
        query += " ORDER BY updated_at DESC, id DESC"
        with self.connect() as conn:
            rows = conn.execute(query, params).fetchall()
        return [self._row_to_template(row) for row in rows]

    def save_template(self, template: TaskTemplate) -> TaskTemplate:
        with self.connect() as conn:
            if template.id is None:
                cursor = conn.execute(
                    "INSERT INTO templates (name, task_type, rows_json) VALUES (?, ?, ?)",
                    (template.name, template.task_type.value, template.rows_json),
                )
                template_id = int(cursor.lastrowid)
            else:
                conn.execute(
                    """
                    UPDATE templates
                    SET name=?, task_type=?, rows_json=?, updated_at=CURRENT_TIMESTAMP
                    WHERE id=?
                    """,
                    (template.name, template.task_type.value, template.rows_json, template.id),
                )
                template_id = template.id
            row = conn.execute("SELECT * FROM templates WHERE id=?", (template_id,)).fetchone()
        return self._row_to_template(row)

    def delete_template(self, template_id: int) -> None:
        with self.connect() as conn:
            conn.execute("DELETE FROM templates WHERE id=?", (template_id,))

    def duplicate_template(self, template_id: int, new_name: str) -> TaskTemplate:
        template = self.get_template(template_id)
        template.id = None
        template.name = new_name
        return self.save_template(template)

    def get_template(self, template_id: int) -> TaskTemplate:
        with self.connect() as conn:
            row = conn.execute("SELECT * FROM templates WHERE id=?", (template_id,)).fetchone()
        if row is None:
            raise KeyError(f"Template {template_id} not found")
        return self._row_to_template(row)

    def save_execution(self, record: ExecutionRecord) -> ExecutionRecord:
        with self.connect() as conn:
            if record.id is None:
                cursor = conn.execute(
                    """
                    INSERT INTO executions
                    (task_type, started_at, finished_at, status, row_count, success_count, failure_count, source_template_id, source_execution_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        record.task_type.value,
                        record.started_at,
                        record.finished_at,
                        record.status,
                        record.row_count,
                        record.success_count,
                        record.failure_count,
                        record.source_template_id,
                        record.source_execution_id,
                    ),
                )
                record_id = int(cursor.lastrowid)
            else:
                conn.execute(
                    """
                    UPDATE executions
                    SET finished_at=?, status=?, row_count=?, success_count=?, failure_count=?, source_template_id=?, source_execution_id=?
                    WHERE id=?
                    """,
                    (
                        record.finished_at,
                        record.status,
                        record.row_count,
                        record.success_count,
                        record.failure_count,
                        record.source_template_id,
                        record.source_execution_id,
                        record.id,
                    ),
                )
                conn.execute("DELETE FROM execution_rows WHERE execution_id=?", (record.id,))
                record_id = record.id

            for row in record.rows:
                conn.execute(
                    """
                    INSERT INTO execution_rows
                    (execution_id, row_index, session_name, success, error_code, error_message, raw_error, row_payload_json)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        record_id,
                        row.row_index,
                        row.session_name,
                        1 if row.success else 0,
                        row.error_code,
                        row.error_message,
                        row.raw_error,
                        row.row_payload_json,
                    ),
                )
            row = conn.execute("SELECT * FROM executions WHERE id=?", (record_id,)).fetchone()
            record_rows = conn.execute(
                "SELECT * FROM execution_rows WHERE execution_id=? ORDER BY row_index ASC",
                (record_id,),
            ).fetchall()
        loaded = self._row_to_execution(row)
        loaded.rows = [self._row_to_execution_row(item) for item in record_rows]
        return loaded

    def list_executions(self, limit: int = 100) -> list[ExecutionRecord]:
        with self.connect() as conn:
            rows = conn.execute(
                "SELECT * FROM executions ORDER BY id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [self._row_to_execution(row) for row in rows]

    def get_execution(self, execution_id: int) -> ExecutionRecord:
        with self.connect() as conn:
            row = conn.execute("SELECT * FROM executions WHERE id=?", (execution_id,)).fetchone()
            row_items = conn.execute(
                "SELECT * FROM execution_rows WHERE execution_id=? ORDER BY row_index ASC",
                (execution_id,),
            ).fetchall()
        if row is None:
            raise KeyError(f"Execution {execution_id} not found")
        execution = self._row_to_execution(row)
        execution.rows = [self._row_to_execution_row(item) for item in row_items]
        return execution

    def clear_history(self) -> None:
        with self.connect() as conn:
            conn.execute("DELETE FROM execution_rows")
            conn.execute("DELETE FROM executions")

    def save_export_record(self, record: ExportHistoryRecord) -> ExportHistoryRecord:
        with self.connect() as conn:
            if record.id is None:
                cursor = conn.execute(
                    """
                    INSERT INTO export_records
                    (export_kind, title, export_folder, exported_count, summary_path, detail_json)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        record.export_kind,
                        record.title,
                        record.export_folder,
                        record.exported_count,
                        record.summary_path,
                        record.detail_json,
                    ),
                )
                record_id = int(cursor.lastrowid)
            else:
                conn.execute(
                    """
                    UPDATE export_records
                    SET export_kind=?, title=?, export_folder=?, exported_count=?, summary_path=?, detail_json=?
                    WHERE id=?
                    """,
                    (
                        record.export_kind,
                        record.title,
                        record.export_folder,
                        record.exported_count,
                        record.summary_path,
                        record.detail_json,
                        record.id,
                    ),
                )
                record_id = record.id
            row = conn.execute("SELECT * FROM export_records WHERE id=?", (record_id,)).fetchone()
        return self._row_to_export_record(row)

    def list_export_records(self, limit: int = 200) -> list[ExportHistoryRecord]:
        with self.connect() as conn:
            rows = conn.execute(
                "SELECT * FROM export_records ORDER BY id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [self._row_to_export_record(row) for row in rows]

    def clear_export_history(self) -> None:
        with self.connect() as conn:
            conn.execute("DELETE FROM export_records")

    @staticmethod
    def _row_to_template(row: sqlite3.Row) -> TaskTemplate:
        return TaskTemplate(
            id=row["id"],
            name=row["name"],
            task_type=TaskType(row["task_type"]),
            rows_json=row["rows_json"],
            created_at=row["created_at"],
            updated_at=row["updated_at"],
        )

    @staticmethod
    def _row_to_execution(row: sqlite3.Row) -> ExecutionRecord:
        return ExecutionRecord(
            id=row["id"],
            task_type=TaskType(row["task_type"]),
            started_at=row["started_at"],
            finished_at=row["finished_at"],
            status=row["status"],
            row_count=row["row_count"],
            success_count=row["success_count"],
            failure_count=row["failure_count"],
            source_template_id=row["source_template_id"],
            source_execution_id=row["source_execution_id"],
        )

    @staticmethod
    def _row_to_execution_row(row: sqlite3.Row) -> ExecutionRowResult:
        return ExecutionRowResult(
            row_index=row["row_index"],
            session_name=row["session_name"],
            success=bool(row["success"]),
            error_code=row["error_code"],
            error_message=row["error_message"],
            raw_error=row["raw_error"],
            row_payload_json=row["row_payload_json"],
        )

    @staticmethod
    def _row_to_export_record(row: sqlite3.Row) -> ExportHistoryRecord:
        return ExportHistoryRecord(
            id=row["id"],
            export_kind=row["export_kind"],
            title=row["title"],
            export_folder=row["export_folder"],
            exported_count=row["exported_count"],
            summary_path=row["summary_path"],
            detail_json=row["detail_json"],
            created_at=row["created_at"],
        )
