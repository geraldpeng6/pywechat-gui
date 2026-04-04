from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtWidgets import QApplication, QMessageBox, QPushButton
except ModuleNotFoundError:  # pragma: no cover - optional GUI dependency in local test env
    QApplication = None  # type: ignore[assignment]
    QMessageBox = None  # type: ignore[assignment]
    QPushButton = None  # type: ignore[assignment]

from pyweixin_gui.models import RelayCollectMode, RelayItemType, RelayPackageRow, RelayRecentRange, RelayRouteRow
if QApplication is not None:
    from pyweixin_gui.ui.widgets import ExportPage, RelayWorkbenchPage, RemovableChip, ResourceToolsPage
else:  # pragma: no cover - optional GUI dependency in local test env
    ExportPage = None  # type: ignore[assignment]
    RelayWorkbenchPage = None  # type: ignore[assignment]
    RemovableChip = None  # type: ignore[assignment]
    ResourceToolsPage = None  # type: ignore[assignment]


@unittest.skipIf(QApplication is None, "当前环境未安装 PySide6")
class RelayWorkbenchPageTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_template_payload_roundtrip_restores_collect_settings(self):
        page = RelayWorkbenchPage()
        page.source_session_input.setText("上游A")
        page.package_name_input.setText("周报发送")
        page.message_limit_spin.setValue(42)
        page.file_limit_spin.setValue(9)
        page.collect_mode_combo.setCurrentIndex(
            page.collect_mode_combo.findData(RelayCollectMode.PERIOD.value)
        )
        page.collect_recent_combo.setCurrentIndex(
            page.collect_recent_combo.findData(RelayRecentRange.MONTH.value)
        )
        page.collect_sender_chips.set_values(["张三", "李四"])
        page.append_package_rows([RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="第一条")])
        page.append_route_rows([RelayRouteRow(downstream_session="客户群")])

        restored = RelayWorkbenchPage()
        restored.load_template_payload(page.template_payload())

        self.assertEqual(restored.source_session_input.text(), "上游A")
        self.assertEqual(restored.package_name_input.text(), "周报发送")
        self.assertEqual(restored.message_limit_spin.value(), 42)
        self.assertEqual(restored.file_limit_spin.value(), 9)
        self.assertEqual(restored.current_collect_mode(), RelayCollectMode.PERIOD)
        self.assertEqual(restored.current_recent_range(), RelayRecentRange.MONTH)
        self.assertEqual(restored.collect_sender_chips.values(), ["张三", "李四"])
        self.assertEqual(len(restored.package_rows()), 1)
        self.assertEqual(len(restored.route_rows()), 1)

    def test_running_state_locks_tables_chip_and_row_delete_buttons(self):
        page = RelayWorkbenchPage()
        page.collect_sender_chips.set_values(["张三"])
        page.append_package_rows([RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="第一条")])
        page.append_route_rows([RelayRouteRow(downstream_session="客户群")])

        page.set_running_state(True)

        self.assertFalse(page.package_table.isEnabled())
        self.assertFalse(page.route_table.isEnabled())
        chip_item = page.collect_sender_chips.chip_layout.itemAt(0)
        chip_widget = chip_item.widget() if chip_item is not None else None
        self.assertIsInstance(chip_widget, RemovableChip)
        self.assertFalse(chip_widget.delete_button.isEnabled())
        package_remove_button = self._row_remove_button(page.package_table, page.PACKAGE_COLUMNS)
        route_remove_button = self._row_remove_button(page.route_table, page.ROUTE_COLUMNS)
        self.assertIsNotNone(package_remove_button)
        self.assertIsNotNone(route_remove_button)
        self.assertFalse(package_remove_button.isEnabled())
        self.assertFalse(route_remove_button.isEnabled())

        page.set_running_state(False)

        self.assertTrue(page.package_table.isEnabled())
        self.assertTrue(page.route_table.isEnabled())
        self.assertTrue(chip_widget.delete_button.isEnabled())
        self.assertTrue(package_remove_button.isEnabled())
        self.assertTrue(route_remove_button.isEnabled())

    def test_export_page_ignores_rapid_repeat_submit(self):
        page = ExportPage()
        page.session_name_input.setText("项目群")
        page.target_folder_input.setText("C:/exports")
        emitted = []
        page.export_requested.connect(emitted.append)

        page._request_export()
        page._request_export()

        self.assertEqual(len(emitted), 1)
        self.assertIn("不必重复点击", page.summary_label.text())

    def test_resource_export_page_ignores_rapid_repeat_submit(self):
        page = ResourceToolsPage()
        page.target_folder_input.setText("C:/exports")
        emitted = []
        page.export_requested.connect(emitted.append)

        page._request_export()
        page._request_export()

        self.assertEqual(len(emitted), 1)
        self.assertIn("不必重复点击", page.summary_label.text())

    def test_relay_package_export_ignores_rapid_repeat_submit(self):
        page = RelayWorkbenchPage()
        page.append_package_rows([RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="第一条")])
        emitted = []
        page.export_package_requested.connect(emitted.append)

        with patch("pyweixin_gui.ui.widgets.QFileDialog.getExistingDirectory", return_value="/tmp"):
            page._request_export_package()
            page._request_export_package()

        self.assertEqual(len(emitted), 1)
        self.assertIn("不必重复点击", page.result_text.toPlainText())

    def test_focus_package_sequence_selects_matching_row(self):
        page = RelayWorkbenchPage()
        page.append_package_rows(
            [
                RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="第一条"),
                RelayPackageRow(sequence=2, item_type=RelayItemType.IMAGE, content="海报.png", file_path="/tmp/海报.png"),
            ]
        )

        focused = page.focus_package_sequence(2, "已定位到失败内容。")

        self.assertTrue(focused)
        self.assertEqual(page.package_table.currentRow(), 1)
        self.assertIn("已定位到失败内容", page.result_text.toPlainText())

    def test_request_send_blocks_when_route_validation_failed(self):
        page = RelayWorkbenchPage()
        page.append_package_rows([RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="第一条")])
        page.append_route_rows([RelayRouteRow(downstream_session="客户群", validation_status="未找到")])
        emitted = []
        page.send_requested.connect(emitted.append)

        with patch("pyweixin_gui.ui.widgets.QMessageBox.warning") as warning:
            page._request_send()

        self.assertEqual(len(emitted), 0)
        warning.assert_called_once()
        self.assertIn("已拦截正式发送", page.result_text.toPlainText())

    def test_request_send_confirms_when_route_not_validated(self):
        page = RelayWorkbenchPage()
        page.append_package_rows([RelayPackageRow(sequence=1, item_type=RelayItemType.TEXT, content="第一条")])
        page.append_route_rows([RelayRouteRow(downstream_session="客户群", validation_status="未验证")])
        emitted = []
        page.send_requested.connect(emitted.append)

        with patch("pyweixin_gui.ui.widgets.QMessageBox.question", return_value=QMessageBox.StandardButton.No) as question:
            page._request_send()
        self.assertEqual(len(emitted), 0)
        question.assert_called_once()
        self.assertIn("已取消正式发送", page.result_text.toPlainText())

        with patch("pyweixin_gui.ui.widgets.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            page._request_send()
        self.assertEqual(len(emitted), 1)

    @staticmethod
    def _row_remove_button(table, columns) -> QPushButton | None:
        remove_column = next(index for index, (key, _, _) in enumerate(columns) if key == "remove")
        widget = table.cellWidget(0, remove_column)
        return widget if isinstance(widget, QPushButton) else None


if __name__ == "__main__":
    unittest.main()
