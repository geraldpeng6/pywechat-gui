from __future__ import annotations

import unittest

from pyweixin_gui.models import GroupSummaryRow, RuntimeOptions, SessionScanRequest, SessionSummaryRow
from pyweixin_gui.session_tools_service import SessionToolsService


class FakeAdapter:
    def dump_sessions(self, chatted_only, options):
        return [
            SessionSummaryRow(session_name="项目群", last_time="昨天", last_message="已收到"),
            SessionSummaryRow(session_name="客户A", last_time="今天", last_message="请报价"),
        ]

    def dump_groups(self, options):
        return [
            GroupSummaryRow(group_name="项目群", member_count="28"),
            GroupSummaryRow(group_name="通知群", member_count="56"),
        ]


class SessionToolsServiceTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.service = SessionToolsService(FakeAdapter())
        self.options = RuntimeOptions()

    def test_scan_sessions(self):
        result = self.service.scan_sessions(SessionScanRequest(chatted_only=True), self.options)
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.rows[0].session_name, "项目群")

    def test_scan_groups(self):
        result = self.service.scan_groups(self.options)
        self.assertEqual(result.rows[1].group_name, "通知群")

if __name__ == "__main__":
    unittest.main()
