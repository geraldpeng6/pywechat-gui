from __future__ import annotations

import unittest

from pyweixin_gui.models import GroupMemberRow, GroupMembersRequest, GroupSummaryRow, RuntimeOptions, SessionScanRequest, SessionSummaryRow
from pyweixin_gui.session_tools_service import SessionToolsService


class FakeAdapter:
    def dump_sessions(self, chatted_only, no_official, options):
        return [
            SessionSummaryRow(session_name="项目群", last_time="昨天", last_message="已收到"),
            SessionSummaryRow(session_name="客户A", last_time="今天", last_message="请报价"),
        ]

    def dump_groups(self, options):
        return [
            GroupSummaryRow(group_name="项目群", member_count="28"),
            GroupSummaryRow(group_name="通知群", member_count="56"),
        ]

    def dump_group_members(self, group_name, options):
        return [
            GroupMemberRow(group_name=group_name, alias="张三", nickname="张总"),
            GroupMemberRow(group_name=group_name, alias="李四", nickname="李经理"),
        ]


class SessionToolsServiceTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.service = SessionToolsService(FakeAdapter())
        self.options = RuntimeOptions()

    def test_scan_sessions(self):
        result = self.service.scan_sessions(SessionScanRequest(chatted_only=True, no_official=True), self.options)
        self.assertEqual(len(result.rows), 2)
        self.assertEqual(result.rows[0].session_name, "项目群")

    def test_scan_groups(self):
        result = self.service.scan_groups(self.options)
        self.assertEqual(result.rows[1].group_name, "通知群")

    def test_load_group_members(self):
        result = self.service.load_group_members(GroupMembersRequest(group_name="项目群"), self.options)
        self.assertEqual(result.member_count, 2)
        self.assertEqual(result.rows[0].alias, "张三")

    def test_load_group_members_requires_group_name(self):
        with self.assertRaises(ValueError):
            self.service.load_group_members(GroupMembersRequest(group_name=""), self.options)


if __name__ == "__main__":
    unittest.main()
