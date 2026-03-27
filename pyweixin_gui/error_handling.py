from __future__ import annotations

import traceback
from dataclasses import dataclass


@dataclass
class UiError:
    code: str
    title: str
    message: str
    suggestion: str
    raw_error: str
    traceback_text: str

    @property
    def user_summary(self) -> str:
        return f"{self.title}\n\n{self.message}\n\n建议：{self.suggestion}"

    @property
    def diagnostic_text(self) -> str:
        return (
            f"[{self.code}] {self.title}\n"
            f"{self.message}\n\n"
            f"建议: {self.suggestion}\n\n"
            f"原始错误:\n{self.raw_error}\n\n"
            f"Traceback:\n{self.traceback_text}"
        )


_KNOWN_ERRORS = {
    "UnsupportedPlatformError": (
        "UNSUPPORTED_PLATFORM",
        "当前平台不受支持",
        "该 GUI 只能在 Windows 10/11 桌面环境中执行微信自动化。",
    ),
    "NotStartError": ("WECHAT_NOT_STARTED", "微信未启动", "请先启动 PC 微信后再执行任务。"),
    "NotLoginError": ("WECHAT_NOT_LOGGED_IN", "微信未登录", "请先登录微信，再回到工具中重试。"),
    "NotFoundError": (
        "WECHAT_UI_NOT_FOUND",
        "无法识别微信主界面",
        "请尝试在微信登录前开启 Windows 讲述人/无障碍服务，登录后再重新连接。",
    ),
    "NoSuchFriendError": ("SESSION_NOT_FOUND", "会话不存在", "请检查会话名称是否与微信中的备注或群名完全一致。"),
    "NotFriendError": ("SESSION_TYPE_UNSUPPORTED", "当前目标不支持该操作", "请确认目标是可聊天的好友或群聊。"),
    "NoFilesToSendError": ("FILES_INVALID", "没有可发送的文件", "请检查文件是否存在、非空且小于 1GB。"),
    "TimeNotCorrectError": ("INVALID_TIME", "时间格式不合法", "请使用库支持的时间格式。"),
    "NetWorkError": ("NETWORK_ERROR", "网络不可用", "请确认网络连接正常，再重新执行。"),
    "ValueError": ("INVALID_INPUT", "输入无效", "请检查当前任务行中的必填项和字段格式。"),
}


def map_exception(exc: BaseException) -> UiError:
    exc_name = exc.__class__.__name__
    code, title, suggestion = _KNOWN_ERRORS.get(
        exc_name,
        ("UNKNOWN_ERROR", "未知错误", "请复制诊断信息并查看日志定位问题。"),
    )
    return UiError(
        code=code,
        title=title,
        message=str(exc) or title,
        suggestion=suggestion,
        raw_error=repr(exc),
        traceback_text="".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
    )
