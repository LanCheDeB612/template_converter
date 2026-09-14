"""在后台线程执行转换并把状态传回 Tkinter 主线程。"""

from __future__ import annotations

from dataclasses import dataclass
from queue import Queue
from threading import Thread

from converter_core.contracts import ConversionIssue, ConversionRequest, ConversionResult
from converter_core.service import ConversionService


@dataclass(frozen=True)
class WorkerMessage:
    """后台线程发送给界面的进度或完成消息。"""

    kind: str
    progress: int = 0
    text: str = ""
    result: ConversionResult | None = None


class ConversionWorker:
    """一次只运行一个转换任务，避免 Excel 操作阻塞界面。"""

    def __init__(self, messages: Queue[WorkerMessage]) -> None:
        self.messages = messages
        self._thread: Thread | None = None

    @property
    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self, request: ConversionRequest) -> None:
        if self.is_running:
            raise RuntimeError("已有转换任务正在运行")
        self._thread = Thread(target=self._run, args=(request,), daemon=True)
        self._thread.start()

    def _run(self, request: ConversionRequest) -> None:
        service = ConversionService()

        def report(progress: int, text: str) -> None:
            self.messages.put(WorkerMessage(kind="progress", progress=progress, text=text))

        try:
            result = service.convert(request, progress=report)
        except Exception as error:  # Tkinter 线程不能静默退出，否则界面会一直显示“正在处理”。
            result = ConversionResult(
                success=False,
                issues=(ConversionIssue(level="error", message=f"转换程序异常：{error}"),),
            )
        self.messages.put(WorkerMessage(kind="complete", result=result))
