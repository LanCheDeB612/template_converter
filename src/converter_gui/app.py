"""模板转换工具桌面入口。"""

from datetime import datetime
from pathlib import Path
from queue import Empty, Queue
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk

from converter_core.contracts import ConversionRequest, ConversionResult
from converter_core.output import normalise_output_filename
from converter_gui.worker import ConversionWorker, WorkerMessage

TEMPLATE_TYPES = ("DDS", "SOME/IP", "路由")


class TemplateConverterApp:
    """构建主界面并处理第一阶段的文件路径选择。"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.template_type = tk.StringVar(value=TEMPLATE_TYPES[0])
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.output_name = tk.StringVar()
        self.worker_messages: Queue[WorkerMessage] = Queue()
        self.worker = ConversionWorker(self.worker_messages)

        self._build_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_layout(self) -> None:
        """创建类型、路径、输出文件名和结果区域。"""
        content = ttk.Frame(self.root, padding=(28, 24, 28, 20))
        content.pack(fill="both", expand=True)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(2, weight=1)

        config_frame = ttk.Labelframe(
            content,
            text="转换配置",
            padding=(20, 14),
            bootstyle="primary",
        )
        config_frame.grid(row=0, column=0, columnspan=3, sticky="nsew")
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="模板类型", font=("TkDefaultFont", 11)).grid(
            row=0, column=0, sticky="e", padx=(0, 14), pady=8
        )
        self.type_box = ttk.Combobox(
            config_frame,
            textvariable=self.template_type,
            values=TEMPLATE_TYPES,
            state="readonly",
            width=18,
            bootstyle="primary",
        )
        self.type_box.grid(row=0, column=1, sticky="w", pady=8)
        self.type_box.bind("<<ComboboxSelected>>", self._on_type_selected)

        ttk.Label(config_frame, text="输入文件", font=("TkDefaultFont", 11)).grid(
            row=1, column=0, sticky="e", padx=(0, 14), pady=8
        )
        ttk.Entry(
            config_frame,
            textvariable=self.input_path,
            state="readonly",
            bootstyle="primary",
        ).grid(row=1, column=1, sticky="ew", pady=8)
        self.input_button = ttk.Button(
            config_frame,
            text="选择文件…",
            width=11,
            command=self._browse_input,
            bootstyle="primary",
        )
        self.input_button.grid(row=1, column=2, padx=(12, 0), pady=8)

        ttk.Label(config_frame, text="输出目录", font=("TkDefaultFont", 11)).grid(
            row=2, column=0, sticky="e", padx=(0, 14), pady=8
        )
        ttk.Entry(
            config_frame,
            textvariable=self.output_path,
            state="readonly",
            bootstyle="primary",
        ).grid(row=2, column=1, sticky="ew", pady=8)
        self.output_button = ttk.Button(
            config_frame,
            text="选择目录…",
            width=11,
            command=self._browse_output,
            bootstyle="primary-outline",
        )
        self.output_button.grid(row=2, column=2, padx=(12, 0), pady=8)

        ttk.Label(config_frame, text="输出文件名", font=("TkDefaultFont", 11)).grid(
            row=3, column=0, sticky="e", padx=(0, 14), pady=8
        )
        self.output_name_entry = ttk.Entry(
            config_frame,
            textvariable=self.output_name,
            bootstyle="primary",
        )
        self.output_name_entry.grid(row=3, column=1, columnspan=2, sticky="ew", pady=8)

        ttk.Label(
            config_frame,
            text="仅支持 .xlsx 文件；输出文件名可省略后缀，重名时自动增加数字",
            bootstyle="secondary",
        ).grid(row=4, column=1, columnspan=2, sticky="w", pady=(2, 0))

        self.start_button = ttk.Button(
            content,
            text="开始转换",
            width=20,
            command=self._start_conversion,
            bootstyle="success",
        )
        self.start_button.grid(row=1, column=0, columnspan=3, pady=(22, 16))

        progress_frame = ttk.Labelframe(
            content,
            text="转换进度",
            padding=(20, 14),
            bootstyle="primary",
        )
        progress_frame.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(2, 0))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)

        log_header = ttk.Frame(progress_frame)
        log_header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(log_header, text="运行日志", font=("TkDefaultFont", 11)).grid(
            row=0, column=0, sticky="w"
        )

        self.log_frame = ttk.Frame(progress_frame)
        self.log_frame.grid(row=1, column=0, sticky="nsew")
        self.log_frame.columnconfigure(0, weight=1)
        self.log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            self.log_frame,
            height=7,
            wrap="word",
            state="disabled",
            relief="flat",
            borderwidth=0,
            padx=10,
            pady=8,
            font=("TkDefaultFont", 10),
            foreground="#2c3e50",
            background="#f1f4f5",
            highlightthickness=1,
            highlightbackground="#2c3e50",
            highlightcolor="#2c3e50",
        )
        log_scrollbar = ttk.Scrollbar(
            self.log_frame,
            orient="vertical",
            command=self.log_text.yview,
            bootstyle="primary",
        )
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.tag_configure("error", foreground="#c0392b")
        self.log_text.tag_configure("warning", foreground="#b9770e")
        self.log_text.tag_configure("success", foreground="#148f77")

    def _on_type_selected(self, _event: tk.Event) -> None:
        """选择模板类型后清除只读下拉框的文字选中背景。"""
        self.root.after_idle(self._clear_type_selection)

    def _clear_type_selection(self) -> None:
        """清除只读下拉框的文字选中背景。"""
        self.type_box.selection_clear()
        self.root.focus_set()

    def _browse_input(self) -> None:
        """选择一个 xlsx 输入文件。"""
        selected_type = self.template_type.get()
        selected = filedialog.askopenfilename(
            parent=self.root,
            title=f"选择 {selected_type} 模板文件",
            filetypes=(("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")),
        )
        if selected:
            selected_path = Path(selected)
            self.input_path.set(str(selected_path))
            # 默认把结果保存在输入文件旁边，用户仍可通过“选择目录”覆盖。
            self.output_path.set(str(selected_path.parent))
            if not self.output_name.get().strip():
                self.output_name.set(f"{selected_path.stem}_DDS通信矩阵.xlsx")
            self._append_log(f"已选择输入文件：{selected_path}")
            self._append_log(f"已自动设置输出目录：{selected_path.parent}")

    def _browse_output(self) -> None:
        """选择转换结果的保存目录。"""
        selected = filedialog.askdirectory(parent=self.root, title="选择输出目录")
        if selected:
            self.output_path.set(selected)
            self._append_log(f"已选择输出目录：{selected}")

    def _start_conversion(self) -> None:
        """校验用户输入，并启动不会阻塞界面的后台转换任务。"""
        selected_type = self.template_type.get()
        input_file = Path(self.input_path.get()) if self.input_path.get() else None
        output_dir = Path(self.output_path.get()) if self.output_path.get() else None
        output_filename = self.output_name.get().strip() or None

        if input_file is None:
            messagebox.showwarning(
                "缺少输入文件",
                f"请先选择待转换的 {selected_type} 模板文件。",
                parent=self.root,
            )
            return
        if not input_file.is_file() or input_file.suffix.lower() != ".xlsx":
            messagebox.showerror("输入文件无效", "请选择有效的 .xlsx 文件。", parent=self.root)
            return
        if output_dir is None:
            messagebox.showwarning("缺少输出目录", "请选择转换结果的保存目录。", parent=self.root)
            return
        if not output_dir.is_dir():
            messagebox.showerror("输出目录无效", "请选择有效的输出目录。", parent=self.root)
            return
        try:
            output_filename = normalise_output_filename(output_filename)
        except ValueError as error:
            messagebox.showerror("输出文件名无效", str(error), parent=self.root)
            self.output_name_entry.focus_set()
            return

        if selected_type != "DDS":
            messagebox.showinfo(
                "功能开发中",
                f"{selected_type} 转换功能尚未实现。",
                parent=self.root,
            )
            return

        self._clear_log()
        self._append_log(f"开始转换：{input_file.name}")
        self._set_controls_enabled(False)
        self.worker.start(
            ConversionRequest(
                input_file=input_file,
                output_directory=output_dir,
                template_type=selected_type,
                output_filename=output_filename,
            )
        )
        self.root.after(80, self._poll_worker)

    def _poll_worker(self) -> None:
        """只在 Tkinter 主线程更新控件，避免后台线程直接操作界面。"""
        complete = False
        while True:
            try:
                message = self.worker_messages.get_nowait()
            except Empty:
                break
            if message.kind == "progress":
                self._append_log(f"[{message.progress}%] {message.text}")
            elif message.kind == "complete" and message.result is not None:
                complete = True
                self._show_result(message.result)
        # complete 消息在线程退出前入队；即使恰好发生在线程状态切换瞬间，
        # 再轮询一次也能取到最终结果，不会让界面永久停留在“正在处理”。
        if not complete:
            self.root.after(80, self._poll_worker)

    def _show_result(self, result: ConversionResult) -> None:
        """展示转换数量、输出路径或能够定位到客户 Sheet 的问题。"""
        self._set_controls_enabled(True)
        if result.success and result.output_file is not None:
            self._append_log(
                f"转换完成：节点 {result.node_count}，通信记录 {result.matrix_count}",
                "success",
            )
            self._append_log(f"输出文件：{result.output_file}", "success")
            messagebox.showinfo(
                "转换完成",
                f"已生成：\n{result.output_file}\n\n节点：{result.node_count}\n通信记录：{result.matrix_count}",
                parent=self.root,
            )
            return

        details = []
        for issue in result.issues:
            location = " / ".join(
                part
                for part in (
                    issue.sheet,
                    f"第 {issue.row} 行" if issue.row is not None else None,
                    issue.column,
                )
                if part
            )
            detail = f"{location + '：' if location else ''}{issue.message}"
            details.append(detail)
            self._append_log(detail, "warning" if issue.level == "warning" else "error")
        messagebox.showerror(
            "转换失败",
            "\n".join(details) if details else "转换失败，未生成文件。",
            parent=self.root,
        )

    def _append_log(self, message: str, tag: str | None = None) -> None:
        """将带时间的转换阶段或问题写入只读日志，并保持最新消息可见。"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n", tag or ())
        self.log_text.configure(state="disabled")
        self.log_text.see("end")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.input_button.configure(state=state)
        self.output_button.configure(state=state)
        self.start_button.configure(state=state, text="开始转换" if enabled else "转换中…")
        self.output_name_entry.configure(state=state)
        self.type_box.configure(state="readonly" if enabled else "disabled")

    def _on_close(self) -> None:
        """转换仍在执行时提示用户，避免误以为关闭窗口等同于正常完成。"""
        if self.worker.is_running and not messagebox.askyesno(
            "转换仍在执行",
            "转换尚未完成，确定要关闭程序吗？",
            parent=self.root,
        ):
            return
        self.root.destroy()


def _center_window(root: tk.Tk, width: int, height: int) -> None:
    """将主窗口放到当前屏幕中央。"""
    root.update_idletasks()
    x = max((root.winfo_screenwidth() - width) // 2, 0)
    y = max((root.winfo_screenheight() - height) // 2, 0)
    root.geometry(f"{width}x{height}+{x}+{y}")


def build_app() -> tk.Tk:
    """创建工程初始化阶段使用的主窗口。"""
    # 禁用 ttkbootstrap 的彩色圆点图标，保留 Tkinter 原生默认程序图标。
    root = ttk.Window(themename="flatly", iconphoto=None)
    root.title("模板转换工具")
    root.minsize(680, 460)
    root.resizable(True, True)
    TemplateConverterApp(root)
    _center_window(root, 760, 560)
    return root


def main() -> None:
    """启动桌面程序事件循环。"""
    build_app().mainloop()
