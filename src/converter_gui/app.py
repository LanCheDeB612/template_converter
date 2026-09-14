"""模板转换工具桌面入口。"""

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk

TEMPLATE_TYPES = ("DDS", "SOME/IP", "路由")


class TemplateConverterApp:
    """构建主界面并处理第一阶段的文件路径选择。"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.template_type = tk.StringVar(value=TEMPLATE_TYPES[0])
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="请选择待转换的 DDS 模板文件")
        self.result_text = tk.StringVar(value="尚未开始转换")

        self._build_layout()

    def _build_layout(self) -> None:
        """创建类型、路径、执行进度和结果区域。"""
        content = ttk.Frame(self.root, padding=(28, 24, 28, 20))
        content.pack(fill="both", expand=True)
        content.columnconfigure(1, weight=1)

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
        ttk.Button(
            config_frame,
            text="选择文件…",
            width=11,
            command=self._browse_input,
            bootstyle="primary",
        ).grid(row=1, column=2, padx=(12, 0), pady=8)

        ttk.Label(config_frame, text="输出目录", font=("TkDefaultFont", 11)).grid(
            row=2, column=0, sticky="e", padx=(0, 14), pady=8
        )
        ttk.Entry(
            config_frame,
            textvariable=self.output_path,
            state="readonly",
            bootstyle="primary",
        ).grid(row=2, column=1, sticky="ew", pady=8)
        ttk.Button(
            config_frame,
            text="选择目录…",
            width=11,
            command=self._browse_output,
            bootstyle="primary-outline",
        ).grid(row=2, column=2, padx=(12, 0), pady=8)

        ttk.Label(
            config_frame,
            text="仅支持 .xlsx 文件，文件路径通过按钮选择",
            bootstyle="secondary",
        ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(2, 0))

        ttk.Button(
            content,
            text="开始转换",
            width=20,
            command=self._start_conversion,
            bootstyle="success",
        ).grid(row=1, column=0, columnspan=3, pady=(22, 16))

        self.progress = ttk.Progressbar(
            content,
            mode="determinate",
            maximum=100,
            value=0,
            bootstyle="success-striped",
        )
        self.progress.grid(row=2, column=0, columnspan=3, sticky="ew")

        status_frame = ttk.Frame(content)
        status_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(
            status_frame,
            textvariable=self.status_text,
            bootstyle="secondary",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            status_frame,
            textvariable=self.result_text,
            bootstyle="secondary",
        ).grid(row=0, column=1, sticky="e")

    def _on_type_selected(self, _event: tk.Event) -> None:
        """根据下拉框选择更新当前转换提示。"""
        selected_type = self.template_type.get()
        self.status_text.set(f"请选择待转换的 {selected_type} 模板文件")
        self.result_text.set("尚未开始转换")
        self.progress.configure(value=0)
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
            self.input_path.set(selected)
            self.status_text.set("已选择输入文件，请确认输出目录")

    def _browse_output(self) -> None:
        """选择转换结果的保存目录。"""
        selected = filedialog.askdirectory(parent=self.root, title="选择输出目录")
        if selected:
            self.output_path.set(selected)
            self.status_text.set("路径已准备，可以开始转换")

    def _start_conversion(self) -> None:
        """校验用户输入，并说明当前阶段的功能边界。"""
        selected_type = self.template_type.get()
        input_file = Path(self.input_path.get()) if self.input_path.get() else None
        output_dir = Path(self.output_path.get()) if self.output_path.get() else None

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

        # 初始化阶段只验证界面输入，避免在转换核心接入前产生伪成功结果。
        self.progress.configure(value=0)
        self.status_text.set(f"路径校验通过，{selected_type} 转换核心尚未接入。")
        self.result_text.set("未生成文件")
        messagebox.showinfo(
            "功能开发中",
            f"输入和输出路径校验已通过。\n{selected_type} 转换核心尚未接入。",
            parent=self.root,
        )


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
    root.minsize(680, 380)
    root.resizable(True, False)
    TemplateConverterApp(root)
    _center_window(root, 760, 380)
    return root


def main() -> None:
    """启动桌面程序事件循环。"""
    build_app().mainloop()
