"""定位源码运行和 PyInstaller 打包环境中的标准模板。"""

from pathlib import Path
import sys

DDS_TEMPLATE_NAME = "DDS通信矩阵.xlsx"


def dds_template_path() -> Path:
    """返回随程序发布的 DDS 标准输出模板路径。"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base_directory = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        base_directory = Path(__file__).resolve().parents[2]
    return base_directory / "resources" / "templates" / DDS_TEMPLATE_NAME
