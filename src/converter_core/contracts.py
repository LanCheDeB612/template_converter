"""GUI 与转换核心之间使用的数据结构。"""

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class ConversionRequest:
    """描述一次转换所需的用户输入。"""

    input_file: Path
    output_directory: Path
    template_type: str = "DDS"
    output_filename: str | None = None


@dataclass(frozen=True)
class ConversionIssue:
    """记录能够定位到客户 Excel 的转换问题。"""

    level: str
    message: str
    sheet: str | None = None
    row: int | None = None
    column: str | None = None


@dataclass(frozen=True)
class ConversionCount:
    """描述一种输出记录的名称和数量，供不同模板共用结果界面。"""

    label: str
    value: int


@dataclass(frozen=True)
class ConversionResult:
    """向界面返回转换结果和问题清单。"""

    success: bool
    output_file: Path | None = None
    issues: tuple[ConversionIssue, ...] = field(default_factory=tuple)
    counts: tuple[ConversionCount, ...] = field(default_factory=tuple)
    # 保留 DDS 专用字段，避免已有调用方在接入通用数量摘要后失效。
    node_count: int = 0
    matrix_count: int = 0
