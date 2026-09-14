"""转换任务统一入口。"""

from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Callable
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from converter_core.contracts import ConversionIssue, ConversionRequest, ConversionResult
from converter_core.converters.dds.converter import build_conversion_data
from converter_core.converters.dds.errors import DdsConversionError
from converter_core.converters.dds.reader import read_customer_workbook
from converter_core.excel_io import MATRIX_SHEET, NODE_SHEET, write_dds_template_copy
from converter_core.output import available_output_path
from converter_core.resources import dds_template_path

ProgressCallback = Callable[[int, str], None]


class ConversionService:
    """隔离界面和协议转换逻辑，并组织 DDS 输出文件的安全发布。"""

    def convert(
        self,
        request: ConversionRequest,
        progress: ProgressCallback | None = None,
    ) -> ConversionResult:
        """执行一次 DDS 转换，成功验证临时文件后才发布最终结果。"""
        notify = progress or (lambda _value, _message: None)
        problem = self._validate_request(request)
        if problem is not None:
            return ConversionResult(success=False, issues=(problem,))

        temporary_path: Path | None = None
        try:
            notify(10, "正在读取客户 DDS 模板")
            deployments, topics = read_customer_workbook(request.input_file)
            notify(35, "正在关联服务、Topic 和 Participant")
            data = build_conversion_data(deployments, topics)

            template = dds_template_path()
            if not template.is_file():
                return ConversionResult(
                    success=False,
                    issues=(ConversionIssue(level="error", message=f"找不到标准模板：{template}"),),
                )

            output_file = available_output_path(
                request.input_file,
                request.output_directory,
                request.output_filename,
            )
            with NamedTemporaryFile(
                prefix=".dds-converting-",
                suffix=".xlsx",
                dir=request.output_directory,
                delete=False,
            ) as temporary:
                temporary_path = Path(temporary.name)

            notify(60, "正在复制标准模板并写入转换数据")
            write_dds_template_copy(template, temporary_path, data)
            notify(85, "正在重新打开并检查转换结果")
            self._verify_output(temporary_path, len(data.nodes), len(data.matrix_rows))
            temporary_path.replace(output_file)
            temporary_path = None
            notify(100, "DDS 转换完成")
            return ConversionResult(
                success=True,
                output_file=output_file,
                node_count=len(data.nodes),
                matrix_count=len(data.matrix_rows),
            )
        except DdsConversionError as error:
            return ConversionResult(success=False, issues=error.issues)
        except (BadZipFile, InvalidFileException, OSError, ParseError, ValueError) as error:
            return ConversionResult(
                success=False,
                issues=(ConversionIssue(level="error", message=f"DDS 转换失败：{error}"),),
            )
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)

    @staticmethod
    def _validate_request(request: ConversionRequest) -> ConversionIssue | None:
        if request.template_type != "DDS":
            return ConversionIssue(level="error", message=f"{request.template_type} 转换功能尚未实现。")
        if not request.input_file.is_file() or request.input_file.suffix.casefold() != ".xlsx":
            return ConversionIssue(level="error", message="请选择有效的 .xlsx 输入文件。")
        if not request.output_directory.is_dir():
            return ConversionIssue(level="error", message="请选择有效的输出目录。")
        return None

    @staticmethod
    def _verify_output(path: Path, expected_nodes: int, expected_matrix_rows: int) -> None:
        workbook = load_workbook(path, read_only=False, data_only=False)
        if workbook.sheetnames != [NODE_SHEET, MATRIX_SHEET]:
            raise ValueError("输出文件的 Sheet 结构发生变化。")
        node_sheet = workbook[NODE_SHEET]
        matrix_sheet = workbook[MATRIX_SHEET]
        node_count = sum(
            1 for row in range(2, node_sheet.max_row + 1) if node_sheet.cell(row, 1).value is not None
        )
        matrix_count = sum(
            1 for row in range(2, matrix_sheet.max_row + 1) if matrix_sheet.cell(row, 1).value is not None
        )
        if node_count != expected_nodes or matrix_count != expected_matrix_rows:
            raise ValueError(
                f"输出重新打开后的记录数不一致：节点 {node_count}/{expected_nodes}，"
                f"通信矩阵 {matrix_count}/{expected_matrix_rows}。"
            )
