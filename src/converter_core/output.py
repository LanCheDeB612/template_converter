"""生成不覆盖已有结果的输出文件名。"""

from pathlib import Path

_INVALID_FILENAME_CHARACTERS = frozenset('<>:"/\\|?*')


def normalise_output_filename(output_filename: str | None) -> str | None:
    """规范自定义文件名，并拒绝 Windows 无法保存的名称。"""
    if output_filename is None or not output_filename.strip():
        return None

    filename = output_filename.strip()
    invalid = sorted(set(filename) & _INVALID_FILENAME_CHARACTERS)
    if invalid:
        rendered = " ".join(invalid)
        raise ValueError(f"输出文件名不能包含以下字符：{rendered}")
    if filename in {".", ".."} or filename.endswith((".", " ")):
        raise ValueError("输出文件名不能以点或空格结尾。")
    if not filename.casefold().endswith(".xlsx"):
        filename = f"{filename}.xlsx"
    if not filename[:-5].strip():
        raise ValueError("输出文件名不能为空。")
    return filename


def available_output_path(
    input_file: Path,
    output_directory: Path,
    output_filename: str | None = None,
) -> Path:
    """结果文件已存在时增加数字后缀，避免覆盖人工补填过的文件。"""
    custom_filename = normalise_output_filename(output_filename)
    base_name = (
        Path(custom_filename).stem
        if custom_filename is not None
        else f"{input_file.stem}_DDS通信矩阵"
    )
    candidate = output_directory / f"{base_name}.xlsx"
    index = 1
    while candidate.exists():
        candidate = output_directory / f"{base_name}_{index}.xlsx"
        index += 1
    return candidate
