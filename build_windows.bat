@echo off
setlocal
chcp 65001 >nul

rem 始终切换到脚本所在目录，避免从其他目录启动时找不到工程文件和模板。
cd /d "%~dp0"

where uv >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 uv，请先安装 uv 并重新打开命令行。
    goto :failed
)

if not exist "pyproject.toml" (
    echo [错误] 当前目录缺少 pyproject.toml：%CD%
    goto :failed
)

if not exist "resources\templates\DDS通信矩阵.xlsx" (
    echo [错误] 缺少 DDS 标准模板：resources\templates\DDS通信矩阵.xlsx
    goto :failed
)

if not exist "resources\templates\SOMEIP通信矩阵模板.xlsx" (
    echo [错误] 缺少 SOME/IP 标准模板：resources\templates\SOMEIP通信矩阵模板.xlsx
    goto :failed
)

if not exist "resources\templates\路由输入模板.xlsx" (
    echo [错误] 缺少路由标准模板：resources\templates\路由输入模板.xlsx
    goto :failed
)

echo [1/2] 正在按 uv.lock 同步构建依赖...
uv sync --locked
if errorlevel 1 goto :failed

echo [2/2] 正在生成 Windows 单文件程序...
rem 三个模板必须保持 resources\templates 目录结构，程序解压后才能按相同路径找到它们。
uv run pyinstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name 模板转换工具 ^
    --paths src ^
    --collect-all ttkbootstrap ^
    --add-data "resources\templates\DDS通信矩阵.xlsx:resources\templates" ^
    --add-data "resources\templates\SOMEIP通信矩阵模板.xlsx:resources\templates" ^
    --add-data "resources\templates\路由输入模板.xlsx:resources\templates" ^
    main.py
if errorlevel 1 goto :failed

if not exist "dist\模板转换工具.exe" (
    echo [错误] PyInstaller 未报告错误，但未生成 dist\模板转换工具.exe。
    goto :failed
)

rem 新的中文名称生成成功后再删除旧名称，避免打包失败时丢失原有可执行文件。
if exist "dist\TemplateConverter.exe" del /q "dist\TemplateConverter.exe"

echo.
echo 打包完成：%CD%\dist\模板转换工具.exe
echo 请在 Windows 上分别执行一次 DDS、SOME/IP 和路由转换进行验收。
pause
exit /b 0

:failed
echo.
echo 打包失败，请检查上方错误信息。
pause
exit /b 1
