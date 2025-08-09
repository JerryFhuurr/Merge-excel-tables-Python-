@echo off
:: Excel 表格合并批处理脚本
:: 此脚本运行 mergeTable.py Python 脚本

echo ======================================================
echo           Excel 表格合并工具
echo ======================================================
echo 正在开始合并过程...
echo.

:: 检查是否安装了 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未安装 Python 或 Python 未添加到 PATH 环境变量
    echo 请安装 Python 后重试
    pause
    exit /b 1
)

:: 检查 mergeTable.py 是否存在
if not exist "mergeTable.py" (
    echo 错误：当前目录中未找到 mergeTable.py
    echo 请确保 mergeTable.py 与此批处理文件在同一文件夹中
    pause
    exit /b 1
)

:: 运行 Python 脚本
echo 正在运行 mergeTable.py...
echo.
python mergeTable.py

:: 检查脚本是否成功运行
if errorlevel 1 (
    echo.
    echo 错误：脚本执行失败
    echo 请查看上方的错误信息
) else (
    echo.
    echo ======================================================
    echo           合并过程完成！
    echo ======================================================
    echo 请查看 logs 文件夹获取详细信息
    echo 输出文件：1.xlsx
)

echo.
pause
