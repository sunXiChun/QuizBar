@echo off
chcp 65001 >nul
where pythonw >nul 2>nul
if errorlevel 1 (
  echo [ExcelFloat] 未检测到 Python，请先安装并勾选 "Add to PATH"
  pause
  exit /b 1
)
python -c "import openpyxl, PIL, winsdk" 2>nul
if errorlevel 1 (
  echo [ExcelFloat] 正在安装依赖 openpyxl / pillow / winsdk ...
  python -m pip install openpyxl pillow winsdk
)
start "" pythonw "%~dp0excelfloat.pyw" %*
