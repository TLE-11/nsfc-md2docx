@echo off
rem Windows 入口。用法与 md2docx.py 完全一致，例如：
rem   md2docx.bat 本子.md
rem   md2docx.bat 本子.md -o 本子.docx --number-style chapter
rem   md2docx.bat 本子.md --math-mode latex
rem
rem 依赖：Python 3（python.org 安装时勾上 Add to PATH）+ pandoc
rem   winget install --id Python.Python.3.12
rem   winget install --id JohnMacFarlane.Pandoc

setlocal
set "HERE=%~dp0"

rem 优先用 py 启动器，没有再退回 python
where py >nul 2>nul
if %ERRORLEVEL%==0 (
  py -3 "%HERE%md2docx.py" %*
) else (
  where python >nul 2>nul
  if %ERRORLEVEL%==0 (
    python "%HERE%md2docx.py" %*
  ) else (
    echo 找不到 Python。装法: winget install --id Python.Python.3.12
    exit /b 1
  )
)
exit /b %ERRORLEVEL%
