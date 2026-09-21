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

rem 优先用 py 启动器，没有再退回 python。
rem 注意：嵌套 if 里不能再用 %ERRORLEVEL%——cmd 会在解析外层括号块时把所有
rem %VAR% 一次性展开冻结，内层读到的是 where py 之后的旧值（导致只有 python、
rem 没有 py 启动器的机器被误报"找不到 Python"）。if errorlevel N 是运行时
rem 求值（>=N 为真），没有这个陷阱。
where py >nul 2>nul
if %ERRORLEVEL%==0 (
  py -3 "%HERE%md2docx.py" %*
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo 找不到 Python。装法: winget install --id Python.Python.3.12
    exit /b 1
  ) else (
    python "%HERE%md2docx.py" %*
  )
)
exit /b %ERRORLEVEL%
