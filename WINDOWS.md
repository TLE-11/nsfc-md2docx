# 在 Windows 上使用

整套工具就是几个 Python 脚本 + pandoc，把整个仓库目录拷到 Windows 机器上就能跑。

## 一次性装环境

```powershell
winget install --id Python.Python.3.12
winget install --id JohnMacFarlane.Pandoc
```

装完**重开一个终端**（PATH 要刷新）。验证：

```powershell
python --version
pandoc --version
```

装不上 winget 的话：Python 去 python.org 下载（安装时务必勾 **Add python.exe to PATH**），
pandoc 去 GitHub releases 下 `.msi`。

不需要 bash、不需要 WSL、不需要装任何 Python 第三方包（只用标准库）。

## 用法

```powershell
cd nsfc-md2docx
md2docx.bat 本子.md
md2docx.bat 本子.md -o 本子.docx
md2docx.bat 本子.md --number-style chapter
md2docx.bat 本子.md --math-mode latex
md2docx.bat 本子.md --assets D:\Obsidian\vault\attachments
```

`md2docx.bat` 只是个壳，等价于 `python md2docx.py ...`。全部选项看 `python md2docx.py --help`。

Windows 上 `--fonts` 保持默认的 `windows` 就对了（宋体/黑体/楷体本来就在）。

## 转成 MathType / AxMath 公式

**只有需要它们的时候才做这一步。** 默认 `--math-mode omml` 产出的是 Word 原生公式，
Word 和 WPS 都能直接双击编辑、不需要任何插件，绝大多数场合够用。

要转的话有两条路。

### 路线 A：从 OMML 转（用默认产物）

不需要任何脚本。Word 里打开默认模式生成的 docx：

MathType 选项卡 → **Convert Equations** → 输入勾 `Word 2007 and later (OMML) equations`，
输出选 `MathType equations`，范围选 `Whole document`。

**这条路不是无损的。** 转完必须抽查符号——OMML 里视觉相同的符号可能编码不同，
MathType 会解释错，官方专门有故障文档记录这个问题。

### 路线 B：从 LaTeX 源码转（保真度更高）

```powershell
md2docx.bat 本子.md -o 待转换.docx --math-mode latex
```

产物里公式是**蓝色等宽的 LaTeX 源码**（`MathSource` 字符样式），一眼能看出哪些还没转。
好处是绕开 pandoc 的 texmath，`\boxed`、复杂 `aligned` 这些原样交给更强的解析器。

然后在 Word 里：

1. 按 `Alt+F11` 打开 VBA 编辑器
2. 文件 → 导入文件 → 选仓库根目录下的 `mathtype_axmath.bas`
3. 回到 Word 按 `Alt+F8`，可先运行 **`ProbeEquationMacros`** 查看加载项和宏名配置

此报告不会执行任何插件宏，也不验证配置名称是否可用。

- 没有本机已确认的宏名：运行 `SelectLatexEquations` 标记源码，使用插件 GUI 批量转换。
  MathType 可使用 Convert Equations 的 TeX 输入；AxMath 使用插件菜单的批量处理功能。
- 已从插件文档或受控试运行确认宏名：填入模块顶部的
  `MATHTYPE_MACRO_OVERRIDE` / `AXMATH_MACRO_OVERRIDE`，再运行对应转换宏。
  转换宏会在执行前提示备份并请求确认。

计数、高亮和转换均覆盖正文、脚注、尾注及其他 story。每处源码只尝试一次；
插件报错或未消耗源码时记为失败，保留源码、样式和定界符供重试。
`CleanupAfterConvert` 只清理已转换内容的残留样式，不删除待转换源码的 `$`。
最后运行 `UpdateAllFields` 刷新公式编号与引用。

### 为什么不自动试执行宏

`Application.Run` 会实际运行宏，不能作为只读的“接口探测”。候选宏可能修改当前选区、
打开对话框或启动批量转换，因此工具不再通过试执行猜测插件接口。
未配置已确认宏名时，转换宏会明确提示改用 GUI。

## 宏清单

| 宏 | 作用 | 是否依赖插件 |
|---|---|---|
| `UpdateAllFields` | 更新全文域：公式 SEQ 编号 + 正文 REF 引用。会跑两遍（两者互相依赖） | 不依赖插件 |
| `CountLatexEquations` | 数还剩多少处未转换的 LaTeX 源码 | 不依赖 |
| `SelectLatexEquations` | 把全部 LaTeX 源码标黄，便于确认范围 / GUI 批量转换 | 不依赖 |
| `ProbeEquationMacros` | 报告加载项和已配置宏名，不执行宏 | — |
| `ConvertLatexToMathType` | 逐个转成 MathType 公式 | 依赖 MathType 且配置已确认宏名 |
| `ConvertLatexToAxMath` | 逐个转成 AxMath 公式 | 依赖 AxMath 且配置已确认宏名 |
| `CleanupAfterConvert` | 清理已转换内容的样式；保留待转换源码 | 不依赖 |

转换宏会先弹确认框并提醒备份。**动手前请另存一份。**

## AxMath 的平台限制

AxMath 只有 Windows 版，所以这台 Windows 机器正好是它唯一能用的地方。
macOS 上装不了，那边只能用默认的 OMML 模式。

## 公式编号与交叉引用

- 编号用 `SEQ` 域，在 Word 里增删公式后按 `Ctrl+A` 再 `F9` 可整篇重编号
- 正文引用用 `REF` 域指向编号书签，跟着编号一起变
- 直接用宏 `UpdateAllFields` 更省事，它会正确地跑两遍

## VBA 回归验证

普通 `example/run_tests.sh` 包含静态契约检查，不会启动 Word。
Windows + Word 下还可运行：

```powershell
powershell -File example/test_word_macros.ps1
```

前提是用户已允许访问 VBA 工程对象模型；脚本不修改该设置。它新建临时文档，
使用受控宏替身验证确认取消、正文/脚注/尾注转换、失败重试和清理，完成后关闭临时文档。
该测试不验证真实 MathType/AxMath 插件的接口与符号保真度；转换后仍需抽查。
