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
3. 回到 Word 按 `Alt+F8`，先运行 **`ProbeEquationMacros`**

`ProbeEquationMacros` 会列出本机装了哪些 Word 加载项，以及 MathType / AxMath
到底暴露了哪些能从 VBA 调用的宏。**必须先跑这个**，原因见下面「为什么要探测」。

- 探测**到**可用宏 → 直接运行 `ConvertLatexToMathType` 或 `ConvertLatexToAxMath`，
  它会倒序逐个转换并报告成功/失败数
- 探测**不到** → 运行 `SelectLatexEquations` 把待转换内容标成黄色高亮，
  然后用插件自己的 GUI 批量功能：
  - MathType：Convert Equations → 输入勾 `Text using a translator` → 选 TeX 翻译器 → 范围 Whole document
  - AxMath：插件菜单的批量处理功能

转完运行 `CleanupAfterConvert` 去高亮、清残留样式和 `$` 定界符，
再运行 `UpdateAllFields` 刷新公式编号与交叉引用。

### 为什么要探测，而不是直接给你一条命令

MathType 和 AxMath **都没有公开、稳定的 VBA 接口文档**。能查到的唯一线索是一条
2009 年的资料提到 MathType 6.5 的 `MTCommand_TeXToggle`，而现在是 MathType 7，
这个名字还在不在没法确认；AxMath 官方文档只写了插件按钮，没写 VBA 接口。

我不在这台 macOS 上验证不了的东西上硬编码猜测值，所以做成探测 + 报告 + 优雅降级：
探测得到就自动用，探测不到就明确告诉你走 GUI（那条路是官方文档写明的、可靠的）。

跑完 `ProbeEquationMacros` 如果发现了可用的宏名，填到 `mathtype_axmath.bas` 顶部的
`MATHTYPE_MACRO_OVERRIDE` / `AXMATH_MACRO_OVERRIDE` 常量里，以后就不用再探测了。

## 宏清单

| 宏 | 作用 | 是否依赖插件 |
|---|---|---|
| `UpdateAllFields` | 更新全文域：公式 SEQ 编号 + 正文 REF 引用。会跑两遍（两者互相依赖） | 不依赖，永远可用 |
| `CountLatexEquations` | 数还剩多少处未转换的 LaTeX 源码 | 不依赖 |
| `SelectLatexEquations` | 把全部 LaTeX 源码标黄，便于确认范围 / GUI 批量转换 | 不依赖 |
| `ProbeEquationMacros` | 探测本机可调用的插件宏，只报告不改文档 | — |
| `ConvertLatexToMathType` | 逐个转成 MathType 公式 | 依赖 MathType 且探测到接口 |
| `ConvertLatexToAxMath` | 逐个转成 AxMath 公式 | 依赖 AxMath 且探测到接口 |
| `CleanupAfterConvert` | 去高亮、清残留 `MathSource` 样式与 `$` 定界符 | 不依赖 |

转换宏会先弹确认框并提醒备份。**动手前请另存一份。**

## AxMath 的平台限制

AxMath 只有 Windows 版，所以这台 Windows 机器正好是它唯一能用的地方。
macOS 上装不了，那边只能用默认的 OMML 模式。

## 公式编号与交叉引用

- 编号用 `SEQ` 域，在 Word 里增删公式后按 `Ctrl+A` 再 `F9` 可整篇重编号
- 正文引用用 `REF` 域指向编号书签，跟着编号一起变
- 直接用宏 `UpdateAllFields` 更省事，它会正确地跑两遍
