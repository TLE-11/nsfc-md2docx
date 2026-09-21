# nsfc-md2docx

[English](README.en.md) | **简体中文**

[![License: MIT](https://img.shields.io/badge/License-MIT-informational.svg)](LICENSE)
[![python 3+](https://img.shields.io/badge/python%203%2B-stdlib%20only-success.svg)](#依赖)
[![pandoc](https://img.shields.io/badge/pandoc-required-orange.svg)](https://pandoc.org)
[![platform](https://img.shields.io/badge/platform-macOS%20%7C%20Linux%20%7C%20Windows-60707f.svg)](#用法)

**写基金本子 / 学位论文用的 Markdown 转 Word。**

把带大量 LaTeX 公式的 Markdown 转成**能直接交稿**的 .docx：

- 公式是 **Word 原生公式对象**，双击可编辑。不是 LaTeX 代码，不是图片，不需要装任何插件
- **所有行间公式自动右编号**，用制表位排版（不是无边框表格，所以没有屏幕虚框）
- 正文里的「式（n）」自动跟着重编号，并做成 **Word 的 REF 域**，按 F9 联动更新
- 中文排版规范一次配好：宋体小四正文、首行缩进 2 字符、1.5 倍行距、黑体标题、五号居中题注
- 需要 MathType / AxMath 的场合有专门的中转模式

单纯的「LaTeX 转 OMML」已经有不少工具了。这个项目的重点是**中文本子的排版规范**
（公式右编号、交叉引用联动、题注样式）和**一堆会静默丢内容的坑**——后者见下面的踩坑表，
那部分比代码值钱。

依赖只有 Python 3 + pandoc。

## 用法

跨平台，逻辑都在 `md2docx.py` 里。Windows 详见 **[WINDOWS.md](WINDOWS.md)**。

```bash
python3 md2docx.py 输入.md [-o 输出.docx] [选项]     # 三个平台通用
./md2docx.sh 输入.md ...                            # macOS/Linux 薄壳
md2docx.bat 输入.md ...                             # Windows 薄壳
```

例：

```bash
python3 md2docx.py example/sample.md
python3 md2docx.py 本子.md -o 本子.docx --assets ~/Obsidian/vault/attachments
python3 md2docx.py 本子.md --number-style chapter    # 编号改成 (2-1)(3-1)
python3 md2docx.py 本子.md --math-mode latex         # 给 MathType/AxMath 用
python3 md2docx.py 本子.md --fonts macos --pdf       # 本机渲染 PDF 预览
```

### 选项

| 选项 | 说明 |
|---|---|
| `--math-mode omml\|latex` | 公式形态。默认 omml |
| `--number all\|tag\|none` | 行间公式右编号范围。默认 all（所有行间公式都编号） |
| `--number-style plain\|chapter` | `(1)(2)...` 还是 `(2-1)(3-1)...`。默认 plain |
| `--fonts windows\|macos` | 样式模板字体档位。默认 windows（出稿用） |
| `--assets 目录...` | 图片附件搜索目录（Obsidian 附件目录） |
| `--reference 文件` | 用自己的 reference.docx，跳过自动生成 |
| `--pdf` | 额外用 LibreOffice 渲染 PDF 便于肉眼检查 |
| `--keep-temp` | 保留中间文件便于排查 |

依赖只有 Python 3 标准库 + pandoc，不需要装任何 Python 第三方包。

## 试一下

仓库自带一份样例，刻意覆盖了下面所有已知陷阱：

```bash
python3 md2docx.py example/sample.md -o /tmp/sample.docx
```

## 回归测试

```bash
./example/run_tests.sh                  # 样例跑遍 12 种参数组合
./example/run_tests.sh 我的本子.md        # 顺带测自己的文档
```

靠 `verify.py` 断言，任何一处回归都会非零退出。`example/sample.md` 每个小节末尾用
HTML 注释标了它专门触发哪个陷阱，改动时请保留那些结构。

### 两种公式模式

| 模式 | 公式形态 | 适用场景 |
|---|---|---|
| `omml`（默认） | Word 原生公式对象 | WPS / Word 直接可编辑，**不需要装任何插件**。绝大多数情况用这个 |
| `latex` | 蓝色等宽的 LaTeX 源码文本 | 单位模板要求必须用 MathType / AxMath 公式时的中间产物，见下节 |

## 依赖

Python 3 + pandoc。**不需要任何 Python 第三方包**（只用标准库）。

```bash
# macOS
brew install pandoc
# Windows
winget install --id Python.Python.3.12
winget install --id JohnMacFarlane.Pandoc
# Debian/Ubuntu
sudo apt install python3 pandoc
```

`--pdf` 预览需要 LibreOffice，可选。开发时用的是 pandoc 3.11。

## 流水线

```
md ──pre.py──> 规整后的 md ──pandoc──> step1.docx ──post.py──> 成品 docx
```

### pre.py 解决的问题

| 问题 | 说明 |
|---|---|
| `\rm` 等老式字体命令 | pandoc 的 texmath **不支持** `\rm`/`\bf`/`\it`，遇到就把**整块公式退化成 LaTeX 原文**。这是"公式变成代码"的首要原因。自动转成 `\mathrm{}` 等 |
| `\tag{n}` | texmath 静默吞掉，编号完全丢失。这里摘出来交给 post.py |
| **公式块内部的空行** | markdown 里空行断段，pandoc 拿不到闭合的 `$$`，**整块公式被静默吞掉**且不报警告。空行在 LaTeX 数学里无意义，直接压掉 |
| `\boxed{}` | OMML 无对应结构，摘掉方框、改由段落边框实现 |
| `---` 分隔线 | pandoc 的 `multiline_tables` 会把它当表格头分隔符，**把后面整段正文吞成表格里的纯文本**（`##` 都不解析）。统一改写为 `***` |
| Obsidian `![[图片]]` | 非标准语法，pandoc 当纯文本。改写为 `![](真实路径)`，并在附件目录里递归搜文件；找不到就留醒目占位 |
| 标题层级倒挂 | 原文有 `###` 下面挂 `##` 的情况，转 Word 后大纲和自动编号会乱。按标题文字形态（`x.y` / `x.y.z` / `（n）`）归一化 |
| `【图占位N】` | 转成醒目占位段落 |

### post.py 解决的问题

| 问题 | 说明 |
|---|---|
| `aligned` 的 `&` | texmath 把对齐符当普通字符写进 OMML，公式里出现可见的 `&`。改成 OMML 正确的对齐点 `m:rPr/m:aln` |
| 公式编号排版 | 编号公式合成一个段落：`TAB 公式 TAB （n）`，制表位排版（公式居中、编号右对齐），无表格 |
| 图表题注 | `图 N：xxx` 段落套 Caption 样式（居中、五号、无缩进），可用于交叉引用 |
| 块公式段落 | 统一套 `EquationPara` 样式：居中、无首行缩进 |

### 公式右编号与交叉引用

所有行间公式默认都给右编号，用**制表位**排版：一个段落里 `TAB 公式 TAB （n）`。

早先用的是 1×3 无边框表格，已弃用——Word/WPS 会给无边框表格画屏幕**虚框**，
虽然不打印但碍眼，而且「查看虚框」是应用级开关、存不进文件，每台机器都要手动关。
制表位方案没有表格，也就没有虚框。`verify.py` 会检查文档里不该出现表格。

编号用 `SEQ` 域，正文引用用 `REF` 域指向编号书签，Word 里按 F9 联动更新。

**重编号会打乱原有引用**，所以 pre.py 建了 `旧 \tag 号 → 新编号` 的映射，
自动回改正文里的「式（n）」。对不上任何 `\tag` 的引用会原样保留并**报警**，
不硬猜——源文件本身可能有悬空引用。

### verify.py — 转换校验（流水线自动调用，不通过就非零退出）

必须有这一步。转换失败很多时候是**静默**的，不报错但内容没了。实测就抓到一次：
原文某个 `$$` 块内部有一个空行，markdown 里空行断段，pandoc 拿不到闭合的 `$$`，
整块公式（几十行的优化问题）凭空消失，输出里只剩两个孤立的 `$$`，而且不产生任何警告。

校验项：

- docx 包内所有 XML 格式合法
- **公式数量**：源文件数出来的块公式 / 行内公式数，与 docx 里实际数量必须一致
- **中文正文逐字比对**（剔除公式与占位文字），检出被吞的段落
- 公式编号连续性、`\tag` 个数与编号表格个数一致
- 无残留内部标记、无泄漏的 markdown 语法、公式里无可见 `&`
- OOXML 结构：表格单元格必须含段落、相邻表格不得直接拼接

## 支持 MathType / AxMath

**不能直接生成。** 两者的公式在 docx 里都是 OLE 嵌入对象——MathType 是私有的 MTEF
二进制，AxMath 是自己的格式（[AxMath 官方文档](https://axmath.gitbooks.io/axmath-docs-en/6._equation_output_and_word_plugin.html)
明确写走 OLE）。都没有公开的写入库，还得配 EMF 预览图。而且生成出来对方没装软件也编辑不了。

**正确路径是让它们自己转。** 两条路：

### 路线 A：OMML 中转（用默认的 omml 模式产物）

Word 里 MathType 选项卡 → Convert Equations → 输入选 OMML equations、输出选
MathType equations、范围选 Whole document。官方文档写明支持整篇或选区
（[Typefi 说明](https://help.typefi.com/hc/en-us/articles/360001608675-Add-and-edit-equations-with-MathType-Writer)）。

**这条路不是无损的。** WIRIS 官方有专门的故障文档：
[OMML 转换报错](https://wiris.helpjuice.com/en_US/conversion-and-compatibility/error-message-problem-converting-omml-to-mathml)、
[符号丢失](https://docs.wiris.com/en_US/conversion-and-compatibility/symbols-missing-in-equations-converted-from-words-equation-editor-to-mathtype)
——OMML 里视觉相同的符号编码可能不同，MathType 会解释错。转完必须抽查。

### 路线 B：LaTeX 源码中转（`--math-mode latex`）

跳过 texmath，公式以 LaTeX 源码形式落进 docx（蓝色等宽的 `MathSource` 字符样式，
肉眼一眼能认出哪些还没转），再让 MathType / AxMath 自己解析。

**优势**：绕开 texmath 的短板。`\boxed`、复杂 `aligned` 这些在 omml 模式下会降级处理的
结构，这里原样交给更强的解析器。`--math-mode latex` 下 `\boxed` 也不再被摘除。

配套 VBA 宏在 `mathtype_axmath.bas`，操作步骤见 **[WINDOWS.md](WINDOWS.md)**。
关键一点：**先运行 `ProbeEquationMacros` 探测**，别直接跑转换宏。
MathType / AxMath 都没有公开稳定的 VBA 接口文档，宏名靠探测确定，探测不到就走 GUI。

### 平台现实

- **AxMath 只有 Windows 版**，GitHub 上只发 .exe。macOS 装不了。
- MathType 有 Mac 版，但 Convert Equations 是 **MS Word 加载项**功能，WPS 用不了。
- 所以这两条路的最后一步都得在 **Windows + MS Word** 的机器上做。

### 先想清楚要不要走

OMML 是 Word 原生格式，WPS 和 Word 都能直接双击编辑，不需要装插件，评审专家打开也不会
缺插件缺字体。MathType / AxMath 的正当理由通常只有三个：单位模板硬性规定、要用它们的
公式编号+交叉引用体系、协作者习惯。只是想要"能编辑的公式"的话，默认 omml 模式已经够了。

### make_reference.py — 排版样式模板

生成 `reference.docx`。改 `SPEC` 字典即可适配不同单位的模板要求，
不需要手工在 Word 里调样式。当前设置：

| 元素 | 规格 |
|---|---|
| 正文 | 宋体 / Times New Roman，小四(12pt)，1.5 倍行距，首行缩进 2 字符 |
| 一级标题 | 黑体，三号(16pt)，居中 |
| 二级标题 | 黑体，四号(14pt)，左对齐 |
| 三/四级标题 | 黑体，小四(12pt)，左对齐 |
| 引用块 | 楷体，小四 |
| 图表题注 | 宋体，五号(10.5pt)，居中 |
| 页面 | A4，上下 2.54cm，左右 3.17cm |
| 公式字体 | Cambria Math（`--math-font` 可改） |

改完样式后重新生成：

```bash
python3 make_reference.py reference.docx
```

## 已验证结果

`example/sample.md`（205 行）在 **12 种参数组合**下全部通过 verify.py：
`omml`/`latex` × `plain`/`chapter` 编号样式 × `all`/`tag`/`none` 编号范围。

- 转换过程零 pandoc 警告
- 块公式 10/10，行内公式 18/18
- 右编号 10/10，编号连续无缺
- 正文交叉引用全部回改并做成 REF 域
- 中文正文逐字比对 825/825，零内容丢失
- 所有 XML 通过格式合法性校验；无公式编号表格（因而无屏幕虚框）

另外在一份 1700 余行、含 350 多个公式的真实申报书上验证通过（该文档未公开）。

开发过程中 verify.py 抓到过四次**静默**内容丢失，都是不报错但东西没了：

1. `$$` 块内部有空行 → 整块公式消失，只剩两个孤立的 `$$`
2. 给所有公式编号后，「正文 + 公式同段」的段落被整段丢弃 → 正文和行内公式一起没了
3. latex 模式下同类混排段落的公式拿不到编号
4. 文档正文里用反引号写了 `` `$$` ``（在讲语法本身），被当成公式定界符，
   导致后面所有公式配对错位

前两次都是因为只比对了中文正文、没校验公式数量才漏掉的。**这就是 verify.py 存在的理由。**

## 已知限制

- omml 模式下 `\boxed` 用段落边框近似，边框是整段宽度，不是紧贴公式。
  latex 模式下 `\boxed` 原样保留，交给 MathType / AxMath 处理。
- 重编号时，正文里对不上任何原有 `\tag` 的引用会**原样保留并报警**，不做猜测。
  源文档本身可能有悬空引用（比如某节公式在 md 里整节缺失），这种必须人工处理。
- `--number-style chapter` 用静态编号文本，在 Word 里按 F9 不会自动重编号
  （`plain` 样式用 SEQ 域，可以）。因为「2-1」这种带章号的格式要叠 STYLEREF，
  兼容性和可读性都更差。
- 找不到的图片会输出红黄底的醒目占位，需手工回填；给了 `--assets` 目录可自动解析。

## 尚未程序化验证的部分

- **OMML 在 WPS / Word 里的实际渲染。** 开发机是 macOS，LibreOffice 在上面找不到
  能绘制 CJK 的字体，栅格化时中文全空（PDF 文字层里汉字都在，是纯栅格化问题）。
  结构层面全部校验通过，但字体、字号、行距的观感需要人工打开确认。
- **`mathtype_axmath.bas` 里的 VBA 一行都没跑过**（macOS 跑不了 Word 宏）。
  不依赖插件的四个宏（`UpdateAllFields` / `SelectLatexEquations` /
  `CountLatexEquations` / `CleanupAfterConvert`）用的都是标准 Word 对象模型；
  依赖插件的两个转换宏靠运行时探测确定接口，探测失败会明确降级到 GUI 流程。
  欢迎在 Windows 上跑 `ProbeEquationMacros` 并把结果反馈到 issue。

## 其他踩过的坑

- **Word 按样式「名称」匹配，不是 styleId。** `Find.Style` 和 `Styles()` 都用 name。
  把 `MathSource` 的 name 写成 `Math Source (LaTeX)` 会让整套 VBA 一个都找不到。
- **pandoc 输出的自闭合标签带空格**（`<w:pStyle w:val="X" />`），自己拼的通常不带。
  两种混在一个文档里，任何按字面写 `w:val="X"/>` 的正则都会静默漏掉一半。
  `post.normalize_xml()` 在入口统一归一化，根治这一类 bug。
- **`grep -c` 返回 0 不代表成功。** 有一次 pandoc 因 YAML 解析失败根本没产出文件，
  但 grep 计数是 0，看着像通过。**必须验证产物存在。**
- bash 里变量后紧跟全角字符要写 `${VAR}`。`"$MODE）"` 会把全角括号的 UTF-8 字节
  当成变量名的一部分，配合 `set -u` 直接报 unbound variable。

## 反馈与贡献

- **转换出问题**：附上能复现的最小 .md 片段、完整命令与控制台输出；加 `--keep-temp`
  把中间文件一并贴出来，定位会快很多。
- **Windows 用户**：`mathtype_axmath.bas` 里 `ProbeEquationMacros` 的探测报告对固化
  MathType / AxMath 宏名非常有价值，直接开 issue 贴出来即可。
- **提 PR 前**先跑 `./example/run_tests.sh`，12 种参数组合全过再提。
- 英文文档支持（`Eq. (n)` 引用回改、`Figure 1:` 题注识别）是范围明确的欢迎贡献，
  缺口清单见 [README.en.md](README.en.md) 的 Why "NSFC" 一节。

## 开源声明

- 代码以 **[MIT](LICENSE)** 许可。仓库不含任何 GPL 数据：自动生成的 `reference-*.docx`
  样式模板派生自 pandoc 自带模板（pandoc 为 GPL v2+），已在 `.gitignore` 排除、
  不随仓库分发，首次运行时由你本机的 pandoc 现场生成。把本项目打包成闭源产品时
  **不要附带这些生成物**。
- 本项目与国家自然科学基金委员会、Microsoft（Word）、金山（WPS）、Wiris（MathType）、
  AxMath 均无关联，相关名称仅用于描述兼容性。
- 本流水线建立在 [pandoc](https://pandoc.org) 之上，感谢 John MacFarlane
  与 texmath 的贡献者。
