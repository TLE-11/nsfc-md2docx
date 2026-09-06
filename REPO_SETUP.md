# 建仓速查（建完可删掉本文件）

## 建仓

```bash
cd nsfc-md2docx
git init
git add .
git commit -m "初始提交：Markdown 转 Word，公式原生化 + 行间公式自动右编号"
git branch -M main
gh repo create TLE-11/nsfc-md2docx --public --source=. --push
```

没装 `gh` 就先在 GitHub 网页上建空仓库，再：

```bash
git remote add origin https://github.com/TLE-11/nsfc-md2docx.git
git push -u origin main
```

提交前确认 `.gitignore` 生效，`reference-*.docx` 不应出现在 `git status` 里
（它派生自 pandoc 自带模板，而 pandoc 是 GPL，不随仓库分发）：

```bash
git status --short
```

## Description（填到仓库设置的 About）

```
写中文基金本子/学位论文用的 Markdown 转 Word：Word 原生公式、行间公式自动右编号、交叉引用 F9 联动、中文排版规范一次配好
```

## Topics

```
markdown  docx  pandoc  word  latex  omml
equation-numbering  cross-reference
chinese  academic-writing  thesis  nsfc  mathtype
```

前两行是功能词，后一行是人群词。`equation-numbering`、`omml`、`nsfc`、`mathtype`
比 `markdown`、`docx` 这类大词更容易被对的人搜到。

## 同类项目

写 README 或发帖时可以主动对比，说清自己的边界：

| 项目 | 定位 |
|---|---|
| pandoc | 底层转换引擎，本项目就是包在它外面 |
| `md2star`、`markdocx`、`mdstyledocx` | 同样是「pandoc + 样式层」，但不做中文本子的公式编号与交叉引用 |
| MathType / AxMath | 公式编辑器，本项目提供转换到它们的中转模式 |

本项目的差异点只有一句：**中文本子的排版规范 + 公式右编号 + 交叉引用联动 + 静默丢内容的坑**。

## 发布后

- `example/run_tests.sh` 是回归入口，可直接接到 GitHub Actions 上
- Windows 用户跑 `ProbeEquationMacros` 的结果值得收集，用来固化 MathType/AxMath 的宏名
- 中文用户主要从知乎/公众号/搜索引擎过来，GitHub 站内搜索不是主要入口
