# nsfc-md2docx

**English** | [简体中文](README_cn.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-informational.svg)](LICENSE)
[![python 3+](https://img.shields.io/badge/python%203%2B-stdlib%20only-success.svg)](#dependencies)
[![pandoc](https://img.shields.io/badge/pandoc-required-orange.svg)](https://pandoc.org)
[![platform](https://img.shields.io/badge/platform-macOS%20%7C%20Linux%20%7C%20Windows-60707f.svg)](#usage)

**Markdown → editable .docx for grant proposals and theses.**

Turns a Markdown file full of LaTeX math into an editable Word document.
Check the layout in your target Word/WPS application before submission:

- Equations become **native Word equation objects** — double-click to edit in Word or WPS.
  Not LaTeX code, not images, no plugins required
- **Every display equation gets an automatic right-aligned number**, laid out with tab
  stops (not borderless tables, so Word/WPS never draws those on-screen gridlines)
- In-text references such as `式（6）` follow the renumbering and become **Word REF
  fields** — press F9 and everything stays in sync
- Chinese typographic conventions configured in one pass: SimSun 12 pt body text,
  2-character first-line indent, 1.5× line spacing, bold-face headings, centered
  10.5 pt captions
- A relay mode for MathType / AxMath for when a template mandates them

Converting "LaTeX → OMML" is already a solved problem. This project is about **the
typographic conventions of Chinese academic documents** (right-aligned equation
numbers, live cross-references, caption styles) and about **the traps that silently
eat content** — the pitfall tables below are arguably worth more than the code.

Dependencies: Python 3 (standard library only) + pandoc. Nothing else.

## Why "NSFC"?

NSFC is the National Natural Science Foundation of China. Its grant proposals — like
Chinese theses — are submitted as Word files under strict conventions: SimSun body
text, right-aligned equation numbers such as （2-1）, cross-references that must
survive renumbering, centered captions. This tool produces documents like that
straight from Markdown.

The pipeline itself is language-agnostic. What is currently Chinese-specific:

| Behavior | Chinese-specific part | With an English document |
|---|---|---|
| Typography presets | SimSun / SimHei / KaiTi (reviewers' machines always have them) | Fonts are just presets — edit `SPEC` in `make_reference.py` |
| In-text reference rewriting | matches `式（n）` / `式(n)` | `Eq. (n)` is **not** rewritten yet |
| Caption detection | matches `图 N：` / `表 N：` | `Figure 1:` is **not** styled yet |
| `--number-style chapter` | chapter number read from digit-leading headings (`# 2 …`) | falls back to plain numbering |
| `verify.py` content check | character-by-character comparison of CJK text | equation-count checks still apply; the text diff does not |

Everything else — equation conversion, numbering, tab-stop layout, field codes,
image handling — works regardless of language. Patches for the gaps above are
welcome and well-scoped.

## Usage

Same on all platforms; all logic lives in `md2docx.py`. Step-by-step for Windows
(including the MathType/AxMath macros): **[WINDOWS.md](WINDOWS.md)** (in Chinese;
the short version: `winget install python pandoc`, then `md2docx.bat input.md`).

```bash
python3 md2docx.py input.md [-o output.docx] [options]   # all platforms
./md2docx.sh input.md ...                                # macOS/Linux wrapper
md2docx.bat input.md ...                                 # Windows wrapper
```

Examples:

```bash
python3 md2docx.py example/sample.md
python3 md2docx.py proposal.md -o proposal.docx --assets ~/Obsidian/vault/attachments
python3 md2docx.py proposal.md --number-style chapter    # numbers become (2-1)(3-1)
python3 md2docx.py proposal.md --math-mode latex         # for MathType/AxMath
python3 md2docx.py proposal.md --fonts macos --pdf       # render a PDF preview locally
```

### Options

| Option | Description |
|---|---|
| `--math-mode omml\|latex` | Equation form. Default `omml` |
| `--number all\|tag\|none` | Which display equations get a number. Default `all` |
| `--number-style plain\|chapter` | `(1)(2)…` or `(2-1)(3-1)…`. Default `plain` |
| `--fonts windows\|macos` | Font preset for the style template. Default `windows` (SimSun/SimHei/KaiTi, for submission); `macos` uses Songti SC etc. for local preview |
| `--assets DIR…` | Directories searched for image attachments (e.g. an Obsidian vault) |
| `--reference FILE` | Use your own reference.docx, skip auto-generation |
| `--pdf` | Additionally render a PDF via LibreOffice for visual inspection |
| `--keep-temp` | Keep intermediate files for troubleshooting |

## Try it in one minute

The repo ships a sample that deliberately triggers every known trap:

```bash
python3 md2docx.py example/sample.md -o /tmp/sample.docx
```

## Regression tests

```bash
./example/run_tests.sh                 # 24 sample/encoding combinations + scenario regressions
./example/run_tests.sh my-doc.md       # also runs your own document
```

`run_tests.sh` checks process exit codes, then runs conversion regressions and
`example/test_vba_contract.py` static VBA checks.
The conversion regressions independently check equation characters, REF targets, images, tables,
footnotes, and deliberately damaged outputs. Run it alone with
`python3 example/test_regressions.py`; inputs are in `example/fixtures/`. Each section
of `example/sample.md` ends with an HTML comment naming the trap it exists to
trigger — keep those structures when modifying it.

### The two math modes

| Mode | Equations become | When to use |
|---|---|---|
| `omml` (default) | Native Word equation objects | Word/WPS edit them natively, **no plugins needed** — covers almost every case |
| `latex` | LaTeX source as blue monospace text | Intermediate product for MathType / AxMath batch conversion, see below |

## Dependencies

Python 3 + pandoc. **No third-party Python packages** — standard library only.

```bash
# macOS
brew install pandoc
# Windows
winget install --id Python.Python.3.12
winget install --id JohnMacFarlane.Pandoc
# Debian/Ubuntu
sudo apt install python3 pandoc
```

The optional `--pdf` preview needs LibreOffice. Developed against pandoc 3.11.

## Pipeline

```
md ──pre.py──> normalized md ──pandoc──> step1.docx ──post.py──> final docx
```

### Problems pre.py solves

| Problem | What actually happens without it |
|---|---|
| `\rm` and other legacy font commands | texmath **does not support** `\rm`/`\bf`/`\it` and silently degrades the **whole equation into literal LaTeX text** — the #1 cause of "my equations turned into code". Rewritten to `\mathrm{}` etc. |
| `\tag{n}` | texmath drops it silently; the number vanishes completely. Extracted here, applied by post.py |
| **Blank lines inside a `$$` block** | A blank line ends the paragraph in Markdown, so pandoc never sees a closing `$$` — **the entire equation is silently swallowed** with zero warnings. Blank lines are meaningless in LaTeX math, so they are squeezed out |
| `\boxed{}` | Preserved as native OMML `m:borderBox`, including grouping and scripts |
| `---` horizontal rules | pandoc's `multiline_tables` treats them as table separators and **swallows the following text into table cells** (even `##` stops parsing). The reader disables simple/multiline tables while preserving Setext headings |
| Obsidian `![[image]]` embeds | Non-standard syntax that pandoc renders as plain text. Rewritten to `![](resolved-path)` with a recursive search through `--assets` directories; a loud placeholder is left if the file can't be found |
| Inverted heading levels | `##` nested under `###` breaks the Word outline and auto-numbering. Normalized by the shape of the heading text (`x.y` / `x.y.z` / `（n）`) |
| `【图占位N】` markers | Turned into loud placeholder paragraphs |

### Problems post.py solves

| Problem | What it does |
|---|---|
| `&` in `aligned` | texmath writes alignment markers as ordinary characters — a visible `&` appears inside the equation. Rewritten into proper OMML alignment points (`m:rPr/m:aln`) |
| Numbered-equation layout | `TAB equation TAB （n）` in a single paragraph with tab stops — centered equation, right-aligned number, no table involved |
| Figure/table captions | `图 N：…` / `表 N：…` paragraphs get the Caption style (centered, 10.5 pt, no indent) and become cross-referenceable |
| Display-equation paragraphs | Uniformly given the `EquationPara` style (centered, no first-line indent) |

### Right-aligned numbering and cross-references

Every display equation gets a number by default, laid out with **tab stops**: one
paragraph containing `TAB equation TAB （n）`.

An earlier version used a 1×3 borderless table per equation; it was abandoned
because Word/WPS draw **screen gridlines** around borderless tables — they never
print, but they are ugly, and "view gridlines" is an application-level toggle that
cannot be stored in the file, so it had to be turned off manually on every machine.
The tab-stop scheme has no table, hence no gridlines. `verify.py` checks that the
document contains none.

Numbers are `SEQ` fields; in-text references are `REF` fields pointing at numbered
bookmarks, so pressing F9 in Word renumbers everything consistently.

**Renumbering would scramble existing references**, so pre.py builds an
`old \tag → new number` map and rewrites the `式（n）` references in the body text.
References that match no `\tag` are left untouched and **reported as warnings** —
no guessing, because the source document may genuinely contain dangling references.

### verify.py — post-conversion checks (run automatically; non-zero exit on failure)

This step is non-negotiable. Conversion failures are frequently **silent** — no
error, the content is simply gone. One real case caught in development: a `$$`
block contained a blank line, pandoc never saw the closing delimiter, and a
multi-line optimization equation vanished without a trace — the output just had
two orphaned `$$`, and not a single warning was emitted.

Checks:

- every XML part inside the docx package is well-formed
- **equation counts**: block / inline equation counts from the source must match
  what is actually in the docx
- **character-by-character comparison of visible CJK text** (parsed from the
  Pandoc document tree, with footnotes expanded in reading order; code, math and
  link destinations excluded) — catches swallowed paragraphs
- image counts and REF targets, counts and cached values against prepared input
- equation-number continuity; `\tag` count matches the number of numbered equations
- no leftover internal markers, no leaked Markdown syntax, no unprocessed alignment markers; literal `\&` characters are preserved
- OOXML structure: table cells must contain a paragraph; adjacent tables must not
  be glued together

## MathType / AxMath support

**These cannot be generated directly.** Both store equations as OLE-embedded
objects — MathType uses the proprietary MTEF binary format, AxMath its own format
(the [AxMath docs](https://axmath.gitbooks.io/axmath-docs-en/6._equation_output_and_word_plugin.html)
confirm OLE). There is no public library for writing either, an EMF preview image
is required alongside, and the result wouldn't be editable anyway on machines
without the software installed.

**The right approach is to let them do the converting.** Two routes:

### Route A: via OMML (using the default output)

In Word: MathType tab → Convert Equations → input "Word 2007 and later (OMML)
equations", output "MathType equations", scope "Whole document". This is officially
documented ([Typefi notes](https://help.typefi.com/hc/en-us/articles/360001608675-Add-and-edit-equations-with-MathType-Writer)).

**This route is not lossless.** WIRIS maintains dedicated troubleshooting pages for
[OMML conversion errors](https://wiris.helpjuice.com/en_US/conversion-and-compatibility/error-message-problem-converting-omml-to-mathml)
and [missing symbols](https://docs.wiris.com/en_US/conversion-and-compatibility/symbols-missing-in-equations-converted-from-words-equation-editor-to-mathtype)
— visually identical symbols can have different encodings in OMML, and MathType may
misinterpret them. Spot-check after converting.

### Route B: via LaTeX source (`--math-mode latex`)

Skips texmath entirely: equations land in the docx as LaTeX source (a blue
monospace `MathSource` character style, so unconverted ones are visible at a
glance), and MathType / AxMath parse them themselves.

**Advantage**: preserves source for commands unsupported by texmath. `\boxed`
is preserved in both modes; OMML uses a native mathematical box.

The companion VBA macros are in `mathtype_axmath.bas`; see [WINDOWS.md](WINDOWS.md).
`ProbeEquationMacros` reports installed add-ins and configured macro names without
executing them. Conversion requires an explicitly configured, locally verified
macro name and confirmation. Without one, use the plugin GUI. Counting, highlighting
and conversion include footnotes and endnotes; failed sources remain available for retry.

### Platform reality

- **AxMath is Windows-only** — its GitHub releases ship only .exe files. It cannot
  be installed on macOS.
- MathType has a Mac version, but Convert Equations is an **MS Word add-in**
  feature and doesn't exist in WPS.
- So the final step of both routes has to happen on a **Windows + MS Word** machine.

### Think before you go

OMML is Word's native format. Word and WPS edit it directly with no plugins, and
reviewers won't hit missing plugins or fonts when they open the file. The
legitimate reasons for MathType / AxMath usually boil down to three: the template
mandates it, you want their equation-numbering + cross-reference system, or your
collaborators are used to it. If you just want "editable equations", the default
omml mode already delivers that.

### make_reference.py — the style template

Generates `reference.docx`. Adapt the `SPEC` dict to your institution's template
without hand-tuning styles in Word. Current settings:

| Element | Spec |
|---|---|
| Body text | SimSun / Times New Roman, 12 pt, 1.5× line spacing, 2-character first-line indent |
| Heading 1 | SimHei, 16 pt, centered |
| Heading 2 | SimHei, 14 pt, left |
| Headings 3/4 | SimHei, 12 pt, left |
| Block quotes | KaiTi, 12 pt |
| Captions | SimSun, 10.5 pt, centered |
| Page | A4, 2.54 cm top/bottom, 3.17 cm left/right |
| Equation font | Cambria Math (change with `--math-font`) |

Regenerate after changing styles:

```bash
python3 make_reference.py reference.docx
```

## Verified results

`example/sample.md` and its BOM/CRLF variant pass verify.py under **24 conversions**: `omml`/`latex` × `plain`/`chapter` numbering × `all`/`tag`/`none`
numbering scope.

- zero pandoc warnings during conversion
- block equations 11/11, inline equations 18/18
- right-aligned numbers 11/11, consecutive, none missing
- resolvable references remapped to REF fields; the sample also retains an intentional dangling-reference warning
- CJK body text compared character by character: 972/972, zero content loss
- all XML well-formed; no equation-numbering tables (hence no screen gridlines)

A historical version was also validated on a real 1700+ line proposal containing
350+ equations (not public and not rerun in this round).

During development verify.py caught **four** silent content losses — no error, the
content was simply gone:

1. a blank line inside a `$$` block → the whole equation vanished, leaving two
   orphaned `$$`
2. after numbering all equations, paragraphs mixing "body text + equation" were
   dropped entirely → both the text and its inline equations disappeared
3. in latex mode, equations in such mixed paragraphs didn't get numbers
4. a backtick-quoted `` `$$` `` in the body text (the document was explaining the
   syntax itself) was treated as a math delimiter, mispairing every equation
   after it

The first two slipped through precisely because only the Chinese text was compared
and equation counts weren't checked. **That is why verify.py exists.**

## Automatic wrapping of long equations

OMML display equations use the section/page margins, table cell width and equation
font size to estimate available space, reserving room for the number. Long sums
and relations wrap before top-level operators into an editable `m:eqArr` inside
one native equation; numbering and cross-references remain intact. Short and
inline equations are not reflowed. Fractions, radicals, scripts and grouped
expressions stay intact. Oversized existing `aligned` rows can also wrap.
LaTeX relay mode keeps its source representation.

Width estimation is conservative, not a complete font-layout engine. An oversized
indivisible expression fails with an actionable error and preserves any existing
output file, rather than truncating content or reducing the font size. Equations
whose estimated height exceeds one page are rejected too; split them into separate
equations or adjust the page height. Rewrite
such expressions with explicit `aligned` breaks or adjust the page layout.
Check the target Word/WPS application when fonts are substituted.

Optional PDF layout regression requires LibreOffice and `pdfplumber` (neither is
an additional requirement for ordinary conversion):

```bash
python3 example/test_layout.py --output-dir /tmp/nsfc-layout
```

It renders ten scenarios, checks horizontal and vertical PDF bounds, equation
terms, pagination, text following boxed formulas, links and numbering, and verifies rejection of an
oversized 200-term equation. DOCX/PDF outputs are retained for visual inspection.

DOCX output is replaced only after validation succeeds. PDF previews are exported
in isolation; failed exports preserve an existing PDF and report that it was not updated.

## Known limitations

- Programmatic checks and LibreOffice layout tests do not replace visual review
  in the target Word/WPS application.
- Unsupported inline OMML math (for example `\ce{H2O}`) now fails conversion,
  rather than warning and reporting success. Rewrite it using supported syntax
  or use `--math-mode latex` for a later conversion.

- Native `\boxed` output is covered by regression tests with Pandoc 3.11.
  Unsupported formulas fail validation instead of silently losing structure.
  The LibreOffice renderer used for these checks omits native `m:borderBox`
  borders in PDF previews. Box structure and script grouping are verified in DOCX
  XML; their appearance still needs checking in the target Word/WPS application.
- When renumbering, references that match no original `\tag` are **left as-is and
  reported** — no guessing. The source may genuinely contain dangling references
  (e.g. a whole section missing from the md); those need manual attention.
- `--number-style chapter` uses static number text; F9 does not renumber it in
  Word (`plain` uses SEQ fields and does). Chapter-style numbers like "2-1" would
  require stacking STYLEREF, which is worse for compatibility and readability.
- Images that can't be found become a loud red/yellow placeholder for manual
  backfill; `--assets` directories enable automatic resolution.

## Not yet programmatically verified

- **How OMML actually renders in WPS / Word.** The dev machine is macOS, where
  LibreOffice finds no CJK-capable font for rasterization (Chinese comes out
  blank; the PDF text layer is intact — purely a rasterization issue). Everything
  structural is verified; the look of fonts, sizes and spacing needs a human to
  open the file.
- **Word VBA integration has not been executed in this macOS environment.**
  Portable checks in `example/test_vba_contract.py` check the no-execution probe
  and shared story traversal contracts, but cannot validate Word's object model.
  On Windows with Word and VBA project access already enabled, run
  `powershell -File example/test_word_macros.ps1`. It uses a temporary document and
  controlled callbacks to check confirmation, footnotes/endnotes, no-op and failed
  conversions, and cleanup. It does not change Office security settings or validate
  a real MathType/AxMath version.

## Other traps we hit

- **Word matches styles by *name*, not styleId.** Both `Find.Style` and `Styles()`
  take the name. Name the `MathSource` style `Math Source (LaTeX)` and none of the
  VBA will find anything.
- **pandoc emits self-closing tags with a trailing space** (`<w:pStyle w:val="X" />`);
  hand-built XML usually doesn't. Mix both in one document and any regex written
  against the literal `w:val="X"/>` silently matches only half of them.
  `post.normalize_xml()` normalizes at the entry point, fixing this class of bug
  at the root.
- **`grep -c` returning 0 is not success.** Once, pandoc failed on a YAML error
  and produced no file at all — but the grep count was 0, which looked like a
  pass. **Always assert the artifact exists.**
- In bash, a variable followed immediately by a full-width character needs
  `${VAR}`. `"$MODE）"` treats the UTF-8 bytes of the full-width paren as part of
  the variable name, which blows up with `set -u`.

## Contributing

- **Bug reports**: attach a minimal .md snippet that reproduces the issue, the
  exact command, and the full console output. Running with `--keep-temp` and
  including the intermediate files makes locating the problem much faster.
- **Windows users**: the report produced by `ProbeEquationMacros` (pasted into an
  issue) is exactly what's needed to pin down the MathType/AxMath macro names for
  everyone else.
- **Pull requests**: run `./example/run_tests.sh` first; all 12 combinations must
  pass.
- **English-language support** (`Eq. (n)` reference rewriting, `Figure 1:`
  caption detection) is a welcome, well-scoped contribution — see the table in
  [Why "NSFC"?](#why-nsfc).

## License

[MIT](LICENSE) — with one caveat: the auto-generated `reference-*.docx` style
templates derive from pandoc's own `data/reference.docx`, and pandoc is GPL (v2+).
Those files are gitignored, not distributed with this repository, and are generated
locally by *your* pandoc on first run. If you bundle this project into a
closed-source product, do not ship those generated templates.

This project is not affiliated with, endorsed by, or sponsored by the National
Natural Science Foundation of China, Microsoft, Kingsoft, Wiris (MathType), or
AxMath. All product names are used solely to describe compatibility.

Built on top of [pandoc](https://pandoc.org) — thanks to John MacFarlane and the
texmath contributors.
