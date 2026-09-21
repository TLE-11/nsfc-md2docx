#!/usr/bin/env python3
"""Markdown 预处理器：把 Obsidian / LaTeX 风格的 md 规整成 pandoc 能干净吃下的形式。

处理项：
  1. \rm / \bf / \it / \sf / \tt  ->  \mathrm / \mathbf / ...   (texmath 不支持老式字体命令，
     遇到就整块公式退化成 LaTeX 原文，这是"公式变成代码"的首要原因)
  2. \tag{n}  ->  从公式里摘掉，改成段落末尾的 ⟦EQNO:n⟧ 标记，交给 post.py 做右对齐编号
  3. \boxed{...} -> 摘掉方框，记录为需要加边框的公式
  4. Obsidian ![[图片]] -> 标准 ![](真实路径)，并在附件目录里搜索文件
  5. 标题层级归一化：修正 ### 下面挂 ## 这类倒挂
  6. 【图占位N：说明】 -> 独立段落 + 图题占位

用法: python3 pre.py 输入.md 输出.md [--assets 附件搜索目录 ...]
"""
import argparse
import json
import os
import re
import sys

# ---------------------------------------------------------------- 数学区域切分

MATH_BLOCK = re.compile(r'(?<!\\)\$\$(.+?)(?<!\\)\$\$', re.S)
MATH_INLINE = re.compile(r'(?<![\$\\])\$(?!\$)((?:[^\$\n]|\\\$)+?)(?<!\\)\$(?!\$)')

# 代码区域：围栏代码块 + 行内代码。必须先摘出来，否则文档里只要出现讨论 $ 或 $$
# 的代码段（技术文档里很常见），就会被当成公式定界符，导致后面所有公式配对错位。
CODE_RE = re.compile(
    r'(?P<fence>^(?P<f>```+|~~~+)[^\n]*\n.*?^(?P=f)[ \t]*$)'
    r'|(?P<inline>(?P<t>`+)(?!`).+?(?<!`)(?P=t)(?!`))',
    re.S | re.M)


def split_code(text):
    """把文本切成 [(是否代码, 片段), ...]，代码片段不参与公式识别。"""
    parts, pos = [], 0
    for m in CODE_RE.finditer(text):
        if m.start() > pos:
            parts.append((False, text[pos:m.start()]))
        parts.append((True, m.group(0)))
        pos = m.end()
    if pos < len(text):
        parts.append((False, text[pos:]))
    return parts


def map_noncode(text, fn):
    """对非代码片段套用函数，代码片段原样保留。

    图片 wikilink、【图占位】、"式（n）"回改这类针对正文的替换都必须走这里：
    代码块/行内代码里的同样字样是字面文字，改了就是静默篡改，而且 verify 的
    正文比对两侧都剔代码区，查不出来。
    """
    return ''.join(chunk if is_code else fn(chunk)
                   for is_code, chunk in split_code(text))


# 围栏代码块开关：与 CODE_RE 口径一致（行首 ``` 或 ~~~，闭合围栏不短于开启）。
FENCE_RE = re.compile(r'^(`{3,}|~{3,})')


def iter_noncode_lines(text):
    """逐行产出 (是否代码行, 行文本)。围栏行本身算代码行。

    行级扫描（标题归一化、分隔线消歧、编号分配）都用它跳过代码块：
    代码里的 `# 5 注释`、`---`、公式标记都只是字面文字。
    """
    fence = None
    for ln in text.split('\n'):
        fm = FENCE_RE.match(ln)
        if fm:
            mark = fm.group(1)
            if fence is None:
                fence = mark
            elif mark[0] == fence[0] and len(mark) >= len(fence):
                fence = None
            yield True, ln
        else:
            yield fence is not None, ln


def map_math(text, fn_block, fn_inline):
    """分别对块公式与行内公式套用函数，不触碰正文，也不触碰代码区域。"""
    def on_plain(seg):
        out, pos = [], 0
        for m in MATH_BLOCK.finditer(seg):
            out.append(MATH_INLINE.sub(lambda i: fn_inline(i.group(1)), seg[pos:m.start()]))
            out.append(fn_block(m.group(1)))
            pos = m.end()
        out.append(MATH_INLINE.sub(lambda i: fn_inline(i.group(1)), seg[pos:]))
        return ''.join(out)

    return ''.join(chunk if is_code else on_plain(chunk)
                   for is_code, chunk in split_code(text))


# ------------------------------------------------------- 1. 老式字体命令转换

FONT_CMDS = {'rm': 'mathrm', 'bf': 'mathbf', 'it': 'mathit',
             'sf': 'mathsf', 'tt': 'mathtt', 'cal': 'mathcal'}
FONT_RE = re.compile(r'\\(' + '|'.join(FONT_CMDS) + r')(?![A-Za-z])')


def _match_brace(s, i):
    """s[i] == '{'，返回配对 '}' 的下标。"""
    depth = 0
    while i < len(s):
        if s[i] == '\\':
            i += 2
            continue
        if s[i] == '{':
            depth += 1
        elif s[i] == '}':
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def fix_font_cmds(m, stats):
    """{\rm abc} -> {\mathrm{abc}}；\rm 作用到所在分组末尾，需要花括号配对。"""
    while True:
        hit = FONT_RE.search(m)
        if not hit:
            return m
        cmd = FONT_CMDS[hit.group(1)]
        # 向左跳过空白，看是否紧跟一个 '{'
        j = hit.start() - 1
        while j >= 0 and m[j] in ' \t\n':
            j -= 1
        if j >= 0 and m[j] == '{':
            close = _match_brace(m, j)
            if close == -1:
                # 括号不配对，保守处理：只换命令名，避免死循环
                m = m[:hit.start()] + '\\' + cmd + '{}' + m[hit.end():]
                stats['font_unbalanced'] += 1
                continue
            body = m[hit.end():close].strip()
            m = m[:j] + '{\\' + cmd + '{' + body + '}}' + m[close + 1:]
        else:
            # 裸 \rm：作用到当前片段末尾，取到下一个空白/结束
            tail = m[hit.end():]
            tok = re.match(r'\s*([A-Za-z0-9]+)', tail)
            if tok:
                m = m[:hit.start()] + '\\' + cmd + '{' + tok.group(1) + '}' + tail[tok.end():]
            else:
                m = m[:hit.start()] + m[hit.end():]
        stats['font_fixed'] += 1


# --------------------------------------------------------------- 2. \tag 提取

TAG_RE = re.compile(r'\\tag\{([^}]*)\}')
BOXED_RE = re.compile(r'\\boxed\s*\{')


def strip_boxed(m, stats):
    """\boxed{X} -> X，并记账（OMML 无对应结构，改由 post.py 给段落加边框）。"""
    boxed = False
    while True:
        hit = BOXED_RE.search(m)
        if not hit:
            return m, boxed
        open_brace = hit.end() - 1
        close = _match_brace(m, open_brace)
        if close == -1:
            return m, boxed
        m = m[:hit.start()] + m[open_brace + 1:close] + m[close + 1:]
        boxed = True
        stats['boxed'] += 1


# ------------------------------------------------------------------ 主处理流程

def process(text, assets_dirs, stats, math_mode='omml', math_sink=None,
            number='all', number_style='plain'):
    """math_mode:
         omml  —— 公式交给 pandoc 编译成 Word 原生公式对象（默认）
         latex —— 公式保留 LaTeX 源码，作为带样式的文本落到 docx 里，
                  再由 MathType / AxMath 的批量转换功能就地转成它们自己的公式对象。
                  好处是绕开 texmath 的短板（\boxed、复杂 aligned 等），
                  由更强的解析器处理；代价是必须装对应软件才能变成公式。
    """
    latex_mode = (math_mode == 'latex')

    # --- 2/3: 块公式里摘 \tag 与 \boxed ---------------------------------
    def on_block(body):
        tags = TAG_RE.findall(body)
        body = TAG_RE.sub('', body)
        # 块公式内部的空行会在 markdown 层断段，导致 pandoc 拿不到闭合的 $$，
        # 整个公式被静默吞掉（只剩两个孤立的 $$）。空行在 LaTeX 数学里无意义，直接压掉。
        # 先去掉首尾空白（摘掉 \tag 常留下尾部空行，那种无害），再统计真正危险的内部空行。
        body = body.strip()
        body, n = re.subn(r'\n[ \t]*\n+', '\n', body)
        if n:
            stats['blank_in_math'] += n
        # \boxed 在 latex 模式下原样保留，MathType/AxMath 认得
        if latex_mode:
            boxed = False
        else:
            body, boxed = strip_boxed(body, stats)
        body = fix_font_cmds(body, stats)
        body = body.strip()

        marks = ''
        if boxed:
            marks += '⟦BOX⟧'
        # 每个块公式都打一个占位标记，原有 \tag 号记在里面（没有就留空）。
        # 最终编号在后面的编号扫描里统一分配，这样才能给所有行间公式编号，
        # 同时保留 旧tag号 -> 新编号 的映射用于回改正文里的"式（n）"。
        if tags:
            stats['tags'] += len(tags)
        marks += '⟦EQ:' + (tags[-1] if tags else '') + '⟧'

        stats['math_blocks'] += 1
        if latex_mode:
            # 用不含 markdown 特殊字符的令牌占位，避免 $ _ ^ * 被 pandoc 解析
            idx = len(math_sink)
            math_sink.append({'kind': 'display', 'tex': body})
            stats['math_latex'] += 1
            return '⟦MATHD:%d⟧' % idx + marks
        # 标记紧跟在 $$ 之后不空行 -> pandoc 会与公式合成同一段落
        return '$$\n' + body + '\n$$' + marks

    def on_inline(body):
        if not latex_mode:
            body, _ = strip_boxed(body, stats)
        body = fix_font_cmds(body, stats)
        stats['math_inline'] += 1
        if latex_mode:
            idx = len(math_sink)
            math_sink.append({'kind': 'inline', 'tex': body.strip()})
            stats['math_latex'] += 1
            return '⟦MATHI:%d⟧' % idx
        return '$' + body + '$'

    text = map_math(text, on_block, on_inline)

    # --- 4: Obsidian wikilink 图片 --------------------------------------
    def resolve(name):
        for d in assets_dirs:
            for root, _dirs, files in os.walk(d):
                if name in files:
                    return os.path.join(root, name)
        return None

    def on_wikiimg(m):
        name = m.group(1).split('|')[0].strip()
        path = resolve(name)
        if path:
            stats['img_found'] += 1
            # 尖括号包裹路径：含空格（中文文件名/目录很常见）或 Windows 反斜杠时
            # 裸写会被 markdown 从空格处截断或把 \ 当转义符，<> 内一律按字面解析
            return '![](<%s>)' % path
        stats['img_missing'].append(name)
        return '⟦MISSINGIMG:%s⟧' % name

    text = map_noncode(text, lambda s: re.sub(r'!\[\[([^\]]+)\]\]', on_wikiimg, s))

    # --- 6: 图占位 ------------------------------------------------------
    def on_ph(m):
        stats['placeholders'] += 1
        return '⟦FIGPH⟧' + m.group(1)

    text = map_noncode(text, lambda s: re.sub(r'【(图占位[^】]*)】', on_ph, s))

    # --- 5/7: 标题层级归一化 + 分隔线消歧 --------------------------------
    fixed = []
    for in_code, ln in iter_noncode_lines(text):
        if in_code:
            fixed.append(ln)
            continue
        # 独立的 --- / ___ 分隔线：pandoc 的 multiline_tables 会把它当表格头
        # 分隔符，从而把后面整段正文吞成表格里的纯文本。统一改写为 ***。
        if re.match(r'^\s*(-{3,}|_{3,})\s*$', ln):
            stats['hr_normalized'] += 1
            fixed.append('***')
            continue
        h = re.match(r'^(#{1,6})\s+(.*)$', ln)
        if not h:
            fixed.append(ln)
            continue
        lvl, title = len(h.group(1)), h.group(2)
        # 按标题文字形态推断应有层级：# 章 / ## x.y / ### x.y.z / #### （n）
        if re.match(r'^[（(]\s*\d+\s*[)）]', title):
            want = 4
        elif re.match(r'^\d+\.\d+\.\d+', title):
            want = 3
        elif re.match(r'^\d+\.\d+', title) or re.match(r'^\d+\.\d+\s', title):
            want = 2
        elif re.match(r'^\d+\s*[、.]', title):
            want = 1
        else:
            want = lvl
        if want != lvl:
            stats['heading_fixed'].append('%s -> %s  %s' % ('#' * lvl, '#' * want, title[:40]))
            lvl = want
        fixed.append('#' * lvl + ' ' + title)
    text = '\n'.join(fixed)

    # --- 8: 公式编号分配 + 正文交叉引用回改 ------------------------------
    text = assign_numbers(text, stats, number, number_style)
    return text


# ------------------------------------------------- 8. 编号分配与交叉引用

EQ_MARK = re.compile(r'⟦EQ:([^⟧]*)⟧')
H1_NUM = re.compile(r'^#\s+(\d+)')
# 正文里的公式引用：式（6）/ 式(6) / 式 (6)，也支持"由式（6）—（8）"这类区间的端点
REF_RE = re.compile(r'式\s*[（(]\s*(\d+)\s*[)）]')


def assign_numbers(text, stats, number='all', style='plain'):
    """给块公式分配最终编号，并把正文里的"式（旧号）"改写成指向新编号的引用标记。

    number: all=所有行间公式都编号 / tag=只给原来有 \tag 的编号 / none=不编号
    style : plain=全文连续 (1)(2)... / chapter=按一级标题分章 (2-1)(3-1)...
    """
    if number == 'none':
        stats['eq_numbered'] = 0
        return EQ_MARK.sub('', text)

    tag2num = {}          # 旧 \tag 号 -> 新编号文本
    seq_global = 0
    seq_chapter = 0
    chapter = None
    out_lines = []

    for in_code, ln in iter_noncode_lines(text):
        if in_code:
            out_lines.append(ln)
            continue
        h = H1_NUM.match(ln)
        if h:
            chapter = h.group(1)
            seq_chapter = 0

        def on_mark(m):
            nonlocal seq_global, seq_chapter
            origtag = m.group(1)
            if number == 'tag' and not origtag:
                return ''                      # 该公式不编号
            seq_global += 1
            seq_chapter += 1
            if style == 'chapter' and chapter:
                num = '%s-%d' % (chapter, seq_chapter)
            else:
                num = str(seq_global)
            bmk = 'eq_%d' % seq_global
            if origtag:
                if origtag in tag2num:
                    stats['dup_tags'].append(origtag)
                tag2num[origtag] = (num, bmk)
            return '⟦EQNO:%s|%s|%d⟧' % (num, bmk, seq_global)

        out_lines.append(EQ_MARK.sub(on_mark, ln))

    text = '\n'.join(out_lines)
    stats['eq_numbered'] = seq_global

    # 正文交叉引用回改。只改能对上原 \tag 的，改不上的原样留下并报告，
    # 避免把作者本来就悬空的引用（比如某节公式在源文件里整节缺失）改错。
    def on_ref(m):
        old = m.group(1)
        if old in tag2num:
            num, bmk = tag2num[old]
            stats['ref_remapped'] += 1
            return '式⟦EQREF:%s|%s⟧' % (bmk, num)
        stats['ref_unresolved'].append(m.group(0))
        return m.group(0)

    # 代码区里的"式（n）"是字面文字，不能改成 REF 域
    text = map_noncode(text, lambda s: REF_RE.sub(on_ref, s))
    return text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src')
    ap.add_argument('dst')
    ap.add_argument('--assets', nargs='*', default=[])
    ap.add_argument('--report')
    ap.add_argument('--math-mode', choices=['omml', 'latex'], default='omml',
                    help='omml=Word 原生公式(默认); latex=保留 LaTeX 源码待 MathType/AxMath 批量转换')
    ap.add_argument('--math-json', help='latex 模式下公式源码的落盘位置')
    ap.add_argument('--number', choices=['all', 'tag', 'none'], default='all',
                    help='all=所有行间公式右编号(默认); tag=只给原有 \\tag 的编号; none=不编号')
    ap.add_argument('--number-style', choices=['plain', 'chapter'], default='plain',
                    help='plain=全文连续(1)(2)...; chapter=按章 (2-1)(3-1)...')
    a = ap.parse_args()

    stats = {'font_fixed': 0, 'font_unbalanced': 0, 'tags': 0, 'boxed': 0,
             'img_found': 0, 'img_missing': [], 'placeholders': 0,
             'heading_fixed': [], 'hr_normalized': 0, 'math_latex': 0,
             'blank_in_math': 0, 'math_blocks': 0, 'math_inline': 0,
             'eq_numbered': 0, 'ref_remapped': 0, 'ref_unresolved': [],
             'dup_tags': []}
    math_sink = []
    text = open(a.src, encoding='utf-8').read()
    out = process(text, a.assets, stats, a.math_mode, math_sink,
                  a.number, a.number_style)
    open(a.dst, 'w', encoding='utf-8').write(out)

    if a.math_mode == 'latex':
        path = a.math_json or (a.dst + '.math.json')
        json.dump(math_sink, open(path, 'w', encoding='utf-8'), ensure_ascii=False)
        print('[pre] 公式模式 latex：保留 %d 个公式源码 -> %s'
              % (stats['math_latex'], os.path.basename(path)))

    print('[pre] 字体命令修正 %d 处（未配对 %d）' % (stats['font_fixed'], stats['font_unbalanced']))
    print('[pre] 行间公式编号 %d 个（源文件原有 \\tag %d 个，其余为新增）'
          % (stats['eq_numbered'], stats['tags']))
    print('[pre] 正文引用回改 %d 处' % stats['ref_remapped'])
    if stats['ref_unresolved']:
        print('[pre] 注意：%d 处引用对不上任何原有 \\tag，已原样保留（源文件本身可能悬空）：%s'
              % (len(stats['ref_unresolved']), '、'.join(stats['ref_unresolved'])))
    if stats['dup_tags']:
        print('[pre] 注意：源文件里重复的 \\tag 号 %s' % '、'.join(stats['dup_tags']))
    print('[pre] \\boxed 摘除 %d 处' % stats['boxed'])
    print('[pre] 图片解析成功 %d 张，缺失 %d 张 %s'
          % (stats['img_found'], len(stats['img_missing']), stats['img_missing'] or ''))
    print('[pre] 图占位 %d 处' % stats['placeholders'])
    print('[pre] 分隔线消歧 %d 处（--- -> ***，防被当表格吞正文）' % stats['hr_normalized'])
    print('[pre] 公式内空行压缩 %d 处（否则该公式会被 pandoc 整块吞掉）' % stats['blank_in_math'])
    print('[pre] 公式清点：块公式 %d 个，行内公式 %d 个，合计 %d'
          % (stats['math_blocks'], stats['math_inline'],
             stats['math_blocks'] + stats['math_inline']))
    if stats['heading_fixed']:
        print('[pre] 标题层级修正 %d 处：' % len(stats['heading_fixed']))
        for s in stats['heading_fixed']:
            print('       ' + s)
    if a.report:
        json.dump(stats, open(a.report, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)


if __name__ == '__main__':
    sys.exit(main())
