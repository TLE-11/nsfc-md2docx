#!/usr/bin/env python3
"""转换结果校验。数量或内容对不上就非零退出，防止静默丢东西。

检查项：
  1. docx 包内所有 XML 格式合法
  2. 公式数量：源文件里数出来的块公式/行内公式，与 docx 里的实际数量一致
     （omml 模式比 <m:oMath>/<m:oMathPara>；latex 模式比 MathSource 文本 run）
  3. 中文正文逐字比对（剔除公式与占位提示），检出被吞的段落
  4. 公式编号连续性
  5. 无残留内部标记、无泄漏的 markdown 语法、公式里无可见 &
  6. OOXML 结构：表格单元格必须含段落、相邻表格不得直接拼接
"""
import argparse
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import post  # noqa: E402
from source_model import markdown_model, docx_text


def _console_safe():
    """Windows GBK 控制台打印 GBK 之外的字符会 UnicodeEncodeError，掩盖真实报错。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors='replace')
        except Exception:
            pass


def cjk(t):
    return ''.join(c for c in t if '\u4e00' <= c <= '\u9fff')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src_md')
    ap.add_argument('out_docx')
    ap.add_argument('--math-mode', default='omml')
    ap.add_argument('--number', default='all', choices=['all', 'tag', 'none'])
    ap.add_argument('--prepared-md', help='预处理 Markdown，用于核对已解析图片及引用')
    a = ap.parse_args()
    _console_safe()

    fail = []
    warn = []

    # ---- 源文件侧清点（先规整数学块，再用 Pandoc 文档结构计数）----
    # utf-8-sig 与 pre.py 口径一致，否则带 BOM 时两侧首行标题判定不一致
    raw = open(a.src_md, encoding='utf-8-sig').read()
    source = markdown_model(raw)
    n_block, n_inline, n_tag = source['block'], source['inline'], source['tag']

    z = zipfile.ZipFile(a.out_docx)

    # ---- 1. XML 合法性 ----
    for n in z.namelist():
        if n.endswith(('.xml', '.rels')):
            try:
                ET.fromstring(z.read(n))
            except Exception as e:
                fail.append('XML 非法 %s: %s' % (n, str(e)[:60]))

    # 与 post.py 一致地归一化自闭合标签，否则 pandoc 的 ` />` 会让下面的
    # 字面正则静默漏掉一半匹配
    parts = {n: z.read(n) for n in ('word/document.xml', 'word/footnotes.xml',
                                    'word/endnotes.xml') if n in z.namelist()}
    s = ''.join(post.normalize_xml(value.decode('utf-8')) for value in parts.values())
    z.close()

    # ---- 2. 公式数量 ----
    # 行间公式最终形态：EquationNumbered（带右编号，靠制表位排版）
    #                或 EquationPara（不编号，居中）
    # 两者都含一个内联 m:oMath（omml 模式）或一个 $$ 文本 run（latex 模式）
    paras = re.findall(r'<w:p(?:\s[^>]*)?>.*?</w:p>|<w:p\s*/>', s, re.S)

    def is_disp(p):
        st = re.search(r'<w:pStyle w:val="(EquationNumbered|EquationPara)"/>', p)
        if not st:
            return False
        if a.math_mode == 'omml':
            return '<m:oMath' in p
        return any(t.startswith('$$') for t in re.findall(
            r'<w:rStyle w:val="MathSource"/></w:rPr><w:t[^>]*>(.*?)</w:t>', p, re.S))

    disp_paras = [p for p in paras if is_disp(p)]
    got_block = len(disp_paras)
    if a.math_mode == 'omml':
        got_all = s.count('<m:oMath>')
    else:
        got_all = len(re.findall(r'<w:rStyle w:val="MathSource"/></w:rPr>'
                                 r'<w:t[^>]*>.*?</w:t>', s, re.S))
    got_inline = got_all - got_block
    # 带编号的公式必须保留 m:oMathPara（display 规格容器），否则 Word 把公式
    # 按行内紧凑规格排版：Σ 上下限挪到右下、分数缩成小分式。
    if a.math_mode == 'omml':
        bad_wrap = sum(1 for p in paras
                       if '<w:pStyle w:val="EquationNumbered"/>' in p
                       and '<m:oMathPara' not in p)
        if bad_wrap:
            fail.append('%d 个带编号公式丢失 m:oMathPara，公式被降级为行内规格' % bad_wrap)
    if got_block != n_block:
        fail.append('块公式数量不符: 源 %d -> 产出 %d（很可能有公式被静默吞掉）'
                    % (n_block, got_block))
    if got_inline != n_inline:
        fail.append('行内公式数量不符: 源 %d -> 产出 %d' % (n_inline, got_inline))

    # ---- 3. 中文正文逐字比对 ----
    visible_text = docx_text(parts)
    missing_images = len(re.findall(r'［待插入图片：[^］]*］', visible_text))
    actual_text = re.sub(r'［待插入图片：[^］]*］', '', visible_text)
    x, y = cjk(source['text']), cjk(actual_text)
    if x != y:
        import difflib
        sm = difflib.SequenceMatcher(None, x, y)
        diffs = [(t, x[i1:i2], y[j1:j2])
                 for t, i1, i2, j1, j2 in sm.get_opcodes() if t != 'equal']
        # 后处理主动插入的「图占位」提示文字不算差异
        real = [d for d in diffs if not (d[0] == 'insert' and '图占位' in d[2])]
        if real:
            fail.append('中文正文不一致（相似度 %.4f），前 5 处差异:' % sm.ratio())
            for t, u, v in real[:5]:
                fail.append('    %-7s 源=%r 产出=%r' % (t, u[:40], v[:40]))

    # ---- 4. 公式编号 ----
    nums = []
    for p in paras:
        if '<w:pStyle w:val="EquationNumbered"/>' not in p:
            continue
        m = re.search(r'（([\d\-]+)）', re.sub(r'<[^>]+>', '', p))
        if m:
            nums.append(m.group(1))
    # 公式编号必须靠制表位排版。正文里的内容表格是合法的，只查「疑似公式编号表格」：
    # 单行、含公式、且某个单元格是「（n）」——那种会在 Word/WPS 里显示屏幕虚框。
    eqtbl = 0
    for t in re.findall(r'<w:tbl>.*?</w:tbl>', s, re.S):
        if t.count('<w:tr') == 1 and '<m:oMath' in t \
                and re.search(r'（[\d\-]+）', re.sub(r'<[^>]+>', '', t)):
            eqtbl += 1
    if eqtbl:
        fail.append('有 %d 个疑似公式编号表格，Word/WPS 会显示屏幕虚框，应改用制表位' % eqtbl)
    want = {'all': n_block, 'tag': n_tag, 'none': 0}[a.number]
    if len(nums) != want:
        fail.append('公式编号个数不符: --number %s 期望 %d 个 -> 实际 %d 个'
                    % (a.number, want, len(nums)))
    # 编号书签与 REF 引用必须一一对得上，否则 Word 里会显示「错误!未定义书签」
    bmks = set(re.findall(r'<w:bookmarkStart w:id="\d+" w:name="(eq_\d+)"', s))
    refs = set(re.findall(r'w:instr=" REF (eq_\d+) ', s))
    if len(bmks) != len(nums):
        fail.append('编号书签 %d 个，编号 %d 个' % (len(bmks), len(nums)))
    dangling = refs - bmks
    if dangling:
        fail.append('REF 域指向不存在的书签（Word 会显示"错误!未定义书签"）: %s'
                    % sorted(dangling)[:5])
    # 编号连续性（plain 样式下应为 1..N）。脚注存储在另一 XML 部件中，
    # 拼接后的段落顺序不等于源文档编号顺序；同时检查编号与书签绑定，防止换号漏检。
    if nums and all(x.isdigit() for x in nums):
        ints = [int(x) for x in nums]
        if sorted(ints) != list(range(1, len(ints) + 1)):
            fail.append('公式编号非 1..N 连续，前 20 个: %s' % ints[:20])
        for p in paras:
            if '<w:pStyle w:val="EquationNumbered"/>' not in p:
                continue
            label = re.search(r'（(\d+)）', re.sub(r'<[^>]+>', '', p))
            bookmark = re.search(r'w:name="eq_(\d+)"', p)
            if label and bookmark and int(label.group(1)) != int(bookmark.group(1)):
                fail.append('公式书签与编号不符: eq_%s -> %s' % (bookmark.group(1), label.group(1)))

    # ---- 5. 残留物 ----
    if '⟦' in s:
        fail.append('残留内部标记 %d 处' % s.count('⟦'))
    if any(post.AMP_RUN.search(arr) for arr in re.findall(r'<m:eqArr>.*?</m:eqArr>', s, re.S)):
        fail.append('公式里有未处理的 aligned 对齐符')
    if a.math_mode == 'omml':
        literal_amps = sum((e.text or '').count('&') for value in parts.values()
                           for e in ET.fromstring(value).iter(
                               '{http://schemas.openxmlformats.org/officeDocument/2006/math}t'))
        if literal_amps < source['ampersands']:
            fail.append('公式中的字面量 & 丢失: 源 %d -> 产出 %d'
                        % (source['ampersands'], literal_amps))
    expected_images = source['images'] + max(0, source['wiki_images'] - missing_images)
    if a.prepared_md:
        prepared = open(a.prepared_md, encoding='utf-8').read()
        expected_images = markdown_model(prepared)['images']
        expected_refs = re.findall(r'⟦EQREF:(eq_\d+)\\?\|([^⟧]+)⟧', prepared)
        actual_refs = []
        for value in parts.values():
            for field in ET.fromstring(value).iter(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldSimple'):
                instruction = field.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instr', '')
                match = re.search(r'\bREF (eq_\d+)\b', instruction)
                if match:
                    actual_refs.append((match.group(1), ''.join(field.itertext())))
        if sorted(expected_refs) != sorted(actual_refs):
            fail.append('REF 引用目标、数量或缓存值与预处理结果不符')
    drawings = sum(len(ET.fromstring(value).findall(
        './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'))
        for value in parts.values())
    if drawings != expected_images:
        fail.append('图片数量不符: 源 %d -> 产出 %d' % (expected_images, drawings))
    # 检查 LaTeX 是否泄漏成纯文本前，先剔除行内代码/代码块（pandoc 给它们套
    # VerbatimChar 样式）。代码段里出现 $$ / \begin{ 是合法的，比如文档在讲语法本身。
    s_nocode = re.sub(r'<w:r>(?:(?!</w:r>).)*?<w:rStyle w:val="VerbatimChar"/>'
                      r'(?:(?!</w:r>).)*?</w:r>', '', s, flags=re.S)
    s_nocode = re.sub(r'<w:p(?:\s[^>]*)?>(?:(?!</w:p>).)*?'
                      r'<w:pStyle w:val="SourceCode"/>(?:(?!</w:p>).)*?</w:p>',
                      '', s_nocode, flags=re.S)
    plain_all = re.sub(r'<[^>]+>', '', s_nocode)
    if a.math_mode == 'omml':
        for pat, why in [(r'\\begin\{', '裸露的 \\begin'),
                         (r'\\mathrm\{', '裸露的 \\mathrm'),
                         (r'\$\$', '裸露的 $$')]:
            k = len(re.findall(pat, plain_all))
            if k:
                fail.append('%s 以纯文本出现 %d 次（公式未被编译）' % (why, k))
    if re.search(r'(^|\n)#{2,}\s', plain_all):
        fail.append('裸露的 markdown 标题语法')

    # ---- 6. OOXML 结构 ----
    tcs = re.findall(r'<w:tc>.*?</w:tc>', s, re.S)
    bad_tc = sum(1 for t in tcs if '<w:p' not in t)
    if bad_tc:
        fail.append('%d 个表格单元格不含段落（Word 会报文件损坏）' % bad_tc)
    if '</w:tbl><w:tbl>' in s:
        warn.append('存在直接相邻的表格，Word 会把它们合并成一个')

    # ---- 输出 ----
    print('[verify] 块公式 %d/%d  行内公式 %d/%d  右编号 %d/%d  REF引用 %d  汉字 %d/%d'
          % (got_block, n_block, got_inline, n_inline, len(nums), want,
             len(refs), len(y), len(x)))
    for w in warn:
        print('[verify] 注意: %s' % w)
    if fail:
        for f in fail:
            print('[verify] 失败: %s' % f)
        return 1
    print('[verify] 全部检查通过')
    return 0


if __name__ == '__main__':
    sys.exit(main())
