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
import pre  # noqa: E402


def cjk(t):
    return ''.join(c for c in t if '\u4e00' <= c <= '\u9fff')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src_md')
    ap.add_argument('out_docx')
    ap.add_argument('--math-mode', default='omml')
    ap.add_argument('--number', default='all', choices=['all', 'tag', 'none'])
    a = ap.parse_args()

    fail = []
    warn = []

    # ---- 源文件侧清点（复用 pre.py 的同一套切分逻辑，同样跳过代码区域）----
    raw = open(a.src_md, encoding='utf-8').read()
    noncode = ''.join(c for is_code, c in pre.split_code(raw) if not is_code)
    n_block = len(pre.MATH_BLOCK.findall(noncode))
    n_inline = len(pre.MATH_INLINE.findall(pre.MATH_BLOCK.sub('', noncode)))
    n_tag = len(pre.TAG_RE.findall(noncode))

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
    s = post.normalize_xml(z.read('word/document.xml').decode('utf-8'))

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
    # 带编号的公式必须已从 m:oMathPara 拆成内联 m:oMath，否则制表位排不进同一段落。
    # 不编号的公式保留 m:oMathPara 是正确的（块级居中）。
    bad_wrap = sum(1 for p in paras
                   if '<w:pStyle w:val="EquationNumbered"/>' in p and '<m:oMathPara' in p)
    if bad_wrap:
        fail.append('%d 个带编号公式仍是块级 m:oMathPara，编号排不到同一行' % bad_wrap)
    if got_block != n_block:
        fail.append('块公式数量不符: 源 %d -> 产出 %d（很可能有公式被静默吞掉）'
                    % (n_block, got_block))
    if got_inline != n_inline:
        warn.append('行内公式数量不符: 源 %d -> 产出 %d' % (n_inline, got_inline))

    # ---- 3. 中文正文逐字比对 ----
    # HTML 注释在 docx 输出里会被 pandoc 丢掉，不能算进比对
    md_body = re.sub(r'<!--.*?-->', '', noncode, flags=re.S)
    md_body = pre.MATH_BLOCK.sub('', md_body)
    md_body = pre.MATH_INLINE.sub('', md_body)
    md_body = re.sub(r'!\[\[[^\]]*\]\]', '', md_body)
    docx_text = re.sub(r'<m:oMath[^>]*>.*?</m:oMath>', '', s, flags=re.S)
    # latex 模式下公式是文本 run，必须一并剔除，否则 \text{中文} 里的汉字会被误算
    docx_text = re.sub(r'<w:r><w:rPr><w:rStyle w:val="MathSource"/></w:rPr>'
                       r'<w:t[^>]*>.*?</w:t></w:r>', '', docx_text, flags=re.S)
    docx_text = re.sub(r'<[^>]+>', '', docx_text)
    # 剔除后处理主动插入的占位文字
    docx_text = re.sub(r'［待插入图片：[^］]*］', '', docx_text)
    x, y = cjk(md_body), cjk(docx_text)
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
        warn.append('编号书签 %d 个，编号 %d 个' % (len(bmks), len(nums)))
    dangling = refs - bmks
    if dangling:
        fail.append('REF 域指向不存在的书签（Word 会显示"错误!未定义书签"）: %s'
                    % sorted(dangling)[:5])
    # 编号连续性（plain 样式下应为 1..N）
    if nums and all(x.isdigit() for x in nums):
        ints = [int(x) for x in nums]
        if ints != list(range(1, len(ints) + 1)):
            warn.append('公式编号非 1..N 连续，前 20 个: %s' % ints[:20])

    # ---- 5. 残留物 ----
    if '⟦' in s:
        fail.append('残留内部标记 %d 处' % s.count('⟦'))
    if re.findall(r'<m:t[^>]*>[^<]*&amp;[^<]*</m:t>', s):
        fail.append('公式里有可见的 & （aligned 对齐符未处理）')
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
