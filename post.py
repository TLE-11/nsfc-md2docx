#!/usr/bin/env python3
"""docx 后处理器：直接操作 word/document.xml，完成 pandoc 做不到的排版。

处理项：
  1. ⟦EQNO:n⟧  -> 把上一段块公式与编号合成一个段落：
                   TAB 公式 TAB (n)，制表位排版（国标公式编号写法，无表格）
  2. ⟦BOX⟧     -> 给该公式段落加外框
  3. ⟦MISSINGIMG:名⟧ -> 醒目红色占位提示，方便回填
  4. ⟦FIGPH⟧   -> 图占位段落套 Caption 样式
  5. **图 N：xxx** / **表 N：xxx** 独立段落 -> 套 Caption 样式并居中去缩进
  6. 块公式段落统一套 EquationPara 样式（居中、无首行缩进）

用法: python3 post.py 输入.docx 输出.docx
"""
import argparse
import json
import os
import re
import shutil
import sys
import zipfile


def _console_safe():
    """Windows GBK 控制台打印 ⟦⟧ 等私用符号（或 GBK 之外的生僻字）会
    UnicodeEncodeError，把真正的报错/警告变成二次崩溃。降级为替换字符。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors='replace')
        except Exception:
            pass

# ------------------------------------------------------------------ XML 片段

BOXBORDER = ('<w:pBdr>' + ''.join(
    '<w:%s w:val="single" w:sz="8" w:space="4" w:color="auto"/>' % s
    for s in ('top', 'left', 'bottom', 'right')) + '</w:pBdr>')

class BmkIds(object):
    """书签 id 必须是文档内唯一的整数，从已有最大值之上开始发号。"""

    def __init__(self, xml):
        used = [int(v) for v in re.findall(r'<w:bookmark(?:Start|End) w:id="(\d+)"', xml)]
        self.next = (max(used) + 1) if used else 1

    def take(self):
        v = self.next
        self.next += 1
        return v


OMATHPARA_UNWRAP = re.compile(
    r'<m:oMathPara>\s*(?:<m:oMathParaPr>.*?</m:oMathParaPr>)?\s*(.*?)\s*</m:oMathPara>',
    re.S)


def number_xml(num, bmk, ids, seq_field):
    """编号「（n）」，数字包在书签里供正文 REF 域引用。

    seq_field=True 时数字用 SEQ 域生成（Word 里增删公式后按 F9 可整篇重编号），
    同时写入缓存值，没按 F9 也能正常显示。chapter 样式用静态文本，因为「2-1」
    这种带章号的格式要叠 STYLEREF，兼容性和可读性都更差。
    """
    bid = ids.take()
    if seq_field:
        inner = ('<w:fldSimple w:instr=" SEQ eqn \\* ARABIC ">'
                 '<w:r><w:t>%s</w:t></w:r></w:fldSimple>' % num)
    else:
        inner = '<w:r><w:t>%s</w:t></w:r>' % num
    return ('<w:r><w:t>（</w:t></w:r>'
            '<w:bookmarkStart w:id="%d" w:name="%s"/>%s<w:bookmarkEnd w:id="%d"/>'
            '<w:r><w:t>）</w:t></w:r>' % (bid, bmk, inner, bid))


def numbered_eq_p(math_para_xml, num, bmk, ids, seq_field, boxed):
    """把「块公式段落 + 编号」合成单个段落：公式 （n）。

    omml 模式：保留完整 m:oMathPara，编号 run 跟在它后面同段。
    这正是 Word 原生行为（在显示公式后面 Tab 输入编号，存出来的结构就是
    oMathPara + 文本 run 同段）。早期版本把公式从 oMathPara 里拆成内联
    m:oMath，编号是能同段了，但 Word 对内联公式一律套用行内紧凑规格：
    Σ 的上下限被挪到右下、分数缩成小分式——带编号公式整体降级。
    oMathPara 自带居中（defJc=centerGroup），不需要前置 TAB。

    latex 模式：本体是 MathSource 文本 run，没有显示规格一说，左起即可
    （反正随后要在 Word 里交给 MathType/AxMath 转换，居中性由转换结果决定）。
    """
    m = OMATHPARA.search(math_para_xml)
    if m:
        # omml 模式：完整保留显示公式容器
        return ('<w:p><w:pPr><w:pStyle w:val="EquationNumbered"/>%s</w:pPr>'
                '%s<w:r><w:tab/></w:r>%s</w:p>'
                % (BOXBORDER if boxed else '', m.group(0),
                   number_xml(num, bmk, ids, seq_field)))
    # latex 模式：取段落内容（MathSource run）
    body = re.sub(r'^<w:p(?:\s[^>]*)?>|</w:p>$', '', math_para_xml)
    body = re.sub(r'<w:pPr>.*?</w:pPr>|<w:pPr/>', '', body, flags=re.S)
    body = re.sub(r'<w:bookmark(?:Start|End)[^>]*/>', '', body)
    return ('<w:p><w:pPr><w:pStyle w:val="EquationNumbered"/>%s</w:pPr>'
            '%s<w:r><w:tab/></w:r>%s</w:p>'
            % (BOXBORDER if boxed else '', body,
               number_xml(num, bmk, ids, seq_field)))



EQREF_TOKEN = re.compile(r'⟦EQREF:([^|⟧]*)\|([^⟧]*)⟧')


def fill_eqrefs(xml, stats):
    """正文里的 ⟦EQREF:书签|编号⟧ -> Word 的 REF 域，公式编号变动后按 F9 自动同步。"""
    def on_run(run):
        text_m = re.search(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', run, re.S)
        if not text_m or '⟦EQREF:' not in text_m.group(1):
            return run
        text = text_m.group(1)
        rpr_m = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S)
        rpr = rpr_m.group(0) if rpr_m else ''
        out, pos = [], 0
        for t in EQREF_TOKEN.finditer(text):
            if t.start() > pos:
                out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                           % (rpr, text[pos:t.start()]))
            bmk, num = t.group(1), t.group(2)
            out.append('<w:r>%s<w:t>（</w:t></w:r>'
                       '<w:fldSimple w:instr=" REF %s \\h ">'
                       '<w:r>%s<w:t>%s</w:t></w:r></w:fldSimple>'
                       '<w:r>%s<w:t>）</w:t></w:r>' % (rpr, bmk, rpr, num, rpr))
            stats['ref_fields'] += 1
            pos = t.end()
        if pos < len(text):
            out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                       % (rpr, text[pos:]))
        return ''.join(out)

    if '⟦EQREF:' not in xml:
        return xml
    return re.sub(r'<w:r>(?:(?!</w:r>).)*?</w:r>', lambda m: on_run(m.group(0)),
                  xml, flags=re.S)


def set_style(para_xml, style_id, extra_ppr=''):
    """把段落样式换成 style_id，并可追加 pPr 片段。"""
    new_ppr = '<w:pPr><w:pStyle w:val="%s"/>%s</w:pPr>' % (style_id, extra_ppr)
    if re.search(r'<w:pPr>.*?</w:pPr>', para_xml, re.S):
        return re.sub(r'<w:pPr>.*?</w:pPr>', new_ppr, para_xml, count=1, flags=re.S)
    if '<w:pPr/>' in para_xml:
        return para_xml.replace('<w:pPr/>', new_ppr, 1)
    return re.sub(r'(<w:p\b[^>]*>)', r'\1' + new_ppr, para_xml, count=1)


def red_run(text):
    return ('<w:r><w:rPr><w:color w:val="C00000"/><w:b/>'
            '<w:highlight w:val="yellow"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % text)


def replace_token_runs(para_xml, token_re, make_xml):
    """把段落里 token 命中的文本换成自定义 run，保持前后文字与原字符样式。

    直接对段落 XML 做字符串替换会把 <w:r> 嵌进 <w:t> 里（非法嵌套），
    必须按 run 拆开：前文 run + 新 run + 后文 run。token 跨 run 时匹配不到，
    调用方应检查替换后标记是否还在，再决定 fallback。
    """
    def on_run(m):
        run = m.group(0)
        tm = re.search(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', run, re.S)
        if not tm or not token_re.search(tm.group(1)):
            return run
        text = tm.group(1)
        rpr_m = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S)
        rpr = rpr_m.group(0) if rpr_m else ''
        out, pos = [], 0
        for t in token_re.finditer(text):
            if t.start() > pos:
                out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                           % (rpr, text[pos:t.start()]))
            out.append(make_xml(t))
            pos = t.end()
        if pos < len(text):
            out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                       % (rpr, text[pos:]))
        return ''.join(out)

    return re.sub(r'<w:r>(?:(?!</w:r>).)*?</w:r>', on_run, para_xml, flags=re.S)


# ------------------------------------------------------------------- 主逻辑

PARA_RE = re.compile(r'<w:p(?:\s[^>]*)?>.*?</w:p>|<w:p\s*/>', re.S)
# 只认冒号形态「图 N：/ 表 N:」。不能放行 '.'：「图 3.1 展示了……」是论文里
# 极常见的正文句式，会连同小数点一起被误判成题注（居中、小五号、丢首行缩进），
# 而逐字校验查不出样式层面的错误。
CAPTION_RE = re.compile(r'^\s*(图|表)\s*\d+\s*[：:]')


def plain(x):
    return re.sub(r'<[^>]+>', '', x)


MATHSRC_RUN = re.compile(
    r'<w:r><w:rPr><w:rStyle w:val="MathSource"/></w:rPr>'
    r'<w:t[^>]*>(.*?)</w:t></w:r>', re.S)


OMATHPARA = re.compile(r'<m:oMathPara>.*?</m:oMathPara>', re.S)


def is_display_math(para_xml):
    """判断段落是否为「独占一段」的块公式。

    只有独占整段的公式才该套 EquationPara（居中、无首行缩进）。原文里存在正文与
    $$ 写在同一段的情况，那种段落若整段居中，正文也会跟着居中并丢掉首行缩进。

    omml 模式  -> 含 <m:oMathPara> 且去掉公式后无其它正文
    latex 模式 -> 含 $$ 包裹的 MathSource run 且去掉公式后无其它正文
    """
    if '<m:oMathPara' in para_xml:
        rest = OMATHPARA.sub('', para_xml)
    elif any(t.startswith('$$') for t in MATHSRC_RUN.findall(para_xml)):
        rest = MATHSRC_RUN.sub('', para_xml)
    else:
        return False
    rest = re.sub(r'⟦[A-Z]+(?::[^⟧]*)?⟧', '', plain(rest))
    return rest.strip() == ''


def strip_marks(para_xml):
    """去掉段落里残留的 ⟦...⟧ 标记文本。"""
    return re.sub(r'⟦[A-Z]+(?::[^⟧]*)?⟧', '', para_xml)


RUN_RE = re.compile(r'<w:r>(?:(?!</w:r>).)*?</w:r>', re.S)
MARK_ONLY = re.compile(r'^(?:⟦[A-Z]+(?::[^⟧]*)?⟧)+$')


def split_inline_display(para_xml):
    """latex 模式下拆开「正文 + 行间公式」混排的段落。

    原文存在把 $$ 与正文写在同一段（中间无空行）的写法。omml 模式下 pandoc 会
    自动把公式拆成独立段落，latex 模式下公式只是个文本 run，不拆的话这个公式
    既不会居中也拿不到右编号。返回段落 XML 列表（不需要拆就原样返回单元素列表）。
    """
    if 'w:val="MathSource"' not in para_xml:
        return [para_xml]
    runs = RUN_RE.findall(para_xml)
    # 找出承载 $$ 块公式的 run
    idx = -1
    for k, r in enumerate(runs):
        if 'w:val="MathSource"' in r:
            t = re.search(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', r, re.S)
            if t and t.group(1).startswith('$$'):
                idx = k
                break
    if idx < 0:
        return [para_xml]

    ppr_m = re.search(r'<w:pPr>.*?</w:pPr>|<w:pPr/>', para_xml, re.S)
    ppr = ppr_m.group(0) if ppr_m else ''
    # 公式后面紧跟的、只含标记的 run 要留在公式段落里（⟦EQNO⟧ / ⟦BOX⟧）
    end = idx + 1
    while end < len(runs):
        t = re.search(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', runs[end], re.S)
        if t and MARK_ONLY.match(t.group(1)):
            end += 1
        else:
            break

    before, eq, after = runs[:idx], runs[idx:end], runs[end:]
    if not before and not after:
        return [para_xml]

    def mk(rs, keep_ppr):
        return '<w:p>%s%s</w:p>' % (keep_ppr, ''.join(rs))

    out = []
    if before and keep_para(''.join(before)):
        out.append(mk(before, ppr))
    out.append(mk(eq, ppr))
    if after:
        # 同一段里可能还有第二个行间公式，对剩余部分递归处理
        rest = mk(after, ppr)
        if 'w:val="MathSource"' in rest:
            out.extend(split_inline_display(rest))
        elif keep_para(rest):
            # 这里不能 strip_marks：⟦EQNO⟧ 标记通常和后续正文挤在同一个 run 里
            # （pandoc 会把连续文本合成一个 run），提前剥掉就等于把那个公式的
            # 编号丢了。交给主循环处理，最后统一清理残留标记。
            out.append(rest)
    return out


def keep_para(para_xml):
    """段落是否需要保留：有实际文字，或带着还没处理的 ⟦...⟧ 标记。"""
    return bool(plain(strip_marks(para_xml)).strip()) or '⟦' in para_xml


# pandoc/texmath 把 aligned 的对齐符 & 直接当普通字符写进 OMML，公式里会出现
# 可见的 "&"。OMML 里对齐点的正确表达是 m:rPr/m:aln，这里做替换。
AMP_RUN = re.compile(r'<m:r>(?:(?!</m:r>).)*?<m:t[^>]*>&amp;</m:t>\s*</m:r>', re.S)
ALN_RUN = '<m:r><m:rPr><m:aln/></m:rPr></m:r>'


def fix_align_markers(xml, stats):
    xml, n = AMP_RUN.subn(ALN_RUN, xml)
    stats['align_fixed'] = n
    return xml


SELFCLOSE_WS = re.compile(r'(<[^<>]*?)\s+/>')


def normalize_xml(xml):
    """把自闭合标签的 ` />` 统一收成 `/>`。

    pandoc 输出 `<w:pStyle w:val="BodyText" />`（带空格），本工具自己拼的是
    `<w:pStyle w:val="BodyText"/>`（不带）。两种混在一个文档里，任何按字面写
    `w:val="X"/>` 的正则都会静默漏掉一半，非常难查。入口统一归一化一次即可根治。

    `[^<>]*?` 保证只在标签内部生效：合法 XML 的文本内容里不会出现裸的 < 或 >，
    所以不会误伤正文里形如 "a /> b" 的文字。
    """
    return SELFCLOSE_WS.sub(r'\1/>', xml)


XML_ESC = {'&': '&amp;', '<': '&lt;', '>': '&gt;'}


def xesc(s):
    return ''.join(XML_ESC.get(c, c) for c in s)


def mathsrc_run(tex, display):
    """把 LaTeX 源码写成带 MathSource 字符样式的文本 run。

    保留 $...$ / $$...$$ 定界符，MathType 的 Convert Equations（输入选
    "Text using a translator" -> TeX）和 MTCommand_TeXToggle 宏都靠它识别边界；
    AxMath 侧则可直接框选后走插件的批量处理。
    """
    d = '$$' if display else '$'
    body = tex.replace('\n', ' ') if not display else tex
    return ('<w:r><w:rPr><w:rStyle w:val="MathSource"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>'
            % xesc(d + body + d))


MATH_TOKEN = re.compile(r'⟦MATH([DI]):(\d+)⟧')


def fill_math_tokens(xml, math_list, stats):
    """把 ⟦MATHD:n⟧ / ⟦MATHI:n⟧ 令牌换成 LaTeX 源码 run。

    令牌可能被 pandoc 拆到多个 w:t 里，所以先把相邻 run 的文本合并再替换。
    """
    if not math_list:
        return xml

    # pandoc 不会在令牌中间插标签（纯 ASCII+私用符号，无 markdown 语义），
    # 但保险起见先做一次跨 run 的令牌拼合检查。
    def repl(m):
        kind, idx = m.group(1), int(m.group(2))
        if idx >= len(math_list):
            stats['math_orphan'] += 1
            return ''
        stats['math_filled'] += 1
        return mathsrc_run(math_list[idx]['tex'], kind == 'D')

    # 令牌所在的 w:t 需要整体替换：先把 <w:t>...⟦MATHD:1⟧...</w:t> 拆开
    def on_run(m):
        run = m.group(0)
        if '⟦MATH' not in run:
            return run
        # 取出 run 里的文本，按令牌切段，逐段重建
        tm = re.search(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', run, re.S)
        if not tm:
            return run
        text = tm.group(1)
        rpr = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S)
        rpr = rpr.group(0) if rpr else ''
        out, pos = [], 0
        for t in MATH_TOKEN.finditer(text):
            if t.start() > pos:
                out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                           % (rpr, text[pos:t.start()]))
            out.append(repl(t))
            pos = t.end()
        if pos < len(text):
            out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                       % (rpr, text[pos:]))
        return ''.join(out)

    return re.sub(r'<w:r>(?:(?!</w:r>).)*?</w:r>', on_run, xml, flags=re.S)


def process(xml, stats, math_list=None, seq_field=True):
    xml = normalize_xml(xml)
    xml = fix_align_markers(xml, stats)
    xml = fill_math_tokens(xml, math_list or [], stats)
    xml = fill_eqrefs(xml, stats)
    ids = BmkIds(xml)
    body_m = re.search(r'(<w:body>)(.*)(</w:body>)', xml, re.S)
    head, body, tail = xml[:body_m.start(2)], body_m.group(2), xml[body_m.end(2):]

    # 切成 [段落 | 其它节点] 的有序块列表。latex 模式下顺手拆开「正文 + 行间公式」
    # 混排的段落，让它与 omml 模式（pandoc 会自动拆）表现一致。
    blocks, pos = [], 0
    for m in PARA_RE.finditer(body):
        if m.start() > pos:
            blocks.append(('raw', body[pos:m.start()]))
        parts = split_inline_display(m.group(0))
        if len(parts) > 1:
            stats['para_split'] += 1
        for p in parts:
            blocks.append(('p', p))
        pos = m.end()
    if pos < len(body):
        blocks.append(('raw', body[pos:]))

    out = []
    i = 0
    while i < len(blocks):
        kind, x = blocks[i]
        if kind != 'p':
            out.append(x)
            i += 1
            continue

        txt = plain(x)

        # --- 1/2: 编号公式 -> 制表位段落 -----------------------------
        eq = re.search(r'⟦EQNO:([^|⟧]*)\|([^|⟧]*)\|(\d+)⟧', txt)
        if eq:
            boxed = '⟦BOX⟧' in txt
            num, bmk = eq.group(1), eq.group(2)
            # 情况 A（latex 模式）：公式与编号标记在同一段落
            if is_display_math(x):
                out.append(numbered_eq_p(strip_marks(x), num, bmk, ids,
                                         seq_field, boxed))
                stats['eqnum'] += 1
                if boxed:
                    stats['boxed'] += 1
                i += 1
                continue
            # 情况 B（omml 模式）：公式在上一段落
            j = len(out) - 1
            while j >= 0 and not out[j].strip():
                j -= 1
            if j >= 0 and is_display_math(out[j]):
                out[j] = numbered_eq_p(out[j], num, bmk, ids, seq_field, boxed)
                stats['eqnum'] += 1
                stats['eq_plain'] -= 1  # 该公式已并入编号段落，不再计入未编号
                if boxed:
                    stats['boxed'] += 1
                # 编号标记所在的段落可能还带着正文：原文存在把 $$ 与正文写在同一段的
                # 写法，pandoc 会把公式拆成独立段落，标记则落在后半段正文里。
                # 剥掉标记后若仍有内容必须保留，否则整段正文被吞掉。
                rest = strip_marks(x)
                if plain(rest).strip():
                    out.append(rest)
                    stats['tail_kept'] += 1
                i += 1
                continue
            stats['eqnum_orphan'] += 1

        # 孤立的 ⟦BOX⟧（无编号）
        if '⟦BOX⟧' in txt and 'EQNO' not in txt and out:
            j = len(out) - 1
            while j >= 0 and '<m:oMathPara' not in out[j]:
                j -= 1
            if j >= 0:
                out[j] = set_style(out[j], 'EquationPara', BOXBORDER)
                stats['boxed'] += 1
                i += 1
                continue

        # --- 3: 缺失图片 --------------------------------------------
        mi = re.search(r'⟦MISSINGIMG:([^⟧]*)⟧', txt)
        if mi:
            # 段落里还有其他正文（行内插图缺图）：就地换红色占位 run。
            # 整段替换会把同段落的正文一起吞掉。
            if plain(strip_marks(x)).strip():
                new_x = replace_token_runs(
                    x, re.compile(r'⟦MISSINGIMG:([^⟧]*)⟧'),
                    lambda m: red_run('［待插入图片：%s］' % m.group(1)))
                if '⟦MISSINGIMG' not in new_x:
                    out.append(new_x)
                    stats['missing_img'] += 1
                    i += 1
                    continue
            p = ('<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>%s</w:p>'
                 % red_run('［待插入图片：%s］' % mi.group(1)))
            out.append(p)
            stats['missing_img'] += 1
            i += 1
            continue

        # --- 4: 图占位 ----------------------------------------------
        if '⟦FIGPH⟧' in txt:
            label = txt.replace('⟦FIGPH⟧', '').strip()
            p = ('<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>%s</w:p>'
                 % red_run('［%s］' % label))
            out.append(p)
            stats['figph'] += 1
            i += 1
            continue

        # --- 5: 图表题注 --------------------------------------------
        # SourceCode 段落排除：代码块里的「图 1：xxx」是字面文字，不是题注
        if CAPTION_RE.match(txt) and len(txt) < 120 and 'SourceCode' not in x:
            out.append(set_style(x, 'Caption'))
            stats['caption'] += 1
            i += 1
            continue

        # --- 6: 未编号块公式 ----------------------------------------
        if is_display_math(x):
            out.append(set_style(strip_marks(x), 'EquationPara'))
            stats['eq_plain'] += 1
            i += 1
            continue

        out.append(x)
        i += 1

    new_body = ''.join(out)
    # 清掉可能残留的标记
    left = re.findall(r'⟦[A-Z]+[^⟧]*⟧', new_body)
    if left:
        stats['leftover'] = left[:10]
        new_body = re.sub(r'⟦[A-Z]+[^⟧]*⟧', '', new_body)
    return head + new_body + tail


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src')
    ap.add_argument('dst')
    ap.add_argument('--math-json', help='latex 模式下 pre.py 落盘的公式源码')
    ap.add_argument('--number-style', choices=['plain', 'chapter'], default='plain')
    a = ap.parse_args()
    _console_safe()

    math_list = []
    if a.math_json and os.path.exists(a.math_json):
        math_list = json.load(open(a.math_json, encoding='utf-8'))

    shutil.copy(a.src, a.dst + '.tmp')
    stats = {'eqnum': 0, 'eqnum_orphan': 0, 'boxed': 0, 'missing_img': 0,
             'figph': 0, 'caption': 0, 'eq_plain': 0, 'align_fixed': 0,
             'math_filled': 0, 'math_orphan': 0, 'ref_fields': 0, 'tail_kept': 0, 'para_split': 0,
             'leftover': []}

    zin = zipfile.ZipFile(a.dst + '.tmp')
    with zipfile.ZipFile(a.dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                data = process(data.decode('utf-8'), stats, math_list,
                               seq_field=(a.number_style == 'plain')).encode('utf-8')
            zout.writestr(item, data)
    zin.close()
    os.remove(a.dst + '.tmp')

    if math_list:
        print('[post] LaTeX 源码回填 %d 个（未匹配 %d）待 MathType/AxMath 批量转换'
              % (stats['math_filled'], stats['math_orphan']))
    else:
        print('[post] aligned 对齐符 & 修正 %d 处（改为 OMML m:aln 对齐点）' % stats['align_fixed'])
    print('[post] 公式右编号 %d 个（孤立标记 %d）' % (stats['eqnum'], stats['eqnum_orphan']))
    print('[post] 正文 REF 交叉引用域 %d 个（Word 里按 F9 联动更新）' % stats['ref_fields'])
    print('[post] 加框公式 %d 个' % stats['boxed'])
    print('[post] 未编号块公式套样式 %d 个' % stats['eq_plain'])
    print('[post] 图表题注 %d 条' % stats['caption'])
    print('[post] 缺图占位 %d 处，图占位 %d 处' % (stats['missing_img'], stats['figph']))
    if stats['leftover']:
        print('[post] 警告：残留标记 %s' % stats['leftover'])


if __name__ == '__main__':
    main()
