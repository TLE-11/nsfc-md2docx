#!/usr/bin/env python3
"""生成中文项目申报书（本子）排版用的 pandoc reference.docx。

基于 pandoc 自带模板改样式，可重复生成，不需要手工调 Word。
默认规范（可按单位模板改 SPEC 字典）：
    正文      宋体 / Times New Roman，小四(12pt)，行距 1.5 倍，首行缩进 2 字符
    一级标题  黑体，三号(16pt)，居中
    二级标题  黑体，四号(14pt)，左对齐
    三/四级   黑体，小四(12pt)，左对齐
    图表题注  宋体，五号(10.5pt)，居中，无缩进
    公式段落  居中，无缩进
    页面      A4，上下 2.54cm，左右 3.17cm
"""
import argparse
import re
import shutil
import subprocess
import sys
import zipfile
import os
import tempfile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def which_pandoc():
    """与 md2docx.py 同一套查找逻辑：Windows 上装完 pandoc 没重启终端时，
    PATH 里可能还没有，直接 subprocess 'pandoc' 会抛 FileNotFoundError。"""
    p = shutil.which('pandoc')
    if p:
        return p
    for c in (os.path.expandvars(r'%LOCALAPPDATA%\Pandoc\pandoc.exe'),
              r'C:\Program Files\Pandoc\pandoc.exe'):
        if os.path.exists(c):
            return c
    sys.exit('找不到 pandoc。装法：\n'
             '  Windows : winget install --id JohnMacFarlane.Pandoc\n'
             '            或 choco install pandoc\n'
             '  macOS   : brew install pandoc\n'
             '  Linux   : apt install pandoc')

# ---------------------------------------------------------------- 排版规范

SPEC = {
    # styleId: (中文字体角色, 西文字体, 磅值, 对齐, 首行缩进字符数, 段前pt, 段后pt, 行距倍数, 加粗)
    # 中文字体写角色名（song/hei/kai），最终字体名由 --font-profile 决定
    'Normal':         ('song', 'Times New Roman', 12,   None,     2, 0, 0, 1.5, False),
    'BodyText':       ('song', 'Times New Roman', 12,   None,     2, 0, 0, 1.5, False),
    'FirstParagraph': ('song', 'Times New Roman', 12,   None,     2, 0, 0, 1.5, False),
    'Compact':        ('song', 'Times New Roman', 12,   None,     2, 0, 0, 1.5, False),
    'BlockText':      ('kai',  'Times New Roman', 12,   None,     2, 6, 6, 1.5, False),
    'Heading1':       ('hei',  'Times New Roman', 16,   'center', 0, 12, 12, 1.5, True),
    'Heading2':       ('hei',  'Times New Roman', 14,   None,     0, 12, 6,  1.5, True),
    'Heading3':       ('hei',  'Times New Roman', 12,   None,     0, 6,  6,  1.5, True),
    'Heading4':       ('hei',  'Times New Roman', 12,   None,     0, 6,  6,  1.5, True),
    'Heading5':       ('hei',  'Times New Roman', 12,   None,     0, 6,  6,  1.5, True),
    'Caption':        ('song', 'Times New Roman', 10.5, 'center', 0, 6,  6,  1.0, False),
    'ImageCaption':   ('song', 'Times New Roman', 10.5, 'center', 0, 6,  6,  1.0, False),
    'TableCaption':   ('song', 'Times New Roman', 10.5, 'center', 0, 6,  6,  1.0, False),
    'Figure':         ('song', 'Times New Roman', 12,   'center', 0, 6,  6,  1.0, False),
    'CaptionedFigure':('song', 'Times New Roman', 12,   'center', 0, 6,  6,  1.0, False),
}

# 字体档位。中文本子出稿一律用 windows 档（宋体/黑体/楷体是评审方机器上一定有的）。
# macOS 上没有这三个字体，本机用 LibreOffice 渲染预览时中文会整片消失，
# 所以另备一档用 macOS 自带的对应字体，仅供肉眼检查排版用。
FONT_PROFILES = {
    'windows': {'song': '宋体', 'hei': '黑体', 'kai': '楷体'},
    'macos':   {'song': 'Songti SC', 'hei': 'Heiti SC', 'kai': 'Kaiti SC'},
}

# 新增样式：公式段落（居中、无缩进）、公式编号、带编号公式、LaTeX 源码字符样式
EQ_STYLE_ID = 'EquationPara'
EQNUM_STYLE_ID = 'EquationNumber'
EQNUMBERED_STYLE_ID = 'EquationNumbered'
MATHSRC_STYLE_ID = 'MathSource'

# 版面可用文字宽度：A4 宽 21cm - 左右页边距 3.17cm x2 = 14.66cm
# 1cm = 567 twips，制表位就按这个宽度摆
PAGE_W_CM, MARGIN_LR_CM = 21.0, 3.17
TEXT_WIDTH_TWIPS = int(round((PAGE_W_CM - 2 * MARGIN_LR_CM) * 567))


def rpr(cn, en, pt, bold):
    sz = int(round(pt * 2))
    b = '<w:b/><w:bCs/>' if bold else ''
    return ('<w:rPr>'
            '<w:rFonts w:ascii="{en}" w:hAnsi="{en}" w:eastAsia="{cn}" w:cs="{en}"/>'
            '{b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            '</w:rPr>').format(en=en, cn=cn, b=b, sz=sz)


def ppr(align, indent_chars, before, after, line_mult, extra=''):
    parts = ['<w:pPr>']
    parts.append('<w:spacing w:before="{b}" w:after="{a}" w:line="{l}" w:lineRule="auto"/>'
                 .format(b=int(before * 20), a=int(after * 20), l=int(line_mult * 240)))
    if indent_chars:
        parts.append('<w:ind w:firstLineChars="{c}" w:firstLine="0"/>'
                     .format(c=int(indent_chars * 100)))
    else:
        parts.append('<w:ind w:firstLineChars="0" w:firstLine="0"/>')
    if align:
        parts.append('<w:jc w:val="%s"/>' % align)
    parts.append(extra)
    parts.append('</w:pPr>')
    return ''.join(parts)


def patch_style(xml, style_id, cn, en, pt, align, ind, before, after, line, bold):
    """替换指定 styleId 的 <w:pPr> 与 <w:rPr>。"""
    pat = re.compile(r'(<w:style\b[^>]*w:styleId="%s"[^>]*>)(.*?)(</w:style>)'
                     % re.escape(style_id), re.S)
    m = pat.search(xml)
    if not m:
        return xml, False
    head, body, tail = m.group(1), m.group(2), m.group(3)
    body = re.sub(r'<w:pPr>.*?</w:pPr>|<w:pPr/>', '', body, flags=re.S)
    body = re.sub(r'<w:rPr>.*?</w:rPr>|<w:rPr/>', '', body, flags=re.S)
    # w:name / w:basedOn / w:next / w:link 等要留在前面，pPr 与 rPr 追加到末尾即可
    body = body + ppr(align, ind, before, after, line) + rpr(cn, en, pt, bold)
    return xml[:m.start()] + head + body + tail + xml[m.end():], True


def add_style(xml, style_id, name, based_on, body_inner, stype='paragraph'):
    if 'w:styleId="%s"' % style_id in xml:
        return xml
    based = '<w:basedOn w:val="%s"/>' % based_on if based_on else ''
    st = ('<w:style w:type="{st}" w:customStyle="1" w:styleId="{sid}">'
          '<w:name w:val="{nm}"/>{based}<w:qFormat/>'
          '{inner}</w:style>').format(st=stype, sid=style_id, nm=name,
                                      based=based, inner=body_inner)
    return xml.replace('</w:styles>', st + '</w:styles>')


def patch_styles_xml(xml, spec, fonts):
    missing = []
    for sid, (role, en, pt, align, ind, before, after, line, bold) in spec.items():
        cn = fonts.get(role, role)
        xml, ok = patch_style(xml, sid, cn, en, pt, align, ind, before, after, line, bold)
        if not ok:
            missing.append(sid)
    song = fonts['song']
    # 公式段落样式：居中、无首行缩进、单倍行距
    xml = add_style(xml, EQ_STYLE_ID, 'Equation Para', 'Normal',
                    ppr('center', 0, 6, 6, 1.0) + rpr(song, 'Cambria Math', 12, False))
    xml = add_style(xml, EQNUM_STYLE_ID, 'Equation Number', 'Normal',
                    ppr('right', 0, 6, 6, 1.0) + rpr(song, 'Times New Roman', 12, False))
    # 带编号的行间公式：靠制表位实现「公式居中 + 编号右对齐」。
    # 不用无边框表格，因为 Word/WPS 会给无边框表格画屏幕虚框（不打印但碍眼），
    # 而"查看虚框"是应用级开关、存不进文件，每台机器都得手动关。
    tabs = ('<w:tabs><w:tab w:val="center" w:pos="%d"/>'
            '<w:tab w:val="right" w:pos="%d"/></w:tabs>'
            % (TEXT_WIDTH_TWIPS // 2, TEXT_WIDTH_TWIPS))
    xml = add_style(xml, EQNUMBERED_STYLE_ID, 'Equation Numbered', 'Normal',
                    ppr(None, 0, 6, 6, 1.0, extra=tabs)
                    + rpr(song, 'Cambria Math', 12, False))
    # 字符样式：--math-mode latex 时承载 LaTeX 源码。等宽 + 蓝色，肉眼一眼能认出
    # 哪些还没转成公式；MathType / AxMath 批量转换后这些样式自然消失。
    # 注意：样式「名称」必须与 styleId 一致。Word 的对象模型和查找功能都按名称匹配
    # （ActiveDocument.Styles("MathSource") / Find.Style = "MathSource"），
    # 名称写成 "Math Source (LaTeX)" 会让配套 VBA 宏整条路线失效。
    xml = add_style(xml, MATHSRC_STYLE_ID, MATHSRC_STYLE_ID, None,
                    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" '
                    'w:eastAsia="Consolas" w:cs="Consolas"/>'
                    '<w:color w:val="1F4E79"/><w:sz w:val="21"/>'
                    '<w:szCs w:val="21"/></w:rPr>', stype='character')
    return xml, missing


def patch_settings_xml(xml, math_font='Cambria Math'):
    """设置 Word 公式默认字体（settings.xml 里是 m: 命名空间的 m:mathPr）。"""
    mathpr = ('<m:mathPr><m:mathFont m:val="{f}"/><m:brkBin m:val="before"/>'
              '<m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/>'
              '<m:lMargin m:val="0"/><m:rMargin m:val="0"/>'
              '<m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/>'
              '<m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>'
              ).format(f=math_font)
    xml, n = re.subn(r'<m:mathPr>.*?</m:mathPr>|<m:mathPr\s*/>', mathpr, xml, flags=re.S)
    if n == 0:
        xml = xml.replace('</w:settings>', mathpr + '</w:settings>')
    return xml


def patch_page(xml):
    """A4 + 中文公文常用页边距（上下 2.54cm，左右 3.17cm）。"""
    def cm(v):
        return int(round(v * 567))
    sect = ('<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="{t}" w:right="{r}" w:bottom="{b}" w:left="{l}" '
            'w:header="851" w:footer="992" w:gutter="0"/>'
            ).format(t=cm(2.54), b=cm(2.54), l=cm(3.17), r=cm(3.17))
    if '<w:sectPr' in xml:
        xml = re.sub(r'(<w:sectPr[^>]*>)(.*?)(</w:sectPr>)',
                     lambda m: m.group(1) + sect + m.group(3), xml, flags=re.S)
    return xml


def _console_safe():
    """Windows GBK 控制台打印 GBK 之外的字符会 UnicodeEncodeError，掩盖真实报错。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors='replace')
        except Exception:
            pass


def main():
    _console_safe()
    ap = argparse.ArgumentParser()
    ap.add_argument('out')
    ap.add_argument('--math-font', default='Cambria Math')
    ap.add_argument('--font-profile', choices=sorted(FONT_PROFILES), default='windows',
                    help='windows=宋体/黑体/楷体（出稿用）; macos=Songti SC 等（本机预览用）')
    a = ap.parse_args()
    fonts = FONT_PROFILES[a.font_profile]

    base = subprocess.run([which_pandoc(), '--print-default-data-file', 'reference.docx'],
                          capture_output=True, check=True).stdout
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, 'base.docx')
    open(src, 'wb').write(base)

    zin = zipfile.ZipFile(src)
    missing = []
    with zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/styles.xml':
                x, missing = patch_styles_xml(data.decode('utf-8'), SPEC, fonts)
                data = x.encode('utf-8')
            elif item.filename == 'word/settings.xml':
                data = patch_settings_xml(data.decode('utf-8'), a.math_font).encode('utf-8')
            elif item.filename == 'word/document.xml':
                data = patch_page(data.decode('utf-8')).encode('utf-8')
            zout.writestr(item, data)
    zin.close()
    shutil.rmtree(tmp, ignore_errors=True)

    print('[ref] 生成 %s' % a.out)
    print('[ref] 已改写样式 %d 个' % (len(SPEC) - len(missing)))
    if missing:
        print('[ref] 模板中未找到（已跳过）: %s' % ', '.join(missing))
    print('[ref] 字体档位: %s (%s / %s / %s)'
          % (a.font_profile, fonts['song'], fonts['hei'], fonts['kai']))
    print('[ref] 公式字体: %s' % a.math_font)


if __name__ == '__main__':
    main()
