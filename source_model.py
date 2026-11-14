"""用 Pandoc 的文档结构提取可见正文，避免把链接地址、代码或注释当成正文。"""
import json
import re
import subprocess
from xml.etree import ElementTree as ET

import pre
from md2docx import PANDOC_FROM, which_pandoc

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'


def markdown_model(text):
    model = {'block': 0, 'inline': 0, 'tag': 0, 'images': 0,
             'wiki_images': 0, 'ampersands': 0}
    # 只规整数学块的空行和容器前缀，不使用转换器的编号、字体或正文改写。
    text = pre.map_math(text, lambda s: '$$' + pre.compact_math(s) + '$$',
                        lambda s: '$' + s + '$')
    def skip_wiki_image(match):
        model['wiki_images'] += 1
        return ''
    text = pre.map_noncode(text, lambda s: re.sub(r'!\[\[[^\]]*\]\]', skip_wiki_image, s))
    result = subprocess.run([which_pandoc(), '--from', PANDOC_FROM, '--to', 'json'],
                            input=text, encoding='utf-8', capture_output=True)
    if result.returncode:
        raise ValueError('源文档解析失败: ' + result.stderr.strip())

    def visit(node):
        if isinstance(node, list):
            return ''.join(visit(v) for v in node)
        if not isinstance(node, dict):
            return ''
        kind, content = node.get('t'), node.get('c')
        if kind == 'Str':
            return content
        if kind in ('Code', 'CodeBlock', 'RawInline', 'RawBlock'):
            return ''
        if kind == 'Math':
            display = content[0]['t'] == 'DisplayMath'
            model['block' if display else 'inline'] += 1
            model['tag'] += int(display and bool(pre.TAG_RE.search(content[1])))
            model['ampersands'] += len(re.findall(r'(?<!\\)(?:\\\\)*\\&', content[1]))
            return ''
        if kind == 'Image':
            model['images'] += 1
            return ''  # alt 属性不是正文；独立图片的题注位于 Figure/Caption 中。
        if kind == 'Link':
            return visit(content[1])
        return visit(content)

    model['text'] = visit(json.loads(result.stdout)['blocks'])
    return model


def docx_text(parts):
    """按正文阅读顺序展开脚注，忽略代码和公式；只读取可见的 w:t。"""
    notes = {}
    for filename, tag in (('word/footnotes.xml', 'footnote'),
                          ('word/endnotes.xml', 'endnote')):
        if filename in parts:
            for note in ET.fromstring(parts[filename]):
                notes[(tag, note.get(W + 'id'))] = note

    def visit(node, active=frozenset()):
        if node.tag == M + 'oMath':
            return ''
        if node.tag == W + 'p':
            style = node.find(W + 'pPr/' + W + 'pStyle')
            if style is not None and style.get(W + 'val') == 'SourceCode':
                return ''
        if node.tag == W + 'r':
            style = node.find(W + 'rPr/' + W + 'rStyle')
            if style is not None and style.get(W + 'val') in ('VerbatimChar', 'MathSource'):
                return ''
        for tag in ('footnote', 'endnote'):
            if node.tag == W + tag + 'Reference':
                key = (tag, node.get(W + 'id'))
                return visit(notes[key], active | {key}) if key in notes and key not in active else ''
        if node.tag == W + 't':
            return node.text or ''
        return ''.join(visit(child, active) for child in node)

    return visit(ET.fromstring(parts['word/document.xml']))
