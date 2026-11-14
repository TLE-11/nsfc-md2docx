"""给原生显示公式分配实际版心宽度，在最外层运算符前生成 OMML 多行结构。

宽度是保守估计（pt），不是字体排版引擎。复杂不可拆分项超宽时明确失败，
不缩小字号、栅格化或截断数学内容。m:eqArr 结构见 Microsoft Open XML 文档：
https://learn.microsoft.com/dotnet/api/documentformat.openxml.math.equationarray
"""
from xml.dom import minidom
from bisect import bisect_right
import unicodedata


class LayoutError(ValueError):
    pass


def children(node, tag=None):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE
            and (tag is None or n.tagName == tag)]


def child(node, tag):
    return next(iter(children(node, tag)), None) if node is not None else None


def number(node, attr, default=0):
    value = node.getAttribute(attr) if node is not None else ''
    return float(value) if value else default


def element(doc, tag, **attrs):
    node = doc.createElement(tag)
    for key, value in attrs.items():
        node.setAttribute(key.replace('_', ':'), str(value))
    return node


def text(node):
    return ''.join(n.data if n.nodeType == n.TEXT_NODE else text(n)
                   for n in node.childNodes)


def glyph_width(value, em):
    width = 0
    for char in value:
        if unicodedata.combining(char):
            continue
        if char in 'il., ': factor = .35
        elif char in 'MW': factor = 1
        elif char in '+−-=<>≤≥≈≠≡∈': factor = 1
        elif unicodedata.east_asian_width(char) in ('W', 'F'): factor = 1.1
        else: factor = .7
        width += factor * em
    return width


def math_width(node, em):
    tag = node.tagName
    if tag.endswith('Pr'):
        return 0
    if tag == 'm:r':
        size = child(child(node, 'w:rPr'), 'w:sz')
        em = number(size, 'w:val', em * 2) / 2
        return sum(glyph_width(text(t), em) for t in children(node, 'm:t'))
    if tag in ('m:sSub', 'm:sSup', 'm:sSubSup'):
        base = child(node, 'm:e')
        scripts = [math_width(n, em * .75) for n in children(node)
                   if n.tagName in ('m:sub', 'm:sup')]
        return (math_width(base, em) if base else 0) + max(scripts, default=0) + .15 * em
    if tag == 'm:f':
        return max((math_width(n, em) for n in children(node)
                    if n.tagName in ('m:num', 'm:den')), default=0) + .7 * em
    if tag == 'm:eqArr':
        columns = []
        for row in children(node, 'm:e'):
            segments = [0]
            for part in children(row):
                if child(child(part, 'm:rPr'), 'm:aln') is not None:
                    segments.append(0)
                segments[-1] += math_width(part, em)
            for i, value in enumerate(segments):
                if i == len(columns): columns.append(value)
                else: columns[i] = max(columns[i], value)
        return sum(columns)
    if tag == 'm:m':
        rows = [children(n, 'm:e') for n in children(node, 'm:mr')]
        cols = max((len(row) for row in rows), default=0)
        return sum(max((math_width(row[i], em) for row in rows if i < len(row)),
                       default=0) for i in range(cols)) + max(0, cols - 1) * em
    extra = {'m:d': 2, 'm:rad': 1, 'm:nary': 2, 'm:func': .5}.get(tag, 0) * em
    return sum(math_width(n, em) for n in children(node)) + extra


def split_rows(nodes, limit, em):
    """只在最外层二元运算符前断行，结构节点及成对括号保持完整。"""
    widths = [math_width(n, em) for n in nodes]
    cumulative = [0]
    for width in widths:
        cumulative.append(cumulative[-1] + width)
    breaks, depth = [], 0
    for i, node in enumerate(nodes):
        if node.tagName != 'm:r' or node.getElementsByTagName('m:nor'):
            continue
        value = ''.join(text(t) for t in children(node, 'm:t'))
        if depth == 0 and value in ('+', '−', '-', '=', '<', '>', '≤', '≥', '≈', '≠', '≡') and i:
            breaks.append(i)
        for char in value:
            if char in '([{': depth += 1
            elif char in ')]}': depth = max(0, depth - 1)
    rows, start = [], 0
    break_widths = [cumulative[i] for i in breaks]
    while cumulative[-1] - cumulative[start] > limit:
        choice = bisect_right(break_widths, cumulative[start] + limit) - 1
        if choice < 0 or breaks[choice] <= start:
            raise LayoutError('存在无法安全自动换行的公式项；请用 aligned 显式分行或调整页面宽度')
        end = breaks[choice]
        rows.append(nodes[start:end])
        start = end
    rows.append(nodes[start:])
    return rows


def wrap_math(math, limit, em):
    if math_width(math, em) <= limit:
        return False
    nodes = children(math)
    # 已有 aligned 保留行序；只有超宽时才重新分行并统一左侧对齐点。
    if len(nodes) == 1 and nodes[0].tagName == 'm:eqArr':
        rows = []
        for row in children(nodes[0], 'm:e'):
            rows.extend(split_rows(children(row), limit, em))
    else:
        rows = split_rows(nodes, limit, em)
    doc = math.ownerDocument
    array = element(doc, 'm:eqArr')
    props = element(doc, 'm:eqArrPr')
    for tag, value in (('baseJc', 'center'), ('maxDist', 0), ('objDist', 0)):
        props.appendChild(element(doc, 'm:' + tag, m_val=value))
    array.appendChild(props)
    for row in rows:
        base = element(doc, 'm:e')
        marker = element(doc, 'm:r')
        rpr = element(doc, 'm:rPr')
        rpr.appendChild(element(doc, 'm:aln'))
        marker.appendChild(rpr)
        base.appendChild(marker)
        for node in row:
            # 旧行的对齐点可能把新行扩张；只移除对齐属性，不移除文本。
            for alignment in list(node.getElementsByTagName('m:aln')):
                alignment.parentNode.removeChild(alignment)
            base.appendChild(node)
        array.appendChild(base)
    for node in list(math.childNodes):
        math.removeChild(node)
    math.appendChild(array)
    return True


def section_width(section):
    page, margin = child(section, 'w:pgSz'), child(section, 'w:pgMar')
    width = number(page, 'w:w', 11906) - number(margin, 'w:left', 1797) - number(margin, 'w:right', 1797)
    width -= number(margin, 'w:gutter')
    cols = child(section, 'w:cols')
    explicit = children(cols, 'w:col') if cols else []
    if explicit:
        width = min(number(c, 'w:w', width) for c in explicit)
    else:
        count = max(1, number(cols, 'w:num', 1))
        width = (width - (count - 1) * number(cols, 'w:space', 720)) / count
    return width / 20


def document_width(xml):
    sections = minidom.parseString(xml).getElementsByTagName('w:sectPr')
    return min((section_width(s) for s in sections), default=415.6)


def paragraph_widths(doc, fallback):
    widths, pending = {}, []
    for p in doc.getElementsByTagName('w:p'):
        pending.append(p)
        section = child(child(p, 'w:pPr'), 'w:sectPr')
        if section is not None:
            width = section_width(section)
            widths.update((n, width) for n in pending)
            pending = []
    body = child(doc.documentElement, 'w:body')
    section = child(body, 'w:sectPr')
    widths.update((p, section_width(section) if section else fallback) for p in pending)
    return widths


def container_width(node, page):
    """表格用单元格/网格宽度，普通段落用所在节的版心宽度。"""
    parent = node.parentNode
    while parent is not None and parent.nodeType != parent.DOCUMENT_NODE:
        if parent.tagName == 'w:tc':
            tc = parent
            table = tc.parentNode
            while table.tagName != 'w:tbl':
                table = table.parentNode
            outer = container_width(table, page)
            table_props = child(table, 'w:tblPr')
            outer -= max(0, number(child(table_props, 'w:tblInd'), 'w:w')) / 20
            preferred = child(table_props, 'w:tblW')
            preferred_value = number(preferred, 'w:w')
            if preferred is not None and preferred_value > 0:
                if preferred.getAttribute('w:type') == 'dxa':
                    outer = min(outer, preferred_value / 20)
                elif preferred.getAttribute('w:type') == 'pct':
                    outer *= min(1, preferred_value / 5000)
            cell_width = child(child(tc, 'w:tcPr'), 'w:tcW')
            value = number(cell_width, 'w:w')
            kind = cell_width.getAttribute('w:type') if cell_width else ''
            grid = child(table, 'w:tblGrid')
            sizes = [number(c, 'w:w') for c in children(grid, 'w:gridCol')] if grid else []
            start = 0
            for other in children(tc.parentNode, 'w:tc'):
                span = int(number(child(child(other, 'w:tcPr'), 'w:gridSpan'), 'w:val', 1))
                if other is tc:
                    break
                start += span
            grid_width = (sum(sizes[start:start + span]) / max(sum(sizes), outer * 20) * outer
                          if sum(sizes) > 0 else None)
            if value > 0 and kind == 'dxa':
                width = value / 20
            elif value > 0 and kind == 'pct':
                width = outer * value / 5000
            else:
                width = outer / max(1, len(children(tc.parentNode, 'w:tc')))
            if grid_width is not None:
                width = min(width, grid_width)
            # 默认单元格左右内边距各 108 twips，显式配置时使用配置值。
            margins = child(child(tc, 'w:tcPr'), 'w:tcMar')
            if margins is None:
                margins = child(child(table, 'w:tblPr'), 'w:tblCellMar')
            padding = sum(number(child(margins, 'w:' + side), 'w:w', 108)
                          for side in ('left', 'right')) / 20
            return min(width, outer) - padding
        parent = parent.parentNode
    return page


def layout_equations(xml, styles_xml, page_width=415.6):
    doc, styles = minidom.parseString(xml), minidom.parseString(styles_xml)
    style_map = {s.getAttribute('w:styleId'): s for s in styles.getElementsByTagName('w:style')}

    def inherited(style_id, group, prop):
        seen = set()
        while style_id in style_map and style_id not in seen:
            seen.add(style_id)
            st = style_map[style_id]
            value = child(child(st, group), prop)
            if value is not None:
                return value
            based = child(st, 'w:basedOn')
            style_id = based.getAttribute('w:val') if based else ''
        return None

    widths = paragraph_widths(doc, page_width)
    wrapped = 0
    for p in doc.getElementsByTagName('w:p'):
        ppr = child(p, 'w:pPr')
        style = child(ppr, 'w:pStyle')
        style_id = style.getAttribute('w:val') if style else ''
        if style_id not in ('EquationNumbered', 'EquationPara'):
            continue
        displays = children(p, 'm:oMathPara')
        if not displays:
            continue  # latex 中转源码不改写为数学多行结构。
        em = number(inherited(style_id, 'w:rPr', 'w:sz'), 'w:val', 24) / 2
        indent = child(ppr, 'w:ind') or inherited(style_id, 'w:pPr', 'w:ind')
        available = container_width(p, widths[p]) - (number(indent, 'w:left') + number(indent, 'w:right')) / 20
        numbered = style_id == 'EquationNumbered'
        label = ''.join(text(n) for n in p.getElementsByTagName('w:t'))
        reserve = max(3 * em, glyph_width(label, em) + em) if numbered else 0
        budget = (available - reserve) * .82  # 给字体替换、运算符间距留下余量。
        if budget <= em:
            raise LayoutError('公式可用宽度不足，请检查页面或表格列宽')
        for display in displays:
            for math in children(display, 'm:oMath'):
                wrapped += int(wrap_math(math, budget, em))
        if numbered:
            for old in children(ppr, 'w:tabs'):
                ppr.removeChild(old)
            tabs = element(doc, 'w:tabs')
            tabs.appendChild(element(doc, 'w:tab', w_val='center', w_pos=round((available - reserve) * 10)))
            tabs.appendChild(element(doc, 'w:tab', w_val='right', w_pos=round(available * 20)))
            ppr.appendChild(tabs)
            lead = element(doc, 'w:r')
            lead.appendChild(element(doc, 'w:tab'))
            p.insertBefore(lead, displays[0])
    return doc.toxml(), wrapped
