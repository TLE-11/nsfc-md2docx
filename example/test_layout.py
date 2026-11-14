#!/usr/bin/env python3
"""可选的实际排版回归：需要 LibreOffice 和 pdfplumber，不影响转换器标准库依赖。

python3 example/test_layout.py --output-dir /tmp/nsfc-layout
输出保留 DOCX/PDF，便于人工查看；断言基于实际 PDF 字符坐标而非宽度估计器。
"""
import argparse
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from md2docx import which_soffice


def run(command):
    result = subprocess.run(command, capture_output=True, encoding='utf-8')
    if result.returncode:
        raise AssertionError(result.stdout + result.stderr)
    return result


def check_bounds(page, left=1797 / 20, right=1797 / 20, top=1440 / 20, bottom=1440 / 20):
    assert page.chars, 'PDF 中没有文字'
    bad = [c for c in page.chars if c['x0'] < left - 2 or c['x1'] > page.width - right + 2
           or c['top'] < top - 2 or c['bottom'] > page.height - bottom + 2]
    assert not bad, '文字越过版心: ' + repr([(c['text'], c['x0'], c['top'], c['bottom']) for c in bad[:5]])


def check_page(pdf, expected_terms, left, right, numbered):
    import pdfplumber
    with pdfplumber.open(pdf) as document:
        assert len(document.pages) == 1, '意外分页'
        page = document.pages[0]
        chars = page.chars
        check_bounds(page, left, right)
        terms = [c for c in chars if c['text'] == 'a']
        assert len(terms) == expected_terms, '公式项缺失或重复: %d/%d' % (len(terms), expected_terms)
        lines = sorted(set(round(c['top'] / 3) * 3 for c in terms))
        assert len(lines) > 1, '长公式没有实际换行'
        if numbered:
            labels = [c for c in chars if c['text'].isdigit() and c['x0'] > page.width - right - 30]
            assert labels, '右侧编号缺失'
            # 取最后一个编号；它应在多行公式高度内，而不是落到公式下方。
            label = max(labels, key=lambda c: c['top'])
            assert min(c['top'] for c in terms) - 3 <= label['top'] <= max(c['bottom'] for c in terms) + 3, '编号掉到公式之外'
            assert max(c['x1'] for c in terms) < label['x0'] - 8, '公式与编号重叠'
        return {'pages': len(document.pages), 'terms': len(terms), 'lines': len(lines),
                'rightmost_pt': round(max(c['x1'] for c in chars), 2),
                'right_boundary_pt': round(page.width - right, 2)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', required=True)
    args = parser.parse_args()
    try:
        import pdfplumber  # noqa: F401
    except ImportError:
        sys.exit('排版测试需要 pdfplumber；普通转换和结构回归不需要此依赖')
    soffice = which_soffice()
    if not soffice:
        sys.exit('排版测试需要 LibreOffice')
    out = Path(args.output_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='nsfc-layout-') as temp:
        repo = Path(temp) / 'repo'
        repo.mkdir()
        for path in ROOT.glob('*.py'):
            shutil.copy2(path, repo / path.name)
        reference = out / 'reference.docx'
        run([sys.executable, str(repo / 'make_reference.py'), str(reference)])
        narrow = out / 'narrow-reference.docx'
        with zipfile.ZipFile(reference) as z, zipfile.ZipFile(narrow, 'w') as target:
            for name in z.namelist():
                data = z.read(name)
                if name == 'word/document.xml':
                    data = re.sub(rb'w:w="11906"', b'w:w="8500"', data)
                target.writestr(name, data)
        polynomial = 'y=' + '+'.join('a_{%d}x^{%d}' % (i, i) for i in range(1, 21))
        fractions = 'y=' + '+'.join(r'\frac{a_{%d}}{\sqrt{x_{%d}+1}}' % (i, i) for i in range(1, 13))
        # 渲染夹具只用西文标题，隔离本机中文字体缺失，不掩盖数学布局问题。
        cases = [
            ('default', '# Expression\n\n$$x=1$$\n\n$$' + polynomial + '$$\n', reference, 'all', 20),
            ('unnumbered', '# Expression\n\n$$' + polynomial + '$$\n', reference, 'none', 20),
            ('narrow', '# Reduced width\n\n$$' + polynomial + '$$\n', narrow, 'all', 20),
            ('fractions', '# Quotients\n\n$$' + fractions + '$$\n', reference, 'all', 12),
            ('table', '| Model | Expr |\n|---|---|\n| A | $$' + polynomial + '$$ |\n', reference, 'all', 20),
            ('aligned', '# Aligned\n\n$$\\begin{aligned}' + polynomial + r'\\z&=1\end{aligned}$$' + '\n', reference, 'all', 20),
        ]
        results = {}

        def convert_render(name, content, ref=reference, numbering='all', mode='omml'):
            src, docx = out / (name + '.md'), out / (name + '.docx')
            src.write_text(content, encoding='utf-8')
            pdf = out / (name + '.pdf')
            if pdf.exists():
                pdf.unlink()
            run([sys.executable, str(repo / 'md2docx.py'), str(src), '-o', str(docx),
                 '--reference', str(ref), '--number', numbering, '--math-mode', mode, '--pdf'])
            assert pdf.is_file(), 'LibreOffice 没有生成 PDF'
            return pdf

        for name, content, ref, numbering, terms in cases:
            pdf = convert_render(name, content, ref, numbering)
            results[name] = check_page(pdf, terms, 1797 / 20, 1797 / 20, numbering == 'all')
            if name == 'table':
                # 公式所在第二列的左边界；防止只在整页内、却越过自身单元格。
                with pdfplumber.open(pdf) as document:
                    page = document.pages[0]
                    formula_chars = [c for c in page.chars if c['text'] == 'a']
                    assert min(c['x0'] for c in formula_chars) >= page.width / 2 - 2, '公式越过单元格左边界'
            print('PASS', name, results[name], flush=True)

        # 高但可容纳的公式，以及正文占满当前页后自动移到下一页的场景。
        tall = 'y=' + '+'.join('a_{%d}x^{%d}' % (i, i) for i in range(1, 101))
        pdf = convert_render('tall_supported', '$$' + tall + '$$\n')
        results['tall_supported'] = check_page(pdf, 100, 1797 / 20, 1797 / 20, True)
        pdf = convert_render('page_flow', 'Intro.\n\n' * 30 + '$$' + tall + '$$\n')
        with pdfplumber.open(pdf) as document:
            assert len(document.pages) == 2, '公式应整体移到下一页'
            for page in document.pages:
                check_bounds(page)
            terms = [sum(c['text'] == 'a' for c in page.chars) for page in document.pages]
            assert terms == [0, 100], '跨页时公式被裁切或拆散'
            results['page_flow'] = {'pages': 2, 'terms_by_page': terms}

        for mode in ('omml', 'latex'):
            name = 'boxed_link_' + mode
            pdf = convert_render(name, '[Source](https://example.com/source) $$y=\\boxed{x+1}^2$$ Tail 123.\n', mode=mode)
            with pdfplumber.open(pdf) as document:
                assert len(document.pages) == 1
                page = document.pages[0]
                check_bounds(page)
                assert 'Tail 123.' in page.extract_text(), '公式后的尾文丢失'
                assert any(link.get('uri') == 'https://example.com/source' for link in page.hyperlinks), '链接丢失'
                results[name] = {'pages': 1, 'tail_and_link': 'preserved'}

        # 200 项已在旧版实测裁切；新版本必须拒绝，且不能发布新的 DOCX/PDF。
        src, docx, pdf = (out / ('too_tall.' + suffix) for suffix in ('md', 'docx', 'pdf'))
        src.write_text('$$y=' + '+'.join('a_{%d}x^{%d}' % (i, i) for i in range(1, 201)) + '$$\n')
        docx.write_bytes(b'previous DOCX')
        pdf.write_bytes(b'previous PDF')
        result = subprocess.run([sys.executable, str(repo / 'md2docx.py'), str(src), '-o', str(docx), '--pdf'],
                                capture_output=True, encoding='utf-8')
        assert result.returncode != 0 and '公式高度超过' in result.stdout + result.stderr
        assert docx.read_bytes() == b'previous DOCX' and pdf.read_bytes() == b'previous PDF'
        # 哨兵用于失败原子性断言，不作为可打开的文档交付。
        docx.unlink()
        pdf.unlink()
        results['too_tall'] = {'rejected': True, 'previous_outputs_preserved': True}
        for name in ('tall_supported', 'page_flow', 'boxed_link_omml', 'boxed_link_latex', 'too_tall'):
            print('PASS', name, results[name], flush=True)
        (out / 'results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0


if __name__ == '__main__':
    sys.exit(main())
