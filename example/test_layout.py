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


def check_page(pdf, expected_terms, left, right, numbered):
    import pdfplumber
    with pdfplumber.open(pdf) as document:
        assert len(document.pages) == 1, '意外分页'
        page = document.pages[0]
        chars = page.chars
        assert chars, 'PDF 中没有文字'
        bad = [c for c in chars if c['x0'] < left - 2 or c['x1'] > page.width - right + 2]
        assert not bad, '文字越过版心: ' + repr([(c['text'], c['x0'], c['x1']) for c in bad[:5]])
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
        for name, content, ref, numbering, terms in cases:
            src, docx = out / (name + '.md'), out / (name + '.docx')
            src.write_text(content, encoding='utf-8')
            run([sys.executable, str(repo / 'md2docx.py'), str(src), '-o', str(docx),
                 '--reference', str(ref), '--number', numbering])
            profile = Path(temp) / ('lo-' + name)
            pdf = out / (name + '.pdf')
            if pdf.exists():
                pdf.unlink()
            run([soffice, '-env:UserInstallation=' + profile.as_uri(), '--headless',
                 '--convert-to', 'pdf', '--outdir', str(out), str(docx)])
            assert pdf.is_file(), 'LibreOffice 没有生成 PDF'
            results[name] = check_page(pdf, terms, 1797 / 20, 1797 / 20, numbering == 'all')
            if name == 'table':
                # 公式所在第二列的左边界；防止只在整页内、却越过自身单元格。
                with pdfplumber.open(pdf) as document:
                    page = document.pages[0]
                    formula_chars = [c for c in page.chars if c['text'] == 'a']
                    assert min(c['x0'] for c in formula_chars) >= page.width / 2 - 2, '公式越过单元格左边界'
            print('PASS', name, results[name], flush=True)
        (out / 'results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0


if __name__ == '__main__':
    sys.exit(main())
