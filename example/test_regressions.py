#!/usr/bin/env python3
"""真实 pandoc 转换回归；独立检查产物，而非只匹配「全部检查通过」。"""
import re
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
import importlib.util
from unittest import mock
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).with_name('fixtures')
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
W = '{%s}' % NS['w']


class ConversionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix='nsfc-regression-')
        cls.work = Path(cls.temp.name)
        # 缓存模板和产物均留在临时副本中，不污染工作区。
        cls.repo = cls.work / 'repo'
        cls.repo.mkdir()
        for p in ROOT.glob('*.py'):
            shutil.copy2(p, cls.repo / p.name)

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def convert(self, name, mode='omml', text=None, success=True, number='all', reference=None):
        src = FIXTURES / (name + '.md')
        if text is not None:
            src = self.work / (name + '.md')
            src.write_text(text, encoding='utf-8')
        out = self.work / ('%s-%s-%s.docx' % (name, mode, number))
        command = [sys.executable, str(self.repo / 'md2docx.py'), str(src), '-o', str(out),
                   '--assets', str(FIXTURES), '--math-mode', mode, '--number', number]
        if reference:
            command += ['--reference', str(reference)]
        result = subprocess.run(command,
            capture_output=True, text=True)
        log = result.stdout + result.stderr
        if success:
            self.assertEqual(result.returncode, 0, log)
        else:
            self.assertNotEqual(result.returncode, 0, log)
        with zipfile.ZipFile(out) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        return ET.fromstring(parts['word/document.xml']), parts, log

    def test_literal_ampersand_and_alignment(self):
        tree, _, _ = self.convert('literal_amp', text=(
            r'公司名称 $R\&D$ 和 $\text{R\&D}$。' + '\n\n' +
            r'$$\begin{aligned}R\&D&=x\\ a&=b\end{aligned}$$'))
        math = ''.join(e.text or '' for e in tree.findall('.//m:t', NS))
        self.assertEqual(math.count('R&D'), 3)
        self.assertEqual(len(tree.findall('.//m:aln', NS)), 2)

    def test_range_reference_targets(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                tree, _, _ = self.convert('range_refs', mode)
                self.assertIn('式（2）—（3）', ''.join(tree.itertext()))
                refs = [e.get(W + 'instr').strip() for e in tree.findall('.//w:fldSimple', NS)
                        if ' REF ' in e.get(W + 'instr', '')]
                self.assertEqual(refs, ['REF eq_2 \\h', 'REF eq_3 \\h'])

    def test_image_and_formula_preserve_drawing(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                tree, _, _ = self.convert('image_and_math_latex', mode)
                self.assertEqual(len(tree.findall('.//w:drawing', NS)), 1)

    def test_wiki_subdirectory(self):
        tree, _, _ = self.convert('wiki_subpath')
        self.assertEqual(len(tree.findall('.//w:drawing', NS)), 1)
        self.assertNotIn('待插入', ''.join(tree.itertext()))
        # verify.py 的原有独立调用方式仍可使用，无需提供预处理中间文件。
        result = subprocess.run([sys.executable, str(self.repo / 'verify.py'),
                                 str(FIXTURES / 'wiki_subpath.md'),
                                 str(self.work / 'wiki_subpath-omml-all.docx')],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_display_math_in_table(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                tree, _, _ = self.convert('table_display', mode)
                self.assertEqual(len(tree.findall('.//w:tbl', NS)), 1)
                cells = tree.findall('.//w:tc', NS)
                self.assertEqual(len(cells), 4)
                self.assertIn('（1）', ''.join(cells[-1].itertext()))
                self.assertNotIn('⟦', ''.join(tree.itertext()))

    def test_quote_formula(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                tree, _, _ = self.convert('quote_display', mode)
                if mode == 'omml':
                    self.assertEqual(''.join(e.text or '' for e in tree.findall('.//m:t', NS)), 'x+1=2')
                else:
                    self.assertIn('$$x+1=2$$', ''.join(tree.itertext()))

    def test_display_after_inline_code_or_comment(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                tree, _, _ = self.convert('inline_code_math', mode)
                if mode == 'omml':
                    self.assertEqual(''.join(e.text or '' for e in tree.findall('.//m:t', NS)),
                                     'x=1y=2z=3w=4')
                else:
                    for expression in ('x=1', 'y=2', 'z=3', 'w=4'):
                        self.assertIn('$$' + expression + '$$', ''.join(tree.itertext()))
                for num in range(1, 5):
                    self.assertIn('（%d）' % num, ''.join(tree.itertext()))

    def test_footnote_and_link_content(self):
        tree, parts, _ = self.convert('footnote')
        self.assertEqual(len(tree.findall('.//w:footnoteReference', NS)), 1)
        notes = ET.fromstring(parts['word/footnotes.xml'])
        self.assertIn('注释内容', ''.join(notes.itertext()))
        _, parts, _ = self.convert('link_destination')
        rels = ET.fromstring(parts['word/_rels/document.xml.rels'])
        self.assertTrue(any(e.get('Target') == 'https://example.com/中文路径' for e in rels))

    def test_comments_do_not_consume_equation_numbers(self):
        tree, _, _ = self.convert('comment_math')
        self.assertIn('y=1（1）', ''.join(tree.itertext()))
        self.assertEqual(len(tree.findall('.//m:oMath', NS)), 1)
        tree, _, _ = self.convert('tex_comment_tag', number='tag', text=(
            '$$\nx=1 % \\tag{9}\n$$\n'))
        self.assertFalse(tree.findall('.//w:bookmarkStart', NS))

    def test_indented_code_remains_literal(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                tree, _, _ = self.convert('indented_code', mode)
                self.assertIn('print("$x$")', ''.join(tree.itertext()))
                self.assertEqual(len(tree.findall('.//m:oMath', NS)), 0)

    def test_list_math_and_tex_comments(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                tree, _, _ = self.convert('list_math', mode, text=(
                    '1. 推导\n\n    $$\n    x=1 % 注释不会吞掉下一行\n'
                    '    +2\n    $$\n\n2. 结论\n'))
                if mode == 'omml':
                    self.assertEqual(''.join(e.text or '' for e in tree.findall('.//m:t', NS)), 'x=1+2')
                self.assertIn('（1）', ''.join(tree.itertext()))

    def test_long_equation_content_is_preserved(self):
        tree, _, _ = self.convert('long_equation')
        formulas = tree.findall('.//m:oMath', NS)
        self.assertEqual(len(formulas), 2)
        formula = formulas[-1]
        self.assertEqual(len(formula.findall('.//m:sSub', NS)), 20)
        self.assertEqual(len(formula.findall('.//m:sSup', NS)), 20)
        self.assertIsNone(formulas[0].find('m:eqArr', NS))
        rows = formula.findall('m:eqArr/m:e', NS)
        self.assertGreater(len(rows), 1)
        # 比对完整字符序列，换行不得增删运算符或改变项的顺序。
        expected = 'y=' + '+'.join('a%dx%d' % (i, i) for i in range(1, 21))
        self.assertEqual(''.join(t.text or '' for t in formula.findall('.//m:t', NS)), expected)
        for row in rows[1:]:
            self.assertTrue(''.join(row.itertext()).startswith('+'))
        self.assertEqual(len([b for b in tree.findall('.//w:bookmarkStart', NS)
                              if b.get(W + 'name', '').startswith('eq_')]), 2)
        self.assertEqual(len(tree.findall('.//w:tbl', NS)), 0)

    def test_narrow_page_and_table_equation_width(self):
        _, parts, _ = self.convert('long_equation')
        reference = self.work / 'narrow-reference.docx'
        parts['word/document.xml'] = re.sub(rb'w:w="11906"', b'w:w="8500"', parts['word/document.xml'])
        with zipfile.ZipFile(reference, 'w') as z:
            for name, data in parts.items():
                z.writestr(name, data)
        narrow, _, _ = self.convert('long_equation', reference=reference)
        normal, _, _ = self.convert('long_equation')
        self.assertGreater(len(narrow.findall('.//m:eqArr/m:e', NS)),
                           len(normal.findall('.//m:eqArr/m:e', NS)))
        expr = '+'.join('a_{%d}x^{%d}' % (i, i) for i in range(1, 9))
        table, _, _ = self.convert('long_table', text='| Model | Formula |\n|---|---|\n| A | $$y=' + expr + '$$ |\n')
        self.assertGreater(len(table.findall('.//w:tc//m:eqArr/m:e', NS)), 1)
        tabs = table.findall('.//w:tc//w:tab[@w:val="right"]', NS)
        self.assertTrue(tabs)
        self.assertLess(int(tabs[0].get(W + 'pos')), 8312)

    def test_wrapping_preserves_fractions_and_inline_math(self):
        expr = '+'.join(r'\frac{a_{%d}}{\sqrt{x_{%d}+1}}' % (i, i) for i in range(1, 13))
        tree, _, _ = self.convert('long_fractions', text='Inline $x+1$.\n\n$$y=' + expr + '$$\n', number='none')
        self.assertEqual(len(tree.findall('.//m:f', NS)), 12)
        self.assertEqual(len(tree.findall('.//m:rad', NS)), 12)
        self.assertGreater(len(tree.findall('.//m:eqArr/m:e', NS)), 1)
        self.assertFalse(tree.findall('.//m:f//m:eqArr', NS))
        self.assertFalse(tree.findall('.//m:rad//m:eqArr', NS))
        self.assertIsNone(tree.findall('.//m:oMath', NS)[0].find('m:eqArr', NS))

    def test_unbreakable_equation_fails_without_overwriting_output(self):
        src = self.work / 'unbreakable.md'
        src.write_text('$$\\text{' + 'M' * 200 + '}$$\n', encoding='utf-8')
        out = self.work / 'existing.docx'
        out.write_bytes(b'previous result')
        result = subprocess.run([sys.executable, str(self.repo / 'md2docx.py'),
                                 str(src), '-o', str(out)], capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('无法安全自动换行', result.stdout + result.stderr)
        self.assertEqual(out.read_bytes(), b'previous result')

    def test_boxed_group_and_scripts_keep_native_structure(self):
        tree, _, _ = self.convert('boxed_scripts', text=(
            r'Inline $\boxed{a+b}_{i}$.' + '\n\n' + r'$$y=\boxed{x+1}^2$$'))
        self.assertEqual(len(tree.findall('.//m:borderBox', NS)), 2)
        self.assertEqual(''.join(tree.find('.//m:sSup/m:e/m:borderBox/m:e', NS).itertext()), 'x+1')
        self.assertEqual(''.join(tree.find('.//m:sSub/m:e/m:borderBox/m:e', NS).itertext()), 'a+b')
        self.assertFalse(tree.findall('.//w:pBdr', NS))

    def test_unnumbered_box_preserves_tail_and_image(self):
        for number in ('none', 'tag'):
            with self.subTest(number=number):
                tree, _, _ = self.convert('box_tail', number=number, text=(
                    r'$$\boxed{x=1}$$ KEEP_THIS_TEXT 123. ' +
                    '![](<%s>)\n' % (FIXTURES / 'assets/plot.png')))
                self.assertIn('KEEP_THIS_TEXT 123.', ''.join(tree.itertext()))
                self.assertEqual(len(tree.findall('.//w:drawing', NS)), 1)
                self.assertEqual(len(tree.findall('.//m:borderBox', NS)), 1)

    def test_latex_split_preserves_links_and_reference_fields(self):
        tree, _, _ = self.convert('mixed_links', 'latex', text=(
            '[前文](https://example.com/before) 见式（7） $$x=1 \\tag{7}$$ '
            '[后文](https://example.com/after)\n'))
        self.assertEqual(len(tree.findall('.//w:hyperlink', NS)), 2)
        self.assertTrue(any(' REF eq_1 ' in f.get(W + 'instr', '')
                            for f in tree.findall('.//w:fldSimple', NS)))

    def test_failed_verification_preserves_old_output(self):
        self.convert('atomic_output', text='Original good document.\n')
        output = self.work / 'atomic_output-omml-all.docx'
        previous = output.read_bytes()
        _, _, log = self.convert('atomic_output', text=r'Unsupported $\ce{H2O}$.', success=False)
        self.assertIn('行内公式数量不符', log)
        self.assertEqual(output.read_bytes(), previous)
        output.unlink()
        src = self.work / 'atomic_output.md'
        result = subprocess.run([sys.executable, str(self.repo / 'md2docx.py'),
                                 str(src), '-o', str(output)], capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(output.exists())
        self.assertFalse(list(self.work.glob('.md2docx-*.docx')))

    def test_tall_equations_fail_without_publishing(self):
        expressions = [
            'y=' + '+'.join('a_{%d}x^{%d}' % (i, i) for i in range(1, 201)),
            r'\begin{aligned}' + r'\\'.join('x=1' for _ in range(80)) + r'\end{aligned}',
        ]
        for expression in expressions:
            with self.subTest(expression=expression[:30]):
                src, out = self.work / 'tall.md', self.work / 'tall.docx'
                src.write_text('$$' + expression + '$$\n', encoding='utf-8')
                out.write_bytes(b'previous result')
                result = subprocess.run([sys.executable, str(self.repo / 'md2docx.py'),
                                         str(src), '-o', str(out)], capture_output=True, text=True)
                self.assertNotEqual(result.returncode, 0)
                self.assertIn('公式高度超过', result.stdout + result.stderr)
                self.assertEqual(out.read_bytes(), b'previous result')

    def test_pdf_export_never_reuses_stale_output(self):
        spec = importlib.util.spec_from_file_location('pdf_cli', self.repo / 'md2docx.py')
        cli = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(cli)
        docx, pdf = self.work / 'preview.docx', self.work / 'preview.pdf'
        docx.write_bytes(b'docx input')
        pdf.write_bytes(b'old preview')
        # 故障注入：退出失败、成功但无输出、零字节输出均不得引用旧 PDF。
        for mode in ('failed', 'missing', 'empty', 'success'):
            with self.subTest(mode=mode):
                def render(command, **kwargs):
                    folder = Path(command[command.index('--outdir') + 1])
                    self.assertNotEqual(folder, pdf.parent)
                    if mode in ('empty', 'success'):
                        (folder / pdf.name).write_bytes(b'new preview' if mode == 'success' else b'')
                    return subprocess.CompletedProcess(command, 1 if mode == 'failed' else 0)
                with mock.patch.object(cli.subprocess, 'run', side_effect=render):
                    result = cli.export_pdf('test-soffice', docx)
                if mode == 'success':
                    self.assertEqual(result, str(pdf.resolve()))
                    self.assertEqual(pdf.read_bytes(), b'new preview')
                else:
                    self.assertIsNone(result)
                    self.assertEqual(pdf.read_bytes(), b'old preview')
                self.assertFalse(list(self.work.glob('.md2docx-pdf-*')))

    def test_percentage_table_uses_its_own_width(self):
        from xml.dom import minidom
        sys.path.insert(0, str(self.repo))
        try:
            from equation_layout import container_width
            # 表格占版心一半、两等宽列；旧首选列宽不能越过缩窄后的网格。
            doc = minidom.parseString(
                '<w:tbl xmlns:w="' + NS['w'] + '"><w:tblPr>'
                '<w:tblW w:type="pct" w:w="2500"/></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
                '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:p/></w:tc>'
                '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>')
            self.assertAlmostEqual(container_width(doc.getElementsByTagName('w:p')[1], 400),
                                   100 - 10.8)
        finally:
            sys.path.pop(0)

    def test_math_and_references_in_footnotes(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                _, parts, _ = self.convert('note_math', mode, text=(
                    '正文[^1]。\n\n$$x=1 \\tag{7}$$\n\n'
                    '[^1]: 见式（7），其中 $R\\&D$ 表示名称。\n'))
                notes = ET.fromstring(parts['word/footnotes.xml'])
                self.assertTrue(any(' REF eq_1 ' in f.get(W + 'instr', '')
                                    for f in notes.findall('.//w:fldSimple', NS)))

    def test_equation_numbers_across_document_and_notes(self):
        for mode in ('omml', 'latex'):
            with self.subTest(mode=mode):
                tree, parts, _ = self.convert('note_display', mode, text=(
                    '正文[^1]。\n\n$$x=1$$\n\n[^1]: $$y=2$$\n\n$$z=3$$\n'))
                notes = ET.fromstring(parts['word/footnotes.xml'])
                self.assertIn('（1）', ''.join(tree.itertext()))
                self.assertIn('（3）', ''.join(tree.itertext()))
                self.assertIn('（2）', ''.join(notes.itertext()))

    def test_verifier_rejects_swapped_number_caches(self):
        _, parts, _ = self.convert('range_refs')
        # 数字集合仍连续，但交换前两个公式的缓存值必须被拒绝。
        xml = parts['word/document.xml']
        pattern = rb'(<w:fldSimple w:instr=" SEQ [^"]+"><w:r><w:t>)([12])(</w:t>)'
        xml, count = re.subn(pattern, lambda m: m[1] + (b'2' if m[2] == b'1' else b'1') + m[3], xml)
        self.assertEqual(count, 2)
        parts['word/document.xml'] = xml
        damaged = self.work / 'swapped-numbers.docx'
        with zipfile.ZipFile(damaged, 'w') as z:
            for name, data in parts.items():
                z.writestr(name, data)
        result = subprocess.run([sys.executable, str(self.repo / 'verify.py'),
                                 str(FIXTURES / 'range_refs.md'), str(damaged)],
                                capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('公式书签与编号不符', result.stdout)

    def test_image_caption_and_inline_alt(self):
        for inline in (False, True):
            text = '![示意图](<%s>)' % (FIXTURES / 'assets/plot.png')
            if inline:
                text = '前文' + text + '后文。'
            tree, _, _ = self.convert('image_alt', text=text)
            self.assertEqual(len(tree.findall('.//w:drawing', NS)), 1)

    def test_verifier_rejects_deleted_media_and_symbols(self):
        # 故意破坏产物，证明校验能检测错误，而不只是转换器输出能自洽。
        for name, pattern, expected in (
                ('literal_amp', rb'<m:t>&amp;</m:t>', '字面量 & 丢失'),
                ('image_and_math_latex', rb'<w:drawing>.*?</w:drawing>', '图片数量不符'),
                ('footnote', '注释内容'.encode(), '中文正文不一致')):
            with self.subTest(case=name):
                _, parts, _ = self.convert(name)
                part = 'word/footnotes.xml' if name == 'footnote' else 'word/document.xml'
                parts[part], count = re.subn(pattern, b'', parts[part], count=1, flags=re.S)
                self.assertEqual(count, 1)
                damaged = self.work / ('damaged-' + name + '.docx')
                with zipfile.ZipFile(damaged, 'w') as z:
                    for key, value in parts.items():
                        z.writestr(key, value)
                result = subprocess.run([sys.executable, str(self.repo / 'verify.py'),
                                         str(FIXTURES / (name + '.md')), str(damaged)],
                                        capture_output=True, text=True)
                self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
                self.assertIn(expected, result.stdout)

    def test_verifier_rejects_wrong_existing_reference_target(self):
        _, parts, _ = self.convert('range_refs')
        parts['word/document.xml'] = parts['word/document.xml'].replace(
            b' REF eq_2 ', b' REF eq_3 ', 1)
        damaged = self.work / 'wrong-reference.docx'
        with zipfile.ZipFile(damaged, 'w') as z:
            for key, value in parts.items():
                z.writestr(key, value)
        prepared = self.work / 'range_refs.pre.md'
        subprocess.run([sys.executable, str(self.repo / 'pre.py'),
                        str(FIXTURES / 'range_refs.md'), str(prepared)],
                       check=True, capture_output=True, text=True)
        result = subprocess.run([sys.executable, str(self.repo / 'verify.py'),
                                 str(FIXTURES / 'range_refs.md'), str(damaged),
                                 '--prepared-md', str(prepared)], capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('REF 引用目标', result.stdout)

    def test_setext_heading(self):
        tree, _, _ = self.convert('setext_heading')
        headings = [p for p in tree.findall('.//w:p', NS)
                    if p.find('./w:pPr/w:pStyle[@w:val="Heading2"]', NS) is not None]
        self.assertEqual([''.join(p.itertext()) for p in headings], ['研究背景'])

    def test_unsupported_inline_math_fails(self):
        src = FIXTURES / 'unsupported_inline.md'
        out = self.work / 'unsupported-inline.docx'
        result = subprocess.run([sys.executable, str(self.repo / 'md2docx.py'),
                                 str(src), '-o', str(out)], capture_output=True, text=True)
        log = result.stdout + result.stderr
        self.assertNotEqual(result.returncode, 0, log)
        self.assertFalse(out.exists())
        self.assertIn('行内公式数量不符', log)
        self.assertNotIn('全部检查通过', log)
        tree, _, _ = self.convert('unsupported_inline', 'latex')
        self.assertIn(r'\ce{H2O}', ''.join(tree.itertext()))


if __name__ == '__main__':
    unittest.main(verbosity=2)
