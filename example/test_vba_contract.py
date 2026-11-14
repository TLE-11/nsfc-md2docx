#!/usr/bin/env python3
"""VBA 静态契约回归；Word 行为测试另见 test_word_macros.ps1，不混同实机验证。"""
from pathlib import Path
import re
import unittest

SOURCE = Path(__file__).resolve().parents[1] / 'mathtype_axmath.bas'


def procedures(source):
    return {m[1].lower(): m[2] for m in re.finditer(
        r'(?im)^(?:Public |Private )?(?:Sub|Function) (\w+)[^\n]*\n(.*?)^End (?:Sub|Function)',
        source, re.S)}


def code_only(body):
    return '\n'.join(re.sub(r'"(?:[^"]|"")*"', '""', line).split("'")[0]
                     for line in body.splitlines())


class MacroContracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = SOURCE.read_text(encoding='utf-8')
        cls.functions = procedures(cls.source)

    def reachable(self, name):
        todo, seen = [name.lower()], set()
        while todo:
            current = todo.pop()
            if current in seen:
                continue
            seen.add(current)
            code = code_only(self.functions[current]).lower()
            todo.extend(set(re.findall(r'\b\w+\b', code)) & self.functions.keys() - seen)
        return seen

    def test_reporting_counting_and_cleanup_cannot_execute_macros(self):
        for entry in ('ProbeEquationMacros', 'CountLatexEquations',
                      'SelectLatexEquations', 'CleanupAfterConvert'):
            with self.subTest(entry=entry):
                for name in self.reachable(entry):
                    self.assertNotRegex(code_only(self.functions[name]),
                                        r'(?i)Application\s*\.\s*Run|CallByName|WordBasic\s*\.\s*Call')

    def test_conversion_execution_is_after_confirmation(self):
        callers = [name for name, body in self.functions.items()
                   if re.search(r'(?i)Application\s*\.\s*Run', code_only(body))]
        self.assertEqual(callers, ['convertpending'])
        gate = code_only(self.functions['convertwith'])
        confirm = re.search(r'If MsgBox\([\s\S]+?<> vbYes Then Exit Sub', gate, re.I)
        self.assertIsNotNone(confirm)
        self.assertLess(confirm.end(), gate.lower().index('convertpending'))
        for entry in ('ConvertLatexToMathType', 'ConvertLatexToAxMath'):
            self.assertIn('convertwith', self.reachable(entry))

    def test_all_source_operations_share_all_story_traversal(self):
        for entry in ('CountLatexEquations', 'SelectLatexEquations', 'ConvertPending', 'CleanupAfterConvert'):
            self.assertIn('styledranges', self.reachable(entry))
        search = code_only(self.functions['styledranges'])
        self.assertRegex(search, r'For Each \w+ In doc\.StoryRanges')
        self.assertIn('.NextStoryRange', search)
        self.assertNotIn('ActiveDocument.Content', code_only(self.source))
        # 失败不通过删除样式伪装成已完成；有界快照遍历避免无效插件死循环。
        convert = code_only(self.functions['convertpending'])
        self.assertNotRegex(convert, r'(?i)\.Style\s*=')
        self.assertIn('For i = sources.Count To 1 Step -1', convert)


if __name__ == '__main__':
    unittest.main(verbosity=2)
