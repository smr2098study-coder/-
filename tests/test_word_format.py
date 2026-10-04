"""Regression tests for formatting and non-destructive pagination repair."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile
from lxml import etree as E

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'plugins/shimen-format/skills/shimen-word-format/scripts/word_format.py'
spec = importlib.util.spec_from_file_location('word_format', SCRIPT)
fmt = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fmt)


class WordFormatTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def build_sample(self):
        from PIL import Image
        Image.new('RGB', (400, 180), 'white').save(self.path / 'figure.png')
        blocks = []
        for level, text in enumerate(['一、背景', '（一）目的', '1、方法', '（1）步骤', '①范围'], 1):
            blocks.append({'type': 'heading', 'level': level, 'text': text})
            blocks.append({'type': 'paragraph', 'text': '测试正文 English 2026，验证字体与段落格式。'})
        blocks.extend([
            {'type': 'table', 'caption': '表1 汇总', 'rows': [['指标', 'Value'], ['样本', '20']]},
            {'type': 'image', 'path': 'figure.png', 'caption': '图1 流程'}])
        source = self.path / 'input.json'
        source.write_text(json.dumps({'title': '格式验收', 'blocks': blocks}, ensure_ascii=False), encoding='utf-8')
        target = self.path / 'sample.docx'
        fmt.build(source, target)
        return target

    def test_all_profiles_and_image_caption(self):
        path = self.build_sample()
        result = fmt.audit(path)
        self.assertEqual(result['errors'], 0, result['issues'])
        self.assertEqual(result['warnings'], 0, result['issues'])
        pkg = fmt.Package(path)
        for p in pkg.paragraphs:
            if pkg.role(p) == 'body':
                self.assertFalse(fmt.enabled(pkg.value(p, 'keepNext')))
            if pkg.role(p).startswith('h'):
                self.assertTrue(fmt.enabled(pkg.value(p, 'keepNext')))
        images = [p for p in pkg.paragraphs if pkg.role(p) == 'image']
        self.assertEqual(len(images), 1)
        self.assertTrue(fmt.enabled(pkg.value(images[0], 'keepNext')))
        captions = [p for p in pkg.paragraphs if pkg.role(p) == 'caption']
        self.assertTrue(fmt.enabled(pkg.value(captions[0], 'keepNext')))
        self.assertFalse(fmt.enabled(pkg.value(captions[1], 'keepNext')))

    def test_inherited_flags_are_detected_and_overridden(self):
        from docx import Document
        from docx.oxml import OxmlElement
        doc = Document()
        normal = doc.styles['Normal']
        normal.paragraph_format.keep_with_next = True
        normal.paragraph_format.keep_together = True
        doc.add_heading('一、标题', 1)
        doc.add_paragraph('这是足够清晰的普通正文，测试继承的分页设置。')
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = '表内文字'
        doc.add_paragraph('表1 示例', 'Caption')
        # Embedded field must survive the narrow XML repair.
        p = doc.add_paragraph('域内容所在正文')
        r = p.add_run()
        field = OxmlElement('w:fldChar')
        field.set(fmt.tag('fldCharType'), 'begin')
        r._r.append(field)
        path = self.path / 'inherited.docx'
        doc.save(path)
        before = fmt.Package(path)
        self.assertTrue(fmt.enabled(before.value(before.paragraphs[1], 'keepNext')))
        self.assertTrue(any(x['code'] == 'body_pagination' for x in fmt.audit(path)['issues']))
        target = self.path / 'repaired.docx'
        fmt.repair(path, target)
        after = fmt.Package(target)
        self.assertFalse(fmt.enabled(after.value(after.paragraphs[1], 'keepNext')))
        self.assertFalse(fmt.enabled(after.value(after.paragraphs[1], 'keepLines')))
        self.assertTrue(fmt.enabled(after.value(after.paragraphs[0], 'keepNext')))
        self.assertEqual(before.doc.xpath('//w:t/text()', namespaces=fmt.NS), after.doc.xpath('//w:t/text()', namespaces=fmt.NS))
        for name, part in before.parts.items():
            if name != 'word/document.xml':
                self.assertEqual(part, after.parts[name])
        for index in (0, 2, 3, 4):
            self.assertEqual(E.tostring(before.paragraphs[index]), E.tostring(after.paragraphs[index]))
        with self.assertRaises(FileExistsError):
            fmt.repair(path, target)

    def test_style_bullets_and_cancellation(self):
        from docx import Document
        doc = Document()
        p = doc.add_paragraph('项目一', 'List Bullet')
        p2 = doc.add_paragraph('取消继承项目符号', 'List Bullet')
        fmt.sub(fmt.sub(p2._p.get_or_add_pPr(), 'numPr'), 'numId', val='0')
        path = self.path / 'bullets.docx'
        doc.save(path)
        pkg = fmt.Package(path)
        self.assertEqual(pkg.number_format(pkg.paragraphs[0]), 'bullet')
        self.assertIsNone(pkg.number_format(pkg.paragraphs[1]))

    def test_direct_overrides_are_not_hidden_by_style(self):
        from docx import Document
        from docx.shared import Pt
        path = self.build_sample()
        doc = Document(path)
        p = doc.paragraphs[2]
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.name = 'Arial'
        p.paragraph_format.keep_with_next = True
        mutated = self.path / 'mutated.docx'
        doc.save(mutated)
        codes = {x['code'] for x in fmt.audit(mutated)['issues'] if x.get('paragraph') == 2}
        self.assertTrue({'font_size', 'latin_font', 'body_pagination'}.issubset(codes), codes)


if __name__ == '__main__':
    unittest.main()
