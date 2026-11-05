"""Song-only title, dated teacher byline, and explicitly supplied notation."""
import json
import re
import shutil
import unittest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import test_production as fixtures
from test_production import maker, sample_data, SCRIPTS
from audit_docx import inspect_docx


class FrontMatterTests(unittest.TestCase):
    setUp = fixtures.ProductionTests.setUp
    tearDown = fixtures.ProductionTests.tearDown

    def data(self):
        data = sample_data()
        data['blocks'].pop(0)
        data['front_matter'] = {'song': '合成歌名', 'teachers': '示例教师甲&示例教师乙', 'date': '2099.01.01',
                                'notation': ['本例的*是斜体标记，++表示动作重复。', '本例〔〕内为短说明。']}
        return data

    def create(self, data):
        source = self.root/'content.json'
        source.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
        output = self.root/'output.docx'
        maker.make_handout(source, output)
        return source, output

    def test_front_matter_and_only_explicit_notes(self):
        source, output = self.create(self.data())
        doc = Document(output)
        self.assertEqual(doc.paragraphs[0].text, '合成歌名')
        self.assertEqual(doc.paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(doc.paragraphs[1].text, '--2099.01.01 by示例教师甲&示例教师乙')
        self.assertEqual(doc.paragraphs[1].alignment, WD_ALIGN_PARAGRAPH.RIGHT)
        self.assertEqual(doc.paragraphs[2].text, self.data()['front_matter']['notation'][0])
        self.assertFalse(inspect_docx(output, source=source)['errors'])

    def test_wrong_alignment_date_or_extra_artist_is_detected(self):
        source, output = self.create(self.data())
        original = output.read_bytes()
        for change in ('alignment', 'date', 'artist', 'notation'):
            with self.subTest(change=change):
                output.write_bytes(original)
                doc = Document(output)
                if change == 'alignment':
                    doc.paragraphs[1].alignment = WD_ALIGN_PARAGRAPH.LEFT
                elif change == 'date':
                    doc.paragraphs[1].text = '--2099.01.02 by示例教师甲&示例教师乙'
                elif change == 'artist':
                    doc.paragraphs[0].text += ' 某歌手'
                else:
                    doc.paragraphs[2].text = '被旧说明替换'
                doc.save(output)
                self.assertTrue(inspect_docx(output, source=source)['errors'])

    def test_invalid_or_duplicate_front_matter_does_not_write(self):
        for key, value in [('date', '2026.02.30'), ('date', '2026.1.1'), ('teachers', ''),
                           ('song', ''), ('notation', '不是数组'), ('artist', '不应另加歌手')]:
            with self.subTest(key=key, value=value):
                data = self.data(); data['front_matter'][key] = value
                with self.assertRaises(ValueError):
                    self.create(data)
                self.assertFalse((self.root/'output.docx').exists())
        data = self.data(); data['blocks'].insert(0, {'heading': '重复标题', 'level': 1})
        with self.assertRaises(ValueError):
            self.create(data)

    def test_no_notation_does_not_invent_intro(self):
        data = self.data(); data['front_matter'].pop('notation')
        source, output = self.create(data)
        self.assertTrue(Document(output).paragraphs[2].text.startswith('测试画面'))
        self.assertFalse(inspect_docx(output, source=source)['errors'])

    def test_documented_example_is_executable(self):
        text = (SCRIPTS.parent/'docs/tools/讲义制作输入与检查.md').read_text(encoding='utf-8')
        data = json.loads(re.search(r'```json\s*(.*?)\s*```', text, re.S).group(1))
        shutil.copy2(self.root/'synthetic.png', self.root/'已有截图.png')
        source, output = self.create(data)
        self.assertFalse(inspect_docx(output, source=source)['errors'])
