"""Contract and mutation checks; no song data or model context required."""
import json
import shutil
import subprocess
import sys
import os
import unittest
from docx import Document
import test_production as fixtures
from test_production import sample_data, maker, SCRIPTS
from audit_docx import inspect_docx


class SequenceTests(unittest.TestCase):
    # Reuse the temp/image fixture, not the parent's tests.
    setUp = fixtures.ProductionTests.setUp
    tearDown = fixtures.ProductionTests.tearDown
    def prepare(self, data=None, edition='student', layout=None):
        source = self.root / 'input.json'
        source.write_text(json.dumps(data or sample_data(), ensure_ascii=False), encoding='utf-8')
        output = self.root / (edition + '.docx')
        maker.make_handout(source, output, layout=layout, edition=edition)
        return source, output

    def test_teacher_notes_and_print_separation(self):
        data = sample_data()
        data['blocks'][1]['teacher_note'] = '比较 *候选* 后确认；仅老师查看'
        source, teacher = self.prepare(data, 'teacher')
        student = self.root / 'student.docx'
        maker.make_handout(source, student)
        text = '\n'.join(p.text for p in Document(teacher).paragraphs)
        self.assertEqual(text.count('【比较 *候选* 后确认；仅老师查看】'), 3)
        self.assertNotIn('候选', '\n'.join(p.text for p in Document(student).paragraphs))
        self.assertFalse(inspect_docx(teacher, source=source, edition='teacher')['errors'])
        self.assertFalse(inspect_docx(student, source=source)['errors'])
        self.assertTrue(inspect_docx(teacher, source=source)['errors'])
        self.assertTrue(inspect_docx(student, source=source, edition='teacher')['errors'])

    def test_moving_note_across_image_with_same_text_is_detected(self):
        source, output = self.prepare()
        doc = Document(output)
        p = doc.paragraphs[2]
        before = p.text
        note = next(r for r in p.runs if '方向调整' in r.text)
        image = next(r for r in p.runs if r._r.xpath('.//w:drawing'))
        note._r.addprevious(image._r)
        self.assertEqual(p.text, before)
        doc.save(output)
        self.assertTrue(inspect_docx(output, source=source)['errors'])

    def test_marker_frame_repeat_and_dimensions_mutations(self):
        source, output = self.prepare()
        original = output.read_bytes()
        def italic(d):
            next(r for p in d.paragraphs for r in p.runs if r.text == '动作①').italic = False
        def strike(d):
            next(r for p in d.paragraphs for r in p.runs if r.text == '省略②').font.strike = False
        def frame(d):
            node = d._element.xpath('.//a:prstDash')[0]
            node.getparent().remove(node)
        def remove_image(d):
            node = d._element.xpath('.//w:drawing')[-1]
            node.getparent().remove(node)
        def stretch(d):
            d.inline_shapes[0].width *= 2
        for mutate in (italic, strike, frame, remove_image, stretch):
            with self.subTest(mutate=mutate.__name__):
                output.write_bytes(original)
                doc = Document(output)
                mutate(doc)
                doc.save(output)
                self.assertTrue(inspect_docx(output, source=source)['errors'])

    def test_omission_and_section_headings(self):
        data = {'version': 2, 'blocks': [
            {'heading': '歌曲', 'level': 1}, {'heading': '主歌一'},
            {'id': 'v', 'lyrics': '合成', 'transcription': '甲/乙/丙', 'sequence': [
                {'type': 'text', 'at': 1, 'text': '向外'},
                {'type': 'omit', 'at': 2, 'reason': '关系已在动作方向中表达'},
                {'type': 'image', 'at': 3, 'path': 'synthetic.png'}]},
            {'heading': '副歌前段'}, {'heading': '副歌'},
            {'repeat_of': 'v', 'omit_images': True}, {'heading': '副歌'}]}
        source, output = self.prepare(data)
        doc = Document(output)
        self.assertEqual([p.style.name for p in doc.paragraphs if p.text in ('歌曲', '主歌一', '副歌前段', '副歌')],
                         ['Title', 'Heading 1', 'Heading 1', 'Heading 1', 'Heading 1'])
        self.assertNotIn('关系已', '\n'.join(p.text for p in doc.paragraphs))
        self.assertEqual(len(doc.inline_shapes), 1)
        self.assertFalse(inspect_docx(output, source=source)['errors'])

    def test_invalid_data_fails_before_output(self):
        cases = []
        for field, value in [('note', '不许落到图下'), ('images', []), ('sequnce', [])]:
            data = sample_data(); data['blocks'][1][field] = value; cases.append(data)
        for change in ({'at': 0}, {'at': True}, {'at': 99}, {'type': 'unknown'},
                       {'note': '图片尾注'}, {'crop': [False, 0, 1, 1]},
                       {'at': []}, {'at': [2, 1]}, {'at': [1, 3]}):
            data = sample_data(); data['blocks'][1]['sequence'][1].update(change); cases.append(data)
        data = sample_data(); data['blocks'][1]['sequence'] = data['blocks'][1]['sequence'][:2]; cases.append(data)
        data = sample_data(); data['blocks'][1]['sequence'][2]['at'] = 1; data['blocks'][1]['sequence'][3]['at'] = 1; cases.append(data)
        data = sample_data(); data['blocks'][1]['transcription'] = '*未闭合'; cases.append(data)
        data = sample_data(); data['blocks'][1]['teacher_note'] = '【重复括号】'; cases.append(data)
        data = sample_data(); data['version'] = 1; cases.append(data)
        for number, data in enumerate(cases):
            with self.subTest(number=number):
                source = self.root / 'invalid.json'
                source.write_text(json.dumps(data), encoding='utf-8')
                output = self.root / 'invalid.docx'
                with self.assertRaises(ValueError):
                    maker.make_handout(source, output)
                self.assertFalse(output.exists())

    def test_one_image_can_cover_adjacent_words(self):
        data = {'version': 2, 'blocks': [{'lyrics': '合成', 'transcription': '甲/乙',
                'sequence': [{'type': 'image', 'at': [1, 2], 'path': 'synthetic.png'}]}]}
        source, output = self.prepare(data)
        self.assertEqual(len(Document(output).inline_shapes), 1)
        self.assertFalse(inspect_docx(output, source=source)['errors'])

    def test_omission_conflict_fails(self):
        data = sample_data()
        data['blocks'][1]['sequence'].insert(1, {'type': 'omit', 'at': 1, 'reason': '重复声明'})
        with self.assertRaises(ValueError):
            self.prepare(data)

    def test_text_draft_is_warning_not_silent_final_or_error(self):
        data = {'version': 2, 'blocks': [{'lyrics': '草稿', 'transcription': '动作',
                'teacher_note': '候选比较及选词依据；此处待讨论'}]}
        source, output = self.prepare(data)
        result = inspect_docx(output, source=source)
        self.assertFalse(result['errors'])
        self.assertTrue(any('文字草稿' in message for message in result['warnings']))
        self.assertEqual([p.text for p in Document(output).paragraphs], ['草稿（动作）'])
        teacher = self.root / 'discussion.docx'
        maker.make_handout(source, teacher, edition='teacher')
        self.assertEqual([p.text for p in Document(teacher).paragraphs],
                         ['草稿（动作）', '【候选比较及选词依据；此处待讨论】'])
        self.assertFalse(inspect_docx(teacher, source=source, edition='teacher')['errors'])

    def test_explicit_sequence_exception_preserves_order_checks_and_bounds(self):
        data = {'version': 2, 'blocks': [{'lyrics': '合成选择句', 'transcription': '甲/乙',
                'sequence': [{'type': 'text', 'at': 2, 'text': '右侧方案'},
                             {'type': 'text', 'at': 1, 'text': '左侧方案'}]}]}
        with self.assertRaises(ValueError):
            self.prepare(data)
        data['blocks'][0]['sequence_exception'] = '两侧方案先右后左展示，按整体图解组织'
        source, output = self.prepare(data)
        result = inspect_docx(output, source=source)
        self.assertFalse(result['errors'])
        self.assertTrue(any('特殊图示组织' in s for s in result['manual']))
        doc = Document(output)
        self.assertLess(doc.paragraphs[1].text.index('右侧'), doc.paragraphs[1].text.index('左侧'))
        doc.paragraphs[1].text = '〔左侧方案〕 〔右侧方案〕 '
        doc.save(output)
        self.assertTrue(inspect_docx(output, source=source)['errors'])
        repeated = json.loads(json.dumps(data))
        repeated['blocks'][0]['id'] = 'choice'
        repeated['blocks'].append({'repeat_of': 'choice', 'omit_images': True})
        repeated_source = self.root / 'exception-repeat.json'
        repeated_source.write_text(json.dumps(repeated), encoding='utf-8')
        repeated_content = maker.check_content(maker.read_content(repeated_source))
        self.assertEqual(repeated_content['blocks'][1]['sequence'], [])
        self.assertNotIn('sequence_exception', repeated_content['blocks'][1])
        data['blocks'][0]['sequence'][0]['at'] = 0
        source.write_text(json.dumps(data), encoding='utf-8')
        with self.assertRaises(ValueError):
            maker.check_content(maker.read_content(source))

    def test_portable_cli_with_unicode_relative_paths_and_layout(self):
        portable = self.root / '独立项目'
        portable.mkdir()
        for name in ('make_handout.py', 'audit_docx.py', 'handout_content.py', 'handout_images.py', 'handout_layout.py', 'handout_format_audit.py'):
            shutil.copy2(SCRIPTS / name, portable / name)
        shutil.copy2(self.root / 'synthetic.png', portable / 'synthetic.png')
        (portable / 'input.json').write_text(json.dumps(sample_data()), encoding='utf-8')
        (portable / 'layout.json').write_text('{"image_height":0.6,"reference_frame_padding":true}', encoding='utf-8')
        for command in ([str(portable/'make_handout.py'), str(portable/'input.json'), '--out', str(portable/'output.docx'),
                         '--layout', str(portable/'layout.json'), '--edition', 'teacher'],
                        [str(portable/'audit_docx.py'), str(portable/'output.docx'), '--source', str(portable/'input.json'),
                         '--layout', str(portable/'layout.json'), '--edition', 'teacher']):
            result = subprocess.run([sys.executable, *command], cwd=self.root, capture_output=True,
                                    env={**os.environ, 'PYTHONIOENCODING': 'cp1252'})
            self.assertEqual(result.returncode, 0, result.stdout.decode('utf-8') + result.stderr.decode(errors='replace'))
