"""Coverage includes valid custom layouts and deliberately damaged output."""
import json
import unittest
from docx import Document
from docx.shared import Pt,Cm
from docx.oxml.ns import qn
import test_production as fixtures
from test_production import maker,sample_data
from handout_layout import read_layout,LAYOUT_COVERAGE,STYLE_FIELDS
from audit_docx import inspect_docx


class FormatAuditTests(unittest.TestCase):
    setUp=fixtures.ProductionTests.setUp
    tearDown=fixtures.ProductionTests.tearDown
    def create(self, settings=None):
        data=sample_data()
        data['blocks'].pop(0)
        data['front_matter']={'song':'合成','teachers':'老师甲','date':'2026.11.01','notation':['仅此例使用的说明。']}
        data['blocks'].insert(0,{'heading':'主歌一'})
        data['blocks'].insert(1,{'text':'开场','role':'intro'})
        source=self.root/'content.json';source.write_text(json.dumps(data),encoding='utf-8')
        layout=self.root/'layout.json';layout.write_text(json.dumps(settings or {}),encoding='utf-8')
        output=self.root/'out.docx';maker.make_handout(source,output,layout=layout,edition='teacher')
        return source,layout,output

    def test_custom_role_fonts_paragraphs_and_frame_padding_pass(self):
        settings={'font_name':'KaiTi','font_size':12,'title_size':18,'heading_size':15,
                  'margin_horizontal':2.2,'margin_vertical':2.1,'image_height':.6,
                  'line_spacing':1.35,'image_line_spacing':1.5,'image_space_before':0,'image_space_after':2,
                  'reference_frame_padding':True,'reference_frame_padding_pt':3,
                  'content_styles':{'byline':{'font_size':9,'space_after':0},
                   'notation':{'font_size':10,'space_before':0,'line_spacing':1.1},
                   'intro':{'font_name':'SimSun','font_size':11,'space_after':3},
                   'body':{'space_before':2,'space_after':4},'lyrics':{'font_size':10},
                   'transcription':{'font_size':11},'note':{'font_size':9},
                   'teacher_note':{'font_size':10,'line_spacing':1.4}}}
        source,layout,out=self.create(settings)
        self.assertFalse(inspect_docx(out,source=source,layout=layout,edition='teacher')['errors'])

    def test_mutations_to_format_and_invisible_frame_are_detected(self):
        source,layout,out=self.create({'reference_frame_padding':True})
        original=out.read_bytes()
        def size(d):d.paragraphs[0].runs[0].font.size=Pt(36)
        def name(d):d.paragraphs[1].runs[0].font.name='Arial'
        def east_asia(d):
            run=d.paragraphs[0].runs[0];run.font.name='SimSun'
            run._r.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Arial')
        def margin(d):d.sections[0].left_margin=Cm(4)
        def spacing(d):d.paragraphs[4].paragraph_format.space_after=Pt(90)
        def lines(d):d.paragraphs[4].paragraph_format.line_spacing=2.5
        def frame(d):
            for line in d._element.iter(qn('a:ln')):
                line.set('w','0')
                for color in line.iter(qn('a:srgbClr')):color.set('val','FFFFFF')
        for change in (size,name,east_asia,margin,spacing,lines,frame):
            with self.subTest(change=change.__name__):
                out.write_bytes(original);d=Document(out);change(d);d.save(out)
                self.assertTrue(inspect_docx(out,source=source,layout=layout,edition='teacher')['errors'])

    def test_coverage_registry_and_invalid_configuration(self):
        self.assertEqual(set(read_layout())|{'image_space_before','image_space_after','image_line_spacing'},set(LAYOUT_COVERAGE))
        for settings in ({'font_name':123},{'font_size':float('nan')},{'line_spacing':True},
                         {'margin_horizontal':float('inf')},[],
                         {'content_styles':{'note':{'space_after':2}}}):
            with self.subTest(settings=settings):
                path=self.root/'invalid.json';path.write_text(json.dumps(settings),encoding='utf-8')
                with self.assertRaises(ValueError):read_layout(path)
        for field in ('paragraph_overrides','reference_frame_rows'):
            with self.assertRaises(ValueError):
                self.create({field:{'1':{'font_size':10}} if field=='paragraph_overrides' else [1]})

    def test_single_word_transcription_is_not_lost(self):
        d=Document();d.add_paragraph('歌词（甲）');out=self.root/'one.docx';d.save(out)
        self.assertEqual(inspect_docx(out)['transcriptions'],['甲'])
