"""Verify white canvas pixels independently of the shared image helper."""
from io import BytesIO
import json
import math
import unittest
from unittest.mock import patch
from PIL import Image
from docx import Document
import test_production as fixtures
from test_production import maker
from audit_docx import inspect_docx
import handout_images


class PaddingTests(unittest.TestCase):
    setUp = fixtures.ProductionTests.setUp
    tearDown = fixtures.ProductionTests.tearDown

    def source(self, scale, crop=None):
        item = {'type': 'image', 'at': 1, 'path': 'synthetic.png',
                'content_scale': scale, 'reference': True}
        if crop is not None:
            item['crop'] = crop
        data = {'version': 2, 'blocks': [
            {'id': 'first', 'lyrics': '合成', 'transcription': '动作', 'sequence': [item]},
            {'repeat_of': 'first'}]}
        source = self.root/'content.json'
        source.write_text(json.dumps(data), encoding='utf-8')
        return source

    def test_pixels_canvas_crop_cache_and_unchanged_source(self):
        source = self.source(.7, [10, 5, 220, 115])
        original = (self.root/'synthetic.png').read_bytes()
        output = self.root/'padded.docx'
        with patch.object(handout_images, 'padded_picture', wraps=handout_images.padded_picture) as call:
            maker.make_handout(source, output)
            self.assertEqual(call.call_count, 1)
        self.assertEqual((self.root/'synthetic.png').read_bytes(), original)
        d = Document(output)
        shape = d.inline_shapes[0]
        rel = shape._inline.xpath('.//a:blip')[0].get(fixtures.maker.qn('r:embed'))
        embedded = Image.open(BytesIO(d.part.related_parts[rel].blob)).convert('RGB')
        selected = Image.open(BytesIO(original)).convert('RGB').crop((10, 5, 220, 115))
        height = math.ceil(110/.7);top=(height-110)//2
        self.assertEqual(embedded.size, (210, height))
        self.assertEqual(embedded.crop((0,top,210,top+110)).tobytes(), selected.tobytes())
        for box in ((0,0,210,top),(0,top+110,210,height)):
            self.assertEqual(embedded.crop(box).getextrema(), ((255,255),)*3)
        self.assertAlmostEqual(shape.height/914400, .79, places=4)
        self.assertAlmostEqual(shape.width/shape.height, 210/height, places=4)
        self.assertFalse(shape._inline.xpath('.//a:srcRect'))
        self.assertFalse(inspect_docx(output, source=source)['errors'])

    def test_scale_one_keeps_original_and_bad_scales_fail(self):
        source=self.source(1);output=self.root/'same.docx';maker.make_handout(source,output)
        doc=Document(output)
        drawing=doc.inline_shapes[0]._inline
        rel=drawing.xpath('.//a:blip')[0].get(fixtures.maker.qn('r:embed'))
        self.assertEqual(doc.part.related_parts[rel].blob,(self.root/'synthetic.png').read_bytes())
        for value in (0, -1, 1.1, True, '0.7', float('nan'), .000000001):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    maker.make_handout(self.source(value),self.root/'bad.docx')
                self.assertFalse((self.root/'bad.docx').exists())
        with self.assertRaises(ValueError):
            maker.make_handout(self.source(.7,[0.5,0,120,120]),self.root/'bad.docx')

    def test_missing_padding_in_output_is_detected(self):
        source=self.source(1);output=self.root/'unadjusted.docx';maker.make_handout(source,output)
        self.source(.7)
        self.assertTrue(inspect_docx(output,source=source)['errors'])

    def test_transparency_is_flattened_on_white(self):
        image=Image.new('RGBA',(4,2),(0,0,0,0));image.putpixel((1,1),(0,0,0,255))
        path=self.root/'alpha.png';image.save(path)
        payload,size=handout_images.padded_picture(path,None,.5)
        result=Image.open(BytesIO(payload))
        self.assertEqual(size,(4,4))
        self.assertEqual(result.getpixel((0,1)),(255,255,255))
        self.assertEqual(result.getpixel((1,2)),(0,0,0))
