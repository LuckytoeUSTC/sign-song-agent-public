"""One handout entry point: production JSON or a selected DOCX copy."""
from pathlib import Path
import argparse
import sys
from io import BytesIO
import json
from zipfile import ZipFile
from lxml import etree

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from handout_layout import DEFAULT_LAYOUT, read_layout, validate_paragraph_style
from handout_content import read_content, check_content, check_image, marked_text

TOOL_VERSION = "3.2.0"


def text_runs(p, text, font_name='SimSun', font_size=11):
    """Render the transcription markers and run boundaries."""
    for token, italic, strike in marked_text(text):
        run = p.add_run(token)
        run.italic = italic
        run.font.strike = strike
        if token == '++':
            ics = OxmlElement('w:iCs')
            ics.set(qn('w:val'), '0')
            run._element.get_or_add_rPr().append(ics)
        run.font.name = 'Arial' if token == '~' else font_name
        run.font.size = Pt(font_size)
        run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), font_name)


def apply_paragraph_style(p, options):
    for key in ('space_before', 'space_after', 'line_spacing'):
        if key in options:
            setattr(p.paragraph_format, key,
                    options[key] if key == 'line_spacing' else Pt(options[key]))
    for run in p.runs:
        if 'font_name' in options and run.text != '~':
            run.font.name = options['font_name']
            run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), options['font_name'])
        if 'font_size' in options:
            run.font.size = Pt(options['font_size'])


def pad_reference_row(p, padding):
    """Reserve vertical room only for a selected row containing dotted pictures."""
    from docx.enum.text import WD_LINE_SPACING
    heights = p._p.xpath('.//wp:inline/wp:extent/@cy')
    lines = p._p.xpath('.//pic:spPr/a:ln[a:prstDash]')
    if not heights or not lines:
        return
    required = max(int(h) for h in heights) / 12700 + padding
    current = p.paragraph_format.line_spacing
    if isinstance(current, int):
        required = max(required, current / 12700)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    p.paragraph_format.line_spacing = Pt(required)
    for line in lines:
        line.set('algn', 'in')


def copy_reference(source, output, supplied):
    """Patch only document.xml; all media, styles, headers and other parts stay intact."""
    document = Document(source)
    forbidden = set(supplied) - {'paragraph_overrides', 'reference_frame_rows',
                                'reference_frame_padding_pt', 'content_styles'}
    if forbidden:
        raise ValueError('docx副本只接受逐段覆盖和指定参考图行；全局参数用于新稿')
    if supplied.get('content_styles'):
        raise ValueError('docx副本请用paragraph_overrides明确选择段落，不推测角色')
    overrides = supplied.get('paragraph_overrides', {})
    rows = supplied.get('reference_frame_rows', [])
    for number in list(overrides) + rows:
        if int(number) > len(document.paragraphs):
            raise ValueError('段落序号超出原件：' + str(number))
    for number, options in overrides.items():
        apply_paragraph_style(document.paragraphs[int(number)-1], options)
    for number in rows:
        pad_reference_row(document.paragraphs[number-1], supplied.get('reference_frame_padding_pt', 6))
    with ZipFile(source) as original, ZipFile(output, 'w') as result:
        for info in original.infolist():
            data = original.read(info.filename)
            if info.filename == 'word/document.xml' and (overrides or rows):
                data = etree.tostring(document._element, xml_declaration=True,
                                      encoding='UTF-8', standalone=True)
            result.writestr(info, data)


def new_document(settings):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(settings['margin_horizontal'])
    section.top_margin = section.bottom_margin = Cm(settings['margin_vertical'])
    for name in ['Normal', 'Title', 'Heading 1', 'Heading 2']:
        style = document.styles[name]
        style.font.name = settings['font_name']
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), settings['font_name'])
    for style in document.styles:
        for border in style.element.xpath('.//w:pBdr'):
            border.getparent().remove(border)
    document.styles['Normal'].font.size = Pt(settings['font_size'])
    return document


def add_heading(document, block, settings):
    level = block.get('level', 2)
    p = document.add_paragraph(block['heading'], style='Title' if level == 1 else 'Heading ' + str(level - 1))
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    for run in p.runs:
        run.font.size = Pt(settings['title_size'] if level == 1 else settings['heading_size'])
    return p


def add_paragraph(document, settings, images=False, keep_with_next=False):
    p = document.add_paragraph()
    p.paragraph_format.space_after = Pt(7 if images else 5)
    p.paragraph_format.line_spacing = 1.25 if images else settings['line_spacing']
    p.paragraph_format.keep_together = True
    if keep_with_next:
        p.paragraph_format.keep_with_next = True
    if images:
        p.paragraph_format.space_before = Pt(4)
        if 'image_space_before' in settings:
            p.paragraph_format.space_before = Pt(settings['image_space_before'])
        if 'image_space_after' in settings:
            p.paragraph_format.space_after = Pt(settings['image_space_after'])
        if 'image_line_spacing' in settings:
            p.paragraph_format.line_spacing = settings['image_line_spacing']
    return p


def add_picture(p, item, base, settings, cache=None):
    path, image = check_image(item, base)
    padded = item.get('content_scale', 1) < 1
    if padded:
        from handout_images import padded_picture
        key = (str(path.resolve()), tuple(item.get('crop', [])), item['content_scale'])
        if cache is None:
            cache = {}
        if key not in cache:
            cache[key] = padded_picture(path, item.get('crop'), item['content_scale'])
        payload, dimensions = cache[key]
        shape = p.add_run().add_picture(BytesIO(payload), height=Inches(settings['image_height']))
        shape.width = int(shape.height * dimensions[0] / dimensions[1])
    else:
        shape = p.add_run().add_picture(str(path), height=Inches(settings['image_height']))
    # Alt/entry metadata is never turned into a visible caption.
    shape._inline.docPr.set('descr', item.get('alt', item.get('entry', '')))
    if 'crop' in item and not padded:
        left, top, right, bottom = item['crop']
        shape.width = int(shape.height * (right - left) / (bottom - top))
        rect = OxmlElement('a:srcRect')
        for key, fraction in dict(l=left / image.px_width, t=top / image.px_height,
                                  r=1 - right / image.px_width, b=1 - bottom / image.px_height).items():
            rect.set(key, str(round(fraction * 100000)))
        fill = shape._inline.xpath('.//pic:blipFill')[0]
        stretch = fill.find(qn('a:stretch'))
        if stretch is None:
            fill.append(rect)
        else:
            fill.insert(list(fill).index(stretch), rect)
    if item.get('reference', False):
        for effect in shape._inline.xpath('./wp:effectExtent'):
            for key, value in dict(t='25400', b='25400', l='12700', r='12700').items():
                effect.set(key, value)
        props = shape._inline.xpath('.//pic:spPr')[0]
        line = OxmlElement('a:ln')
        line.set('w', '12700')
        if settings.get('reference_frame_padding', False):
            line.set('algn', 'in')
        fill = OxmlElement('a:solidFill')
        color = OxmlElement('a:srgbClr')
        color.set('val', '000000')
        fill.append(color)
        line.append(fill)
        dash = OxmlElement('a:prstDash')
        dash.set('val', 'sysDot')
        line.append(dash)
        props.append(line)
    return shape


def render_content(content, settings, edition='student'):
    if edition not in ('student', 'teacher'):
        raise ValueError('edition须为student或teacher')
    document = new_document(settings)
    image_cache = {}
    if 'front_matter' in content:
        front = content['front_matter']
        add_heading(document, {'heading': front['song'], 'level': 1}, settings)
        byline = add_paragraph(document, settings)
        byline.add_run('--' + front['date'] + ' by' + front['teachers'])
        apply_paragraph_style(byline, settings.get('content_styles', {}).get('byline', {}))
        byline.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for explanation in front.get('notation', []):
            p = add_paragraph(document, settings)
            p.add_run(explanation)
            apply_paragraph_style(p, settings.get('content_styles', {}).get('notation', {}))
    write = lambda p, value: text_runs(p, value, settings['font_name'], settings['font_size'])
    for block in content['blocks']:
        if 'heading' in block:
            add_heading(document, block, settings)
            continue
        sequence = block.get('sequence', [])
        visible = [item for item in sequence if item['type'] != 'omit']
        text = block.get('text', block.get('lyrics', ''))
        if 'transcription' in block:
            text += '（' + block['transcription'] + '）'
        p = add_paragraph(document, settings, keep_with_next=bool(visible))
        roles = settings.get('content_styles', {})
        if 'transcription' in block and any(k in roles for k in ('lyrics', 'transcription')):
            options = roles.get('lyrics', {})
            text_runs(p, block.get('lyrics', ''), options.get('font_name', settings['font_name']),
                      options.get('font_size', settings['font_size']))
            options = roles.get('transcription', {})
            text_runs(p, '（' + block['transcription'] + '）',
                      options.get('font_name', settings['font_name']),
                      options.get('font_size', settings['font_size']))
        else:
            write(p, text)
        options = roles.get(block.get('role', 'body'), {})
        if 'transcription' in block and any(k in roles for k in ('lyrics', 'transcription')):
            options = {k: v for k, v in options.items() if k not in ('font_name', 'font_size')}
        apply_paragraph_style(p, options)
        if visible:
            picture_row = add_paragraph(document, settings, images=True)
            for item in visible:
                if item['type'] == 'image':
                    add_picture(picture_row, item, content['base'], settings, image_cache)
                else:
                    # A text item occupies its own position; no implicit image caption.
                    run = picture_row.add_run('〔' + item['text'] + '〕')
                    options = roles.get('note', {})
                    run.font.name = options.get('font_name', settings['font_name'])
                    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), run.font.name)
                    run.font.size = Pt(options.get('font_size', settings['font_size']))
                    run.italic = False
                    run.font.strike = False
                write(picture_row, ' ')
            if settings.get('reference_frame_padding'):
                pad_reference_row(picture_row, settings['reference_frame_padding_pt'])
        if edition == 'teacher' and block.get('teacher_note'):
            note_row = add_paragraph(document, settings)
            note_row.add_run('【' + block['teacher_note'] + '】')
            apply_paragraph_style(note_row, roles.get('teacher_note', roles.get('note', {})))
    return document


def make_handout(source, output, height=None, layout=None, input_format='json', edition='student'):
    output = Path(output)
    if output.exists():
        raise ValueError('输出文件已存在；请选择新文件名')
    if edition not in ('student', 'teacher'):
        raise ValueError('edition须为student或teacher')
    settings = read_layout(layout, height)
    if input_format == 'docx':
        if edition != 'student':
            raise ValueError('docx副本不按edition筛选内容；只按明确授权局部修改')
        if height is not None:
            raise ValueError('docx副本保留原图尺寸，不接受--height')
        supplied = json.loads(Path(layout).read_text(encoding='utf-8')) if layout else {}
        output.parent.mkdir(parents=True, exist_ok=True)
        copy_reference(source, output, supplied)
        return output
    if settings.get('paragraph_overrides') or settings.get('reference_frame_rows'):
        raise ValueError('逐段覆盖及指定参考图行只用于docx副本模式，不用于JSON新稿')
    content = check_content(read_content(source, input_format))
    document = render_content(content, settings, edition)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    return output


def main():
    # CLI output is UTF-8 even when a Windows parent redirects it through a pipe.
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--version', action='version', version=TOOL_VERSION)
    ap.add_argument('source', type=Path, nargs='?')
    ap.add_argument('--out', type=Path)
    ap.add_argument('--height', type=float)
    ap.add_argument('--edition', choices=('student', 'teacher'), default='student',
                    help='student excludes teacher_note; teacher wraps it in 【】')
    ap.add_argument('--layout', type=Path, help='JSON layout settings')
    ap.add_argument('--input-format', choices=('json', 'docx'), default='json')
    args = ap.parse_args()
    if not args.source or not args.out:
        ap.error('制作需要source与--out')
    make_handout(args.source, args.out, args.height, args.layout, args.input_format, args.edition)
    print('生成', args.out, '；请逐页预览后使用。')


if __name__ == '__main__':
    main()
