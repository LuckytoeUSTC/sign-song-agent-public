"""Read-only Word checks: confirmed errors, warnings and human review items."""
from pathlib import Path
from collections import Counter
import argparse
import sys
import io
import hashlib
import re
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.image.image import Image
from docx.oxml.ns import qn
from handout_format_audit import check_layout


def effective(r, p, attribute):
    direct = getattr(r.font, attribute)
    if direct is not None:
        return direct
    for style in [r.style, p.style]:
        while style:
            value = getattr(style.font, attribute)
            if value is not None:
                return value
            style = style.base_style
    return None


def _crop_values(drawing):
    found = drawing.xpath('.//a:srcRect')
    return tuple(int(found[0].get(k, '0')) for k in ('l', 't', 'r', 'b')) if found else (0, 0, 0, 0)


def _selected_bytes(drawing, document):
    blips = drawing.xpath('.//a:blip')
    part = document.part.related_parts.get(blips[0].get(qn('r:embed'))) if blips else None
    return part.blob if part is not None else None


def _characters(text, marked=True):
    from handout_content import marked_text
    parts = marked_text(text) if marked else [(text, False, False)]
    return [('text', char, italic, strike) for token, italic, strike in parts for char in token]


def _image_event(data, crop, reference):
    return ('image', hashlib.sha256(data).hexdigest(), crop, reference)


def _actual_events(p, document):
    events = []
    for run in p.runs:
        for node in run._r:
            if node.tag == qn('w:t'):
                italic, strike = bool(effective(run, p, 'italic')), bool(effective(run, p, 'strike'))
                events.extend(('text', char, italic, strike) for char in (node.text or ''))
            elif node.tag in (qn('w:tab'), qn('w:br'), qn('w:cr')):
                events.extend(_characters('\t' if node.tag == qn('w:tab') else '\n', False))
            elif node.tag == qn('w:drawing'):
                events.append(_image_event(_selected_bytes(node, document) or b'', _crop_values(node),
                              bool(node.xpath('.//a:prstDash[@val="sysDot"]'))))
    return events


def _expected_rows(content, edition):
    # Build the contract from data, never call the Word renderer as an oracle.
    from handout_content import check_image
    rows, geometry = [], []
    image_cache = {}
    if 'front_matter' in content:
        front = content['front_matter']
        rows.append((_characters(front['song'], False), 'Title'))
        rows.append((_characters('--' + front['date'] + ' by' + front['teachers'], False), None))
        rows.extend((_characters(text, False), None) for text in front.get('notation', []))
    for block in content['blocks']:
        if 'heading' in block:
            level = block.get('level', 2)
            rows.append((_characters(block['heading'], False), 'Title' if level == 1 else 'Heading ' + str(level - 1)))
            continue
        text = block.get('text', block.get('lyrics', ''))
        if 'transcription' in block:
            text += '（' + block['transcription'] + '）'
        rows.append((_characters(text), None))
        events = []
        for item in block.get('sequence', []):
            if item['type'] == 'omit':
                continue
            if item['type'] == 'text':
                events.extend(_characters('〔' + item['text'] + '〕', False))
            else:
                path, image = check_image(item, content['base'])
                crop = item.get('crop')
                padded = item.get('content_scale', 1) < 1
                if padded:
                    from handout_images import padded_picture
                    key = (str(path.resolve()), tuple(crop or []), item['content_scale'])
                    if key not in image_cache:
                        image_cache[key] = padded_picture(path, crop, item['content_scale'])
                    payload, dimensions = image_cache[key]
                    rect = (0, 0, 0, 0)
                    ratio = dimensions[0] / dimensions[1]
                elif crop:
                    l, t, r, b = crop
                    rect = tuple(round(v * 100000) for v in
                                 (l / image.px_width, t / image.px_height,
                                  1 - r / image.px_width, 1 - b / image.px_height))
                    ratio = (r-l)/(b-t)
                else:
                    rect = (0, 0, 0, 0)
                    ratio = (image.px_width / image.horz_dpi) / (image.px_height / image.vert_dpi)
                events.append(_image_event(payload if padded else path.read_bytes(), rect, item.get('reference', False)))
                geometry.append(ratio)
            events.extend(_characters(' ', False))
        if events:
            rows.append((events, None))
        if edition == 'teacher' and block.get('teacher_note'):
            rows.append((_characters('【' + block['teacher_note'] + '】', False), None))
    return rows, geometry


def inspect_docx(path, height=None, source=None, input_format='json', edition='student', layout=None):
    """No writes. Missing optional production fields are not errors."""
    d = Document(path)
    errors, warnings = [], []
    manual = ['逐页查看导出PDF或页面图，核对分页、图框、文字和留白。',
              '手形、箭头是否被裁掉、借图是否适当及动作自然度须看原图并由老师试打；静态检查不能确认。']
    fonts, sizes = Counter(), Counter()
    image_headers = {}  # Per-document cache; never retained across file changes.
    transcriptions, images = [], 0
    width = min(s.page_width - s.left_margin - s.right_margin for s in d.sections)
    paragraphs = list(d.paragraphs) + [p for t in d.tables for row in t.rows for c in row.cells for p in c.paragraphs]
    for i, p in enumerate(paragraphs, 1):
        text = p.text
        if not text.startswith('【') and '（' in text and '）' in text:
            transcriptions.append(text.split('（', 1)[1].rsplit('）', 1)[0])
        flags = []
        for r in p.runs:
            flags.extend([bool(effective(r, p, 'italic'))] * len(r.text))
            if r.text:
                fonts[effective(r, p, 'name') or '未显式指定'] += len(r.text)
                size = effective(r, p, 'size')
                sizes[round(size.pt, 1) if size else '未显式指定'] += len(r.text)
        for match in re.finditer(r'\+\+', text):
            if any(flags[match.start():match.end()]):
                errors.append(f'段落{i}：++含斜体')
        if '**' in text or '\\n' in text:
            warnings.append(f'段落{i}：可能残留Markdown或转义字符')
        for drawing in p._p.xpath('.//w:drawing'):
            images += 1
            extents = drawing.xpath('.//wp:extent')
            if drawing.xpath('./wp:anchor'):
                warnings.append(f'段落{i}：浮动图片，检查位置')
            if not extents:
                continue
            h, w = int(extents[0].get('cy')), int(extents[0].get('cx'))
            if height and abs(h / 914400 - height) > .002:
                warnings.append(f'段落{i}：图高{h / 914400:.3f}英寸，检查是否为照片等例外')
            if w > width:
                warnings.append(f'段落{i}：单图宽度超过正文宽度')
            if p.paragraph_format.line_spacing_rule == WD_LINE_SPACING.EXACTLY:
                line = p.paragraph_format.line_spacing
                if line and line < h:
                    warnings.append(f'段落{i}：固定行距小于图高，可能截框')
            # Compare stored image geometry to the visible crop; no semantic claims.
            blips = drawing.xpath('.//a:blip')
            if blips and h > 0:
                rel = blips[0].get(qn('r:embed'))
                if rel and rel in d.part.related_parts:
                    try:
                        if rel not in image_headers:
                            image_headers[rel] = Image.from_file(io.BytesIO(d.part.related_parts[rel].blob))
                        image = image_headers[rel]
                        left, top, right, bottom = _crop_values(drawing)
                        visible_w = 1 - (left + right) / 100000
                        visible_h = 1 - (top + bottom) / 100000
                        if visible_w <= 0 or visible_h <= 0:
                            errors.append(f'段落{i}：图片裁图范围为空')
                        else:
                            expected = (image.px_width / image.horz_dpi) / (image.px_height / image.vert_dpi) * visible_w / visible_h
                            # JSON pixel crops use square-pixel display; unusual DPI needs review.
                            pixel_expected = image.px_width / image.px_height * visible_w / visible_h
                            if min(abs(w / h / expected - 1), abs(w / h / pixel_expected - 1)) > .03:
                                warnings.append(f'段落{i}：图片显示比例与原图及裁图不一致，核对是否有意变形')
                    except (ValueError, KeyError, ZeroDivisionError, TypeError):
                        warnings.append(f'段落{i}：此图片类型未完成比例检查，人工核对')
    if source:
        from handout_content import read_content, check_content
        from handout_layout import read_layout
        if edition not in ('student', 'teacher'):
            raise ValueError('edition须为student或teacher')
        content = check_content(read_content(source, input_format))
        settings = read_layout(layout, height)
        for block in content['blocks']:
            if block.get('sequence_exception'):
                manual.append('特殊图示组织：' + block['lyrics'][:28] + '；' + block['sequence_exception'])
            if 'transcription' in block and not block.get('sequence') and not block.get('omit_images'):
                warnings.append('尚未提供图示链，按文字草稿核对：' + block['lyrics'][:28])
        expected, geometry = _expected_rows(content, edition)
        errors.extend(check_layout(d, content, settings, edition))
        if 'front_matter' in content:
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            if len(d.paragraphs) < 2 or d.paragraphs[1].alignment != WD_ALIGN_PARAGRAPH.RIGHT:
                errors.append('第二行授课日期与老师署名未右对齐')
            if not d.paragraphs or d.paragraphs[0].alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append('歌名总标题未居中')
            manual.append('核对歌名未混入歌手、实际授课日期和老师署名；逐条确认开头说明只列本稿使用的记号且含义正确。')
        if d.tables:
            errors.append('当前制作输入未包含表格，但成品含表格；核对是否选错文件')
        if len(d.paragraphs) != len(expected):
            errors.append('正文段落数与输入/所选老师版或学生版不一致')
        for number, (actual, (events, style)) in enumerate(zip(d.paragraphs, expected), 1):
            if _actual_events(actual, d) != events:
                errors.append(f'段落{number}：文字、斜体/删除线、图片与短说明的交错顺序、选图、裁图或参考框不符（{actual.text[:28]}）')
            if style and actual.style.name != style:
                errors.append(f'段落{number}：标题层级不符')
        drawings = d._element.xpath('.//w:drawing')
        if len(drawings) != len(geometry):
            errors.append('图片数量与输入的重复引用/省图关系不一致')
        for number, (drawing, ratio) in enumerate(zip(drawings, geometry), 1):
            extent = drawing.xpath('./wp:inline/wp:extent')
            if not extent:
                errors.append(f'图{number}：不是嵌入型图片或尺寸缺失')
                continue
            h, w = int(extent[0].get('cy')), int(extent[0].get('cx'))
            if abs(h / 914400 - settings['image_height']) > .002 or not h or abs(w/h/ratio-1) > .002:
                errors.append(f'图{number}：图高或所取区域比例与输入/版式不符；自定义版式须一并传入--layout')
    return dict(path=str(path), errors=errors, warnings=warnings, manual=manual,
                fonts=dict(fonts), sizes=dict(sizes), transcriptions=transcriptions, images=images)


def print_findings(result):
    print(result['path'], '\n转写', len(result['transcriptions']), '处；图片', result['images'],
          '幅\n字体分布', result['fonts'], '\n字号分布', result['sizes'])
    for label, key in (('错误', 'errors'), ('警告', 'warnings'), ('人工核对', 'manual')):
        print(label + '：')
        print('\n'.join(result[key]) if result[key] else '无')


def main():
    # CLI output is UTF-8 even when a Windows parent redirects it through a pipe.
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('files', nargs='+', type=Path)
    ap.add_argument('--height', type=float)
    ap.add_argument('--layout', type=Path, help='same layout used by production')
    ap.add_argument('--edition', choices=('student', 'teacher'), default='student')
    ap.add_argument('--source', type=Path, help='optional approved production input')
    ap.add_argument('--input-format', choices=('json',), default='json')
    args = ap.parse_args()
    if args.source and len(args.files) != 1:
        ap.error('--source只核对一个明确对应的文件')
    results = [inspect_docx(path, args.height, args.source, args.input_format, args.edition, args.layout) for path in args.files]
    for result in results:
        print_findings(result)
    values = [result['transcriptions'] for result in results]
    if len(values) == 2:
        if values[0] == values[1]:
            print('两版正文转写文字一致。')
        else:
            print('警告：两版正文转写有差异：')
            from itertools import zip_longest
            for i, (x, y) in enumerate(zip_longest(*values), 1):
                if x != y:
                    print(i, repr(x), '<>', repr(y))
    return 2 if any(result['errors'] for result in results) else 0


if __name__ == '__main__':
    raise SystemExit(main())
