"""Independently inspect declared layout; never render a comparison document."""
from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING
from handout_content import marked_text


def font_value(run, paragraph, key):
    value = getattr(run.font, key)
    if value is not None:
        return value
    for style in (run.style, paragraph.style):
        while style:
            value = getattr(style.font, key)
            if value is not None:
                return value
            style = style.base_style


def font_contract(text, settings, options=None, marked=True):
    options = options or {}
    name = options.get('font_name', settings['font_name'])
    size = int(options.get('font_size', settings['font_size']) * 2) / 2
    parts = marked_text(text) if marked else [(text, False, False)]
    return [(c, 'Arial' if token == '~' and marked else name, size)
            for token, _, _ in parts for c in token]


def format_contract(content, settings, edition):
    rows = []
    roles = settings.get('content_styles', {})
    def row(spans, options=None, images=False, heading=False, framed=False):
        options = options or {}
        paragraph = {'space_after': 10 if heading else 7 if images else 5}
        if not heading:
            paragraph['line_spacing'] = 1.25 if images else settings['line_spacing']
        if images:
            paragraph.update(space_before=settings.get('image_space_before', 4),
                             space_after=settings.get('image_space_after', 7),
                             line_spacing=settings.get('image_line_spacing', 1.25))
        paragraph.update({k:v for k,v in options.items() if k in ('space_before','space_after','line_spacing')})
        if framed and settings.get('reference_frame_padding'):
            paragraph.pop('line_spacing', None)
            paragraph['minimum_line_pt'] = settings['image_height']*72 + settings['reference_frame_padding_pt']
        rows.append((spans, paragraph))
    front = content.get('front_matter')
    if front:
        row(font_contract(front['song'], settings, {'font_size':settings['title_size']},False),heading=True)
        options=roles.get('byline',{})
        row(font_contract('--'+front['date']+' by'+front['teachers'],settings,options,False),options)
        for note in front.get('notation',[]):
            options=roles.get('notation',{})
            row(font_contract(note,settings,options,False),options)
    for block in content['blocks']:
        if 'heading' in block:
            size=settings['title_size'] if block.get('level',2)==1 else settings['heading_size']
            row(font_contract(block['heading'],settings,{'font_size':size},False),heading=True)
            continue
        options=roles.get(block.get('role','body'),{})
        if 'transcription' in block and any(k in roles for k in ('lyrics','transcription')):
            spans=font_contract(block['lyrics'],settings,roles.get('lyrics',{}))
            spans+=font_contract('（'+block['transcription']+'）',settings,roles.get('transcription',{}))
        else:
            text=block.get('text',block.get('lyrics',''))
            if 'transcription' in block:text+='（'+block['transcription']+'）'
            spans=font_contract(text,settings,options)
        row(spans,options)
        visible=[i for i in block.get('sequence',[]) if i['type']!='omit']
        if visible:
            spans=[]
            for item in visible:
                if item['type']=='text':spans+=font_contract('〔'+item['text']+'〕',settings,roles.get('note',{}),False)
                spans+=font_contract(' ',settings)
            row(spans,images=True,framed=any(i.get('reference') for i in visible))
        if edition=='teacher' and block.get('teacher_note'):
            options=roles.get('teacher_note',roles.get('note',{}))
            row(font_contract('【'+block['teacher_note']+'】',settings,options,False),options)
    return rows


def check_layout(document, content, settings, edition):
    errors=[]
    for section in document.sections:
        for key, expected in [('left_margin',settings['margin_horizontal']),('right_margin',settings['margin_horizontal']),
                              ('top_margin',settings['margin_vertical']),('bottom_margin',settings['margin_vertical']),
                              ('page_width',21),('page_height',29.7)]:
            actual=getattr(section,key)
            if actual is None or abs(actual.cm-expected)>.005:
                errors.append('页面设置不符：'+key)
    for number,(p,(spans,paragraph)) in enumerate(zip(document.paragraphs,format_contract(content,settings,edition)),1):
        actual=[]
        for run in p.runs:
            size=font_value(run,p,'size')
            name=font_value(run,p,'name')
            # Resolve eastAsia separately; Word's western font name is insufficient.
            east=None
            for props in [run._r.rPr]+[style.element.rPr for style in (p.style,document.styles['Normal'])]:
                if props is not None:
                    fonts=props.find(qn('w:rFonts'))
                    if fonts is not None and fonts.get(qn('w:eastAsia')):
                        east=fonts.get(qn('w:eastAsia'));break
            actual.extend((c,east if east and '\u2e80'<=c<='\uffff' else name,size.pt if size else None) for c in run.text)
        if actual!=spans:
            errors.append(f'段落{number}：文字字体/字号或对应内容与版式不符')
        for key,value in paragraph.items():
            if key=='minimum_line_pt':
                current=p.paragraph_format.line_spacing
                if p.paragraph_format.line_spacing_rule!=WD_LINE_SPACING.AT_LEAST or current is None or abs(current.pt-value)>.06:
                    errors.append(f'段落{number}：参考框最小行高不符')
            elif key=='line_spacing':
                if p.paragraph_format.line_spacing_rule not in (None,WD_LINE_SPACING.SINGLE,WD_LINE_SPACING.ONE_POINT_FIVE,WD_LINE_SPACING.DOUBLE,WD_LINE_SPACING.MULTIPLE) or p.paragraph_format.line_spacing is None or abs(p.paragraph_format.line_spacing-value)>.01:
                    errors.append(f'段落{number}：行距不符')
            else:
                current=getattr(p.paragraph_format,key)
                if current is None or abs(current.pt-value)>.06:errors.append(f'段落{number}：{key}不符')
    for number,drawing in enumerate(document._element.xpath('.//w:drawing'),1):
        for line in drawing.xpath('.//pic:spPr/a:ln[a:prstDash[@val="sysDot"]]'):
            colors=list(line.iter(qn('a:srgbClr')))
            transparent=any(int(a.get('val','100000'))<100000 for a in line.iter(qn('a:alpha')))
            if int(line.get('w','0'))!=12700 or not colors or colors[0].get('val')!='000000' or list(line.iter(qn('a:noFill'))) or transparent:
                errors.append(f'图{number}：参考图框颜色、线宽或可见性不符')
    return errors
