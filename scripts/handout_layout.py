"""Layout settings and coverage registry shared by production and audit."""
from pathlib import Path
import json
import math

DEFAULT_LAYOUT = dict(font_name='SimSun', font_size=11, title_size=16,
                      heading_size=13, image_height=.79, margin_horizontal=2,
                      margin_vertical=1.8, line_spacing=1.15)
STYLE_FIELDS = {'font_name', 'font_size', 'space_before', 'space_after', 'line_spacing'}
LAYOUT_COVERAGE = {
    **{key: 'program+page' for key in ('font_name','font_size','title_size','heading_size',
       'image_height','margin_horizontal','margin_vertical','line_spacing')},
    **{key: 'program+page' for key in ('image_space_before', 'image_space_after',
       'image_line_spacing', 'reference_frame_padding', 'reference_frame_padding_pt', 'content_styles')},
    'paragraph_overrides': 'copy-test+page', 'reference_frame_rows': 'copy-test+page'
}

def validate_paragraph_style(options):
    if not isinstance(options, dict) or set(options) - STYLE_FIELDS:
        raise ValueError('段落格式只支持字体、字号、段前段后及行距')
    for key, value in options.items():
        if key == 'font_name':
            if not isinstance(value, str) or not value.strip():
                raise ValueError('字体须为非空文字')
        elif type(value) not in (int, float) or not math.isfinite(value) or (
                value < 0 if key.startswith('space_') else value <= 0):
            raise ValueError('无效段落格式：' + key)


def read_layout(path=None, height=None):
    settings = DEFAULT_LAYOUT.copy()
    settings['reference_frame_padding'] = False
    settings.update(content_styles={}, paragraph_overrides={}, reference_frame_rows=[],
                    reference_frame_padding_pt=6)
    if path:
        supplied = json.loads(Path(path).read_text(encoding='utf-8'))
        if not isinstance(supplied, dict):
            raise ValueError('版式JSON须为对象')
        unknown = set(supplied) - set(settings) - {'image_space_before', 'image_space_after', 'image_line_spacing'}
        if unknown:
            raise ValueError('未知版式参数：' + ', '.join(sorted(unknown)))
        settings.update(supplied)
    for key, value in settings.items():
        if key in ('content_styles', 'paragraph_overrides'):
            if not isinstance(value, dict):
                raise ValueError(key + '须为对象')
            for name, options in value.items():
                if key == 'paragraph_overrides' and (not str(name).isdigit() or int(name) < 1):
                    raise ValueError('段落序号从1开始')
                validate_paragraph_style(options)
                if key == 'content_styles' and name in ('lyrics', 'transcription', 'note') and set(options) - {'font_name', 'font_size'}:
                    raise ValueError(name + '为同段内文字角色，只接受字体字号；段距行距请设置body或image_*')
            continue
        if key == 'reference_frame_rows':
            if not isinstance(value, list) or any(type(n) is not int or n < 1 for n in value):
                raise ValueError('reference_frame_rows须为从1开始的段落序号数组')
            continue
        if key in ('image_space_before', 'image_space_after', 'reference_frame_padding_pt'):
            if type(value) not in (int, float) or not math.isfinite(value) or value < 0:
                raise ValueError(key + '须为非负有限数')
            continue
        if key == 'reference_frame_padding':
            if not isinstance(value, bool):
                raise ValueError('reference_frame_padding须为true或false')
            continue
        if key == 'font_name':
            if not isinstance(value, str) or not value.strip():
                raise ValueError('字体须为非空文字')
            continue
        if type(value) not in (int, float) or not math.isfinite(value) or value <= 0:
            raise ValueError('版式参数必须是正数：' + key)
    if height is not None:
        if type(height) not in (int, float) or not math.isfinite(height):
            raise ValueError('图高须为有限数')
        settings['image_height'] = height
    if not .1 <= settings['image_height'] <= 3:
        raise ValueError('图片高度应为0.1至3英寸')
    return settings
