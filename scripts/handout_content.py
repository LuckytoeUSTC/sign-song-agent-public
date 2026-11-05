"""Production data contract, independent of Word rendering and inspection."""
from pathlib import Path
import copy
import json
import math
from datetime import datetime
from docx.image.image import Image


def marked_text(text):
    """Yield visible text, italic and strike; ++ is always upright."""
    italic = strike = False
    index = 0
    while index < len(text):
        if text.startswith('~~', index):
            strike = not strike
            index += 2
            continue
        if text[index] == '*':
            italic = not italic
            index += 1
            continue
        if text.startswith('++', index):
            token = '++'
        elif text[index] == '~':
            token = '~'
        else:
            end = index + 1
            while end < len(text) and text[end] not in '*~' and not text.startswith('++', end):
                end += 1
            token = text[index:end]
        yield token, italic and token != '++', strike
        index += len(token)
    if italic or strike:
        raise ValueError('未闭合的斜体或删除线标记：' + text)


def check_image(item, base):
    if not isinstance(item, dict) or not isinstance(item.get('path'), str) or not item['path']:
        raise ValueError('每张选图须有path；不能用词条名代替已选文件')
    path = Path(base) / item['path']
    if not path.is_file():
        raise ValueError('选图不存在：' + str(path))
    image = Image.from_file(str(path))
    crop = item.get('crop')
    if crop is not None:
        if not isinstance(crop, list) or len(crop) != 4 or any(
                type(v) not in (int, float) or not math.isfinite(v) for v in crop):
            raise ValueError('crop须为原图像素坐标[left, top, right, bottom]')
        left, top, right, bottom = crop
        if not (0 <= left < right <= image.px_width and 0 <= top < bottom <= image.px_height):
            raise ValueError('裁图范围越界或为空：' + str(path))
    return path, image


BLOCK_FIELDS = {'id', 'repeat_of', 'omit_images', 'heading', 'level', 'text', 'role',
                'lyrics', 'transcription', 'sequence', 'sequence_exception', 'teacher_note', 'record'}
IMAGE_FIELDS = {'type', 'at', 'path', 'crop', 'reference', 'entry', 'sense',
                'source', 'action', 'alt', 'record', 'content_scale'}


def read_content(source, input_format='json'):
    source = Path(source).resolve()
    if input_format != 'json':
        raise ValueError('制作输入支持JSON；既有Word请使用docx副本模式')
    data = json.loads(source.read_text(encoding='utf-8'))
    if not isinstance(data, dict) or type(data.get('version')) is not int or data.get('version') != 2 or set(data) - {'version', 'blocks', 'front_matter'} or not isinstance(data.get('blocks'), list):
        raise ValueError('新制作JSON须含version: 2和blocks数组；旧制作数据不自动转换')
    resolved, by_id = [], {}
    for original in data['blocks']:
        if not isinstance(original, dict) or set(original) - BLOCK_FIELDS:
            raise ValueError('未知制作字段；使用sequence放顺次图示、teacher_note放老师版备注，不支持块级note/images')
        block = copy.deepcopy(original)
        if 'repeat_of' in block:
            target = block['repeat_of']
            if not isinstance(target, str) or target not in by_id:
                raise ValueError('重复段须引用前面已定义的id')
            if set(block) - {'id', 'repeat_of', 'omit_images', 'teacher_note', 'record'}:
                raise ValueError('重复段不另存正文或图示；修改原块或取消repeat_of')
            inherited = copy.deepcopy(by_id[target])
            inherited.pop('id', None)
            inherited.update(block)
            block = inherited
            if block.get('omit_images', False):
                block['sequence'] = []
                block.pop('sequence_exception', None)
        elif 'omit_images' in block:
            raise ValueError('omit_images仅用于repeat_of重复段')
        if 'id' in block:
            if not isinstance(block['id'], str) or not block['id'] or block['id'] in by_id:
                raise ValueError('制作块id须为唯一的非空字符串')
            by_id[block['id']] = copy.deepcopy(block)
        resolved.append(block)
    result = dict(base=source.parent, blocks=resolved)
    if 'front_matter' in data:
        result['front_matter'] = data['front_matter']
    return result


def check_front_matter(front):
    if not isinstance(front, dict) or set(front) - {'song', 'teachers', 'date', 'notation'}:
        raise ValueError('front_matter只含song、teachers、date和可选notation')
    for key in ('song', 'teachers', 'date'):
        value = front.get(key)
        if not isinstance(value, str) or not value.strip() or any(c in value for c in '\r\n'):
            raise ValueError('front_matter.' + key + '须为非空单行文字；姓名和授课日期须确认，不编造')
    try:
        date = datetime.strptime(front['date'], '%Y.%m.%d')
        if date.strftime('%Y.%m.%d') != front['date']:
            raise ValueError()
    except ValueError:
        raise ValueError('授课日期须为有效的YYYY.MM.DD，如2026.10.17') from None
    notation = front.get('notation', [])
    if not isinstance(notation, list) or any(not isinstance(s, str) or not s.strip() or '\n' in s or '\r' in s for s in notation):
        raise ValueError('notation为记号说明的单行文字数组；不用的说明不填写')
    return front


def check_content(content):
    if 'front_matter' in content:
        check_front_matter(content['front_matter'])
        if any('heading' in b and b.get('level', 2) == 1 for b in content['blocks']):
            raise ValueError('front_matter已生成歌名总标题，blocks不再放level:1标题')
    for block in content['blocks']:
        for key in ('lyrics', 'transcription', 'text', 'heading', 'teacher_note', 'record', 'role', 'sequence_exception'):
            if key in block and not isinstance(block[key], str):
                raise ValueError(key + '须为文字')
        if 'heading' in block:
            if set(block) - {'heading', 'level', 'id'} or type(block.get('level', 2)) is not int or block.get('level', 2) not in (1, 2, 3):
                raise ValueError('标题只含heading、level和id；level为1至3，默认2（段落小标题）')
        elif 'text' in block:
            if any(k in block for k in ('lyrics', 'transcription', 'sequence')):
                raise ValueError('text正文块与歌词转写图示块分开')
        elif not all(k in block for k in ('lyrics', 'transcription')):
            raise ValueError('歌曲句块须同时提供lyrics与transcription')
        if 'heading' not in block and 'level' in block:
            raise ValueError('level仅用于标题块')
        if 'heading' in block and any(c in block['heading'] for c in '*~'):
            raise ValueError('heading填写普通标题文字，不使用转写标记')
        for key in ('heading', 'text', 'lyrics', 'transcription'):
            if key in block:
                list(marked_text(block[key]))
        if 'teacher_note' in block and any(c in block['teacher_note'] for c in '【】'):
            raise ValueError('teacher_note填写正文，程序统一加【】')
        if 'omit_images' in block and type(block['omit_images']) is not bool:
            raise ValueError('omit_images须为true或false')
        sequence = block.get('sequence', [])
        if not isinstance(sequence, list):
            raise ValueError('sequence须为有序数组')
        if 'sequence_exception' in block:
            if 'transcription' not in block or not block['sequence_exception'].strip() or not sequence:
                raise ValueError('sequence_exception只用于有图示链的歌曲句，须说明具体组织例外')
        slots = block.get('transcription', '').split('/')
        previous, covered, omitted, displayed = 0, set(), set(), set()
        for item in sequence:
            if not isinstance(item, dict):
                raise ValueError('图示项目须为对象')
            kind = item.get('type')
            fields = IMAGE_FIELDS if kind == 'image' else {'type', 'at', 'text', 'record'} if kind == 'text' else {'type', 'at', 'reason'} if kind == 'omit' else set()
            if not fields or set(item) - fields:
                raise ValueError('sequence项目为image、text或omit，不支持图片note；短说明用独立text项目')
            at = item.get('at')
            positions = at if isinstance(at, list) else [at]
            if (not positions or any(type(n) is not int or not 1 <= n <= len(slots) for n in positions)
                    or positions != list(range(positions[0], positions[-1] + 1)) or (positions[0] < previous and not block.get('sequence_exception'))):
                raise ValueError('at为从1起的转写词位或连续词位数组，按非递减顺序排列')
            previous = positions[0]
            covered.update(positions)
            (omitted if kind == 'omit' else displayed).update(positions)
            for key in fields - {'type', 'at', 'crop', 'reference', 'content_scale'}:
                if key in item and not isinstance(item[key], str):
                    raise ValueError('图示' + key + '须为文字')
            if kind == 'image':
                check_image(item, content['base'])
                scale = item.get('content_scale', 1)
                if type(scale) not in (int, float) or not math.isfinite(scale) or not 0 < scale <= 1:
                    raise ValueError('content_scale须为0至1之间的有限数（不含0），如0.7')
                if scale < 1 and any(v != int(v) for v in item.get('crop', [])):
                    raise ValueError('补白时crop须为整数像素边界，避免隐式舍入')
                if 'reference' in item and type(item['reference']) is not bool:
                    raise ValueError('reference须为true或false')
            elif not item.get('text' if kind == 'text' else 'reason', '').strip():
                raise ValueError('text须有短说明；omit须有内部省略理由')
        if sequence and (covered != set(range(1, len(slots) + 1)) or any(not s.strip() for s in slots)):
            raise ValueError('图示须交代每个转写词位；无图说明用text，有意省略用omit；不留空词位')
        if omitted & displayed:
            raise ValueError('同一词位不能同时声明omit和可见图文')
    return content
