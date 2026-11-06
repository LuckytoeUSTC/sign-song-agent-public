"""Build the public workspace or the public offline rule browser."""
import argparse
import hashlib
import json
import re
from html import escape
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

RELEASE = Path(__file__).resolve().parent
ROOT = RELEASE.parents[1]
RULES = '资料/规则与例句数据库/公共转写规则.md'
COURSE = '资料/小班讲义/26春手语小班知识纯享版'


def source(name):
    path = ROOT / name
    if not path.is_file() or not path.resolve().is_relative_to(ROOT):
        raise ValueError(f'缺少文件或文件位于项目外：{name}')
    return path.read_bytes()


def write_zip(output, files):
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, 'x', ZIP_DEFLATED) as archive:
        for name, data in sorted(files.items()):
            archive.writestr(name, data)
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    output.with_suffix('.zip.sha256').write_text(f'{digest}  {output.name}\n', encoding='ascii')


def workspace(output, preview):
    names = json.loads((RELEASE / 'public-files.json').read_text(encoding='utf-8'))
    if not preview:
        materials = [COURSE + ext for ext in ('.pdf', '.docx') if (ROOT / (COURSE + ext)).is_file()]
        if len(materials) != 1:
            raise ValueError('正式包需要且只接受一份审核通过的纯享版 PDF 或 DOCX。')
        names += materials
    files = {name: source(name) for name in names}
    if not {'AGENTS.md', RULES, '导入自己的备课资料.md'} <= files.keys():
        raise ValueError('公开清单缺少备课必需文件。')
    manifest = {'version': 'v3.2.0-preview' if preview else 'v3.2.0',
                'files': {name: {'sha256': hashlib.sha256(data).hexdigest()}
                          for name, data in files.items()}}
    files['发行文件清单.json'] = json.dumps(manifest, ensure_ascii=False, indent=2).encode()
    suffix = '-preview' if preview else ''
    write_zip(output / f'sign-song-agent-v3.2.0-public{suffix}.zip',
              {'手语歌智能工坊/' + name: data for name, data in files.items()})


def offline(output):
    raw = source(RULES)
    rules = re.findall(r'^## (R\d+) ([^\n]+)\n(.*?)(?=^## |\Z)',
                       raw.decode('utf-8-sig').replace('\r\n', '\n'), re.M | re.S)
    if [number for number, _, _ in rules] != [f'R{i:02d}' for i in range(1, 15)]:
        raise ValueError('公开离线包应包含 R01–R14。')
    cards = '\n'.join(f'<article data-rule="{number}"><h2>{number} {escape(title)}</h2>'
                      f'<pre>{escape(body.strip())}</pre></article>' for number, title, body in rules)
    page = (RELEASE / 'no_ai_template.html').read_text(encoding='utf-8').replace('__RULES__', cards)
    readme = ('# NO-AI-ZONE v3.2.0\n\n解压后双击「公共转写规则查询.html」。'
              '页面支持离线查阅 14 条公共转写规则及正文中的例句，可按编号或关键词筛选。'
              'Windows、macOS、Linux 共用一份 ZIP，无需 Python。\n')
    write_zip(output / 'no-ai-zone-v3.2.0.zip',
              {'NO-AI-ZONE/公共转写规则查询.html': page.encode(),
               'NO-AI-ZONE/公共转写规则.md': raw,
               'NO-AI-ZONE/README.md': readme.encode()})


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('kind', choices=('preview', 'workspace', 'no-ai'))
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    if args.kind == 'no-ai':
        offline(args.out)
    else:
        workspace(args.out, preview=args.kind == 'preview')
