"""File ownership and unified archive tests, with tiny synthetic files only."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
TOOL = ROOT / 'scripts/release/build_packages.py'
if not TOOL.exists():
    raise unittest.SkipTest('开发打包工具不随用户工作空间发行')
spec = importlib.util.spec_from_file_location('package_builder_test', TOOL)
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)


class PackageTests(unittest.TestCase):
    def test_ownership_and_no_stale_private_or_platform_content(self):
        policy = json.loads((ROOT/'项目文件归属.json').read_text(encoding='utf-8'))
        self.assertEqual(builder.file_category('工作台/本曲/截图.png',policy),'personal')
        self.assertEqual(builder.file_category('个人设置/本地要求.md',policy),'personal')
        self.assertEqual(builder.file_category('资料/老师的修改.docx',policy),'shared_materials')
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp)/'source';root.mkdir()
            files=['AGENTS.md','scripts/make_handout.py','scripts/handout_content.py',
                   'scripts/handout_images.py','scripts/handout_layout.py','scripts/handout_format_audit.py','项目文件归属.json','资料/规则与例句查询程序.pyw',
                   'docs/agent/转写与查词.md','docs/agent/讲义制作与复核.md',
                   'docs/agent/资料维护与经验审核.md','docs/tools/讲义制作输入与检查.md',
                   'docs/tools/构建与打包.md','工作台/使用说明.md','工作台/本曲/截图.png','个人资料/经验.md',
                   '个人设置/本地要求.md','dictionaries/自加词典.mdx','debug.log',
                   'scripts/__pycache__/old.pyc','.github/config.yml','scripts/release/private.txt',
                   '规则与例句查询.exe','STOP-AI.jpg','资料/公共原件.pdf']
            for name in files:
                path=root/name;path.parent.mkdir(parents=True,exist_ok=True);path.write_text('fixture',encoding='utf-8')
            (root/'项目文件归属.json').write_text(json.dumps(policy),encoding='utf-8')
            (root/'资料/未跟踪文件.pdf').write_text('not tracked',encoding='utf-8')
            with patch.object(builder,'SRC',root),patch.object(builder,'OUTPUT',Path(temp)/'artifacts'),patch.object(builder.subprocess,'check_output',return_value=('\0'.join(files)+'\0').encode()):
                archive=builder.build_full()
                self.assertEqual(archive.name,'sign-song-agent-development-workspace.zip')
                with ZipFile(archive) as z:
                    names={n.removeprefix('手语歌智能工坊/') for n in z.namelist()}
                    self.assertEqual(names,{'AGENTS.md','scripts/make_handout.py','scripts/handout_content.py',
                                           'scripts/handout_images.py','scripts/handout_layout.py','scripts/handout_format_audit.py','项目文件归属.json','资料/规则与例句查询程序.pyw',
                                           'docs/agent/转写与查词.md','docs/agent/讲义制作与复核.md',
                                           'docs/agent/资料维护与经验审核.md','docs/tools/讲义制作输入与检查.md',
                                           'docs/tools/构建与打包.md','工作台/使用说明.md','资料/公共原件.pdf','发行文件清单.json'})
                    manifest=json.loads(z.read('手语歌智能工坊/发行文件清单.json'))
                    self.assertEqual(manifest['files']['资料/公共原件.pdf']['category'],'shared_materials')
                # Rebuilding from a smaller file set cannot retain old stage files.
                with patch.object(builder.subprocess,'check_output',return_value=('\0'.join(files[:-1])+'\0').encode()):
                    builder.build_full()
                with ZipFile(archive) as z:self.assertNotIn('手语歌智能工坊/资料/公共原件.pdf',z.namelist())

                # A missing stage guide must fail before producing an incomplete archive.
                with patch.object(builder.subprocess,'check_output',return_value=('\0'.join(n for n in files if n!='docs/agent/转写与查词.md')+'\0').encode()):
                    with self.assertRaisesRegex(ValueError,'docs/agent/转写与查词'):
                        builder.build_full()
