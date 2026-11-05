"""Instruction entry points and portable local links; no song workspace access."""
from pathlib import Path
import re
import tempfile
import unittest
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parents[2]
GUIDES = (
    'docs/agent/转写与查词.md', 'docs/agent/讲义制作与复核.md',
    'docs/agent/资料维护与经验审核.md',
    'docs/tools/讲义制作输入与检查.md', 'docs/tools/构建与打包.md',
)


def missing_links(root, names):
    failures = []
    for name in names:
        path = root / name
        if not path.is_file():
            failures.append(name)
            continue
        text = re.sub(r'```.*?```', '', path.read_text(encoding='utf-8'), flags=re.S)
        for href in re.findall(r'\]\(([^)]+)\)', text):
            href = unquote(href.strip('<>'))
            if re.match(r'[a-zA-Z][a-zA-Z0-9+.-]*:', href):
                continue
            target, _, anchor = href.partition('#')
            dest = (path.parent / target).resolve() if target else path
            if not dest.is_file():
                failures.append(f'{name}: {href}')
            elif anchor and dest.suffix == '.md':
                headings = re.findall(r'^#{1,6}\s+(.+?)\s*$', dest.read_text(encoding='utf-8'), re.M)
                slugs = [re.sub(r'[^\w\-\s]', '', h.lower()).replace(' ', '-') for h in headings]
                if anchor not in slugs:
                    failures.append(f'{name}: #{anchor}')
    return failures


class DocumentationTests(unittest.TestCase):
    def test_entrypoints_and_links_resolve(self):
        names = ('AGENTS.md', 'scripts/TOOLING.md', 'README.md', '使用说明.md',
                 '开发参与与改进路线.md') + GUIDES
        self.assertEqual(missing_links(ROOT, names), [])

    def test_all_stage_guides_are_reachable_from_root(self):
        agent = (ROOT / 'AGENTS.md').read_text(encoding='utf-8')
        for guide in GUIDES:
            self.assertIn('](' + guide + ')', agent)
        self.assertIn('进入另一阶段前补读', agent)

    def test_missing_guide_and_broken_anchor_are_reported(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root/'entry.md').write_text('[a](阶段.md) [b](present.md#不存在)', encoding='utf-8')
            (root/'present.md').write_text('# 已存在', encoding='utf-8')
            self.assertEqual(len(missing_links(root, ['entry.md'])), 2)

    def test_workflow_uses_lf(self):
        workflow = ROOT / '.github/workflows/platform-packages.yml'
        if not workflow.exists():
            self.skipTest('构建工作流不随用户工作空间发行')
        self.assertNotIn(b'\r', workflow.read_bytes())
        self.assertTrue(workflow.read_bytes().endswith(b'\n'))
