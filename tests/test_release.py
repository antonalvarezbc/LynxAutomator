import hashlib
from pathlib import Path
import tempfile
import unittest
from scripts.prepare_release import PACKAGES, prepare


class ReleaseTests(unittest.TestCase):
    def test_all_platforms_and_checksums_are_required(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            artifacts = root / 'downloads'
            (root / 'docs').mkdir()
            for name in ('manual.md', 'manual.en.md', 'manual.pt.md'):
                (root / 'docs' / name).write_text('Manual')
            for runner, name in PACKAGES.items():
                directory = artifacts / f'LynxAutomator-{runner}'
                directory.mkdir(parents=True)
                (directory / name).write_bytes(b'archive fixture')
                (directory / 'dependencies.txt').write_text(runner)
            missing = artifacts / 'LynxAutomator-windows-2022' / PACKAGES['windows-2022']
            missing.unlink()
            with self.assertRaises(ValueError):
                prepare(artifacts, root / 'incomplete', root)
            self.assertFalse((root / 'incomplete').exists())
            missing.write_bytes(b'windows fixture')
            output = root / 'ready'
            prepare(artifacts, output, root)
            self.assertEqual(len(list(output.iterdir())), 12)
            for line in (output / 'SHA256SUMS.txt').read_text().splitlines():
                digest, name = line.split('  ', 1)
                self.assertEqual(digest, hashlib.sha256((output / name).read_bytes()).hexdigest())
