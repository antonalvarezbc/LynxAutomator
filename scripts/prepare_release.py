"""Collect all four CI distributions and supporting files for a pre-release."""
import hashlib
from pathlib import Path
import shutil
import sys

PACKAGES = {
    'ubuntu-22.04': 'LynxAutomator-Linux-x86_64.tar.gz',
    'windows-2022': 'LynxAutomator-Windows-AMD64.zip',
    'macos-15': 'LynxAutomator-Darwin-arm64.zip',
    'macos-15-intel': 'LynxAutomator-Darwin-x86_64.zip',
}


def prepare(artifacts, destination, root):
    sources = []
    for runner, package in PACKAGES.items():
        folder = artifacts / f'LynxAutomator-{runner}'
        sources.extend([(folder / package, package),
                        (folder / 'dependencies.txt', f'dependencies-{runner}.txt')])
    for manual in ('manual.md', 'manual.en.md', 'manual.pt.md'):
        sources.append((root / 'docs' / manual, manual))
    for source, _ in sources:
        if not source.is_file() or source.stat().st_size == 0:
            raise ValueError(f'Missing or empty release asset: {source}')
    destination.mkdir(parents=True, exist_ok=False)
    hashes = []
    for source, name in sources:
        target = destination / name
        shutil.copyfile(source, target)
        with target.open('rb') as stream:
            digest = hashlib.file_digest(stream, 'sha256').hexdigest()
        hashes.append(f'{digest}  {name}\n')
    (destination / 'SHA256SUMS.txt').write_text(''.join(hashes), encoding='utf-8')


if __name__ == '__main__':
    prepare(Path(sys.argv[1]), Path(sys.argv[2]), Path(__file__).resolve().parents[1])
