"""Acquire selected WI media, preserving source references and actual local names."""
from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote, urlparse
import csv
import os
import subprocess
import tempfile

from PIL import Image
from lynx_bulk import group_rows
from lynx_tasks import TaskCancelled


@dataclass
class Acquisition:
    available: dict = field(default_factory=dict)
    errors: list = field(default_factory=list)
    failed: list = field(default_factory=list)
    manifest: str = ''


def media_name(source):
    name = Path(unquote(urlparse(source.replace('\\', '/')).path)).name
    if name in ('', '.', '..') or any(char in name for char in '/\\\r\n\0'):
        raise ValueError('Referencia de fotografía inválida: ' + source)
    if Path(name).suffix.lower() not in {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.webp', '.bmp', '.gif'}:
        raise ValueError('Extensión de fotografía no admitida: ' + name)
    return name


def sources_for(rows):
    sources = sorted({str(source) for row in rows for source in row['media']})
    names = {}
    for source in sources:
        name = media_name(source).casefold()
        if name in names and names[name] != source:
            raise ValueError('Referencias distintas tienen el mismo nombre de fotografía: ' + media_name(source))
        names[name] = source
    return sources


def verify_photo(path):
    with Image.open(path) as image:
        image.verify()


def acquire_wi(task, sources, directory, executable=None):
    """Use local files, or download into a fresh batch without overwriting files."""
    sources = sources_for([{'media': sources}])
    root = Path(directory).resolve()
    if not root.is_dir():
        raise ValueError('La carpeta no existe.')
    result = Acquisition()
    names = defaultdict(list)
    if executable is None:
        task.report('Buscando fotografías locales…')
        for path in root.rglob('*'):
            task.checkpoint()
            if path.is_file():
                names[path.name.casefold()].append(path)
    else:
        root = Path(tempfile.mkdtemp(prefix='wi-photos-', dir=root))
        result.manifest = str(root / 'manifest.csv')
    manifest = []
    try:
        for index, source in enumerate(sources):
            task.checkpoint()
            task.report(f'{index + 1}/{len(sources)}: {media_name(source)}', index / max(1, len(sources)))
            entry = {'source': source, 'file': '', 'status': 'failed', 'error': ''}
            try:
                name = media_name(source)
                if executable is None:
                    candidates = names[name.casefold()]
                    if len(candidates) != 1:
                        raise ValueError('Fotografía ausente' if not candidates else 'Nombre ambiguo: varias fotografías locales')
                    target = candidates[0]
                    verify_photo(target)
                else:
                    if not source.startswith('gs://') or any(char in source for char in '\r\n\0*?[]'):
                        raise ValueError('Se requiere una referencia gs:// sin comodines.')
                    target = root / name
                    if target.exists():
                        raise ValueError('Nombre de destino duplicado: ' + name)
                    with tempfile.TemporaryDirectory(dir=root, prefix='.transfer-') as staging:
                        temporary = Path(staging) / name
                        subprocess.run([executable, 'cp', source, str(temporary)], check=True,
                                       capture_output=True, text=True, timeout=300)
                        verify_photo(temporary)
                        os.replace(temporary, target)
                result.available[source] = str(target.resolve())
                entry.update(file=target.name, status='completed')
            except TaskCancelled:
                raise
            except Exception as exc:
                detail = str(getattr(exc, 'stderr', None) or exc)
                result.errors.append(source + ': ' + detail)
                result.failed.append(source)
                entry['error'] = detail
            manifest.append(entry)
    finally:
        if result.manifest:
            with Path(result.manifest).open('w', encoding='utf-8', newline='') as stream:
                writer = csv.DictWriter(stream, fieldnames=['source', 'file', 'status', 'error'])
                writer.writeheader()
                writer.writerows(manifest)
    return result


def acquired_rows(task, rows, available, group=False, threshold=3):
    output, missing = [], []
    for original in rows:
        task.checkpoint()
        row = deepcopy(original)
        paths = []
        for source in row['media']:
            path = available.get(source)
            if path and Path(path).is_file():
                paths.append(path)
            else:
                missing.append(source)
        if paths:
            row['media'] = paths
            row['media_local'] = True
            output.append(row)
    return (group_rows(task, output, threshold) if group else output), sorted(set(missing))
