"""Resolve an explicitly selected local photo folder without ambiguous matches."""
from collections import defaultdict
from pathlib import Path
from urllib.parse import unquote, urlsplit


def photo_index(task, folder):
    root = Path(folder).resolve()
    if not root.is_dir():
        raise ValueError('La carpeta de fotografías no existe.')
    names = defaultdict(list)
    for path in root.rglob('*'):
        task.checkpoint()
        if path.is_file():
            names[path.name.casefold()].append(path)
    return names


def find_photo(index, name):
    candidates = index.get(name.casefold(), [])
    if len(candidates) != 1:
        raise ValueError('Fotografía ausente' if not candidates else 'Nombre ambiguo: varias fotografías locales')
    return candidates[0]


def dp_local_photo(folder, index, media):
    reference = media['filePath']
    if not urlsplit(reference).scheme:
        root = Path(folder).resolve()
        path = (root / reference).resolve()
        if path.is_relative_to(root) and path.is_file():
            return path
    name = media.get('fileName') or Path(unquote(urlsplit(reference).path)).name
    return find_photo(index, name)
