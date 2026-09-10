"""Folder/catalog input adapters for the common Bulk Import engine."""
from pathlib import Path
from lynx_bulk import normalize, attach_metadata, group_rows
from lynx_processing import read_exif, date_taken


def folder_rows(task, folder, species, recursive=False, identity='none', fallback_year=None, group=False, threshold=3, allow_undated=False):
    root = Path(folder)
    rows, missing = [], []
    for path in sorted(root.rglob('*') if recursive else root.glob('*')):
        task.checkpoint()
        if not path.is_file() or path.suffix.lower() not in ('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.webp'):
            continue
        exif = read_exif(path)
        date = date_taken(exif)
        if date is None and fallback_year is None and not allow_undated:
            missing.append(str(path.relative_to(root)) + ' (sin fecha EXIF)')
            continue
        individual = path.stem.split()[0] if identity == 'filename' else path.parent.name if identity == 'folder' else ''
        row = normalize(species, date or (f'{fallback_year:04d}-01-01' if fallback_year else '2000-01-01'), path, str(path.parent), path.name,
                        str(root.resolve()), individual, locality='', latitude='', longitude='')
        if date is None:
            for field in ('month', 'day', 'hour', 'minutes'):
                row[field] = ''
            if fallback_year is None:
                row['year'] = ''
                row['timestamp'] = ''
                row['_allow_undated'] = True
            row['_multiple'] = True  # Year alone cannot establish a temporal sequence.
        attach_metadata(row, deployment=[{'deploymentID': str(path.parent), 'locationName': path.parent.name}],
                        media=[{'filename': path.name, 'relativePath': str(path.relative_to(root)),
                                'dateSource': 'EXIF' if date else 'user year' if fallback_year else 'unknown'}])
        rows.append(row)
    if group:
        dated = [row for row in rows if row['timestamp']]
        undated = [row for row in rows if not row['timestamp']]
        rows = group_rows(task, dated, threshold) + undated
    return rows, missing
