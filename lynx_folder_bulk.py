"""Folder/catalog input adapters for the common Bulk Import engine."""
from pathlib import Path
from lynx_bulk import normalize, attach_metadata, group_rows
from lynx_processing import read_exif, date_taken


def folder_rows(task, folder, species, recursive=False, identity='none', fallback_year=None, group=False, threshold=3, allow_undated=False, levels=None, stations=None):
    root = Path(folder)
    rows, missing = [], []
    stations = validate_stations(stations or {})
    for path in photo_paths(task, root, recursive or bool(levels)):
        task.checkpoint()
        if not path.is_file() or path.suffix.lower() not in ('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.webp'):
            continue
        exif = read_exif(path)
        date = date_taken(exif)
        if date is None and fallback_year is None and not allow_undated:
            missing.append(str(path.relative_to(root)) + ' (sin fecha EXIF)')
            continue
        values = layout_values(path, root, levels) if levels else {}
        station_key = values.get('stationPath', str(path.parent.relative_to(root))) or '.'
        station = stations.get(station_key, {})
        locality = station.get('locality', values.get('locality', ''))
        row_species = values.get('species') or species
        individual = path.stem.split()[0] if identity == 'filename' else path.parent.name if identity == 'folder' else ''
        individual = values.get('individual') or individual
        row = normalize(row_species, date or (f'{fallback_year:04d}-01-01' if fallback_year else '2000-01-01'), path, station_key, str(path.relative_to(root)),
                        str(root.resolve()), individual, locality=locality, latitude=station.get('latitude', ''), longitude=station.get('longitude', ''))
        if date is None:
            for field in ('month', 'day', 'hour', 'minutes'):
                row[field] = ''
            if fallback_year is None:
                row['year'] = ''
                row['timestamp'] = ''
                row['_allow_undated'] = True
            row['_multiple'] = True  # Year alone cannot establish a temporal sequence.
        row['_source_kind'] = 'folder'
        folder_meta = dict(root=str(root.resolve()), stationPath=station_key, station=values.get('station', path.parent.name), locality=locality)
        folder_meta.update({f'level{index}': name for index, name in enumerate(path.relative_to(root).parts[:-1], 1)})
        attach_metadata(row, folder=[folder_meta], station=[station],
                        file=[{'filename': path.name, 'relativePath': str(path.relative_to(root)),
                               'dateSource': 'EXIF' if date else 'user year' if fallback_year else 'unknown'}])
        rows.append(row)
    if group:
        dated = [row for row in rows if row['timestamp']]
        undated = [row for row in rows if not row['timestamp']]
        rows = group_rows(task, dated, threshold) + undated
    return rows, missing


PHOTO_SUFFIXES = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.webp'}


def photo_paths(task, folder, recursive=True):
    root = Path(folder)
    if not root.is_dir():
        raise ValueError('Selecciona una carpeta existente.')
    paths = []
    for path in sorted(root.rglob('*') if recursive else root.glob('*')):
        task.checkpoint()
        if path.is_file() and path.suffix.lower() in PHOTO_SUFFIXES:
            if not path.resolve().is_relative_to(root.resolve()):
                continue
            paths.append(path)
    return paths


def layout_values(path, root, levels):
    parts = path.relative_to(root).parts[:-1]
    def value(key):
        level = int(levels.get(key, 0))
        if level < 0 or level > len(parts):
            raise ValueError(f'{path.relative_to(root)}: falta el nivel {level} de {key}.')
        return parts[level - 1] if level else ''
    result = {key: value(key) for key in ('locality', 'station', 'species', 'individual')}
    level = int(levels.get('station', 0))
    result['stationPath'] = '/'.join(parts[:level]) if level else '/'.join(parts) or '.'
    return result


def scan_layout(task, folder, levels):
    root = Path(folder)
    stations, examples, errors, depth = {}, [], [], 0
    paths = photo_paths(task, root)
    for path in paths:
        depth = max(depth, len(path.relative_to(root).parts) - 1)
        if len(examples) < 8:
            examples.append(str(path.relative_to(root)))
        try:
            values = layout_values(path, root, levels)
        except ValueError as exc:
            errors.append(str(exc))
            continue
        entry = stations.setdefault(values['stationPath'], dict(locality=values['locality'], station=values['station'], latitude='', longitude='', count=0))
        if entry['locality'] != values['locality']:
            errors.append(f"{values['stationPath']}: varias localidades para la misma estación; revisa los niveles.")
        entry['count'] += 1
    return dict(stations=stations, depth=depth, examples=examples, errors=errors, count=len(paths))


def validate_stations(stations):
    import math
    result = {}
    for key, values in stations.items():
        item = dict(values)
        lat, lon = (str(values.get(name, '')).strip() for name in ('latitude', 'longitude'))
        if bool(lat) != bool(lon):
            raise ValueError(f'{key}: completa latitud y longitud, o deja ambas vacías.')
        for name, value, limit in (('latitude', lat, 90), ('longitude', lon, 180)):
            if value:
                try:
                    number = float(value.replace(',', '.'))
                except ValueError:
                    raise ValueError(f'{key}: coordenada inválida ({name}).') from None
                if not math.isfinite(number) or abs(number) > limit:
                    raise ValueError(f'{key}: coordenada fuera de rango ({name}).')
                item[name] = number
            else:
                item[name] = ''
        result[key] = item
    return result
