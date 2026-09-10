"""Shared, template-free Wildbook export. No GUI dependencies."""
import csv
import hashlib
import json
import math
import os
import sys
import zipfile
import io
import re
from pathlib import Path
from urllib.parse import urlparse, unquote

import pandas as pd
from openpyxl import Workbook

if getattr(sys, 'frozen', False):
    settings = Path(os.environ.get('LOCALAPPDATA') or os.environ.get('XDG_CONFIG_HOME') or (Path.home() / '.config')) / 'LynxAutomator'
else:
    settings = Path(__file__).resolve().parent / '.local-settings'
PROFILE = settings / 'bulk-import.json'
SOURCES = ['genus', 'specificEpithet', 'latitude', 'longitude', 'locality', 'year',
           'month', 'day', 'hour', 'minutes', 'eventID', 'individualID', 'species']


def default_fields():
    pairs = [('Encounter.genus', 'genus'), ('Encounter.specificEpithet', 'specificEpithet'),
             ('Encounter.decimalLatitude', 'latitude'), ('Encounter.decimalLongitude', 'longitude'),
             ('Encounter.verbatimLocality', 'locality')]
    pairs += [('Encounter.' + name, name) for name in ['year', 'month', 'day', 'hour', 'minutes']]
    pairs += [('Encounter.sightingID', 'eventID'), ('MarkedIndividual.individualID', 'individualID')]
    integers = {'year', 'month', 'day', 'hour', 'minutes'}
    fields = [dict(name=n, source=s, value='', type='integer' if s in integers else
                   'decimal' if s in {'latitude', 'longitude'} else 'text', enabled=n != 'MarkedIndividual.individualID') for n, s in pairs]
    fields += [dict(name=n, source='fixed', value='', type='text', enabled=True) for n in
               ['Encounter.locationID', 'Encounter.country', 'Encounter.submitterID']]
    return fields


def read_profiles(path=PROFILE):
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding='utf-8'))
    if data.get('version') == 1:
        profiles = {'Default': data['fields']}
    elif data.get('version') == 2 and isinstance(data.get('profiles'), dict):
        profiles = data['profiles']
    else:
        raise ValueError('Perfil Bulk Import incompatible.')
    for fields in profiles.values():
        validate_fields(fields)
    return profiles


def load_profile(path=PROFILE, name='Default'):
    return read_profiles(path).get(name, default_fields())


def save_profile(fields, path=PROFILE, name='Default'):
    validate_fields(fields)
    name = name.strip()
    if not name:
        raise ValueError('Escribe un nombre para el perfil.')
    profiles = read_profiles(path)
    profiles[name] = fields
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps({'version': 2, 'profiles': profiles}, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(temp, path)


def validate_fields(fields):
    names = set()
    for field in fields:
        if not field.get('enabled', True):
            continue
        name = field['name'].strip()
        if not name or name in names or re.fullmatch(r'Encounter\.mediaAsset\d+', name):
            raise ValueError('Nombre vacío, duplicado o reservado para fotografías: ' + name)
        names.add(name)
        if (field['source'] not in SOURCES + ['fixed', 'template', 'location_map'] and not re.fullmatch(r'[A-Za-z_][\w-]*\.[\w.-]+', field['source'])) or field['type'] not in ['text', 'integer', 'decimal', 'boolean']:
            raise ValueError('Origen o tipo de campo desconocido: ' + name)


def attach_metadata(row, **tables):
    metadata = row.setdefault('metadata', {})
    for prefix, records in tables.items():
        for record in records:
            for key, value in record.items():
                values = metadata.setdefault(prefix + '.' + key, [])
                if value is not None and str(value) != '' and str(value) not in values:
                    values.append(str(value))


def merge_metadata(target, row):
    for name in row.get('_unmatched_tables', []):
        if name not in target.setdefault('_unmatched_tables', []):
            target['_unmatched_tables'].append(name)
    for key, values in row.get('metadata', {}).items():
        dest = target.setdefault('metadata', {}).setdefault(key, [])
        dest.extend(v for v in values if v not in dest)


def source_value(row, source):
    if source in SOURCES:
        return row.get(source, '')
    if source not in row.get('metadata', {}):
        raise ValueError('Metadato no disponible: ' + source)
    return ' | '.join(row['metadata'][source])


def template_value(row, template):
    # Literal substitution only: no evaluation, attribute access or executable expressions.
    def replace(match):
        return str(source_value(row, match.group(1)))
    remaining = re.sub(r'\{([\w.-]+)\}', '', template)
    if '{' in remaining or '}' in remaining:
        raise ValueError('Usa marcadores como {deployment.cameraID}.')
    return re.sub(r'\{([\w.-]+)\}', replace, template)


def required_warnings(out):
    warnings = []
    for key in ['Encounter.genus', 'Encounter.specificEpithet', 'Encounter.year', 'Encounter.mediaAsset0']:
        if out.get(key) in (None, ''):
            warnings.append('Falta ' + key)
    if not (out.get('Encounter.locationID') or out.get('Encounter.verbatimLocality') or
            (out.get('Encounter.decimalLatitude') not in (None, '') and out.get('Encounter.decimalLongitude') not in (None, ''))):
        warnings.append('Falta ubicación: Encounter.verbatimLocality, Encounter.locationID o ambas coordenadas.')
    return warnings


def normalize(species, timestamp, media, deployment, event, namespace, individual='', **extra):
    parts = str(species).split()
    if len(parts) < 2 or parts[1].lower() in ('sp.', 'spp.', 'sp', 'spp'):
        raise ValueError('Se necesita una especie binomial: ' + str(species))
    date = pd.Timestamp(timestamp)
    if pd.isna(date):
        raise ValueError('Fecha de captura ausente.')
    key = json.dumps([namespace, deployment, event], ensure_ascii=False)
    result = dict(species=species, genus=parts[0], specificEpithet=parts[1],
                  year=date.year, month=date.month, day=date.day, hour=date.hour, minutes=date.minute,
                  eventID=hashlib.sha256(key.encode()).hexdigest()[:24], individualID=individual,
                  media=[str(media)], timestamp=str(date), **extra)
    return result


def camtrap_rows(task, package, records, batches, group=True, threshold=3):
    extras = {}
    for resource in package.descriptor.get('resources', []):
        name = resource.get('name', '')
        if name in ('deployments', 'media', 'observations') or not name:
            continue
        if isinstance(resource.get('data'), list):
            extras[name] = resource['data']
        elif str(resource.get('path', '')).lower().endswith('.csv'):
            with package.open_local(resource['path']) as stream:
                raw = stream.read(100 * 1024 * 1024 + 1)
            if len(raw) > 100 * 1024 * 1024:
                raise ValueError('Tabla adicional demasiado grande: ' + name)
            extras[name] = pd.read_csv(io.BytesIO(raw), dtype=str).fillna('').to_dict('records')
    available = {}
    for batch in batches:
        with (Path(batch) / 'manifest.csv').open(encoding='utf-8', newline='') as stream:
            for item in csv.DictReader(stream):
                path = Path(batch) / item['file']
                if item['status'] == 'completed' and path.is_file() and path.resolve().is_relative_to(Path(batch).resolve()):
                    available[item['mediaID']] = path
    deployments = {d['deploymentID']: d for d in package.deployments}
    observations = {o['observationID']: o for o in package.observations}
    rows, missing, groups = [], [], {}
    namespace = package.descriptor.get('id') or hashlib.sha256(json.dumps(
        [package.descriptor, package.deployments, [(m['mediaID'], m.get('timestamp')) for m in package.media]],
        sort_keys=True, ensure_ascii=False).encode()).hexdigest()
    for record in sorted(records, key=lambda r: r['media'].get('timestamp', '')):
        task.checkpoint()
        media = record['media']
        mid = media['mediaID']
        if mid not in available:
            missing.append(mid)
            continue
        dep = deployments[media['deploymentID']]
        # Associate each species with its own event/individual, never cross species provenance.
        associations = set()
        singles = set()
        for oid in record['observation_ids']:
            obs = observations[oid]
            species = obs.get('scientificName')
            if species not in record['species']:
                continue
            event = obs.get('eventID') or (json.dumps([obs.get('eventStart'), obs.get('eventEnd')])
                                                  if obs.get('observationLevel') == 'event' else mid)
            individual = obs.get('individualID') or ''
            if float(obs.get('count') or 1) > 1:
                singles.add((species, individual))
            associations.add((species, event, individual))
        associations = {(sp, mid if (sp, ind) in singles else ev, ind) for sp, ev, ind in associations}
        # A direct observation and its event may refer to the same photo/species.
        associations = {a for a in associations if a[1] != mid or not any(
            b[0] == a[0] and b[2] == a[2] and b[1] != mid for b in associations)}
        for species, event, individual in sorted(associations):
            row = normalize(species, media.get('timestamp'), available[mid], media['deploymentID'],
                            event, namespace, individual,
                            latitude=dep.get('latitude', ''), longitude=dep.get('longitude', ''),
                            locality=dep.get('locationName', ''))
            attach_metadata(row, deployment=[dep], media=[media], observation=[observations[oid] for oid in sorted(record['observation_ids']) if observations[oid].get('scientificName') == species and (observations[oid].get('individualID') or '') == individual])
            attach_extra_tables(row, extras, dict(dep, **media))
            attach_metadata(row, package=[{k: v for k, v in package.descriptor.items() if k != 'resources'}])
            row['_event'] = event if event != mid else None
            row['_multiple'] = (species, individual) in singles
            key = (mid, species, individual, event)
            if key in groups:
                target = groups[key]
                merge_metadata(target, row)
                if row['media'][0] not in target['media']:
                    target['media'].extend(row['media'])
            else:
                groups[key] = row
                rows.append(row)
    return (group_rows(task, rows, threshold) if group else rows), missing


def read_wi_tables(images_path, deployments_path=None):
    """Read two explicit CSVs or exactly one pair in a ZIP, without extraction."""
    if str(images_path).lower().endswith('.zip'):
        with zipfile.ZipFile(images_path) as archive:
            tables = []
            folders = set()
            for name in ('images.csv', 'deployments.csv'):
                matches = [i for i in archive.infolist() if (re.fullmatch(r'images(?:_[\w-]+)?\.csv', Path(i.filename).name.lower()) if name == 'images.csv' else Path(i.filename).name.lower() == name)
                           and not i.is_dir() and not i.filename.startswith('__MACOSX/')]
                if not matches or (name != 'images.csv' and len(matches) != 1):
                    raise ValueError('El ZIP debe contener images.csv o images_<proyecto>.csv y un único deployments.csv.')
                if len({Path(i.filename).parent for i in matches}) > 1 or len({Path(i.filename).name.lower() for i in matches}) != len(matches):
                    raise ValueError('El ZIP debe contener un único conjunto de CSV, sin carpetas duplicadas.')
                folders.update(Path(i.filename).parent for i in matches)
                if len(folders) != 1:
                    raise ValueError('Los CSV de imágenes y deployments deben pertenecer a la misma carpeta del ZIP.')
                if sum(i.file_size for i in matches) > 100 * 1024 * 1024:
                    raise ValueError('CSV demasiado grande: ' + name)
                pieces = [pd.read_csv(io.BytesIO(archive.read(i)), dtype=str).fillna('') for i in sorted(matches, key=lambda i: i.filename)]
                tables.append(pd.concat(pieces, ignore_index=True).fillna(''))
            return tables
    if not deployments_path:
        raise ValueError('Selecciona images.csv y deployments.csv, o un ZIP con ambos.')
    return [pd.read_csv(path, dtype=str).fillna('') for path in (images_path, deployments_path)]


def read_extra_tables(images_path, extra_paths=()):
    tables = {}
    def add(name, stream):
        name = re.sub(r'[^\w-]', '_', name)
        if name in tables:
            raise ValueError('Nombre de tabla adicional duplicado: ' + name)
        table = pd.read_csv(stream, dtype=str).fillna('')
        tables[name] = table.to_dict('records')
    if str(images_path).lower().endswith('.zip'):
        with zipfile.ZipFile(images_path) as archive:
            total = 0
            for item in archive.infolist():
                name = Path(item.filename).name
                if item.is_dir() or item.filename.startswith('__MACOSX/') or not name.lower().endswith('.csv') or re.fullmatch(r'images(?:_[\w-]+)?\.csv|deployments\.csv', name.lower()):
                    continue
                total += item.file_size
                if total > 100 * 1024 * 1024:
                    raise ValueError('Tablas adicionales demasiado grandes.')
                add(Path(name).stem, io.BytesIO(archive.read(item)))
    for path in extra_paths:
        if Path(path).stat().st_size > 100 * 1024 * 1024:
            raise ValueError('CSV adicional demasiado grande.')
        add(Path(path).stem, path)
    return tables


def attach_extra_tables(row, tables, context):
    # Follow explicit foreign keys in either naming convention. Never join on names.
    def ids(record):
        return {re.sub(r'[^a-z0-9]', '', key.lower()): str(value)
                for key, value in record.items()
                if value not in ('', None) and (key.endswith('_id') or key.endswith('ID'))}
    known = {key: {value} for key, value in ids(context).items()}
    matched = {name: [] for name in tables}
    seen = set()
    for _ in range(len(tables) + 1):
        candidates = []
        for name, records in tables.items():
            for index, record in enumerate(records):
                if (name, index) in seen:
                    continue
                identifiers = ids(record)
                common = identifiers.keys() & known.keys()
                # Project is a scope, not proof that a camera/deployment belongs to this row.
                useful = common - {'projectid', 'organizationid'}
                scoped = common and not (identifiers.keys() - {'projectid', 'organizationid'})
                global_row = not identifiers and len(records) == 1
                if not global_row and not ((useful or scoped) and all(identifiers[k] in known[k] for k in common)):
                    continue
                candidates.append((name, index, record, identifiers))
        # Resolve each hop together: table/row order must not pick an arbitrary project.
        ambiguous = set()
        for scope in ('projectid', 'organizationid'):
            values = {identifiers[scope] for _, _, _, identifiers in candidates if scope in identifiers}
            if scope not in known and len(values) > 1:
                ambiguous.add(scope)
        accepted = [candidate for candidate in candidates if not ambiguous.intersection(candidate[3])]
        if not accepted:
            break
        for name, index, record, identifiers in accepted:
            seen.add((name, index))
            matched[name].append(record)
            for key, value in identifiers.items():
                known.setdefault(key, set()).add(value)
    for name, records in tables.items():
        if records and not matched[name]:
            row.setdefault('_unmatched_tables', []).append(name)
        attach_metadata(row, **{name: matched[name]})
        for record in records[:1]:
            for key in record:
                row.setdefault('metadata', {}).setdefault(name + '.' + key, [])


def group_rows(task, rows, threshold):
    from lynx_locations import deployment_key
    import copy
    result, active, events = [], {}, {}
    for original in sorted(rows, key=lambda r: pd.Timestamp(r['timestamp']).value):
        task.checkpoint()
        row = copy.deepcopy(original)
        key = (deployment_key(row), row['species'], row.get('individualID', ''))
        stamp = pd.Timestamp(row['timestamp'])
        target = None
        if row.get('_multiple'):
            active.pop(key, None)
        elif row.get('_event'):
            active.pop(key, None)
            event_key = key + (row['_event'],)
            target = events.get(event_key)
            if target is None:
                events[event_key] = row
        else:
            previous = active.get(key)
            if previous and 0 <= (stamp - previous[1]).total_seconds() <= threshold:
                target = previous[0]
            active[key] = (target if target is not None else row, stamp)
        if target is None:
            result.append(row)
        else:
            target['media'].extend(m for m in row['media'] if m not in target['media'])
            merge_metadata(target, row)
    return result


def wi_scientific_name(item):
    return item.get('scientific_name') or item.get('scientificName') or ' '.join(str(item.get(k, '')) for k in ('genus', 'species')).strip()


def wi_species(task, images_path, deployments_path=None):
    from collections import Counter
    images, _ = read_wi_tables(images_path, deployments_path)
    counts = Counter()
    for row in images.to_dict('records'):
        task.checkpoint()
        name = wi_scientific_name(row)
        parts = name.split()
        if len(parts) >= 2 and parts[1].lower() not in ('sp', 'sp.', 'spp', 'spp.'):
            counts[name] += 1
    return counts


def wi_rows(task, images_path, deployments_path=None, group=False, threshold=3, extra_paths=(), species=None):
    from lynx_core import merge_deployments
    task.checkpoint()
    images, deployments = read_wi_tables(images_path, deployments_path)
    extras = read_extra_tables(images_path, extra_paths)
    for table, required in [(images, ['project_id', 'deployment_id', 'location', 'timestamp']),
                            (deployments, ['project_id', 'deployment_id'])]:
        absent = set(required) - set(table.columns)
        if absent:
            raise ValueError('Faltan columnas CSV: ' + ', '.join(sorted(absent)))
    if species is not None:
        images = images[images.apply(lambda row: wi_scientific_name(row) in species, axis=1)]
    data = merge_deployments(images, deployments)
    deployment_lookup = {(r['project_id'], r['deployment_id']): r for r in deployments.to_dict('records')}
    rows, missing, groups = [], [], {}
    data = data.sort_values('timestamp')
    for _, item in data.iterrows():
        task.checkpoint()
        name = Path(unquote(urlparse(item['location']).path)).name
        if not name:
            raise ValueError('Una fila de images.csv no tiene nombre de fotografía en location.')
        species_name = wi_scientific_name(item)
        stamp = pd.Timestamp(item['timestamp'])
        event = str(item.get('image_id') or item['location'])
        row = normalize(species_name, stamp, str(item['location']), item['deployment_id'], event,
                        item['project_id'], latitude=item.get('latitude', ''),
                        longitude=item.get('longitude', ''), locality=item.get('placename', ''), media_local=False)
        attach_metadata(row, deployment=[deployment_lookup[(item['project_id'], item['deployment_id'])]], media=[item.to_dict()])
        row['individualID'] = item.get('individual_id', '')
        row['_multiple'] = float(item.get('number_of_objects') or 1) > 1
        context = dict(deployment_lookup[(item['project_id'], item['deployment_id'])], **item.to_dict())
        attach_extra_tables(row, extras, context)
        rows.append(row)
    return (group_rows(task, rows, threshold) if group else rows), missing


def location_value(row, field):
    from lynx_locations import deployment_key, encounter_key
    key = deployment_key(row)
    if encounter_key(row) in field.get("locations", {}):
        return field["locations"][encounter_key(row)]
    value = field.get('locations', {}).get(key, field.get('value', ''))
    if not value:
        raise ValueError('Falta locationID para el despliegue ' + key)
    return value


def convert(value, kind):
    if value is None or str(value).strip() == '':
        return None
    if kind == 'text':
        return str(value)
    if kind == 'boolean':
        if str(value).lower() not in ['true', 'false', '1', '0']:
            raise ValueError('Boolean: true/false/1/0')
        return str(value).lower() in ['true', '1']
    number = float(value)
    if not math.isfinite(number) or (kind == 'integer' and not number.is_integer()):
        raise ValueError('Número inválido: ' + str(value))
    return int(number) if kind == 'integer' else number


def build_table(rows, fields, task=None):
    validate_fields(fields)
    if not rows:
        raise ValueError('No hay fotografías disponibles para exportar.')
    table, names = [], {}
    for index, row in enumerate(rows, 1):
        if task:
            task.checkpoint()
        out = {}
        for field in fields:
            if field.get('enabled', True):
                value = (field.get('value', '') if field['source'] == 'fixed' else
                         template_value(row, field.get('value', '')) if field['source'] == 'template' else
                         location_value(row, field) if field['source'] == 'location_map' else source_value(row, field['source']))
                out[field['name'].strip()] = convert(value, field['type'])
                if isinstance(value, str) and len(value) > 32767:
                    raise ValueError(f'Fila {index}: {field["name"]} supera el límite de texto de Excel (32767). Reduce los metadatos seleccionados o desagrupa las fotografías.')
        for n, filename in enumerate(row['media']):
            local = row.get('media_local', True)
            path = Path(filename) if local else Path(unquote(urlparse(filename.replace('\\', '/')).path))
            if local and not path.is_file():
                raise ValueError('Fotografía ausente: ' + filename)
            identity = str(path.resolve()) if local else filename
            if path.name in names and names[path.name] != identity:
                raise ValueError('Dos fotografías distintas tienen el mismo nombre: ' + path.name)
            names[path.name] = identity
            out[f'Encounter.mediaAsset{n}'] = path.name
        warnings = required_warnings(out)
        if row.get('_allow_undated'):
            warnings = [warning for warning in warnings if warning != 'Falta Encounter.year']
        if warnings:
            raise ValueError(f'Fila {index}: ' + '\n'.join(warnings))
        for key, low, high in [('Encounter.decimalLatitude', -90, 90), ('Encounter.decimalLongitude', -180, 180),
                               ('Encounter.year', 1, 9999), ('Encounter.month', 1, 12), ('Encounter.day', 1, 31),
                               ('Encounter.hour', 0, 23), ('Encounter.minutes', 0, 59)]:
            value = out.get(key)
            if value not in (None, '') and not low <= float(value) <= high:
                raise ValueError(f'Fila {index}: {key} fuera de rango.')
        table.append(out)
    return table


def write_excel(task, table, destination):
    destination = Path(destination)
    temp = destination.with_name(destination.name + '.tmp')
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Bulk Import'
    headers = list(dict.fromkeys(k for row in table for k in row))
    sheet.append(headers)
    for cell in sheet[1]:
        cell.data_type = 's'
    sheet.freeze_panes = 'A2'
    for row in table:
        task.checkpoint()
        sheet.append([row.get(k) for k in headers])
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = 's'  # Preserve literal user/source text, never Excel formulas.
    try:
        workbook.save(temp)
        task.checkpoint()
        os.replace(temp, destination)
    finally:
        temp.unlink(missing_ok=True)
