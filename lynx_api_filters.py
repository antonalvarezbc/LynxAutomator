"""Explicit deployment selection shared by the API sources."""
from datetime import datetime, timezone


def normalize_filters(values=None):
    values = values or {}
    result = {key: str(values.get(key, '')).strip() for key in ('year', 'site', 'deployment', 'latest')}
    if result['year'] and (not result['year'].isdigit() or not 1 <= int(result['year']) <= 9999):
        raise ValueError('Año / Year: 1–9999.')
    if result['latest'] and (not result['latest'].isdigit() or int(result['latest']) < 1):
        raise ValueError('Últimos despliegues / Latest deployments: mínimo 1.')
    result['exclude_blank'] = bool(values.get('exclude_blank', False))
    if 'deployment_ids' in values:
        ids = values['deployment_ids']
        if not isinstance(ids, (list, tuple)) or not all(isinstance(value, str) and value for value in ids):
            raise ValueError('Selecciona identificadores de despliegue válidos.')
        result['deployment_ids'] = list(dict.fromkeys(ids))
    return result


def selected_deployments(rows, filters):
    filters = normalize_filters(filters)
    matches = []
    for row in rows:
        if 'deployment_ids' in filters and row.get('deploymentID') not in filters['deployment_ids']:
            continue
        stamp = None
        try:
            stamp = datetime.fromisoformat(str(row.get('deploymentStart', '')).replace('Z', '+00:00'))
        except ValueError:
            pass
        if filters['year'] and (stamp is None or stamp.year != int(filters['year'])):
            continue
        if filters['site'] and filters['site'].casefold() not in str(row.get('locationName', '')).casefold():
            continue
        # Trapper's documented filter uses a substring, not a list of exact IDs.
        if filters['deployment'] and filters['deployment'] not in str(row.get('deploymentID', '')):
            continue
        if filters['latest'] and stamp is None:
            continue
        sort_stamp = stamp.replace(tzinfo=timezone.utc) if stamp is not None and stamp.tzinfo is None else stamp
        matches.append((sort_stamp, row))
    if filters['latest']:
        matches.sort(key=lambda pair: (pair[0], pair[1]['deploymentID']), reverse=True)
        matches = matches[:int(filters['latest'])]
    if not matches:
        raise ValueError('Ningún despliegue coincide con los filtros. Revisa año de inicio, sitio e ID; fechas ausentes no cumplen año/últimos.')
    return [row for _, row in matches]


def filter_package(package, filters):
    """Local refinement of a received ZIP; retain its original path for media access."""
    import copy
    package = copy.copy(package)
    selected = selected_deployments(package.deployments, filters)
    ids = {row['deploymentID'] for row in selected}
    package.deployments = selected
    package.media = [row for row in package.media if row['deploymentID'] in ids]
    package.observations = [row for row in package.observations if row['deploymentID'] in ids]
    return package


# Common deployment facets; absent columns are not offered.
FACET_COLUMNS = ('deploymentID', 'locationName', 'locationID', 'cameraID', 'cameraModel',
                 'habitat', 'setupBy', 'deploymentTags', 'captureMethod', 'baitUse')


def deployment_year(row):
    try:
        return str(datetime.fromisoformat(str(row.get('deploymentStart', '')).replace('Z', '+00:00')).year)
    except ValueError:
        return ''


def deployment_facets(rows):
    fields = ('deploymentStart.year',) + FACET_COLUMNS
    result = {}
    for key in fields:
        counts = {}
        for row in rows:
            value = deployment_year(row) if key == 'deploymentStart.year' else str(row.get(key) or '').strip()
            if value:
                counts[value] = counts.get(value, 0) + 1
        if counts:
            result[key] = dict(sorted(counts.items()))
    return result


def match_facets(rows, choices):
    """OR within one variable; AND between variables. Empty choice means unrestricted."""
    result = []
    for row in rows:
        if all(not selected or (deployment_year(row) if key == 'deploymentStart.year' else str(row.get(key) or '').strip()) in selected
               for key, selected in choices.items()):
            result.append(row)
    return result
