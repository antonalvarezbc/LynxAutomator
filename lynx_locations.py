"""Optional Wildbook location catalogs, explicitly refreshed and cached locally."""
import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import urlopen

BRANCHES = {'Lynx': 'lynx', 'Whiskerbook': 'whiskerbook', 'Giraffe': 'giraffe',
            'Zebra': 'zebra', 'African Carnivore Wildbook': 'africancarnivorewildbook-acw',
            'Wild North': 'wildnorth', 'Deer': 'deer'}
CATALOGS = {name: f'https://raw.githubusercontent.com/WildMeOrg/Wildbook/{branch}/src/main/resources/bundles/locationID.json'
            for name, branch in BRANCHES.items()}
LIMIT = 5 * 1024 * 1024


def normalize_url(value):
    if value.startswith('https://github.com/') and '/blob/' in value:
        value = value.replace('https://github.com/', 'https://raw.githubusercontent.com/', 1).replace('/blob/', '/', 1)
    if urlparse(value).scheme != 'https':
        raise ValueError('El catálogo remoto debe usar HTTPS.')
    return value


def flatten(data):
    entries = []
    def visit(nodes, parents, depth=0):
        if depth > 30 or not isinstance(nodes, list):
            raise ValueError('Jerarquía locationID inválida.')
        for node in nodes:
            if not isinstance(node, dict) or not isinstance(node.get('id'), str) or not node['id'].strip():
                raise ValueError('Cada ubicación necesita un id de texto no vacío.')
            name = node.get('name') or node['id']
            if not isinstance(name, str):
                raise ValueError('Nombre de ubicación inválido.')
            route = parents + [name]
            entries.append({'id': node['id'], 'label': ' → '.join(route)})
            visit(node.get('locationID', []), route, depth + 1)
    if not isinstance(data, dict) or 'locationID' not in data:
        raise ValueError('El JSON no contiene locationID.')
    visit(data['locationID'], [])
    if not entries:
        raise ValueError('El catálogo está vacío.')
    return entries


def read_catalog(task, source, cache_dir, refresh=False):
    remote = source.startswith(('http://', 'https://'))
    source = normalize_url(source) if remote else str(Path(source).resolve())
    cache = Path(cache_dir) / (hashlib.sha256(source.encode()).hexdigest() + '.json')
    if cache.exists() and not refresh:
        data = json.loads(cache.read_text(encoding='utf-8'))
        flatten(data['catalog'])
        return data
    task.checkpoint()
    if remote:
        with urlopen(source, timeout=20) as response:
            if urlparse(response.geturl()).scheme != 'https':
                raise ValueError('Redirección de catálogo insegura.')
            raw = response.read(LIMIT + 1)
    else:
        with open(source, 'rb') as stream:
            raw = stream.read(LIMIT + 1)
    task.checkpoint()
    if len(raw) > LIMIT:
        raise ValueError('Catálogo demasiado grande.')
    catalog = json.loads(raw.decode('utf-8-sig'))
    flatten(catalog)
    result = {'source': source, 'updated': datetime.now(timezone.utc).isoformat(), 'catalog': catalog}
    cache.parent.mkdir(parents=True, exist_ok=True)
    temp = cache.with_suffix('.tmp')
    temp.write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
    os.replace(temp, cache)
    return result


def deployment_key(row):
    metadata = row.get('metadata', {})
    def value(*keys):
        for key in keys:
            if metadata.get(key):
                return metadata[key][0]
        return ''
    return json.dumps([value('deployments.project_id', 'deployment.project_id', 'folder.root'), value('deployments.deployment_id', 'deployment.deploymentID', 'deployment.deployment_id', 'folder.stationPath')], ensure_ascii=False)


def branch_url(branch):
    from urllib.parse import quote
    return 'https://raw.githubusercontent.com/WildMeOrg/Wildbook/' + quote(branch, safe='') + '/src/main/resources/bundles/locationID.json'


def github_branches(task, cache_dir, refresh=False):
    from urllib.request import Request
    path = Path(cache_dir) / 'branches.json'
    if path.exists() and not refresh:
        return json.loads(path.read_text(encoding='utf-8'))
    names = []
    for page in range(1, 101):
        task.checkpoint()
        request = Request(f'https://api.github.com/repos/WildMeOrg/Wildbook/branches?per_page=100&page={page}', headers={'User-Agent': 'LynxAutomator', 'Accept': 'application/vnd.github+json'})
        with urlopen(request, timeout=20) as response:
            data = json.loads(response.read(LIMIT))
        if not isinstance(data, list):
            raise ValueError('No se pudieron obtener las ramas de GitHub.')
        names.extend(item['name'] for item in data)
        if len(data) < 100:
            break
    else:
        raise ValueError('Demasiadas ramas; no se ha guardado una lista incompleta.')
    names = sorted(set(names))
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(names), encoding='utf-8')
    os.replace(temp, path)
    return names


def encounter_key(row):
    value = [row.get('eventID'), row.get('species'), row.get('individualID'), sorted(row.get('media', []))]
    return 'encounter:' + hashlib.sha256(json.dumps(value).encode()).hexdigest()


def hierarchy_paths(catalog):
    """Exact ID paths for offline common-ancestor lookup; ambiguous IDs are omitted."""
    flatten(catalog)
    paths, duplicate = {}, set()
    def visit(nodes, parents):
        for node in nodes:
            route = parents + [node['id']]
            if node['id'] in paths:
                duplicate.add(node['id'])
            paths[node['id']] = route
            visit(node.get('locationID', []), route)
    visit(catalog['locationID'], [])
    return {key: value for key, value in paths.items() if not duplicate.intersection(value)}


def export_filename(table, fields, now=None):
    """Suggest a portable name without changing any exported locationID."""
    import re
    field = next((f for f in fields if f['name'] == 'Encounter.locationID' and f.get('enabled', True)), {})
    values = {str(row.get('Encounter.locationID', '') or '').strip() for row in table}
    paths = field.get('location_hierarchy', {})
    location = 'sin-ubicacion'
    if values and '' not in values:
        if len(values) == 1:
            location = next(iter(values))
        else:
            routes = [paths.get(value, []) for value in values]
            common = []
            for level in zip(*routes):
                if len(set(level)) != 1:
                    break
                common.append(level[0])
            location = common[-1] if common else 'varias-ubicaciones'
    elif any(values):
        location = 'ubicaciones-incompletas'
    location = re.sub(r'[^\w.-]+', '_', location, flags=re.UNICODE).strip('._')[:100] or 'ubicacion'
    while len(location.encode('utf-8')) > 150:
        location = location[:-1]
    stamp = (now or datetime.now()).strftime('%Y-%m-%d_%H-%M-%S')
    return f'wildbook_bulk_import_{location}_{stamp}.xlsx'
