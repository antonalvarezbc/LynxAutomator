"""Read-only Agouti and Trapper exports feeding the common Camtrap DP pipeline."""
import csv
import gzip
import io
import json
import re
import shutil
import tempfile
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, urljoin, urlsplit, parse_qsl, urlunsplit
from lynx_camtrap import read_package, LIMIT
from lynx_api_filters import normalize_filters, selected_deployments, filter_package
from lynx_download_auth import open_download, origin


def api_base(server):
    server = server.strip().rstrip('/')
    if origin(server)[0] != 'https' or urlsplit(server).query or urlsplit(server).fragment:
        raise ValueError('Introduce el servidor HTTPS sin parámetros ni credenciales en la URL.')
    return server


def project_key(project):
    project = project.strip()
    if not re.fullmatch(r'[A-Za-z0-9_-]+', project):
        raise ValueError('Introduce un identificador de proyecto válido.')
    return quote(project, safe='')


def read_remote(task, url, auth, limit=LIMIT, timeout=30):
    task.checkpoint()
    if origin(url)[0] != 'https':
        raise ValueError('Los datos de la API deben usar HTTPS.')
    chunks, size = [], 0
    with open_download(url, auth, timeout=timeout) as stream:
        while True:
            task.checkpoint()
            chunk = stream.read(64 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > limit:
                raise ValueError('La respuesta de la API supera el tamaño admitido.')
            chunks.append(chunk)
    task.checkpoint()
    return b''.join(chunks)


def remote_json(task, url, auth, timeout=30):
    try:
        value = json.loads(read_remote(task, url, auth, limit=10 * 1024 * 1024, timeout=timeout))
    except (ValueError, UnicodeError) as exc:
        raise ValueError('La API no devolvió un documento JSON válido o excedió el límite de tamaño.') from exc
    if not isinstance(value, dict):
        raise ValueError('La API no devolvió un objeto JSON.')
    return value


def fetch_agouti(task, server, project, folder, auth, filters=None):
    prefix = server + '/v1/projects/' + project + '/'
    descriptor_url = prefix + 'datapackage.json'
    task.report('Agouti: leyendo descriptor y tablas del proyecto…')
    descriptor = remote_json(task, descriptor_url, auth)
    if 'resources' not in descriptor and isinstance(descriptor.get('data'), dict):
        descriptor = descriptor['data']
    resources = descriptor.get('resources')
    if not isinstance(resources, list):
        raise ValueError('Agouti no devolvió un descriptor Camtrap DP con recursos.')
    total = 0
    def table_rows(resource, deployment=None):
        nonlocal total
        if 'data' in resource:
            rows = resource['data']
            if not isinstance(rows, list):
                raise ValueError('Recurso API inválido.')
            return rows
        path = resource.get('path')
        if not isinstance(path, str) or not path:
            raise ValueError('Recurso API sin ruta CSV.')
        url = urljoin(prefix, path)
        if deployment is not None:
            parts = urlsplit(url)
            query = [(key, value) for key, value in parse_qsl(parts.query, keep_blank_values=True) if key != 'deploymentID']
            query.append(('deploymentID', deployment))
            url = urlunsplit(parts._replace(query=urlencode(query)))
        raw = read_remote(task, url, auth)
        total += len(raw)
        if total > 300 * 1024 * 1024:
            raise ValueError('Las tablas de Agouti superan el tamaño admitido.')
        if urlsplit(url).path.endswith('.gz'):
            with gzip.GzipFile(fileobj=io.BytesIO(raw)) as stream:
                raw = stream.read(LIMIT + 1)
        if len(raw) > LIMIT:
            raise ValueError('Tabla API demasiado grande.')
        reader = csv.DictReader(io.StringIO(raw.decode('utf-8-sig')))
        rows = []
        for row in reader:
            task.checkpoint()
            rows.append(row)
        return rows
    if filters and any(filters.get(key) for key in ('year', 'site', 'deployment', 'latest')):
        by_name = {resource.get('name'): resource for resource in resources}
        if not all(name in by_name for name in ('deployments', 'media', 'observations')):
            raise ValueError('Faltan las tablas Camtrap DP para filtrar.')
        task.report('Agouti: seleccionando despliegues antes de pedir media y observaciones…')
        deployments = selected_deployments(table_rows(by_name['deployments']), filters)
        ids = {row['deploymentID'] for row in deployments}
        by_name['deployments'].pop('path', None)
        by_name['deployments']['data'] = deployments
        for name in ('media', 'observations'):
            resource = by_name[name]
            if 'data' in resource:
                rows = [row for row in resource['data'] if row.get('deploymentID') in ids]
            else:
                rows = []
                for index, deployment in enumerate(sorted(ids), 1):
                    task.report(f'Agouti: {name}, despliegue {index}/{len(ids)}…')
                    subset = table_rows(resource, deployment)
                    if any(row.get('deploymentID') != deployment for row in subset):
                        raise ValueError('El servidor no respetó el filtro deploymentID. Exporta una selección en Agouti y carga el paquete desde Camtrap DP.')
                    rows.extend(subset)
            resource.pop('path', None)
            resource['data'] = rows
    for index, resource in enumerate(resources):
        task.checkpoint()
        if 'data' in resource:
            continue
        path = resource.get('path')
        if not isinstance(path, str) or not path:
            raise ValueError('El descriptor contiene un recurso sin una ruta única.')
        url = urljoin(prefix, path)
        suffix = '.csv.gz' if urlsplit(url).path.endswith('.csv.gz') else '.csv' if urlsplit(url).path.endswith('.csv') else '.json'
        content = read_remote(task, url, auth)
        total += len(content)
        if total > 300 * 1024 * 1024:
            raise ValueError('Las tablas de Agouti superan el tamaño admitido.')
        filename = f'resource-{index}{suffix}'
        (folder / filename).write_bytes(content)
        resource['path'] = filename
    descriptor_path = folder / 'datapackage.json'
    descriptor_path.write_text(json.dumps(descriptor, ensure_ascii=False), encoding='utf-8')
    package = read_package(task, descriptor_path)
    # A remote export's relative media paths are URLs, not files on this computer.
    changed = False
    for media in package.media:
        if not urlsplit(media['filePath']).scheme:
            media['filePath'] = urljoin(prefix, media['filePath'])
            changed = True
    if changed:
        media_resource = next(resource for resource in resources if resource['name'] == 'media')
        media_resource.pop('path', None)
        media_resource['data'] = package.media
        descriptor_path.write_text(json.dumps(descriptor, ensure_ascii=False), encoding='utf-8')
        return read_package(task, descriptor_path)
    return package


def fetch_trapper(task, server, project, folder, auth, approved_only=True, filters=None):
    task.report('Trapper: solicitando la exportación Camtrap DP…')
    params = {'export_format': 'camtrapdp', 'export_filetype': 'csv.gz',
              'approved_only': str(approved_only).lower(), 'release': 'false'}
    if filters:
        if filters.get('deployment'):
            params['filter_deployments'] = filters['deployment']
        if filters.get('exclude_blank'):
            params['exclude_blank'] = 'true'
    query = urlencode(params)
    routes = ['/media_classification/api/package/', '/api/media-classifications/package/']
    for index, route in enumerate(routes):
        try:
            reply = remote_json(task, server + route + project + '/?' + query, auth, timeout=60)
            break
        except HTTPError as exc:
            if exc.code != 404 or index == len(routes) - 1:
                raise
            if filters and (filters.get('deployment') or filters.get('exclude_blank')):
                raise ValueError('La ruta actual de Trapper no está disponible; no se aplican filtros remotos de semántica distinta en la ruta antigua. Exporta el ZIP desde Trapper y cárgalo en Camtrap DP.') from None
    data = reply.get('data')
    link = data.get('package') if isinstance(data, dict) else None
    if not isinstance(link, str) or not link:
        raise ValueError('Trapper no devolvió el enlace del paquete. Revisa el proyecto y sus identificaciones.')
    url = urljoin(server + '/', link)
    task.report('Trapper: descargando el paquete…')
    # Authentication is origin-scoped even when the response links to another host.
    archive = read_remote(task, url, auth, limit=300 * 1024 * 1024, timeout=60)
    path = folder / 'camtrap.zip'
    path.write_bytes(archive)
    package = read_package(task, path)
    return filter_package(package, filters) if filters and any(filters.get(key) for key in ('year', 'site', 'deployment', 'latest')) else package


def fetch_api_package(task, provider, server, project, destination, auth=None, approved_only=True, filters=None):
    filters = normalize_filters(filters)
    server, project = api_base(server), project_key(project)
    if provider not in ('Agouti API', 'Trapper API'):
        raise ValueError('Origen API no admitido.')
    if not Path(destination).is_dir():
        raise ValueError('Selecciona una carpeta de destino existente.')
    folder = Path(tempfile.mkdtemp(prefix=provider.split()[0].lower() + '-', dir=destination))
    try:
        if provider == 'Agouti API':
            return fetch_agouti(task, server, project, folder, auth, filters)
        return fetch_trapper(task, server, project, folder, auth, approved_only, filters)
    except BaseException as exc:
        shutil.rmtree(folder)
        if isinstance(exc, HTTPError):
            if exc.code in (401, 403):
                raise ValueError(f'HTTP {exc.code}: este servidor requiere autorización o permisos para el proyecto. Configura la autorización y vuelve a cargar.') from None
            raise ValueError(f'La API devolvió HTTP {exc.code}. Revisa el servidor y el identificador del proyecto.') from None
        if isinstance(exc, (URLError, TimeoutError)):
            raise ValueError('No se pudo obtener el paquete: revisa la conexión o vuelve a intentarlo si la exportación tarda demasiado.') from None
        raise
