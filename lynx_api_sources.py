"""Read-only Agouti and Trapper exports feeding the common Camtrap DP pipeline."""
import json
import re
import shutil
import tempfile
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, urljoin, urlsplit
from lynx_camtrap import read_package, LIMIT
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


def fetch_agouti(task, server, project, folder, auth):
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


def fetch_trapper(task, server, project, folder, auth, approved_only=True):
    task.report('Trapper: solicitando la exportación Camtrap DP…')
    query = urlencode({'export_format': 'camtrapdp', 'export_filetype': 'csv.gz',
                       'approved_only': str(approved_only).lower(), 'release': 'false'})
    routes = ['/media_classification/api/package/', '/api/media-classifications/package/']
    for index, route in enumerate(routes):
        try:
            reply = remote_json(task, server + route + project + '/?' + query, auth, timeout=60)
            break
        except HTTPError as exc:
            if exc.code != 404 or index == len(routes) - 1:
                raise
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
    return read_package(task, path)


def fetch_api_package(task, provider, server, project, destination, auth=None, approved_only=True):
    server, project = api_base(server), project_key(project)
    if provider not in ('Agouti API', 'Trapper API'):
        raise ValueError('Origen API no admitido.')
    if not Path(destination).is_dir():
        raise ValueError('Selecciona una carpeta de destino existente.')
    folder = Path(tempfile.mkdtemp(prefix=provider.split()[0].lower() + '-', dir=destination))
    try:
        if provider == 'Agouti API':
            return fetch_agouti(task, server, project, folder, auth)
        return fetch_trapper(task, server, project, folder, auth, approved_only)
    except BaseException as exc:
        shutil.rmtree(folder)
        if isinstance(exc, HTTPError):
            if exc.code in (401, 403):
                raise ValueError(f'HTTP {exc.code}: este servidor requiere autorización o permisos para el proyecto. Configura la autorización opcional y vuelve a cargar.') from None
            raise ValueError(f'La API devolvió HTTP {exc.code}. Revisa el servidor y el identificador del proyecto.') from None
        if isinstance(exc, (URLError, TimeoutError)):
            raise ValueError('No se pudo obtener el paquete: revisa la conexión o vuelve a intentarlo si la exportación tarda demasiado.') from None
        raise
