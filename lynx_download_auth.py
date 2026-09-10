"""In-memory, origin-scoped download credentials and Google browser login."""
from dataclasses import dataclass, field
import shutil
import subprocess
import time
from urllib.parse import urlsplit
from urllib.request import Request, HTTPRedirectHandler, build_opener
from urllib.error import HTTPError


def origin(url):
    parsed = urlsplit(url)
    if parsed.scheme not in ('http', 'https') or not parsed.hostname or parsed.username or parsed.password:
        raise ValueError('Introduce una dirección HTTPS válida, sin usuario ni contraseña en la URL.')
    return parsed.scheme.lower(), parsed.hostname.lower(), parsed.port or (443 if parsed.scheme == 'https' else 80)


@dataclass(frozen=True)
class DownloadAuth:
    server: str
    method: str
    secret: str = field(repr=False)

    def __post_init__(self):
        if origin(self.server)[0] != 'https':
            raise ValueError('El servidor de autenticación debe usar HTTPS.')
        parsed = urlsplit(self.server)
        if parsed.query or parsed.fragment:
            raise ValueError('Introduce la dirección del servidor sin parámetros.')
        if self.method not in ('Agouti API key', 'Agouti Bearer', 'Trapper token'):
            raise ValueError('Método de acceso no admitido.')
        if not self.secret.strip() or any(c.isspace() for c in self.secret):
            raise ValueError('Introduce una clave o token válido, sin espacios.')

    def headers(self, url):
        if origin(url) != origin(self.server):
            return {}
        if self.method == 'Agouti API key':
            return {'X-API-KEY': self.secret}
        scheme = 'Bearer' if self.method == 'Agouti Bearer' else 'Token'
        return {'Authorization': scheme + ' ' + self.secret}


class ScopedRedirect(HTTPRedirectHandler):
    """Never forward credentials to another origin or allow an HTTPS downgrade."""
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        previous, target = origin(req.full_url), origin(newurl)
        if previous[0] == 'https' and target[0] != 'https':
            raise ValueError('Redirección HTTPS a HTTP rechazada.')
        redirected = super().redirect_request(req, fp, code, msg, headers, newurl)
        if redirected and previous != target:
            for key in list(redirected.headers):
                if key.lower() in ('authorization', 'x-api-key', 'cookie'):
                    del redirected.headers[key]
            for key in list(redirected.unredirected_hdrs):
                if key.lower() in ('authorization', 'x-api-key', 'cookie'):
                    del redirected.unredirected_hdrs[key]
        return redirected


def open_download(url, auth=None, timeout=20):
    origin(url)
    request = Request(url, headers=auth.headers(url) if auth else {})
    return build_opener(ScopedRedirect()).open(request, timeout=timeout)


def download_error(exc):
    if isinstance(exc, HTTPError) and exc.code in (401, 403):
        return f'HTTP {exc.code}: inicia sesión en Acceso a descargas o revisa los permisos de tu cuenta.'
    return type(exc).__name__


def google_login(task):
    executable = shutil.which('gcloud')
    if not executable:
        raise ValueError('Instala Google Cloud CLI para iniciar sesión; puedes usar fotos locales mientras tanto.')
    task.report('Completa el inicio de sesión en la ventana de Google del navegador.')
    # Google Cloud CLI owns the OAuth flow and credential store. No tokens enter app logs.
    with subprocess.Popen([executable, 'auth', 'login', '--brief'], stdin=subprocess.DEVNULL,
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL) as process:
        deadline = time.monotonic() + 300
        try:
            while process.poll() is None:
                task.checkpoint()
                if time.monotonic() >= deadline:
                    raise ValueError('El inicio de sesión ha agotado el tiempo. Vuelve a intentarlo.')
                task.stop.wait(0.1)
            task.checkpoint()
            if process.returncode:
                raise ValueError('Google no completó el inicio de sesión. Revisa el navegador y Google Cloud CLI.')
        finally:
            if process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait()
    return True
