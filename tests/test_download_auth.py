import io
from copy import deepcopy
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock
from urllib.request import Request
from urllib.error import HTTPError
from PIL import Image
from lynx_download_auth import DownloadAuth, ScopedRedirect, open_download, google_login
from lynx_tasks import TaskContext, TaskCancelled
from lynx_camtrap import read_package, select_media, acquire_media
from camtrap_factory import make_package


class DownloadAuthTests(unittest.TestCase):
    def test_credentials_scoped_to_exact_https_origin_and_redacted_repr(self):
        for method, header, value in [('Agouti API key', 'X-API-KEY', 'secret'),
                                      ('Agouti Bearer', 'Authorization', 'Bearer secret'),
                                      ('Trapper token', 'Authorization', 'Token secret')]:
            auth = DownloadAuth('https://api.example.org', method, 'secret')
            self.assertEqual(auth.headers('https://api.example.org/photo'), {header: value})
            for url in ('https://other.example.org/photo', 'http://api.example.org/photo', 'https://api.example.org:8443/photo'):
                self.assertEqual(auth.headers(url), {})
            self.assertNotIn('secret', repr(auth))
        for server in ('http://api.example.org', 'https://user:pass@api.example.org', 'https://api.example.org?token=secret'):
            with self.assertRaises(ValueError):
                DownloadAuth(server, 'Trapper token', 'secret')

    def test_redirects_strip_credentials_across_origins_and_reject_downgrade(self):
        request = Request('https://api.example.org/image', headers={
            'Authorization': 'Token secret', 'X-API-KEY': 'secret', 'Cookie': 'session=secret'})
        redirect = ScopedRedirect()
        same = redirect.redirect_request(request, None, 302, '', {}, 'https://api.example.org/other')
        self.assertEqual(same.get_header('Authorization'), 'Token secret')
        other = redirect.redirect_request(request, None, 302, '', {}, 'https://cdn.example.org/other')
        self.assertFalse(other.headers)
        with self.assertRaises(ValueError):
            redirect.redirect_request(request, None, 302, '', {}, 'http://api.example.org/other')

    def test_authenticated_request_uses_scoped_opener(self):
        auth = DownloadAuth('https://api.example.org', 'Agouti API key', 'secret')
        with patch('lynx_download_auth.build_opener') as factory:
            open_download('https://api.example.org/photo', auth)
        request = factory.return_value.open.call_args.args[0]
        self.assertEqual(request.get_header('X-api-key'), 'secret')
        self.assertIsInstance(factory.call_args.args[0], ScopedRedirect)

    def test_dp_private_download_auth_and_unauthorized_retry(self):
        with tempfile.TemporaryDirectory() as folder:
            root = make_package(Path(folder) / 'package')
            task = TaskContext()
            package = read_package(task, root / 'datapackage.json')
            records, _ = select_media(task, package, {'Lynx pardinus'}, True)
            record = next(row for row in records if row['media']['mediaID'] == 'm2')
            record['media']['filePublic'] = 'false'
            record['media']['filePath'] = 'https://api.example.org/photo'
            auth = DownloadAuth('https://api.example.org', 'Trapper token', 'secret')
            denied = HTTPError(record['media']['filePath'], 403, 'secret must not leak', {}, None)
            with patch('lynx_camtrap.open_download', side_effect=denied):
                result = acquire_media(task, package, [record], folder, auth=auth, allow_private=True)
            self.assertEqual(result.failed_records, [record])
            self.assertIn('HTTP 403', result.errors[0])
            self.assertNotIn('secret', result.errors[0])
            encoded = io.BytesIO()
            Image.new('RGB', (3, 3)).save(encoded, 'PNG')
            with patch('lynx_camtrap.open_download', return_value=io.BytesIO(encoded.getvalue())) as request:
                result = acquire_media(task, package, result.failed_records, folder, auth=auth, allow_private=True)
            self.assertEqual(result.completed, 1)
            self.assertEqual(request.call_args.args[1], auth)
            self.assertNotIn('secret', (Path(result.output_directory) / 'manifest.csv').read_text())

    def test_dp_local_folder_replaces_remote_and_private_without_network(self):
        with tempfile.TemporaryDirectory() as folder:
            root = make_package(Path(folder) / 'package')
            photos = Path(folder) / 'photos'
            photos.mkdir()
            Image.new('RGB', (3, 3)).save(photos / 'photo.jpg')
            task = TaskContext()
            package = read_package(task, root / 'datapackage.json')
            records, _ = select_media(task, package, {'Lynx pardinus'}, True)
            record = next(row for row in records if row['media']['mediaID'] == 'm2')
            record['media']['filePublic'] = 'false'
            with patch('lynx_camtrap.urlopen') as public, patch('lynx_camtrap.open_download') as authenticated:
                result = acquire_media(task, package, [record], folder, local_only=True, local_folder=photos)
            self.assertEqual(result.completed, 1)
            public.assert_not_called()
            authenticated.assert_not_called()
            duplicate = deepcopy(record)
            duplicate['media']['mediaID'] = 'different'
            duplicate['media']['filePath'] = 'https://other.example.org/photo.jpg'
            with self.assertRaises(ValueError):
                acquire_media(task, package, [record, duplicate], folder, local_only=True, local_folder=photos)
            nested = photos / 'nested'
            nested.mkdir()
            Image.new('RGB', (3, 3)).save(nested / 'photo.jpg')
            result = acquire_media(task, package, [record], folder, local_only=True, local_folder=photos)
            self.assertEqual(result.failed_records, [record])

    def test_google_browser_login_success_missing_cli_and_cancel(self):
        with patch('lynx_download_auth.shutil.which', return_value=None):
            with self.assertRaises(ValueError):
                google_login(TaskContext())
        process = MagicMock()
        process.__enter__.return_value = process
        process.poll.return_value = 0
        process.returncode = 0
        with patch('lynx_download_auth.shutil.which', return_value='/fake/gcloud'), patch('lynx_download_auth.subprocess.Popen', return_value=process) as launch:
            self.assertTrue(google_login(TaskContext()))
        self.assertEqual(launch.call_args.args[0], ['/fake/gcloud', 'auth', 'login', '--brief'])
        process.poll.return_value = None
        task = TaskContext()
        task.stop.set()
        with patch('lynx_download_auth.shutil.which', return_value='/fake/gcloud'), patch('lynx_download_auth.subprocess.Popen', return_value=process):
            with self.assertRaises(TaskCancelled):
                google_login(task)
        process.terminate.assert_called_once()
