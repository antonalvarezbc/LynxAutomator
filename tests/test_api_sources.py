import io
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.parse import urlsplit, parse_qs
from camtrap_factory import make_package
from lynx_api_sources import fetch_api_package
from lynx_camtrap import read_package, select_media, acquire_media
from lynx_bulk import camtrap_rows, build_table, default_fields
from lynx_download_auth import DownloadAuth
from lynx_tasks import TaskContext, TaskCancelled


class APISourceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = make_package(self.root / 'fixture')
        self.output = self.root / 'output'
        self.output.mkdir()
        self.task = TaskContext()
        self.calls = []

    def agouti(self, url, auth=None, timeout=20):
        self.calls.append((url, auth))
        return io.BytesIO((self.source / urlsplit(url).path.rsplit('/', 1)[-1]).read_bytes())

    def archive(self):
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w') as archive:
            for path in self.source.rglob('*'):
                if path.is_file():
                    archive.writestr('export/' + path.relative_to(self.source).as_posix(), path.read_bytes())
        return output.getvalue()

    def test_agouti_optional_auth_and_complete_excel_flow(self):
        with patch('lynx_api_sources.open_download', side_effect=self.agouti):
            package = fetch_api_package(self.task, 'Agouti API', 'https://api.agouti.eu', 'project-1', self.output)
        self.assertEqual(len(self.calls), 4)
        self.assertTrue(all(auth is None for _, auth in self.calls))
        self.assertTrue(self.calls[0][0].endswith('/v1/projects/project-1/datapackage.json'))
        self.assertTrue(package.media[0]['filePath'].startswith('https://api.agouti.eu/v1/projects/project-1/media/'))
        reopened = read_package(self.task, package.path)
        self.assertEqual(reopened.media, package.media)
        self.assertEqual(reopened.descriptor, package.descriptor)
        records, _ = select_media(self.task, package, {'Lynx pardinus'}, True)
        result = acquire_media(self.task, package, records, self.output, local_only=True, local_folder=self.source)
        rows, missing = camtrap_rows(self.task, package, records, [result.output_directory], False)
        self.assertEqual(len(rows), 2)
        self.assertEqual(missing, ['m2'])
        fields = default_fields()
        next(field for field in fields if field['name'] == 'Encounter.locationID').update(value='test-location')
        self.assertEqual(len(build_table(rows, fields)), 2)

    def test_agouti_sends_selected_auth_and_uses_safe_resource_filenames(self):
        auth = DownloadAuth('https://api.agouti.eu', 'Agouti API key', 'test-secret')
        with patch('lynx_api_sources.open_download', side_effect=self.agouti):
            package = fetch_api_package(self.task, 'Agouti API', 'https://api.agouti.eu', 'project-1', self.output, auth)
        self.assertTrue(all(value is auth for _, value in self.calls))
        self.assertNotIn('test-secret', package.path.read_text())
        for resource in package.descriptor['resources']:
            if 'path' in resource:
                self.assertTrue(resource['path'].startswith('resource-'))

    def test_trapper_current_route_and_unsigned_package_link(self):
        archive = self.archive()
        def response(url, auth=None, timeout=20):
            self.calls.append(url)
            if 'api/package' in url:
                return io.BytesIO(json.dumps({'data': {'package': '/data-package/1/?rt=signed-link'}}).encode())
            return io.BytesIO(archive)
        with patch('lynx_api_sources.open_download', side_effect=response):
            package = fetch_api_package(self.task, 'Trapper API', 'https://trapper.example', '123', self.output, approved_only=False)
        query = parse_qs(urlsplit(self.calls[0]).query)
        self.assertEqual(query['export_filetype'], ['csv.gz'])
        self.assertEqual(query['approved_only'], ['false'])
        self.assertEqual(query['release'], ['false'])
        self.assertEqual(len(package.media), 3)
        self.assertEqual(package.prefix, 'export/')

    def test_trapper_legacy_fallback_only_on_404(self):
        archive = self.archive()
        def response(url, auth=None, timeout=20):
            self.calls.append(url)
            if '/media_classification/' in url:
                raise HTTPError(url, 404, 'not found', {}, None)
            if '/api/media-classifications/' in url:
                return io.BytesIO(b'{"data":{"package":"https://cdn.example/export.zip"}}')
            return io.BytesIO(archive)
        with patch('lynx_api_sources.open_download', side_effect=response):
            package = fetch_api_package(self.task, 'Trapper API', 'https://trapper.example', '123', self.output)
        self.assertEqual(len(self.calls), 3)
        self.assertIn('/api/media-classifications/package/123/', self.calls[1])
        self.assertTrue(package.path.is_file())

    def test_unauthorized_and_invalid_replies_leave_no_partial_package(self):
        for response in (HTTPError('https://private/?secret=x', 401, 'secret', {}, None), io.BytesIO(b'not JSON')):
            with patch('lynx_api_sources.open_download', **({'side_effect': response} if isinstance(response, Exception) else {'return_value': response})):
                with self.assertRaises(ValueError) as error:
                    fetch_api_package(self.task, 'Agouti API', 'https://api.agouti.eu', 'project-1', self.output)
            self.assertNotIn('secret', str(error.exception))
            self.assertEqual(list(self.output.iterdir()), [])

    def test_invalid_server_id_and_cancellation(self):
        with patch('lynx_api_sources.open_download') as request:
            for server, project in [('http://example.org', 'p'), ('https://example.org?token=x', 'p'), ('https://example.org', '../other')]:
                with self.assertRaises(ValueError):
                    fetch_api_package(self.task, 'Agouti API', server, project, self.output)
            request.assert_not_called()
        self.task.stop.set()
        with self.assertRaises(TaskCancelled):
            fetch_api_package(self.task, 'Agouti API', 'https://api.agouti.eu', 'project-1', self.output)
        self.assertEqual(list(self.output.iterdir()), [])

    def two_deployments(self):
        import csv
        from copy import deepcopy
        for name in ('deployments', 'media', 'observations'):
            path = self.source / (name + '.csv')
            with path.open() as stream:
                rows = list(csv.DictReader(stream))
            if name == 'deployments':
                rows[0].update(deploymentStart='2024-01-01T00:00:00Z', locationName='North')
                rows.append(dict(deploymentID='d2', deploymentStart='2025-05-01T00:00:00Z', locationName='South'))
            else:
                for row in deepcopy(rows):
                    row['deploymentID'] = 'd2'
                    for key in ('mediaID', 'observationID', 'eventID'):
                        if row.get(key):
                            row[key] += '-d2'
                    rows.append(row)
            with path.open('w') as stream:
                writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
                writer.writeheader()
                writer.writerows(rows)

    def filtered_agouti(self, url, auth=None, timeout=20):
        import csv
        self.calls.append((url, auth))
        path = self.source / urlsplit(url).path.rsplit('/', 1)[-1]
        ids = parse_qs(urlsplit(url).query).get('deploymentID')
        if not ids:
            return io.BytesIO(path.read_bytes())
        with path.open() as stream:
            reader = csv.DictReader(stream)
            fields = reader.fieldnames
            rows = [row for row in reader if row['deploymentID'] in ids]
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
        return io.BytesIO(output.getvalue().encode())

    def test_agouti_selects_deployments_before_requesting_other_tables(self):
        self.two_deployments()
        with patch('lynx_api_sources.open_download', side_effect=self.filtered_agouti):
            package = fetch_api_package(self.task, 'Agouti API', 'https://api.agouti.eu', 'p', self.output,
                                        filters={'year': '2025', 'site': 'south', 'latest': '1'})
        self.assertEqual([row['deploymentID'] for row in package.deployments], ['d2'])
        self.assertTrue(all(row['deploymentID'] == 'd2' for row in package.media + package.observations))
        self.assertEqual(len(self.calls), 4)
        self.assertTrue(self.calls[1][0].endswith('/deployments.csv'))
        for url, _ in self.calls[2:]:
            self.assertEqual(parse_qs(urlsplit(url).query)['deploymentID'], ['d2'])
        reopened = read_package(self.task, package.path)
        self.assertEqual(reopened.media, package.media)
        self.assertEqual(reopened.deployments, package.deployments)

    def test_empty_agouti_selection_stops_before_media_and_cleans_output(self):
        self.two_deployments()
        with patch('lynx_api_sources.open_download', side_effect=self.filtered_agouti):
            with self.assertRaisesRegex(ValueError, 'Ningún despliegue'):
                fetch_api_package(self.task, 'Agouti API', 'https://api.agouti.eu', 'p', self.output, filters={'year': '2020'})
        self.assertEqual(len(self.calls), 2)
        self.assertEqual(list(self.output.iterdir()), [])

    def test_agouti_rejects_server_ignoring_deployment_filter(self):
        self.two_deployments()
        with patch('lynx_api_sources.open_download', side_effect=self.agouti):
            with self.assertRaisesRegex(ValueError, 'no respetó'):
                fetch_api_package(self.task, 'Agouti API', 'https://api.agouti.eu', 'p', self.output, filters={'year': '2025'})
        self.assertEqual(list(self.output.iterdir()), [])

    def test_trapper_remote_filters_and_explicit_local_refinement(self):
        self.two_deployments()
        archive = self.archive()
        def response(url, auth=None, timeout=20):
            self.calls.append(url)
            if 'api/package' in url:
                return io.BytesIO(b'{"data":{"package":"/package.zip"}}')
            return io.BytesIO(archive)
        with patch('lynx_api_sources.open_download', side_effect=response):
            package = fetch_api_package(self.task, 'Trapper API', 'https://trapper.example', 'p', self.output,
                                        filters={'deployment': 'd', 'year': '2025', 'exclude_blank': True})
        query = parse_qs(urlsplit(self.calls[0]).query)
        self.assertEqual(query['filter_deployments'], ['d'])
        self.assertEqual(query['exclude_blank'], ['true'])
        self.assertNotIn('year', query)
        self.assertEqual([row['deploymentID'] for row in package.deployments], ['d2'])
        self.assertTrue(all(row['deploymentID'] == 'd2' for row in package.media + package.observations))
        self.assertEqual(len(read_package(self.task, package.path).deployments), 2)

    def test_trapper_does_not_silently_change_legacy_filter_semantics(self):
        with patch('lynx_api_sources.open_download', side_effect=HTTPError('https://example', 404, 'missing', {}, None)) as request:
            with self.assertRaisesRegex(ValueError, 'ruta antigua'):
                fetch_api_package(self.task, 'Trapper API', 'https://trapper.example', 'p', self.output, filters={'deployment': 'cam'})
        self.assertEqual(request.call_count, 1)
        self.assertEqual(list(self.output.iterdir()), [])

    def test_filter_validation_missing_dates_and_latest_order(self):
        from lynx_api_filters import selected_deployments, normalize_filters
        rows = [dict(deploymentID='old', deploymentStart='2023-01-01'),
                dict(deploymentID='unknown'), dict(deploymentID='new', deploymentStart='2025-01-01T12:00:00Z')]
        self.assertEqual(selected_deployments(rows, {'latest': '1'}), [rows[-1]])
        with self.assertRaisesRegex(ValueError, 'Ningún despliegue'):
            selected_deployments([rows[1]], {'year': '2025'})
        for filters in ({'year': 'yesterday'}, {'latest': '0'}, {'year': '10000'}):
            with self.assertRaises(ValueError):
                normalize_filters(filters)
