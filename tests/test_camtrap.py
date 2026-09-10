import csv
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile
from PIL import Image
from lynx_camtrap import read_package, select_media, acquire_media
from lynx_tasks import TaskContext, TaskCancelled

from camtrap_factory import make_package


class CamtrapTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.fixture = make_package(temporary.name)

    def test_fixture_selection_and_local_copy(self):
        task = TaskContext()
        package = read_package(task, self.fixture / 'datapackage.json')
        self.assertEqual(package.species(), {'Lynx pardinus': 3, 'Vulpes vulpes': 1})
        records, issues = select_media(task, package, {'Lynx pardinus'}, False)
        self.assertEqual(len(records), 2)
        self.assertEqual(len(issues), 1)
        events, issues = select_media(task, package, {'Lynx pardinus'}, True)
        self.assertEqual(len(events), 3)
        self.assertFalse(issues)
        self.assertEqual(len({r['media']['mediaID'] for r in events}), 3)
        with tempfile.TemporaryDirectory() as output, patch('lynx_camtrap.urlopen') as network:
            result = acquire_media(task, package, records, output, local_only=True)
            network.assert_not_called()
            self.assertEqual(result.completed, 1)
            self.assertFalse(result.errors)
            batch = Path(result.output_directory)
            self.assertEqual(len(list(batch.glob('*.jpg'))), 1)
            with (batch / 'manifest.csv').open() as stream:
                self.assertEqual(len(list(csv.DictReader(stream))), 2)

    def test_default_includes_events_and_manifest_preserves_provenance(self):
        task = TaskContext()
        package = read_package(task, self.fixture / 'datapackage.json')
        records, issues = select_media(task, package, {'Lynx pardinus'})
        self.assertEqual(len(records), 3)
        self.assertFalse(issues)
        with tempfile.TemporaryDirectory() as folder:
            result = acquire_media(task, package, records, folder, local_only=True)
            with (Path(result.output_directory) / 'manifest.csv').open() as stream:
                rows = list(csv.DictReader(stream))
            event = next(row for row in rows if row['mediaID'] == 'm3')
            self.assertEqual(event['association'], 'event')
            self.assertEqual(json.loads(event['eventIDs']), ['event1'])
            self.assertEqual(json.loads(event['observationIDs']), ['o3'])

    def test_species_list_is_dynamic_and_multiselect_deduplicates(self):
        task = TaskContext()
        package = read_package(task, self.fixture / 'datapackage.json')
        row = next(r for r in package.observations if r['observationType'] == 'animal' and r['observationLevel'] == 'media')
        row['scientificName'] = 'New test species'
        self.assertIn('New test species', package.species())
        records, _ = select_media(task, package, {'Lynx pardinus', 'New test species'}, False)
        self.assertEqual(len(records), 2)
        only_new, _ = select_media(task, package, {'New test species'})
        self.assertEqual(len(only_new), 1)
        self.assertEqual(only_new[0]['media']['mediaID'], row['mediaID'])

    def test_zip_and_unsafe_paths(self):
        with tempfile.TemporaryDirectory() as folder:
            archive = Path(folder) / 'input.zip'
            with zipfile.ZipFile(archive, 'w') as output:
                for source in self.fixture.rglob('*'):
                    if source.is_file():
                        output.write(source, 'package/' + source.relative_to(self.fixture).as_posix())
            package = read_package(TaskContext(), archive)
            records, _ = select_media(TaskContext(), package, {'Lynx pardinus'}, False)
            result = acquire_media(TaskContext(), package, records, folder, local_only=True)
            self.assertEqual(result.completed, 1)
            for name in ('../secret', '/tmp/secret', 'C:\\secret'):
                with self.assertRaises(ValueError):
                    package.open_local(name)

    def test_download_validates_bytes_and_respects_private(self):
        package = read_package(TaskContext(), self.fixture / 'datapackage.json')
        record = {'media': {'mediaID': 'remote', 'filePath': 'https://example.invalid/photo?token=secret',
                            'fileMediatype': 'image/jpeg', 'filePublic': 'true'},
                  'species': {'Lynx pardinus'}, 'event': False}
        encoded = io.BytesIO()
        Image.new('RGB', (8, 8)).save(encoded, format='JPEG')
        with tempfile.TemporaryDirectory() as folder:
            with patch('lynx_camtrap.urlopen', return_value=io.BytesIO(encoded.getvalue())):
                result = acquire_media(TaskContext(), package, [record], folder)
                self.assertEqual(result.completed, 1)
            with patch('lynx_camtrap.urlopen', return_value=io.BytesIO(b'<html>login</html>')):
                result = acquire_media(TaskContext(), package, [record], folder)
                self.assertEqual(len(result.errors), 1)
                self.assertNotIn('secret', str(result.errors))
                self.assertFalse(list(Path(result.output_directory).glob('*.jpg')))
                self.assertEqual(len(result.failed_records), 1)
            record['media']['filePublic'] = 'false'
            with patch('lynx_camtrap.urlopen') as request:
                result = acquire_media(TaskContext(), package, [record], folder)
                request.assert_not_called()
                self.assertEqual(result.skipped, 1)

    def test_cancel_during_transfer_cleans_partial_file(self):
        task = TaskContext()
        package = read_package(task, self.fixture / 'datapackage.json')
        record = {'media': {'mediaID': 'remote', 'filePath': 'https://example.invalid/photo',
                            'fileMediatype': 'image/jpeg'}, 'species': {'Lynx pardinus'}, 'event': False}
        class Stream(io.BytesIO):
            def read(self, size=-1):
                task.stop.set()
                return b'partial'
        with tempfile.TemporaryDirectory() as folder, patch('lynx_camtrap.urlopen', return_value=Stream()):
            with self.assertRaises(TaskCancelled):
                acquire_media(task, package, [record], folder)
            self.assertEqual([p.name for p in Path(folder).rglob('*') if p.is_file()], ['manifest.csv'])
