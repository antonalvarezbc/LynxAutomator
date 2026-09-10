import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from lynx_locations import flatten, read_catalog, normalize_url, CATALOGS, deployment_key
from lynx_bulk import default_fields, normalize, attach_metadata, build_table, save_profile, load_profile
from lynx_tasks import TaskContext


class LocationTests(unittest.TestCase):
    def test_hierarchy_exact_id_and_missing_children(self):
        entries = flatten({'locationID': [{'id': 'parent', 'name': 'Region', 'locationID': [{'id': 'EXACT-á', 'name': 'Site'}]}]})
        self.assertEqual(entries[-1], {'id': 'EXACT-á', 'label': 'Region → Site'})
        with self.assertRaises(ValueError):
            flatten({'locationID': [{'name': 'No id'}]})
        self.assertEqual(len(CATALOGS), 7)
        self.assertIn('African Carnivore Wildbook', CATALOGS)
        self.assertIn('/lynx/', normalize_url('https://github.com/WildMeOrg/Wildbook/blob/lynx/src/main/resources/bundles/locationID.json'))

    def test_cache_offline_and_refresh(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / 'locations.json'
            source.write_text(json.dumps({'locationID': [{'id': 'old'}]}))
            result = read_catalog(TaskContext(), str(source), root / 'cache')
            source.write_text(json.dumps({'locationID': [{'id': 'new'}]}))
            self.assertEqual(read_catalog(TaskContext(), str(source), root / 'cache'), result)
            refreshed = read_catalog(TaskContext(), str(source), root / 'cache', True)
            self.assertEqual(flatten(refreshed['catalog'])[0]['id'], 'new')
            source.unlink()
            self.assertEqual(read_catalog(TaskContext(), str(source), root / 'cache'), refreshed)

    def test_profile_per_deployment_and_coordinates_preserved(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            photo = root / 'photo.jpg'
            photo.write_bytes(b'photo')
            rows = []
            for dep in ['d1', 'd2']:
                row = normalize('Lynx pardinus', '2024-01-01', photo, dep, 'e', 'p', latitude=12, longitude=34)
                attach_metadata(row, deployment=[{'deploymentID': dep}])
                rows.append(row)
            fields = default_fields()
            location = next(f for f in fields if f['name'] == 'Encounter.locationID')
            location.update(source='location_map', locations={deployment_key(rows[0]): 'one', deployment_key(rows[1]): 'two'}, catalog_name='Lynx', catalog_source=CATALOGS['Lynx'])
            path = root / 'profiles.json'
            save_profile(fields, path, 'localities')
            restored = load_profile(path, 'localities')
            table = build_table(rows, restored)
            self.assertEqual([r['Encounter.locationID'] for r in table], ['one', 'two'])
            self.assertEqual(table[0]['Encounter.decimalLatitude'], 12)
            self.assertEqual(next(f for f in restored if f['name'] == 'Encounter.locationID')['catalog_name'], 'Lynx')
            location['locations'].pop(deployment_key(rows[1]))
            with self.assertRaisesRegex(ValueError, 'd2'):
                build_table(rows, fields)

    def test_branch_list_pagination_cache_and_encoded_branch(self):
        from lynx_locations import github_branches, branch_url
        from unittest.mock import MagicMock
        first = MagicMock()
        first.__enter__.return_value.read.return_value = json.dumps([{'name': f'branch{i}'} for i in range(100)]).encode()
        second = MagicMock()
        second.__enter__.return_value.read.return_value = b'[{"name":"feature/test"}]'
        with tempfile.TemporaryDirectory() as folder, patch('urllib.request.urlopen'):
            with patch('lynx_locations.urlopen', side_effect=[first, second]) as get:
                names = github_branches(TaskContext(), folder, True)
                self.assertEqual(get.call_count, 2)
            with patch('lynx_locations.urlopen', side_effect=AssertionError('No network')):
                self.assertEqual(github_branches(TaskContext(), folder), names)
        self.assertEqual(len(names), 101)
        self.assertIn('feature%2Ftest', branch_url('feature/test'))

    def test_encounter_assignment_overrides_deployment(self):
        from lynx_locations import encounter_key
        from lynx_bulk import location_value
        row = dict(eventID='e', species='Lynx pardinus', media=['one.jpg'])
        attach_metadata(row, deployment=[{'deploymentID': 'd'}])
        field = dict(value='common', locations={deployment_key(row): 'deployment', encounter_key(row): 'encounter'})
        self.assertEqual(location_value(row, field), 'encounter')
        field['locations'].pop(encounter_key(row))
        self.assertEqual(location_value(row, field), 'deployment')
