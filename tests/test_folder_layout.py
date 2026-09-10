import tempfile
import unittest
from pathlib import Path
from PIL import Image
from lynx_folder_bulk import folder_rows, scan_layout, validate_stations
from lynx_bulk import build_table, default_fields
from lynx_locations import deployment_key
from lynx_tasks import TaskContext


class FolderLayoutTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.task = TaskContext()
        for place, filename in [('North', 'north.jpg'), ('South', 'south.jpg')]:
            path = self.root / place / 'Station 1' / 'Lynx pardinus' / filename
            path.parent.mkdir(parents=True)
            Image.new('RGB', (2, 2)).save(path)
        self.levels = dict(locality=1, station=2, species=3)

    def test_station_paths_keep_repeated_station_names_distinct(self):
        inventory = scan_layout(self.task, self.root, self.levels)
        self.assertEqual(inventory['errors'], [])
        self.assertEqual(set(inventory['stations']), {'North/Station 1', 'South/Station 1'})
        stations = inventory['stations']
        stations['North/Station 1'].update(latitude='0', longitude='-6,5')
        stations['South/Station 1'].update(latitude='38', longitude='-5')
        rows, _ = folder_rows(self.task, self.root, '', allow_undated=True, levels=self.levels, stations=stations)
        table = build_table(rows, default_fields())
        self.assertEqual([row['Encounter.verbatimLocality'] for row in table], ['North', 'South'])
        self.assertEqual(table[0]['Encounter.decimalLatitude'], 0)
        self.assertEqual(table[0]['Encounter.decimalLongitude'], -6.5)
        self.assertNotEqual(deployment_key(rows[0]), deployment_key(rows[1]))
        self.assertIn('folder.level3', rows[0]['metadata'])
        self.assertFalse(any(key.startswith('deployment.') for key in rows[0]['metadata']))

    def test_invalid_depth_and_coordinates_are_not_silently_ignored(self):
        inventory = scan_layout(self.task, self.root, dict(station=7))
        self.assertEqual(len(inventory['errors']), 2)
        for coordinates in ({'latitude': '91', 'longitude': '0'}, {'latitude': 'NaN', 'longitude': '0'}, {'latitude': '0'}, {'longitude': '181', 'latitude': '0'}):
            with self.assertRaises(ValueError):
                validate_stations({'station': coordinates})
        with self.assertRaises(ValueError):
            folder_rows(self.task, self.root, 'Lynx pardinus', allow_undated=True, levels={'station': 7})

    def test_individual_level_and_scientific_name_are_separate(self):
        levels = dict(self.levels, individual=2)
        rows, _ = folder_rows(self.task, self.root, '', levels=levels, fallback_year=2025)
        self.assertTrue(all(row['individualID'] == 'Station 1' for row in rows))
        self.assertTrue(all(row['species'] == 'Lynx pardinus' for row in rows))
        self.assertTrue(all(row['year'] == 2025 for row in rows))
