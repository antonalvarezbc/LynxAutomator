import csv
import json
from pathlib import Path
import unittest

FIXTURE = Path(__file__).parent / 'fixtures' / 'camtrap_dp_lynx_synthetic'


@unittest.skipUnless(FIXTURE.is_dir(), 'Optional local paper example, excluded from Git')
class CamtrapFixtureTests(unittest.TestCase):
    def test_synthetic_taxa_preserve_negative_cases_and_references(self):
        def read(name):
            with (FIXTURE / (name + '.csv')).open(newline='', encoding='utf-8') as stream:
                return list(csv.DictReader(stream))
        observations, media, deployments = map(read, ('observations', 'media', 'deployments'))
        self.assertEqual((len(observations), len(media), len(deployments)), (549, 423, 4))
        animals = [r for r in observations if r['observationType'] == 'animal']
        self.assertEqual(len(animals), 366)
        self.assertEqual({r['scientificName'] for r in animals}, {'Lynx pardinus'})
        self.assertTrue(all(r['scientificName'] == 'Homo sapiens' for r in observations if r['observationType'] == 'human'))
        self.assertTrue(all(r['scientificName'] != 'Lynx pardinus' for r in observations if r['observationType'] != 'animal'))
        self.assertEqual({r['observationLevel'] for r in observations}, {'event', 'media'})
        mids = {r['mediaID'] for r in media}
        dids = {r['deploymentID'] for r in deployments}
        self.assertTrue(all(r['deploymentID'] in dids for r in observations + media))
        self.assertTrue(all(not r['mediaID'] or r['mediaID'] in mids for r in observations))
        local = [r for r in media if not r['filePath'].startswith(('http://', 'https://'))]
        self.assertEqual(len(local), 10)
        self.assertTrue(all((FIXTURE / r['filePath']).is_file() for r in local))
        package = json.loads((FIXTURE / 'datapackage.json').read_text())
        self.assertIn('SYNTHETIC', package['title'])
        self.assertEqual({r['scientificName'] for r in package['taxonomic']}, {'Lynx pardinus', 'Homo sapiens'})
