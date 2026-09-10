"""Tiny generated fixture: no field data or photographs stored in Git."""
import csv
import json
from pathlib import Path
from PIL import Image


def make_package(folder):
    root = Path(folder)
    root.mkdir(parents=True, exist_ok=True)
    (root / 'media').mkdir()
    for name in ('one.jpg', 'event.jpg'):
        Image.new('RGB', (8, 8), 'green').save(root / 'media' / name)
    tables = {
        'deployments': [{'deploymentID': 'd1'}],
        'media': [dict(mediaID=mid, deploymentID='d1', filePath=path,
                       timestamp=date, fileMediatype='image/jpeg', filePublic='true')
                  for mid, path, date in (
                      ('m1', 'media/one.jpg', '2024-01-01T00:00:00Z'),
                      ('m2', 'https://example.invalid/photo.jpg', '2024-01-01T00:00:01Z'),
                      ('m3', 'media/event.jpg', '2024-01-01T00:01:00Z'))],
        'observations': [dict(observationID=oid, deploymentID='d1', mediaID=mid,
                              observationType='animal', observationLevel=level, scientificName=species,
                              eventStart='2024-01-01T00:01:00Z', eventEnd='2024-01-01T00:01:00Z')
                         for oid, mid, level, species in (
                             ('o1','m1','media','Lynx pardinus'), ('o2','m2','media','Lynx pardinus'),
                             ('o3','','event','Lynx pardinus'), ('o4','m2','media','Vulpes vulpes'))]}
    for name, rows in tables.items():
        with (root / (name + '.csv')).open('w', newline='', encoding='utf-8') as stream:
            writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
            writer.writeheader()
            writer.writerows(rows)
    descriptor = {'profile': 'https://raw.githubusercontent.com/tdwg/camtrap-dp/1.0/camtrap-dp-profile.json',
                  'title': 'SYNTHETIC TEST', 'resources': [{'name': name, 'path': name+'.csv'} for name in tables]}
    (root / 'datapackage.json').write_text(json.dumps(descriptor), encoding='utf-8')
    return root
