"""Camtrap DP 1.x package reading and cancellable media acquisition."""
import csv
from dataclasses import dataclass
from datetime import datetime
import gzip
import hashlib
import io
import json
from pathlib import Path, PurePosixPath
import tempfile
from urllib.parse import urlsplit
from urllib.request import urlopen
import zipfile
import os
from collections import Counter
from PIL import Image
from lynx_file_jobs import BatchResult
from lynx_tasks import TaskCancelled

LIMIT = 100 * 1024 * 1024


def relative_path(value):
    path = PurePosixPath(value)
    if not value or path.is_absolute() or '..' in path.parts or '\\' in value or ':' in value:
        raise ValueError('Invalid local package path')
    return path


@dataclass
class Package:
    path: Path
    prefix: str
    descriptor: dict
    deployments: list
    media: list
    observations: list

    def open_local(self, value):
        relative = relative_path(value)
        if self.path.suffix.lower() == '.zip':
            archive = zipfile.ZipFile(self.path)
            try:
                info = archive.getinfo(self.prefix + str(relative))
                if info.file_size > LIMIT:
                    raise ValueError('Package file exceeds 100 MB')
                data = archive.read(info)
            finally:
                archive.close()
            return io.BytesIO(data)
        base = self.path.parent.resolve()
        target = (base / str(relative)).resolve()
        if not target.is_relative_to(base):
            raise ValueError('File is outside the package directory')
        if target.stat().st_size > LIMIT:
            raise ValueError('Package file exceeds 100 MB')
        return target.open('rb')

    def species(self):
        return Counter(r['scientificName'] for r in self.observations
                       if r.get('observationType') == 'animal' and r.get('scientificName'))


def read_package(task, filename):
    path = Path(filename)
    prefix = ''
    if path.suffix.lower() == '.zip':
        with zipfile.ZipFile(path) as archive:
            matches = [n for n in archive.namelist() if PurePosixPath(n).name == 'datapackage.json']
            if len(matches) != 1:
                raise ValueError('ZIP must contain exactly one datapackage.json')
            relative_path(matches[0])
            if archive.getinfo(matches[0]).file_size > LIMIT:
                raise ValueError('Descriptor too large')
            descriptor = json.loads(archive.read(matches[0]))
            prefix = matches[0][:-len('datapackage.json')]
    else:
        descriptor = json.loads(path.read_text(encoding='utf-8-sig'))
    profile = str(descriptor.get('profile', ''))
    if '/1.' not in profile:
        raise ValueError('This reader supports Camtrap DP 1.x profiles only')
    package = Package(path, prefix, descriptor, [], [], [])
    for name, id_field, required in (
        ('deployments', 'deploymentID', {'deploymentID'}),
        ('media', 'mediaID', {'mediaID', 'deploymentID', 'filePath', 'timestamp'}),
        ('observations', 'observationID', {'observationID', 'deploymentID', 'observationType', 'observationLevel', 'scientificName'})):
        task.checkpoint()
        resources = [r for r in descriptor.get('resources', []) if r.get('name') == name]
        if len(resources) != 1:
            raise ValueError(f'Missing or repeated resource: {name}')
        resource = resources[0]
        if 'data' in resource:
            rows = resource['data']
        else:
            resource_path = resource.get('path')
            if not isinstance(resource_path, str):
                raise ValueError(f'{name}: expected a local CSV path')
            with package.open_local(resource_path) as stream:
                raw = stream.read(LIMIT + 1)
            if resource_path.endswith('.gz'):
                with gzip.GzipFile(fileobj=io.BytesIO(raw)) as stream:
                    raw = stream.read(LIMIT + 1)
            if len(raw) > LIMIT:
                raise ValueError('Table too large')
            reader = csv.DictReader(io.StringIO(raw.decode('utf-8-sig')))
            if not required.issubset(reader.fieldnames or []):
                raise ValueError(f'{name}: missing required columns')
            rows = []
            for row in reader:
                task.checkpoint()
                rows.append(row)
        ids = set()
        for row in rows:
            task.checkpoint()
            if not required.issubset(row) or not row.get(id_field) or row[id_field] in ids:
                raise ValueError(f'{name}: missing fields or duplicate/empty ID')
            ids.add(row[id_field])
        setattr(package, name, rows)
        task.report(name)
    dids = {r['deploymentID'] for r in package.deployments}
    media = {r['mediaID']: r for r in package.media}
    for row in package.media + package.observations:
        if row['deploymentID'] not in dids:
            raise ValueError('Unknown deployment reference')
    for row in package.observations:
        if row.get('observationLevel') == 'media':
            target = media.get(row.get('mediaID'))
            if target is None or target['deploymentID'] != row['deploymentID']:
                raise ValueError('Invalid media observation reference')
        elif row.get('observationLevel') != 'event':
            raise ValueError('Unknown observation level')
    return package


def iso(value):
    return datetime.fromisoformat(value.replace('Z', '+00:00'))


def select_media(task, package, species, include_events=True):
    selected = {}
    by_id = {r['mediaID']: r for r in package.media}
    by_deployment = {}
    for media in package.media:
        by_deployment.setdefault(media['deploymentID'], []).append(media)
    issues = []
    for observation in package.observations:
        task.checkpoint()
        if observation.get('observationType') != 'animal' or observation.get('scientificName') not in species:
            continue
        event = observation['observationLevel'] == 'event'
        if event:
            if not include_events:
                issues.append(f"{observation['observationID']}: event excluded")
                continue
            try:
                start, end = iso(observation['eventStart']), iso(observation['eventEnd'])
                if end < start:
                    raise ValueError('Invalid event interval')
                candidates = [r for r in by_deployment.get(observation['deploymentID'], [])
                              if start <= iso(r['timestamp']) <= end]
            except (KeyError, ValueError, TypeError):
                issues.append(f"{observation['observationID']}: ambiguous event dates")
                continue
        else:
            candidates = [by_id[observation['mediaID']]]
        if not candidates:
            issues.append(f"{observation['observationID']}: no media")
        for media in candidates:
            record = selected.setdefault(media['mediaID'], {'media': media, 'species': set(), 'event': True})
            record.setdefault('observation_ids', set()).add(observation['observationID'])
            if observation.get('eventID'):
                record.setdefault('event_ids', set()).add(observation['eventID'])
            record['species'].add(observation['scientificName'])
            record['event'] = record['event'] and event
    return list(selected.values()), issues


def acquire_media(task, package, records, destination, local_only=False):
    destination = Path(destination)
    result = BatchResult()
    result.failed_records = []
    # New batch directory prevents collisions or reuse of a previous selection.
    batch = Path(tempfile.mkdtemp(prefix='camtrap-', dir=destination))
    result.output_directory = str(batch)
    with (batch / 'manifest.csv').open('w', newline='', encoding='utf-8') as log:
        writer = csv.DictWriter(log, fieldnames=['mediaID', 'species', 'file', 'status', 'association', 'eventIDs', 'observationIDs'])
        writer.writeheader()
        for index, record in enumerate(records):
            task.checkpoint()
            media = record['media']
            status, filename = 'error', ''
            remote = urlsplit(media['filePath']).scheme in ('http', 'https')
            try:
                if media.get('filePublic', 'true') in ('false', False):
                    status = 'private'
                    result.skipped += 1
                elif not media.get('fileMediatype', '').startswith('image/'):
                    status = 'not_image'
                    result.skipped += 1
                elif remote and local_only:
                    status = 'remote_skipped'
                    result.skipped += 1
                else:
                    with tempfile.TemporaryDirectory(dir=batch, prefix='.partial-') as staging:
                        temporary = Path(staging) / 'image'
                        source = urlopen(media['filePath'], timeout=20) if remote else package.open_local(media['filePath'])
                        with source, temporary.open('wb') as out:
                            size = 0
                            while True:
                                task.checkpoint()
                                chunk = source.read(64 * 1024)
                                if not chunk:
                                    break
                                size += len(chunk)
                                if size > LIMIT:
                                    raise ValueError('Image exceeds 100 MB')
                                out.write(chunk)
                        with Image.open(temporary) as image:
                            extension = {'JPEG': '.jpg', 'PNG': '.png', 'TIFF': '.tif', 'WEBP': '.webp'}.get(image.format)
                            image.verify()
                        if not extension:
                            raise ValueError('Unsupported image format')
                        filename = hashlib.sha256(media['mediaID'].encode()).hexdigest() + extension
                        task.checkpoint()
                        os.replace(temporary, batch / filename)
                    status = 'completed'
                    result.completed += 1
            except TaskCancelled:
                writer.writerow(dict(mediaID=media['mediaID'], species=';'.join(sorted(record['species'])),
                                     file='', status='cancelled', association='event' if record['event'] else 'media',
                                     eventIDs=json.dumps(sorted(record.get('event_ids', []))),
                                     observationIDs=json.dumps(sorted(record.get('observation_ids', [])))))
                log.flush()
                raise
            except Exception as exc:
                # URLs can carry credentials: never include them in messages/logs.
                result.errors.append(f"{media['mediaID']}: {type(exc).__name__}")
                result.failed_records.append(record)
            writer.writerow(dict(mediaID=media['mediaID'], species=';'.join(sorted(record['species'])),
                                 file=filename if status == 'completed' else '', status=status, association='event' if record['event'] else 'media',
                                     eventIDs=json.dumps(sorted(record.get('event_ids', []))),
                                     observationIDs=json.dumps(sorted(record.get('observation_ids', [])))))
            log.flush()
            task.report(f'{index + 1}/{len(records)} — {batch.name}', (index + 1) / max(1, len(records)))
    return result
