"""Regression tests use the same UI-independent processors as both desktop apps."""
from datetime import timedelta
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd
from PIL import Image
import piexif

from lynx_core import file_timestamp
from lynx_tasks import TaskContext
from lynx_processing import WIProcessor, LynxProcessor, FolderProcessor, CatalogProcessor
from lynx_file_jobs import change_dates, download_images


class AppLogicTests(unittest.TestCase):
    def test_grouping_never_combines_different_projects(self):
        processor = WIProcessor(TaskContext(), '', '', '', True, False, 60)
        data = pd.DataFrame([
            dict(project_id=project, deployment_id='01', timestamp=pd.Timestamp('2024-01-01'),
                 latitude='40', longitude='-3', placename='site', location=f'gs://bucket/{project}.jpg')
            for project in ['A', 'B']])
        result = processor.process_multiple_images(data, pd.DataFrame())
        self.assertEqual(len(result), 2)
        self.assertEqual(set(result['Occurrence.occurrenceID']), {'A-01', 'B-01'})
        self.assertNotIn('Encounter.mediaAsset1', result.columns)

    def test_copied_dates_come_from_source(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'source'
            destination = Path(folder) / 'copies'
            source.mkdir()
            destination.mkdir()
            video = source / 'video.mp4'
            video.write_bytes(b'video')
            original = file_timestamp(video)
            change_dates(TaskContext(), source, timedelta(seconds=60), destination)
            target = destination / 'video.mp4'
            self.assertEqual(target.read_bytes(), video.read_bytes())
            self.assertAlmostEqual(target.stat().st_mtime, original + 60, places=2)
            change_dates(TaskContext(), source, timedelta(seconds=60), destination)
            self.assertTrue((destination / 'video_1.mp4').exists())

    def test_invalid_csv_location_never_executes(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(subprocess, 'run') as run:
            csv = Path(folder) / 'images.csv'
            pd.DataFrame({'location': ['$(untrusted command)'], 'deployment_id': ['01']}).to_csv(csv, index=False)
            result = download_images(TaskContext(), csv, folder, False, 'gsutil')
            run.assert_not_called()
            self.assertEqual(len(result.errors), 1)

    def test_failed_download_leaves_no_partial_file(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(subprocess, 'run', side_effect=subprocess.CalledProcessError(1, 'gsutil')):
            csv = Path(folder) / 'images.csv'
            pd.DataFrame({'location': ['gs://bucket/photo.jpg'], 'deployment_id': ['01']}).to_csv(csv, index=False)
            result = download_images(TaskContext(), csv, folder, False, 'gsutil')
            self.assertEqual(list(Path(folder).iterdir()), [csv])
            self.assertEqual(len(result.errors), 1)

    def test_lynx_folder_layouts_with_optional_revision_and_linces(self):
        for revision, linces in [(False, False), (True, False), (False, True), (True, True)]:
            with self.subTest(revision=revision, linces=linces), tempfile.TemporaryDirectory() as folder:
                parts = ['Farm', 'Station']
                if revision:
                    parts.append('Revision01')
                if linces:
                    parts.append('Linces')
                parts.append('Individual')
                image_folder = Path(folder).joinpath(*parts)
                image_folder.mkdir(parents=True)
                Image.new('RGB', (4, 4)).save(image_folder / 'photo.jpg')
                result = LynxProcessor(TaskContext(), folder, linces, revision, 0, None, None).run()
                row = result.iloc[0]
                self.assertEqual(row['Finca'], 'Farm')
                self.assertEqual(row['Estación'], 'Station')
                self.assertEqual(row['Individuo'], 'Individual')
                if revision:
                    self.assertEqual(row['Revisión'], 'Revision01')
                else:
                    self.assertNotIn('Revisión', result.columns)

    def test_folder_and_catalog_spreadsheets_use_real_inputs(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            exif = piexif.dump({'Exif': {piexif.ExifIFD.DateTimeOriginal: b'2024:01:01 12:00:00'}})
            Image.new('RGB', (4, 4)).save(folder / 'Nube lateral.jpg', exif=exif)
            template = folder / 'template.xlsx'
            pd.DataFrame({'Encounter.country': ['Spain']}).to_excel(template, index=False)
            result = FolderProcessor(TaskContext(), folder, template, False, 0).run()
            self.assertEqual(result.iloc[0]['Encounter.year'], 2024)
            self.assertEqual(result.iloc[0]['Encounter.country'], 'Spain')
            catalog = CatalogProcessor(TaskContext(), folder, template, False, True).run()
            self.assertEqual(catalog.iloc[0]['MarkedIndividual.individualID'], 'Nube')

    def test_wi_conversion_reads_csvs_and_template_in_worker_service(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            images, deployments, template = folder / 'images.csv', folder / 'deployments.csv', folder / 'template.xlsx'
            pd.DataFrame({'project_id': ['01', '01'], 'deployment_id': ['001', '001'],
                          'location': ['gs://bucket/a.jpg', 'gs://bucket/b.jpg'],
                          'timestamp': ['2024-01-01 10:00:00', '2024-01-01 10:00:20']}).to_csv(images, index=False)
            pd.DataFrame({'project_id': ['01'], 'deployment_id': ['001'], 'latitude': [40],
                          'longitude': [-3], 'placename': ['Site'], 'subproject_name': ['Subproject']}).to_csv(deployments, index=False)
            pd.DataFrame({'Encounter.country': ['Spain']}).to_excel(template, index=False)
            result = WIProcessor(TaskContext(), template, images, deployments, True, False, 60).run()
            self.assertEqual(len(result), 1)
            self.assertEqual(result.iloc[0]['Occurrence.occurrenceID'], '01-001')
            self.assertEqual(result.iloc[0]['Encounter.mediaAsset0'], 'a.JPG')
            self.assertEqual(result.iloc[0]['Encounter.mediaAsset1'], 'b.JPG')
            self.assertEqual(result.iloc[0]['Encounter.country'], 'Spain')
