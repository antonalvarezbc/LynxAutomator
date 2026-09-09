"""Test actual class methods without Tk; CI separately checks full GUI startup."""
import ast
import os
from datetime import datetime, timedelta
from pathlib import Path
import queue
import re
import shutil
import subprocess
import tempfile
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
import pandas as pd
from PIL import Image
from lynx_core import file_timestamp, shift_file_date, unique_path

SOURCE = Path(__file__).resolve().parents[1] / 'LynxAutomator_v001alpha.py'
tree = ast.parse(SOURCE.read_text(encoding='utf-8'))
classes = ast.Module(body=[node for node in tree.body if isinstance(node, ast.ClassDef)
                         and node.name in {'ExcelCombinerApp', 'DateChangerApp', 'GCSDownloaderAndRenamer', 'LynxOne'}], type_ignores=[])
namespace = dict(globals(), ctk=SimpleNamespace(CTkFrame=object), messagebox=Mock())
exec(compile(classes, str(SOURCE), 'exec'), namespace)


class AppLogicTests(unittest.TestCase):
    def test_grouping_never_combines_different_projects(self):
        cls = namespace['ExcelCombinerApp']
        app = cls.__new__(cls)
        app.time_threshold_entry = SimpleNamespace(get=lambda: '60')
        app.lang = 'en'
        app.translations = {'en': {'error_threshold_value': 'Invalid threshold'}}
        data = pd.DataFrame([
            dict(project_id=project, deployment_id='01', timestamp=pd.Timestamp('2024-01-01'),
                 latitude='40', longitude='-3', placename='site', location=f'gs://bucket/{project}.jpg')
            for project in ['A', 'B']])
        result = app.process_multiple_images(data, pd.DataFrame())
        self.assertEqual(len(result), 2)
        self.assertEqual(set(result['Occurrence.occurrenceID']), {'A-01', 'B-01'})
        self.assertNotIn('Encounter.mediaAsset1', result.columns)

    def test_copied_dates_come_from_source(self):
        cls = namespace['DateChangerApp']
        app = cls.__new__(cls)
        app.get_date_difference = lambda: timedelta(seconds=60)
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'source.mp4'
            target = Path(folder) / 'copy.mp4'
            source.write_bytes(b'video')
            original = file_timestamp(source)
            app.apply_date_change_to_copy(source, target)
            self.assertEqual(target.read_bytes(), source.read_bytes())
            self.assertAlmostEqual(target.stat().st_mtime, original + 60, places=2)
            with self.assertRaises(FileExistsError):
                app.apply_date_change_to_copy(source, target)

    def worker(self):
        cls = namespace['GCSDownloaderAndRenamer']
        app = cls.__new__(cls)
        app.download_events = queue.Queue()
        app.stop_event = threading.Event()
        return app

    def test_invalid_csv_location_never_executes(self):
        app = self.worker()
        df = pd.DataFrame({'location': ['$(untrusted command)'], 'deployment_id': ['01']})
        with tempfile.TemporaryDirectory() as folder, patch.object(subprocess, 'run') as run:
            app.download_and_rename_files(df, folder, False, 'gsutil')
            run.assert_not_called()
        events = list(app.download_events.queue)
        self.assertEqual(events[0][0], 'error')
        self.assertIn('1 failed', events[-1][1])

    def test_failed_download_leaves_no_partial_file(self):
        app = self.worker()
        df = pd.DataFrame({'location': ['gs://bucket/photo.jpg'], 'deployment_id': ['01']})
        with tempfile.TemporaryDirectory() as folder, patch.object(subprocess, 'run', side_effect=subprocess.CalledProcessError(1, 'gsutil')):
            app.download_and_rename_files(df, folder, False, 'gsutil')
            self.assertEqual(list(Path(folder).iterdir()), [])
        self.assertIn('1 failed', list(app.download_events.queue)[-1][1])


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
                (image_folder / 'photo.jpg').write_bytes(b'placeholder')
                cls = namespace['LynxOne']
                app = cls.__new__(cls)
                app.source_folder = folder
                app.lince_checkbox_var = SimpleNamespace(get=lambda: linces)
                app.revision_checkbox_var = SimpleNamespace(get=lambda: revision)
                app.minutes_entry = SimpleNamespace(get=lambda: '0')
                app.estaciones_file = app.individuos_file = None
                app.get_exif_data = lambda path: None
                app.generate_excel()
                row = app.excel_data.iloc[0]
                self.assertEqual(row['Finca'], 'Farm')
                self.assertEqual(row['Estación'], 'Station')
                self.assertEqual(row['Individuo'], 'Individual')
                if revision:
                    self.assertEqual(row['Revisión'], 'Revision01')
                else:
                    self.assertNotIn('Revisión', app.excel_data.columns)
