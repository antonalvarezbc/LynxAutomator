from datetime import datetime
from lynx_core import file_timestamp
from pathlib import Path
import tempfile
from threading import Event, get_ident
import unittest
from unittest.mock import patch

import cv2
import numpy as np
import pandas as pd
from PIL import Image
import piexif

from lynx_tasks import BackgroundTask, TaskContext, TaskCancelled
from lynx_file_jobs import extract_videos, save_excel, rename_images


class TaskTests(unittest.TestCase):
    def test_worker_is_separate_and_busy_until_result_consumed(self):
        engine = BackgroundTask()
        entered, release = Event(), Event()
        owner = get_ident()
        def work(task):
            entered.set()
            release.wait(3)
            return get_ident()
        engine.start(work)
        try:
            self.assertTrue(entered.wait(2))
            self.assertTrue(engine.busy)
            self.assertIsNone(engine.take_result())
            with self.assertRaises(RuntimeError):
                engine.start(work)
        finally:
            release.set()
            engine.thread.join(3)
        kind, worker = engine.take_result()
        self.assertEqual(kind, 'success')
        self.assertNotEqual(worker, owner)
        self.assertFalse(engine.busy)

    def test_cancellation_and_restart(self):
        engine = BackgroundTask()
        entered, release = Event(), Event()
        def work(task):
            entered.set()
            release.wait(3)
            task.checkpoint()
        engine.start(work)
        self.assertTrue(entered.wait(2))
        engine.cancel()
        release.set()
        engine.thread.join(3)
        self.assertEqual(engine.take_result(), ('cancelled', None))
        engine.start(lambda task: 42)
        engine.thread.join(3)
        self.assertEqual(engine.take_result(), ('success', 42))

    def test_failure_is_delivered_to_ui_not_lost_in_thread(self):
        engine = BackgroundTask()
        def fail(task):
            raise ValueError('invalid file')
        engine.start(fail)
        engine.thread.join(3)
        kind, error = engine.take_result()
        self.assertEqual(kind, 'error')
        self.assertEqual(str(error), 'invalid file')

    def test_progress_keeps_latest_update(self):
        task = TaskContext()
        for index in range(10000):
            task.report(str(index), index / 10000)
        self.assertEqual(task.progress(), (0.9999, '9999'))


class MediaTaskTests(unittest.TestCase):
    def fake_capture(self, count=5):
        class Capture:
            released = False
            def isOpened(self): return True
            def get(self, key): return 25 if key == cv2.CAP_PROP_FPS else count
            def read(self):
                return True, np.zeros((8, 8, 3), dtype=np.uint8)
            def release(self): self.released = True
        return Capture()

    def test_cancel_video_preserves_completed_frames_and_releases_capture(self):
        with tempfile.TemporaryDirectory() as folder:
            source, output = Path(folder) / 'input', Path(folder) / 'output'
            source.mkdir()
            output.mkdir()
            (source / 'test.mp4').write_bytes(b'video')
            task = TaskContext()
            capture = self.fake_capture()
            reports = []
            def report(text, fraction=None):
                reports.append(text)
                task.stop.set()
            task.report = report
            with patch('lynx_file_jobs.cv2.VideoCapture', return_value=capture):
                with self.assertRaises(TaskCancelled):
                    extract_videos(task, source, output, 0.01)
            self.assertTrue(capture.released)
            files = list(output.iterdir())
            self.assertEqual(len(files), 1)
            self.assertEqual(files[0].suffix, '.jpg')
            with Image.open(files[0]) as image:
                image.verify()
            self.assertIn(piexif.ExifIFD.DateTimeOriginal, piexif.load(str(files[0]))['Exif'])
            self.assertTrue(reports)

    def test_video_write_failure_is_reported_and_releases_capture(self):
        with tempfile.TemporaryDirectory() as folder:
            source, output = Path(folder) / 'input', Path(folder) / 'output'
            source.mkdir()
            output.mkdir()
            (source / 'test.mp4').touch()
            capture = self.fake_capture()
            with patch('lynx_file_jobs.cv2.VideoCapture', return_value=capture), patch('lynx_file_jobs.cv2.imwrite', return_value=False):
                result = extract_videos(TaskContext(), source, output, 1)
            self.assertEqual(result.completed, 0)
            self.assertEqual(len(result.errors), 1)
            self.assertTrue(capture.released)
            self.assertEqual(list(output.iterdir()), [])

    def test_cancelled_excel_export_preserves_existing_workbook(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / 'result.xlsx'
            target.write_bytes(b'original')
            task = TaskContext()
            task.stop.set()
            with self.assertRaises(TaskCancelled):
                save_excel(task, pd.DataFrame({'a': [1]}), target)
            self.assertEqual(target.read_bytes(), b'original')
            self.assertEqual(list(Path(folder).iterdir()), [target])

    def test_rename_does_not_recurse_into_its_output(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder)
            output = source / 'output'
            output.mkdir()
            Image.new('RGB', (4, 4)).save(source / 'photo.jpg')
            Image.new('RGB', (4, 4)).save(output / 'existing.jpg')
            options = dict(folder=False, custom=False, text='', date=True, original=True, underscores=False)
            result = rename_images(TaskContext(), source, output, options)
            self.assertEqual(result.completed, 1)
            self.assertEqual(len(list(output.iterdir())), 2)
            self.assertTrue((source / 'photo.jpg').exists())
            self.assertFalse(any('None' in p.name for p in output.iterdir()))

    def test_real_video_extracts_expected_frames_and_excel_saves(self):
        with tempfile.TemporaryDirectory() as folder:
            source, output = Path(folder) / 'input', Path(folder) / 'output'
            source.mkdir()
            output.mkdir()
            video = source / 'clip.avi'
            writer = cv2.VideoWriter(str(video), cv2.VideoWriter_fourcc(*'MJPG'), 10, (16, 16))
            self.assertTrue(writer.isOpened(), 'MJPG encoder unavailable')
            try:
                for index in range(10):
                    writer.write(np.full((16, 16, 3), index * 20, dtype=np.uint8))
            finally:
                writer.release()
            result = extract_videos(TaskContext(), source, output, 0.2)
            self.assertEqual(result.errors, [])
            self.assertEqual(result.completed, 5)
            self.assertEqual(len(list(output.glob('*.jpg'))), 5)
            base = file_timestamp(video)
            for index, frame in enumerate(sorted(output.glob('*.jpg'))):
                expected = datetime.fromtimestamp(base + index * 0.2)
                metadata = piexif.load(str(frame))
                for tag in (piexif.ExifIFD.DateTimeOriginal, piexif.ExifIFD.DateTimeDigitized):
                    self.assertEqual(metadata['Exif'][tag], expected.strftime('%Y:%m:%d %H:%M:%S').encode('ascii'))
                self.assertEqual(metadata['0th'][piexif.ImageIFD.DateTime], metadata['Exif'][piexif.ExifIFD.DateTimeOriginal])
                self.assertEqual(metadata['Exif'][piexif.ExifIFD.SubSecTimeOriginal], f'{expected.microsecond:06d}'.encode('ascii'))
                self.assertAlmostEqual(frame.stat().st_mtime, base + index * 0.2, delta=0.001)
                with Image.open(frame) as image:
                    self.assertEqual(image.size, (16, 16))
            workbook = output / 'result.xlsx'
            save_excel(TaskContext(), pd.DataFrame({'frames': [result.completed]}), workbook)
            self.assertEqual(pd.read_excel(workbook).iloc[0]['frames'], 5)
