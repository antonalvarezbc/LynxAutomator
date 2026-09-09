from datetime import datetime
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from lynx_tasks import TaskContext
from lynx_video_dates import date_candidates, inspect_video_dates, parse_capture_date
from lynx_file_jobs import extract_videos


class VideoDatesTests(unittest.TestCase):
    def test_parse_keeps_camera_clock_and_explicit_offset(self):
        self.assertIsNone(parse_capture_date('2024-02-03 12:13:14').tzinfo)
        date = parse_capture_date('2024-02-03 12:13:14+02:00')
        self.assertEqual(date.hour, 12)
        self.assertEqual(date.utcoffset().total_seconds(), 7200)
        for text in ('', '2024-02-03', '0000-00-00 00:00:00', '2024-13-01 12:00:00'):
            with self.assertRaises(ValueError):
                parse_capture_date(text)

    def test_metadata_candidates_keep_provenance(self):
        result = date_candidates({'format': {'tags': {'creation_time': '2024-02-03T12:13:14Z'}},
                                  'streams': [{'tags': {'DateTimeOriginal': '2024-02-03 14:13:14'}}]})
        self.assertEqual(result[0][2], 'stream 0/DateTimeOriginal')
        self.assertEqual(result[1][2], 'format/creation_time')

    def test_missing_and_conflicting_metadata_do_not_use_file_date(self):
        with tempfile.TemporaryDirectory() as folder:
            video = Path(folder) / 'copied.mp4'
            video.write_bytes(b'video')
            with patch('lynx_video_dates.shutil.which', return_value=None):
                rows = inspect_video_dates(TaskContext(), folder)
            self.assertEqual(rows[0]['date'], '')
            data = {'format': {'tags': {'creation_time': '2024-01-01T12:00:00Z'}},
                    'streams': [{'tags': {'creation_time': '2025-01-01T12:00:00Z'}}]}
            with patch('lynx_video_dates.shutil.which', return_value='ffprobe'), \
                    patch('lynx_video_dates.subprocess.run', return_value=SimpleNamespace(stdout=json.dumps(data))):
                rows = inspect_video_dates(TaskContext(), folder)
            self.assertTrue(rows[0]['conflict'])
            self.assertEqual(rows[0]['date'], '')
            with self.assertRaisesRegex(ValueError, 'Confirm'):
                extract_videos(TaskContext(), folder, folder, 1)

    def test_single_metadata_date_is_only_a_proposal(self):
        with tempfile.TemporaryDirectory() as folder:
            (Path(folder) / 'copied.mp4').touch()
            data = {'format': {'tags': {'creation_time': '2024-01-01T12:00:00Z'}}}
            with patch('lynx_video_dates.shutil.which', return_value='ffprobe'), \
                    patch('lynx_video_dates.subprocess.run', return_value=SimpleNamespace(stdout=json.dumps(data))):
                rows = inspect_video_dates(TaskContext(), folder)
            self.assertEqual(rows[0]['date'], '2024-01-01 12:00:00+00:00')

    def test_real_embedded_date_survives_changed_filesystem_date(self):
        import os
        import shutil
        import subprocess
        if not shutil.which('ffmpeg') or not shutil.which('ffprobe'):
            self.skipTest('Requires FFmpeg/ffprobe')
        with tempfile.TemporaryDirectory() as folder:
            video = Path(folder) / 'copied.mov'
            subprocess.run(['ffmpeg', '-v', 'error', '-f', 'lavfi', '-i',
                            'color=black:s=16x16:d=1', '-c:v', 'mpeg4',
                            '-metadata', 'creation_time=2024-02-03T12:13:14Z', str(video)], check=True)
            os.utime(video, (1800000000, 1800000000))
            rows = inspect_video_dates(TaskContext(), folder)
            self.assertEqual(rows[0]['date'], '2024-02-03 12:13:14+00:00')
            import piexif
            destination = Path(folder) / 'frames'
            destination.mkdir()
            with patch('lynx_file_jobs.set_file_timestamp') as set_dates:
                result = extract_videos(TaskContext(), folder, destination, 1,
                                        {str(video): datetime(2024, 2, 3, 12, 13, 14)})
                set_dates.assert_not_called()
            self.assertEqual(result.errors, [])
            metadata = piexif.load(str(next(destination.glob('*.jpg'))))
            self.assertEqual(metadata['Exif'][piexif.ExifIFD.DateTimeOriginal], b'2024:02:03 12:13:14')
            self.assertNotIn(36881, metadata['Exif'])

