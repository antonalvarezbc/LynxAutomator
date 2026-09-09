import os
from pathlib import Path
import tempfile
from datetime import timedelta
import unittest
from unittest.mock import patch
from types import SimpleNamespace

import pandas as pd
import piexif
from PIL import Image

from lynx_core import file_timestamp, frame_step, merge_deployments, shift_file_date, unique_path


class FileDateTests(unittest.TestCase):
    def test_linux_uses_mtime_not_metadata_change_time(self):
        with patch('lynx_core.sys.platform', 'linux'), patch('lynx_core.os.stat', return_value=SimpleNamespace(st_mtime=100, st_ctime=900)):
            self.assertEqual(file_timestamp('unused'), 100)

    def test_mac_uses_editable_mtime(self):
        with patch('lynx_core.sys.platform', 'darwin'), patch('lynx_core.os.stat', return_value=SimpleNamespace(st_mtime=100, st_birthtime=50)):
            self.assertEqual(file_timestamp('unused'), 100)

    def test_shift_preserves_integer_exif_tags_and_updates_mtime_last(self):
        with tempfile.TemporaryDirectory() as folder:
            file = Path(folder) / 'photo.jpg'
            data = {'0th': {piexif.ImageIFD.Orientation: 1}, 'Exif': {
                piexif.ExifIFD.DateTimeOriginal: b'2024:01:02 03:04:05',
                piexif.ExifIFD.ISOSpeedRatings: 400}}
            Image.new('RGB', (8, 8)).save(file, exif=piexif.dump(data))
            shift_file_date(file, timedelta(hours=1), original_timestamp=1700000000)
            result = piexif.load(str(file))
            self.assertEqual(result['0th'][piexif.ImageIFD.Orientation], 1)
            self.assertEqual(result['Exif'][piexif.ExifIFD.ISOSpeedRatings], 400)
            self.assertEqual(result['Exif'][piexif.ExifIFD.DateTimeOriginal], b'2024:01:02 04:04:05')
            self.assertAlmostEqual(os.stat(file).st_mtime, 1700003600, places=2)

    def test_non_jpeg_contents_unchanged(self):
        with tempfile.TemporaryDirectory() as folder:
            file = Path(folder) / 'video.mp4'
            file.write_bytes(b'example video')
            shift_file_date(file, timedelta(seconds=20), original_timestamp=1000)
            self.assertEqual(file.read_bytes(), b'example video')
            self.assertAlmostEqual(file.stat().st_mtime, 1020, places=2)

    def test_multiple_name_collisions(self):
        with tempfile.TemporaryDirectory() as folder:
            file = Path(folder) / 'photo.jpg'
            file.touch()
            file.with_name('photo_1.jpg').touch()
            self.assertEqual(unique_path(file).name, 'photo_2.jpg')


class VideoTests(unittest.TestCase):
    def test_subframe_interval_has_no_zero_divisor(self):
        self.assertEqual(frame_step(25, 0.001), 1)

    def test_invalid_video_rates_and_intervals(self):
        for fps, interval in [(0, 1), (float('nan'), 1), (25, 0), (25, float('inf')), (25, -1)]:
            with self.subTest(fps=fps, interval=interval), self.assertRaises(ValueError):
                frame_step(fps, interval)


class MergeTests(unittest.TestCase):
    def setUp(self):
        self.images = pd.DataFrame({'project_id': ['01', '01'], 'deployment_id': ['001', '001'], 'location': ['a.jpg', 'b.jpg']})
        self.deployments = pd.DataFrame({'project_id': ['01'], 'deployment_id': ['001'], 'latitude': ['40']})

    def test_preserves_all_images_and_identifiers(self):
        merged = merge_deployments(self.images, self.deployments)
        self.assertEqual(merged['location'].tolist(), ['a.jpg', 'b.jpg'])
        self.assertEqual(merged['deployment_id'].tolist(), ['001', '001'])

    def test_missing_deployment_is_reported(self):
        self.images.loc[0, 'deployment_id'] = 'missing'
        with self.assertRaisesRegex(ValueError, 'no matching deployment'):
            merge_deployments(self.images, self.deployments)

    def test_duplicate_deployments_cannot_multiply_images(self):
        with self.assertRaises(pd.errors.MergeError):
            merge_deployments(self.images, pd.concat([self.deployments] * 2))

    def test_null_identifiers_cannot_match(self):
        self.images.loc[0, 'deployment_id'] = None
        with self.assertRaisesRegex(ValueError, 'cannot be empty'):
            merge_deployments(self.images, self.deployments)


if __name__ == '__main__':
    unittest.main()
