"""Cancellable media and filesystem jobs. These functions never access widgets."""
from dataclasses import dataclass, field
from datetime import datetime
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile

import cv2
import pandas as pd
import piexif
from PIL import Image

from lynx_core import file_timestamp, frame_step, set_file_timestamp, shift_file_date, unique_path
from lynx_processing import read_exif, date_taken
from lynx_tasks import TaskCancelled


@dataclass
class BatchResult:
    completed: int = 0
    skipped: int = 0
    errors: list = field(default_factory=list)


def files_in(task, folder):
    files = []
    with os.scandir(folder) as entries:
        for entry in entries:
            task.checkpoint()
            if entry.is_file():
                files.append(Path(entry.path))
    return sorted(files)


def extract_videos(task, folder, destination, interval):
    if not math.isfinite(interval) or interval <= 0:
        raise ValueError('The interval must be finite and greater than zero.')
    videos = [p for p in files_in(task, folder) if p.suffix.lower() in {'.mp4', '.avi', '.mov', '.mkv', '.flv'}]
    if not videos:
        raise ValueError('No videos found in the selected folder.')
    result = BatchResult()
    for video_index, video in enumerate(videos):
        task.checkpoint()
        capture = cv2.VideoCapture(str(video))
        try:
            if not capture.isOpened():
                raise ValueError('Cannot open video.')
            fps = capture.get(cv2.CAP_PROP_FPS)
            step = frame_step(fps, interval)
            total = capture.get(cv2.CAP_PROP_FRAME_COUNT)
            total = total if math.isfinite(total) and total > 0 else 0
            timestamp = file_timestamp(video)
            frame = 0
            output_index = 0
            while True:
                task.checkpoint()
                success, pixels = capture.read()
                if not success:
                    if frame == 0:
                        raise ValueError('The video has no readable frames.')
                    break
                if frame % step == 0:
                    frame_timestamp = timestamp + frame / fps
                    date = datetime.fromtimestamp(frame_timestamp)
                    exif_date = date.strftime('%Y:%m:%d %H:%M:%S').encode('ascii')
                    subseconds = f'{date.microsecond:06d}'.encode('ascii')
                    metadata = piexif.dump({
                        '0th': {piexif.ImageIFD.DateTime: exif_date},
                        'Exif': {
                            piexif.ExifIFD.DateTimeOriginal: exif_date,
                            piexif.ExifIFD.DateTimeDigitized: exif_date,
                            piexif.ExifIFD.SubSecTime: subseconds,
                            piexif.ExifIFD.SubSecTimeOriginal: subseconds,
                            piexif.ExifIFD.SubSecTimeDigitized: subseconds}})
                    target = unique_path(Path(destination) / f'{video.name}_frame{output_index}.jpg')
                    with tempfile.TemporaryDirectory(dir=destination, prefix='.lynx-') as staging:
                        temporary = Path(staging) / 'frame.jpg'
                        if not cv2.imwrite(str(temporary), pixels, [cv2.IMWRITE_JPEG_QUALITY, 100]):
                            raise OSError('Could not write frame.')
                        piexif.insert(metadata, str(temporary))
                        set_file_timestamp(temporary, frame_timestamp)
                        task.checkpoint()
                        os.replace(temporary, target)
                    result.completed += 1
                    output_index += 1
                frame += 1
                fraction = (video_index + min(frame / total, 1)) / len(videos) if total else None
                task.report(f'{video.name}: {result.completed} frames', fraction)
        except TaskCancelled:
            raise
        except Exception as exc:
            result.errors.append(f'{video.name}: {exc}')
        finally:
            capture.release()
    return result


def scan_dates(task, folder):
    oldest = newest = None
    for index, path in enumerate(files_in(task, folder), 1):
        task.checkpoint()
        date = date_taken(read_exif(path))
        timestamp = date.timestamp() if date else file_timestamp(path)
        oldest = timestamp if oldest is None else min(oldest, timestamp)
        newest = timestamp if newest is None else max(newest, timestamp)
        task.report(f'{index}: {path.name}')
    return tuple(datetime.fromtimestamp(t).strftime('%Y-%m-%d %H:%M:%S') if t is not None else ''
                 for t in (oldest, newest))


def change_dates(task, folder, difference, destination=None):
    result = BatchResult()
    # Snapshot before copying, even when the destination is the source folder.
    paths = files_in(task, folder)
    for index, path in enumerate(paths):
        task.checkpoint()
        try:
            if destination:
                timestamp = file_timestamp(path)
                target = unique_path(Path(destination) / path.name)
                with tempfile.TemporaryDirectory(dir=destination, prefix='.lynx-') as staging:
                    temporary = Path(staging) / path.name
                    shutil.copy2(path, temporary)
                    shift_file_date(temporary, difference, timestamp)
                    task.checkpoint()
                    os.replace(temporary, target)
            else:
                # Finish each original-file update before observing cancellation.
                shift_file_date(path, difference)
            result.completed += 1
        except TaskCancelled:
            raise
        except Exception as exc:
            result.errors.append(f'{path.name}: {exc}')
        task.report(f'{result.completed}: {path.name}', (index + 1) / len(paths))
    return result


def rename_images(task, source, destination, options):
    source = Path(source)
    destination = Path(destination) if destination else None
    if destination and source.resolve() == destination.resolve():
        raise ValueError('Select a destination different from the source.')
    if any(char in options['text'] for char in '/\\'):
        raise ValueError('Custom text cannot contain path separators.')
    paths = []
    # Snapshot avoids visiting freshly renamed files or generated copies.
    for root, directories, names in os.walk(source):
        task.checkpoint()
        if destination:
            directories[:] = [d for d in directories if (Path(root) / d).resolve() != destination.resolve()]
        for name in names:
            task.checkpoint()
            path = Path(root) / name
            if path.suffix.lower() in {'.jpg', '.jpeg', '.png', '.gif'}:
                paths.append(path)
    result = BatchResult()
    for index, path in enumerate(paths):
        task.checkpoint()
        try:
            base = path.parent.name.split()[0].title() if options['folder'] else ''
            if options['custom']:
                base += f" {options['text']}"
            date = ''
            if options['date']:
                value = read_exif(path).get('DateTimeOriginal')
                if value:
                    date = ' ' + value.replace(':', '-').replace(' ', '_')
            original = path.stem.title() if options['original'] else ''
            name = f'{base} {date}{original}{path.suffix}'
            if options['underscores']:
                name = name.replace(' ', '_')
            target = (destination or path.parent) / name
            if destination is None and target == path:
                result.skipped += 1
                continue
            target = unique_path(target)
            if destination:
                with tempfile.TemporaryDirectory(dir=destination, prefix='.lynx-') as staging:
                    temporary = Path(staging) / 'copy'
                    shutil.copy2(path, temporary)
                    task.checkpoint()
                    os.replace(temporary, target)
            else:
                path.rename(target)
            result.completed += 1
        except TaskCancelled:
            raise
        except Exception as exc:
            result.errors.append(f'{path.name}: {exc}')
        task.report(f'{result.completed}: {path.name}', (index + 1) / len(paths))
    return result


def save_excel(task, data, path):
    task.report('Writing Excel')
    path = Path(path)
    # Write on the destination filesystem and commit only when fully written.
    with tempfile.TemporaryDirectory(dir=path.parent, prefix='.lynx-') as staging:
        temporary = Path(staging) / 'result.xlsx'
        data.to_excel(temporary, index=False)
        task.checkpoint()
        os.replace(temporary, path)
    return str(path)


def download_images(task, csv_path, folder, multiple_folders, executable):
    task.report('Reading images CSV')
    df = pd.read_csv(csv_path, dtype=str)
    task.checkpoint()
    if not {'location', 'deployment_id'}.issubset(df.columns):
        raise ValueError('The CSV must contain location and deployment_id columns.')
    if df[['location', 'deployment_id']].isna().any().any():
        raise ValueError('Locations and deployment IDs cannot be empty.')
    result = BatchResult()
    targets = {}
    clean = lambda name: re.sub(r'[^a-zA-Z0-9._Ññ\-() ]', '_', name)
    for index, row in enumerate(df.itertuples(index=False), 1):
        task.checkpoint()
        try:
            url = str(row.location)
            if not url.startswith('gs://') or any(c in url for c in '\r\n'):
                raise ValueError('Expected a gs:// image location.')
            filename = url.rsplit('/', 1)[-1]
            if Path(filename).suffix.lower() not in {'.jpg', '.jpeg'}:
                raise ValueError('Expected a JPEG image.')
            deployment = clean(str(row.deployment_id))
            if deployment in {'', '.', '..'}:
                raise ValueError('Invalid deployment folder name.')
            destination = Path(folder) / deployment if multiple_folders else Path(folder)
            destination.mkdir(parents=True, exist_ok=True)
            target = destination / (Path(clean(filename)).stem + '.JPG')
            key = str(target).casefold()
            if key in targets and targets[key] != url:
                raise ValueError(f'Different images map to the same output name: {target.name}')
            targets[key] = url
            if target.exists():
                result.skipped += 1
                continue
            task.report(f'{index}/{len(df)}: {filename}', (index - 1) / len(df))
            with tempfile.TemporaryDirectory(dir=destination, prefix='.lynx-') as staging:
                temporary = Path(staging) / 'download.jpg'
                subprocess.run([executable, 'cp', url, str(temporary)], check=True,
                               capture_output=True, text=True, timeout=300)
                with Image.open(temporary) as image:
                    if image.format != 'JPEG':
                        raise ValueError('The downloaded file is not JPEG.')
                    image.verify()
                # Finish the current transfer, then observe cancellation next iteration.
                os.replace(temporary, target)
            result.completed += 1
        except Exception as exc:
            result.errors.append(f'{row.location}: {getattr(exc, "stderr", None) or exc}')
    task.checkpoint()
    return result
