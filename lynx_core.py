"""File/date operations shared by the desktop variants."""
import math
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import piexif


def file_timestamp(path):
    """Creation time on Windows; editable modification time on Unix (never ctime)."""
    stat = os.stat(path)
    if sys.platform == "win32":
        return stat.st_ctime
    return stat.st_mtime


def set_file_timestamp(path, timestamp):
    """Set mtime/atime everywhere, and creation time on Windows."""
    os.utime(path, (timestamp, timestamp))
    if sys.platform == "win32":
        import win32file
        handle = win32file.CreateFile(str(path), win32file.GENERIC_WRITE, 0, None,
                                     win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
        try:
            value = datetime.fromtimestamp(timestamp, tz=timezone.utc)
            win32file.SetFileTime(handle, value, value, value)
        finally:
            handle.Close()


def shift_file_date(path, difference, original_timestamp=None):
    """Shift EXIF capture dates without changing the types of unrelated tags."""
    timestamp = file_timestamp(path) if original_timestamp is None else original_timestamp
    # piexif.insert supports JPEG, not arbitrary image/video formats.
    if Path(path).suffix.lower() in {".jpg", ".jpeg"}:
        data = piexif.load(str(path))
        original = data["Exif"].get(piexif.ExifIFD.DateTimeOriginal)
        if original:
            date = datetime.strptime(original.decode("ascii"), "%Y:%m:%d %H:%M:%S") + difference
            value = date.strftime("%Y:%m:%d %H:%M:%S").encode("ascii")
            data["Exif"][piexif.ExifIFD.DateTimeOriginal] = value
            data["Exif"][piexif.ExifIFD.DateTimeDigitized] = value
            data["0th"][piexif.ImageIFD.DateTime] = value
            piexif.insert(piexif.dump(data), str(path))
    # Do this last: writing EXIF changes mtime.
    set_file_timestamp(path, timestamp + difference.total_seconds())


def frame_step(fps, interval):
    if not math.isfinite(fps) or fps <= 0:
        raise ValueError("The video has an invalid frame rate.")
    if not math.isfinite(interval) or interval <= 0:
        raise ValueError("The interval must be finite and greater than zero.")
    return max(1, round(fps * interval))


def unique_path(path):
    path = Path(path)
    candidate = path
    counter = 1
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        counter += 1
    return candidate


def merge_deployments(images, deployments):
    keys = ["project_id", "deployment_id"]
    for data in (images, deployments):
        if data[keys].isna().any().any() or data[keys].eq("").any().any():
            raise ValueError("Project and deployment identifiers cannot be empty.")
    merged = images.merge(deployments, on=keys, how="left", validate="many_to_one",
                          indicator=True, suffixes=("_image", "_deployment"))
    missing = merged["_merge"].eq("left_only").sum()
    if missing:
        raise ValueError(f"{missing} images have no matching deployment. Check the CSV files.")
    return merged.drop(columns="_merge")
