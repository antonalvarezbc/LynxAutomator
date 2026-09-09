"""Read embedded video dates without treating filesystem dates as capture dates."""
from datetime import datetime
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys

VIDEO_EXTENSIONS = {'.mp4', '.avi', '.mov', '.mkv', '.flv'}


def parse_capture_date(value):
    value = value.strip()
    if not re.match(r'^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}', value):
        raise ValueError('Use YYYY-MM-DD HH:MM:SS, optionally followed by +HH:MM or Z.')
    date = datetime.fromisoformat(value.replace('Z', '+00:00'))
    if date.year < 1970:
        raise ValueError('Check the capture year (must be 1970 or later).')
    return date


def date_candidates(data):
    candidates = []
    sections = [('format', data.get('format', {}))]
    sections += [(f'stream {i}', stream) for i, stream in enumerate(data.get('streams', []))]
    priorities = {'datetimeoriginal': 0, 'date_time_original': 0, 'com.apple.quicktime.creationdate': 1,
                  'creation_time': 2, 'date': 3}
    for section, values in sections:
        for tag, value in values.get('tags', {}).items():
            if tag.lower() not in priorities:
                continue
            try:
                date = parse_capture_date(str(value))
            except ValueError:
                continue
            candidates.append((priorities[tag.lower()], date, f'{section}/{tag}'))
    return sorted(candidates, key=lambda item: item[0])


def inspect_video_dates(task, folder):
    videos = sorted(p for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower() in VIDEO_EXTENSIONS)
    if not videos:
        raise ValueError('No videos found in the selected folder.')
    executable = shutil.which('ffprobe')
    rows = []
    env = os.environ.copy()
    if getattr(sys, 'frozen', False):
        if 'LD_LIBRARY_PATH_ORIG' in env:
            env['LD_LIBRARY_PATH'] = env['LD_LIBRARY_PATH_ORIG']
        else:
            env.pop('LD_LIBRARY_PATH', None)
    for index, video in enumerate(videos):
        task.checkpoint()
        candidates = []
        reason = 'ffprobe unavailable' if not executable else 'No readable capture date'
        if executable:
            try:
                result = subprocess.run([executable, '-v', 'error', '-show_entries',
                                         'format_tags:stream_tags', '-of', 'json', str(video.resolve())],
                                        capture_output=True, text=True, timeout=20, check=True, env=env,
                                        creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                candidates = date_candidates(json.loads(result.stdout))
            except (OSError, subprocess.SubprocessError, ValueError, TypeError, AttributeError):
                reason = 'Metadata could not be read'
        task.checkpoint()
        distinct = {date.isoformat() for _, date, _ in candidates}
        proposed = candidates[0][1].isoformat(sep=' ') if len(distinct) == 1 else ''
        source = '\n'.join(f'{source}: {date.isoformat(sep=" ")}' for _, date, source in candidates) or reason
        rows.append({'path': str(video), 'date': proposed, 'source': source,
                     'conflict': len(distinct) > 1})
        task.report(video.name, (index + 1) / len(videos))
    return rows
