"""Desktop file selection with portable filters and remembered directories."""
import os
from pathlib import Path
import shutil
import subprocess
import sys

_last_directory = str(Path.home())


def _select(kind, **options):
    global _last_directory
    options.setdefault('initialdir', _last_directory)
    if 'filetypes' in options:
        filters = []
        for label, patterns in options['filetypes']:
            patterns = patterns.replace(';', ' ').split() if isinstance(patterns, str) else patterns
            expanded = tuple(dict.fromkeys(p for pattern in patterns for p in (pattern, pattern.upper())))
            filters.append((label, expanded))
        if not any('*' in patterns for _, patterns in filters):
            filters.append(('All files', ('*',)))
        options['filetypes'] = filters
    executable = shutil.which('zenity') if sys.platform.startswith('linux') else None
    result = None
    if executable:
        command = [executable, '--file-selection', '--width=900', '--height=650',
                   '--filename=' + str(Path(options['initialdir']) / options.get('initialfile', '')) + (os.sep if not options.get('initialfile') else '')]
        if options.get('parent') is not None:
            command.append('--attach=' + str(options['parent'].winfo_toplevel().winfo_id()))
        if kind == 'asksaveasfilename':
            command += ['--save', '--confirm-overwrite']
        if kind == 'askdirectory':
            command.append('--directory')
        if options.get('title'):
            command.append('--title=' + options['title'])
        for label, patterns in options.get('filetypes', []):
            command.append('--file-filter=' + label + ' | ' + ' '.join(patterns))
        env = os.environ.copy()
        # A frozen app's private libraries must not be loaded by system Zenity.
        if getattr(sys, 'frozen', False):
            if 'LD_LIBRARY_PATH_ORIG' in env:
                env['LD_LIBRARY_PATH'] = env['LD_LIBRARY_PATH_ORIG']
            else:
                env.pop('LD_LIBRARY_PATH', None)
        try:
            completed = subprocess.run(command, capture_output=True, text=True, env=env)
            if completed.returncode == 0:
                result = completed.stdout.rstrip('\n')
            elif completed.returncode == 1:
                result = ''
        except OSError:
            pass
    if result is None:
        from tkinter import filedialog as tk_dialogs
        result = getattr(tk_dialogs, kind)(**options)
    if result and kind == 'asksaveasfilename' and options.get('defaultextension') and not Path(result).suffix:
        result += options['defaultextension']
        if executable and Path(result).exists():
            # Zenity checked the unextended name; confirm the actual target too.
            confirm = subprocess.run([executable, '--question', '--no-markup', '--text=Overwrite / Sobrescribir: ' + result + '?'], env=env)
            if confirm.returncode != 0:
                result = ''
    if result:
        _last_directory = str(Path(result) if kind == 'askdirectory' else Path(result).parent)
    return result


def askopenfilename(**options):
    return _select('askopenfilename', **options)


def askdirectory(**options):
    return _select('askdirectory', **options)


def asksaveasfilename(**options):
    return _select('asksaveasfilename', **options)
