"""Tk-side adapters: snapshot inputs, start jobs, and apply results on the UI thread."""
from functools import wraps
import math
import shutil
import lynx_dialogs as filedialog
from tkinter import messagebox, TclError

import customtkinter as ctk

from lynx_tasks import BackgroundTask
from lynx_processing import FolderProcessor, CatalogProcessor, WIProcessor, LynxProcessor
from lynx_file_jobs import (BatchResult, extract_videos, scan_dates, change_dates,
                            rename_images, save_excel, download_images)

TEXT = {
    'es': {'ready': 'Listo', 'cancel': 'Cancelar', 'cancelling': 'Cancelando; esperando a que termine la operación en curso…',
           'cancelled': 'Cancelado. Se conservan los archivos ya completados.', 'done': 'Completado',
           'busy': 'Espera a que termine la tarea actual.', 'failed': 'La tarea no se completó',
           'summary': '{completed} completados, {skipped} omitidos, {failed} fallidos.', 'empty': 'No hay datos para guardar.'},
    'pt': {'ready': 'Pronto', 'cancel': 'Cancelar', 'cancelling': 'A cancelar; a aguardar o fim da operação em curso…',
           'cancelled': 'Cancelado. Os ficheiros já concluídos foram preservados.', 'done': 'Concluído',
           'busy': 'Aguarde a conclusão da tarefa atual.', 'failed': 'A tarefa não foi concluída',
           'summary': '{completed} concluídos, {skipped} ignorados, {failed} falhados.', 'empty': 'Não há dados para guardar.'},
    'en': {'ready': 'Ready', 'cancel': 'Cancel', 'cancelling': 'Cancelling; waiting for the current operation to finish…',
           'cancelled': 'Cancelled. Already completed files are preserved.', 'done': 'Completed',
           'busy': 'Wait for the current task to finish.', 'failed': 'The task did not complete',
           'summary': '{completed} completed, {skipped} skipped, {failed} failed.', 'empty': 'No data to save.'},
}


class JobPanel(ctk.CTkFrame):
    def __init__(self, root, lang):
        super().__init__(root)
        self.root = root
        self.text = TEXT[lang]
        self.engine = BackgroundTask()
        self.closing = False
        self.disabled = []
        self.status = ctk.CTkLabel(self, text=self.text['ready'], anchor='w', wraplength=850)
        self.status.pack(side='top', fill='x', padx=10)
        self.bar = ctk.CTkProgressBar(self)
        self.bar.set(0)
        self.bar.pack(side='left', fill='x', expand=True, padx=10, pady=8)
        self.cancel_button = ctk.CTkButton(self, text=self.text['cancel'], command=self.cancel, state='disabled')
        self.cancel_button.pack(side='right', padx=10, pady=8)
        self.pack(side='bottom', fill='x')
        root.protocol('WM_DELETE_WINDOW', self.close)

    @property
    def busy(self):
        return self.engine.busy

    def _lock_controls(self, parent):
        for widget in parent.winfo_children():
            if widget is self:
                continue
            if isinstance(widget, (ctk.CTkButton, ctk.CTkEntry, ctk.CTkOptionMenu,
                                   ctk.CTkCheckBox, ctk.CTkRadioButton, ctk.CTkSwitch)):
                state = widget.cget('state')
                self.disabled.append((widget, state))
                widget.configure(state='disabled')
            else:
                self._lock_controls(widget)

    def start(self, title, work, on_success=None, on_finished=None):
        if self.busy:
            raise ValueError(self.text['busy'])
        self.title = title
        self.on_success = on_success
        self.on_finished = on_finished
        self.disabled = []
        self._lock_controls(self.root)
        self.cancel_button.configure(state='normal')
        self.status.configure(text=title)
        self.bar.configure(mode='indeterminate')
        self.bar.start()
        try:
            self.engine.start(work)
        except Exception:
            self._restore()
            raise
        self.after(75, self._poll)

    def _restore(self):
        self.bar.stop()
        self.bar.configure(mode='determinate')
        self.cancel_button.configure(state='disabled')
        for widget, state in self.disabled:
            widget.configure(state=state)
        self.disabled = []

    def cancel(self):
        self.engine.cancel()
        self.cancel_button.configure(state='disabled')
        self.status.configure(text=self.text['cancelling'])

    def close(self):
        if self.busy:
            self.closing = True
            self.cancel()
        else:
            self.root.destroy()

    def _poll(self):
        result = self.engine.take_result()
        if result is None:
            fraction, detail = self.engine.context.progress()
            if not self.engine.context.stop.is_set():
                self.status.configure(text=f'{self.title}: {detail}')
            if fraction is not None:
                self.bar.stop()
                self.bar.configure(mode='determinate')
                self.bar.set(max(0, min(fraction, 1)))
            self.after(75, self._poll)
            return
        self._restore()
        if self.closing:
            self.root.destroy()
            return
        on_success, on_finished = self.on_success, self.on_finished
        kind, value = result
        if kind == 'cancelled':
            self.status.configure(text=self.text['cancelled'])
        elif kind == 'error':
            self.status.configure(text=self.text['failed'])
            messagebox.showerror(self.text['failed'], str(value))
        else:
            self.bar.set(1)
            self.status.configure(text=self.text['done'])
            if isinstance(value, BatchResult):
                summary = self.text['summary'].format(completed=value.completed, skipped=value.skipped, failed=len(value.errors))
                self.status.configure(text=summary)
                if value.errors:
                    # One summary, not a blocking dialog per file.
                    messagebox.showwarning(self.text['failed'], summary + '\n\n' + '\n'.join(value.errors[:20]))
            if on_success:
                on_success(value)
        if on_finished:
            on_finished(kind)


def action(method):
    @wraps(method)
    def wrapped(self, *args, **kwargs):
        try:
            if self.root.winfo_toplevel().jobs.busy:
                raise ValueError(TEXT[self.lang]['busy'])
            return method(self, *args, **kwargs)
        except (ValueError, TypeError, OSError, KeyError, TclError) as exc:
            messagebox.showerror('Error', str(exc))
    return wrapped


def start(owner, title, work, callback=None, on_finished=None):
    owner.root.winfo_toplevel().jobs.start(title, work, callback, on_finished)


def integer(value):
    number = int(value)
    if number < 0:
        raise ValueError('Use a nonnegative integer.')
    return number


def require_paths(*paths):
    if not all(paths):
        raise ValueError('Select all input files/folders first.')


class SpreadsheetJobs:
    def _clear_result(self, button):
        self._result = None
        button.configure(state='disabled')

    def _receive_result(self, data, button):
        self._result = data
        button.configure(state='normal' if data is not None and not data.empty else 'disabled')
        if data is None or data.empty:
            messagebox.showinfo('Information', TEXT[self.lang]['empty'])

    @action
    def save_file(self):
        data = getattr(self, '_result', None)
        if data is None or data.empty:
            raise ValueError(TEXT[self.lang]['empty'])
        path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if path:
            start(self, 'Excel', lambda task: save_excel(task, data, path))

    download_file = save_file
    save_excel = save_file


class FolderJobs(SpreadsheetJobs):
    @action
    def process_files(self):
        self._clear_result(self.download_btn)
        folder, template = getattr(self, 'folder_path', None), getattr(self, 'file_path', None)
        require_paths(folder, template)
        group = bool(self.multiple_images_var.get())
        threshold = integer(self.time_threshold_entry.get()) if group else 0
        start(self, 'BIWbE', lambda task: FolderProcessor(task, folder, template, group, threshold).run(),
              lambda data: self._receive_result(data, self.download_btn))


class CatalogJobs(SpreadsheetJobs):
    @action
    def process_files(self):
        self._clear_result(self.download_btn)
        folder, template = getattr(self, 'folder_path', None), getattr(self, 'file_path', None)
        require_paths(folder, template)
        capitalize, collapse = bool(self.capitalize_var.get()), bool(self.collapse_var.get())
        start(self, 'BIWbE Catalog', lambda task: CatalogProcessor(task, folder, template, capitalize, collapse).run(),
              lambda data: self._receive_result(data, self.download_btn))


class WIJobs(SpreadsheetJobs):
    @action
    def process_files(self):
        self._clear_result(self.download_btn)
        paths = (self.initial_excel_path, self.images_csv_path, self.deployments_csv_path)
        require_paths(*paths)
        group = bool(self.multiple_images_var.get())
        separate = bool(self.separate_large_groups_var.get()) if hasattr(self, 'separate_large_groups_var') else False
        threshold = integer(self.time_threshold_entry.get()) if group else 0
        start(self, 'WI → BIWbE', lambda task: WIProcessor(task, *paths, group, separate, threshold).run(),
              lambda data: self._receive_result(data, self.download_btn))


class LynxJobs(SpreadsheetJobs):
    @action
    def generate_excel(self):
        self._clear_result(self.download_button)
        folder = self.source_folder
        require_paths(folder)
        options = (bool(self.lince_checkbox_var.get()), bool(self.revision_checkbox_var.get()),
                   integer(self.minutes_entry.get()), self.estaciones_file, self.individuos_file)
        start(self, 'Lynx', lambda task: LynxProcessor(task, folder, *options).run(),
              lambda data: self._receive_result(data, self.download_button))


class VideoJobs:
    @action
    def start_extraction(self):
        folder = getattr(self, 'folder_path', None)
        require_paths(folder)
        interval = float(self.interval_var.get())
        if not math.isfinite(interval) or interval <= 0:
            raise ValueError('Use a finite interval greater than zero.')
        destination = filedialog.askdirectory(title='Select output folder')
        if destination:
            start(self, 'Video', lambda task: extract_videos(task, folder, destination, interval))


class DateJobs:
    def _receive_dates(self, dates):
        oldest, newest = dates
        self.oldest_date.set(oldest)
        self.newest_date.set(newest)

    @action
    def get_file_dates(self, folder):
        self.oldest_date.set('')
        self.newest_date.set('')
        start(self, 'EXIF', lambda task: scan_dates(task, folder), self._receive_dates)

    @action
    def change_dates(self):
        difference = self.get_date_difference()
        folder = self.selected_folder
        self.oldest_date.set('')
        self.newest_date.set('')
        start(self, 'EXIF', lambda task: change_dates(task, folder, difference),
              on_finished=lambda kind: self.get_file_dates(folder))

    @action
    def copy_to_folder(self):
        difference = self.get_date_difference()
        folder = self.selected_folder
        destination = filedialog.askdirectory(title='Select destination folder')
        if destination:
            start(self, 'EXIF', lambda task: change_dates(task, folder, difference, destination))


class RenamerJobs:
    @action
    def rename_and_copy_photos(self):
        source = self.source_folder
        destination = self.dest_folder if self.copy_photos_var.get() else None
        require_paths(source)
        if self.copy_photos_var.get():
            require_paths(destination)
        options = dict(folder=bool(self.use_folder_name_var.get()), custom=bool(self.use_custom_text_var.get()),
                       text=self.custom_text_var.get().strip(), date=bool(self.add_exif_date_var.get()),
                       original=bool(self.keep_original_name_var.get()), underscores=bool(self.replace_spaces_var.get()))
        start(self, 'Images', lambda task: rename_images(task, source, destination, options))


class DownloadJobs:
    @action
    def start_download(self):
        csv_path = getattr(self, 'csv_path', None)
        require_paths(csv_path)
        executable = shutil.which('gsutil')
        if not executable:
            raise ValueError('Install Google Cloud CLI with gsutil and configure access first.')
        multiple = bool(self.use_multiple_folders.get())
        folder = filedialog.askdirectory(title='Select destination folder')
        if folder:
            start(self, 'WI Download', lambda task: download_images(task, csv_path, folder, multiple, executable))

    def stop_download(self):
        self.root.winfo_toplevel().jobs.cancel()
