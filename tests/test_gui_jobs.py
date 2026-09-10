"""Run with LYNX_GUI_TESTS=1 and a display (xvfb-run on headless Linux)."""
import importlib.util
import os
from pathlib import Path
from threading import Event, get_ident
import time
import unittest
from unittest.mock import patch


@unittest.skipUnless(os.environ.get('LYNX_GUI_TESTS') == '1', 'Requires Tk and a display; run in GUI CI step')
class GuiJobsTests(unittest.TestCase):
    def setUp(self):
        import customtkinter as ctk
        from lynx_ui_jobs import JobPanel
        self.ctk = ctk
        self.root = ctk.CTk()
        self.root.jobs = JobPanel(self.root, 'pt')
        self.button = ctk.CTkButton(self.root, text='Run')
        self.button.pack()
        self.root.update()

    def tearDown(self):
        try:
            if self.root.jobs.busy:
                self.root.jobs.cancel()
                self.wait_until(lambda: not self.root.jobs.busy)
            for timer in self.root.tk.call("after", "info"):
                self.root.after_cancel(timer)
            self.root.destroy()
        except Exception:
            pass

    def wait_until(self, condition, seconds=5):
        deadline = time.monotonic() + seconds
        while not condition() and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.005)
        self.assertTrue(condition(), 'GUI callback did not finish')

    def test_camtrap_multiselect_preview_and_local_download(self):
        from lynx_camtrap_ui import CamtrapTab
        from lynx_camtrap import read_package
        from lynx_tasks import TaskContext
        import tempfile
        tab = CamtrapTab(self.root, lang='es')
        fixture = Path(__file__).parent / 'fixtures' / 'camtrap_dp_lynx_synthetic' / 'datapackage.json'
        tab.receive(read_package(TaskContext(), fixture))
        tab.checks['Lynx pardinus'].select()
        tab.review()
        self.wait_until(lambda: not self.root.jobs.busy)
        self.assertEqual(len(tab.records), 247)
        self.assertEqual(tab.download.cget('state'), 'normal')
        tab.local.select()
        with tempfile.TemporaryDirectory() as folder, patch('lynx_camtrap_ui.filedialog.askdirectory', return_value=folder):
            tab.obtain()
            self.wait_until(lambda: not self.root.jobs.busy)
            self.assertEqual(len(list(Path(folder).rglob('*.jpg'))), 10)
        tab.select_all(False)
        self.assertIsNone(tab.records)
        self.assertEqual(tab.download.cget('state'), 'disabled')

    def test_video_review_requires_explicit_confirmation(self):
        from types import SimpleNamespace
        from lynx_ui_jobs import review_video_dates
        owner = SimpleNamespace(root=self.root, lang='es')
        rows = [{'path': '/tmp/video.mp4', 'date': '', 'source': 'No metadata'}]
        review_video_dates(owner, rows, '/tmp', '/tmp/output', 1)
        window = next(w for w in self.root.winfo_children() if isinstance(w, self.ctk.CTkToplevel))
        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)
        widgets = list(descendants(window))
        entry = next(w for w in widgets if isinstance(w, self.ctk.CTkEntry))
        check = next(w for w in widgets if isinstance(w, self.ctk.CTkCheckBox))
        button = next(w for w in widgets if isinstance(w, self.ctk.CTkButton) and w.cget('text') == 'Confirmar y extraer')
        with patch('lynx_ui_jobs.messagebox.showerror') as error, patch('lynx_ui_jobs.start') as start:
            button.invoke()
            error.assert_called_once()
            start.assert_not_called()
            entry.insert(0, '2024-02-03 12:13:14')
            check.select()
            button.invoke()
            start.assert_called_once()

    def test_window_ticks_and_cancels_without_worker_touching_tk(self):
        entered = Event()
        ticks = []
        threads = []
        owner = get_ident()
        def work(task):
            threads.append(get_ident())
            entered.set()
            task.stop.wait(3)
            task.checkpoint()
        self.root.jobs.start('Video', work)
        self.root.after(10, lambda: ticks.append(get_ident()))
        self.wait_until(lambda: bool(ticks) and entered.is_set())
        self.assertEqual(ticks, [owner])
        self.assertNotEqual(threads, [owner])
        self.assertEqual(self.button.cget('state'), 'disabled')
        self.root.jobs.cancel_button.invoke()
        self.wait_until(lambda: not self.root.jobs.busy)
        self.assertEqual(self.button.cget('state'), 'normal')
        self.assertIn('Cancelado', self.root.jobs.status.cget('text'))
        received = []
        self.root.jobs.start('Excel', lambda task: 123, lambda result: received.append((get_ident(), result)))
        self.wait_until(lambda: bool(received))
        self.assertEqual(received, [(owner, 123)])

    def test_errors_restore_controls_on_main_thread(self):
        received = []
        def fail(task):
            raise ValueError('bad input')
        with patch('lynx_ui_jobs.messagebox.showerror', side_effect=lambda *args: received.append(get_ident())):
            self.root.jobs.start('Excel', fail)
            self.wait_until(lambda: not self.root.jobs.busy)
        self.assertEqual(received, [get_ident()])
        self.assertEqual(self.button.cget('state'), 'normal')

    def test_full_and_mini_inputs_start_async_jobs(self):
        # Pillow/Tk caches can retain the interpreter from a previous CTk root.
        # Exercise application startup in a fresh process, as the smoke test does.
        if os.environ.get('LYNX_ISOLATED_APP_TEST') != '1':
            import subprocess
            import sys
            repo = Path(__file__).resolve().parents[1]
            env = dict(os.environ, LYNX_ISOLATED_APP_TEST='1', PYTHONPATH=str(repo))
            subprocess.run([sys.executable, '-m', 'unittest',
                            'test_gui_jobs.GuiJobsTests.test_full_and_mini_inputs_start_async_jobs'],
                           cwd=repo / 'tests', env=env, check=True, timeout=60)
            return
        from lynx_processing import WIProcessor
        from lynx_ui_jobs import TEXT
        import pandas as pd
        repo = Path(__file__).resolve().parents[1]
        for filename in ('LynxAutomator_v001alpha.py', 'LynxAutomator_v001alpha mini.py'):
            with self.subTest(filename=filename):
                spec = importlib.util.spec_from_file_location('desktop', repo / filename)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                for widget in self.root.winfo_children():
                    widget.destroy()
                desktop = self.root
                app = module.BaseApp(desktop)
                desktop.update()
                try:
                    combiner = app.excel_combiner_app
                    combiner.initial_excel_path = 'template.xlsx'
                    combiner.images_csv_path = 'images.csv'
                    combiner.deployments_csv_path = 'deployments.csv'
                    result = pd.DataFrame({'test': [1]})
                    with patch.object(WIProcessor, 'run', return_value=result):
                        combiner.process_files()
                        deadline = time.monotonic() + 5
                        while desktop.jobs.busy and time.monotonic() < deadline:
                            desktop.update()
                            time.sleep(0.005)
                    self.assertFalse(desktop.jobs.busy)
                    self.assertIs(combiner._result, result)
                    self.assertEqual(combiner.download_btn.cget('state'), 'normal')
                    combiner.multiple_images_var.set(True)
                    combiner.time_threshold_entry.delete(0, 'end')
                    combiner.time_threshold_entry.insert(0, '-1')
                    with patch('lynx_ui_jobs.messagebox.showerror') as error:
                        combiner.process_files()
                        error.assert_called_once()
                    self.assertIsNone(combiner._result)
                    self.assertEqual(combiner.download_btn.cget('state'), 'disabled')

                    app.language.set('Português')
                    app.change_language()
                    desktop.update()
                    self.assertEqual(desktop.jobs.text, TEXT['pt'])
                finally:
                    if desktop.jobs.busy:
                        desktop.jobs.cancel()
                        self.wait_until(lambda: not desktop.jobs.busy)


    def test_close_waits_for_current_write_without_blocking_events(self):
        entered, finish_write, closed = Event(), Event(), Event()
        def work(task):
            entered.set()
            finish_write.wait(3)
            task.checkpoint()
        self.root.jobs.start('Writing', work)
        self.wait_until(entered.is_set)
        with patch.object(self.root, 'destroy', side_effect=closed.set):
            self.root.jobs.close()
            self.assertFalse(closed.is_set())
            self.assertTrue(self.root.jobs.busy)
            self.root.after(10, finish_write.set)
            self.wait_until(closed.is_set)
        self.assertFalse(self.root.jobs.busy)
