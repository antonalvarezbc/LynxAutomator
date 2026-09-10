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

    def test_unified_bulk_sources_and_embedded_editor(self):
        from lynx_bulk_workflow import BulkWorkflow
        from lynx_bulk_ui import BulkEditor
        flow = BulkWorkflow(self.root, 'es')
        for source in flow.sources:
            flow.choice.set(source)
            flow.select(source)
            self.root.update()
            page = flow.pages[source]
            self.assertIs(page.winfo_toplevel(), self.root)
        flow.edit(lambda task, options: ([], []))
        self.assertIsInstance(flow.editor, BulkEditor)
        self.assertIs(flow.editor.winfo_toplevel(), self.root)
        flow.back()
        self.assertIsNone(flow.editor)
        self.assertTrue(flow.pages[flow.choice.get()].winfo_manager())

    def test_camtrap_multiselect_preview_and_local_download(self):
        from lynx_camtrap_ui import CamtrapTab
        from lynx_camtrap import read_package
        from lynx_tasks import TaskContext
        import tempfile
        tab = CamtrapTab(self.root, lang='es')
        from camtrap_factory import make_package
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        fixture = make_package(temporary.name) / 'datapackage.json'
        tab.receive(read_package(TaskContext(), fixture))
        tab.checks['Lynx pardinus'].select()
        tab.review()
        self.wait_until(lambda: not self.root.jobs.busy)
        self.assertEqual(len(tab.records), 3)
        self.assertEqual(tab.download.cget('state'), 'normal')
        tab.local.select()
        with tempfile.TemporaryDirectory() as folder, patch('lynx_camtrap_ui.filedialog.askdirectory', return_value=folder):
            tab.obtain()
            self.wait_until(lambda: not self.root.jobs.busy)
            self.assertEqual(len(list(Path(folder).rglob('*.jpg'))), 2)
        tab.select_all(False)
        self.assertIsNone(tab.records)
        self.assertEqual(tab.download.cget('state'), 'disabled')

    def test_bulk_editor_preview_save_and_invalidation(self):
        import tempfile
        from types import SimpleNamespace
        from lynx_bulk_ui import BulkEditor
        from lynx_bulk import default_fields, normalize
        with tempfile.TemporaryDirectory() as folder:
            photo = Path(folder) / 'photo.png'
            photo.write_bytes(b'test')
            row = normalize('Lynx pardinus', '2024-01-01', photo, 'd', 'e', 'p', locality='test')
            with patch('lynx_bulk_ui.load_profile', return_value=default_fields()), patch('lynx_bulk_ui.save_profile') as profile:
                editor = BulkEditor(SimpleNamespace(root=self.root, lang='pt'), lambda task, group: ([row], []))
                editor.preview()
                self.wait_until(lambda: not self.root.jobs.busy)
                self.assertIsNotNone(editor.snapshot)
                self.assertEqual(len(editor.tree.get_children()), 1)
                profile.assert_not_called()
                destination = Path(folder) / 'bulk.xlsx'
                with patch('lynx_bulk_ui.dialogs.asksaveasfilename', return_value=str(destination)), patch('lynx_bulk_ui.messagebox.showinfo'):
                    editor.save()
                    self.wait_until(lambda: not self.root.jobs.busy)
                self.assertTrue(destination.is_file())
                editor.interval.delete(0, 'end')
                editor.interval.insert(0, '10')
                with patch('lynx_ui_jobs.messagebox.showerror') as error:
                    editor.save()
                    error.assert_called_once()
                editor.destroy()

    def test_wi_zip_loads_species_then_requires_acquisition(self):
        from types import SimpleNamespace
        from lynx_bulk_ui import WIInput
        from wi_factory import make_wi_zip
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            archive = make_wi_zip(folder)
            with patch('lynx_wi_ui.shutil.which', return_value=None):
                window = WIInput(SimpleNamespace(root=self.root, lang='es'))
            page = window.page
            self.assertTrue(page.local.get())
            with patch('lynx_wi_ui.filedialog.askopenfilename', return_value=str(archive)) as dialog:
                page.load()
                self.wait_until(lambda: not self.root.jobs.busy)
            self.assertEqual(dialog.call_args.kwargs['filetypes'], [('Wildlife Insights ZIP', '*.zip')])
            self.assertEqual(set(page.checks), {'Lynx pardinus', 'Vulpes vulpes'})
            self.assertEqual(page.export.cget('state'), 'disabled')
            with patch('lynx_ui_jobs.messagebox.showerror') as error, patch('lynx_bulk_ui.BulkEditor') as editor:
                page.open_bulk_import()
                error.assert_called_once()
                editor.assert_not_called()
            page.checks['Lynx pardinus'].select()
            page.review()
            self.wait_until(lambda: not self.root.jobs.busy)
            page.local.deselect()
            with patch('lynx_wi_ui.shutil.which', return_value=None), patch('lynx_ui_jobs.messagebox.showerror') as error, patch('lynx_wi_ui.filedialog.askdirectory') as folder:
                page.obtain()
                error.assert_called_once()
                folder.assert_not_called()

    def test_bulk_profiles_only_save_explicitly(self):
        import tempfile
        from types import SimpleNamespace
        from lynx_bulk_ui import BulkEditor
        from lynx_bulk import read_profiles, save_profile, load_profile
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'profiles.json'
            with patch('lynx_bulk_ui.read_profiles', side_effect=lambda: read_profiles(path)), patch(
                    'lynx_bulk_ui.save_profile', side_effect=lambda fields, name: save_profile(fields, path, name)), patch(
                    'lynx_bulk_ui.load_profile', side_effect=lambda name: load_profile(path, name)), patch('lynx_bulk_ui.messagebox.showinfo'):
                editor = BulkEditor(SimpleNamespace(root=self.root, lang='es'), lambda task, group: ([], []))
                self.assertFalse(path.exists())
                editor.profile_name.insert(0, 'Doñana')
                editor.rows[-1][3].insert(0, 'submitter')
                editor.store_profile()
                editor.rows[-1][3].delete(0, 'end')
                editor.use_profile()
                self.assertEqual(editor.rows[-1][3].get(), 'submitter')
                self.assertEqual(list(read_profiles(path)), ['Doñana'])
                editor.destroy()

    def test_metadata_comments_picker_and_catalog(self):
        from types import SimpleNamespace
        from lynx_bulk_ui import BulkEditor
        from lynx_wildbook_fields import FIELDS
        from lynx_bulk import normalize, attach_metadata, build_table
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            photo = Path(folder) / 'photo.jpg'
            photo.write_bytes(b'test')
            row = normalize('Lynx pardinus', '2024-01-01', photo, 'd', 'e', 'p', locality='test')
            attach_metadata(row, deployment=[{'setupBy': 'Alice', 'cameraID': 'C1', 'cameraModel': 'Model X'}])
            editor = BulkEditor(SimpleNamespace(root=self.root, lang='es'), lambda task, group: ([row], []))
            self.assertIn('Sighting.comments', editor.rows[0][1].cget('values'))
            self.assertIn('Encounter.project0.researchProjectName', FIELDS)
            editor.render([f for f in editor.collect() if f['name'] != 'Encounter.sightingID'])
            editor.metadata_comments()
            self.wait_until(lambda: not self.root.jobs.busy)
            picker = next(w for w in editor.winfo_children() if isinstance(w, self.ctk.CTkToplevel))
            apply = next(w for w in picker.winfo_children() if isinstance(w, self.ctk.CTkButton))
            apply.invoke()
            table = build_table([row], editor.collect())
            self.assertIn('Encounter.sightingID', table[0])
            self.assertIn('deployment.cameraID: C1', table[0]['Sighting.comments'])
            self.assertIn('deployment.setupBy: Alice', table[0]['Sighting.comments'])
            editor.destroy()

    def test_location_picker_applies_exact_ids_and_keeps_profile_metadata(self):
        from types import SimpleNamespace
        from lynx_bulk_ui import BulkEditor
        from lynx_locations_ui import LocationPicker
        from lynx_locations import CATALOGS, deployment_key, encounter_key
        from lynx_bulk import attach_metadata
        row = {}
        attach_metadata(row, deployment=[{'deploymentID': 'd1'}])
        editor = BulkEditor(SimpleNamespace(root=self.root, lang='es'), lambda task, group: ([row], []))
        other = dict(eventID='other', media=['other.jpg'])
        attach_metadata(other, deployment=[{'deploymentID': 'd2'}])
        picker = LocationPicker(editor, [row, other])
        self.assertEqual(str(picker.transient()), str(editor.winfo_toplevel()))
        self.assertFalse(picker.url.winfo_manager())
        picker.receive({'source': CATALOGS['Lynx'], 'updated': 'test', 'catalog': {'locationID': [{'id': 'parent', 'name': 'Region', 'locationID': [{'id': 'Exact-á', 'name': 'Site'}]}]}})
        picker.tree.selection_set('1')
        picker.apply()
        field = next(f for f in editor.collect() if f['name'] == 'Encounter.locationID')
        self.assertEqual(field['value'], 'Exact-á')
        self.assertEqual(field['catalog_name'], 'Lynx')
        picker.scope_mode.set('Deployment')
        picker.populate_scopes()
        picker.scope.selection_set(next(iid for iid, keys in picker.scope_items.items() if deployment_key(row) in keys))
        picker.tree.selection_set('0')
        picker.apply()
        field = next(f for f in editor.collect() if f['name'] == 'Encounter.locationID')
        self.assertEqual(field['locations'][deployment_key(row)], 'parent')
        self.assertEqual(field['value'], 'Exact-á')
        picker.scope_mode.set('Encounter')
        picker.populate_scopes()
        picker.scope.selection_set('s0', 's1')
        picker.tree.selection_set('1')
        picker.apply()
        field = next(f for f in editor.collect() if f['name'] == 'Encounter.locationID')
        self.assertEqual(field['locations'][encounter_key(row)], 'Exact-á')
        self.assertEqual(field['locations'][encounter_key(other)], 'Exact-á')
        picker.destroy()
        editor.destroy()

    def test_choices_stay_open_and_accept_selection(self):
        from lynx_choices import StableComboBox
        from lynx_windows import foreground
        dialog = self.ctk.CTkToplevel(self.root)
        foreground(dialog, self.root, modal=True)
        choice = StableComboBox(dialog, values=['Lynx', 'Deer'])
        choice.pack()
        choice._open_dropdown_menu()
        popup = choice._choice_popup
        deadline = time.monotonic() + 1.2
        while time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertTrue(popup.winfo_exists())
        popup.list.selection_set('1')
        popup.choose()
        self.assertEqual(choice.get(), 'Deer')
        dialog.destroy()

    def test_wi_shared_flow_local_photos_and_selection_invalidation(self):
        import tempfile
        from PIL import Image
        from wi_factory import make_wi_zip
        from lynx_bulk_workflow import BulkWorkflow
        from lynx_tasks import TaskContext
        with tempfile.TemporaryDirectory() as folder:
            archive = make_wi_zip(folder)
            Image.new('RGB', (4, 4)).save(Path(folder) / 'lynx.jpg')
            flow = BulkWorkflow(self.root, 'es')
            flow.choice.set('Wildlife Insights')
            flow.select('Wildlife Insights')
            page = flow.pages['Wildlife Insights']
            page.local.select()
            with patch('lynx_wi_ui.filedialog.askopenfilename', return_value=str(archive)):
                page.load()
                self.wait_until(lambda: not self.root.jobs.busy)
            page.checks['Lynx pardinus'].select()
            page.review()
            self.wait_until(lambda: not self.root.jobs.busy)
            self.assertEqual(len(page.records), 1)
            with patch('lynx_wi_ui.filedialog.askdirectory', return_value=folder):
                page.obtain()
                self.wait_until(lambda: not self.root.jobs.busy)
            self.assertEqual(page.export.cget('state'), 'normal')
            page.open_bulk_import()
            self.assertIs(flow.editor.winfo_toplevel(), self.root)
            rows, missing = flow.editor.loader(TaskContext(), (False, 3))
            self.assertEqual(len(rows), 1)
            self.assertFalse(missing)
            self.assertTrue(rows[0]['media_local'])
            flow.back()
            page.select_all(False)
            self.assertEqual(page.available, {})
            self.assertEqual(page.export.cget('state'), 'disabled')

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
        with patch('lynx_ui_jobs.messagebox.showerror', side_effect=lambda *args, **kwargs: received.append(get_ident())):
            self.root.jobs.start('Excel', fail)
            self.wait_until(lambda: not self.root.jobs.busy)
        self.assertEqual(received, [get_ident()])
        self.assertEqual(self.button.cget('state'), 'normal')

    def test_application_inputs_start_async_jobs(self):
        # Pillow/Tk caches can retain the interpreter from a previous CTk root.
        # Exercise application startup in a fresh process, as the smoke test does.
        if os.environ.get('LYNX_ISOLATED_APP_TEST') != '1':
            import subprocess
            import sys
            repo = Path(__file__).resolve().parents[1]
            env = dict(os.environ, LYNX_ISOLATED_APP_TEST='1', PYTHONPATH=str(repo))
            subprocess.run([sys.executable, '-m', 'unittest',
                            'test_gui_jobs.GuiJobsTests.test_application_inputs_start_async_jobs'],
                           cwd=repo / 'tests', env=env, check=True, timeout=60)
            return
        from lynx_processing import WIProcessor
        from lynx_ui_jobs import TEXT
        import pandas as pd
        repo = Path(__file__).resolve().parents[1]
        filename = 'LynxAutomator_v001alpha.py'
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
