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
        self.assertEqual(flow.choice.get(), 'Wildlife Insights')
        self.assertIn('Crear desde carpeta', flow.sources)
        self.assertIn('Agouti API (alpha)', flow.sources)
        self.assertIn('Trapper API (alpha)', flow.sources)
        for source in flow.sources:
            flow.choice.set(source)
            flow.select(source)
            self.root.update()
            page = flow.pages[source]
            self.assertIs(page.winfo_toplevel(), self.root)
        flow.edit(lambda task, options: ([], []))
        self.assertIsInstance(flow.editor, BulkEditor)
        self.assertIs(flow.editor.winfo_toplevel(), self.root)
        flow.select(flow.choice.get())
        self.assertIsNone(flow.editor)
        self.assertTrue(flow.pages[flow.choice.get()].winfo_manager())

    def test_api_filter_dialog_applies_without_download(self):
        from lynx_api_ui import APIImportTab
        page = APIImportTab(self.root, 'Agouti API', 'es')
        self.assertEqual(page.auth_button.cget('text'), 'Autorización')
        self.assertIs(page.auth_button.master, page.source_controls)
        page.filter_dialog()
        window = page._filters_window
        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)
        entries = [w for w in descendants(window) if isinstance(w, self.ctk.CTkEntry)]
        entries[0].insert(0, '2025')
        entries[1].insert(0, 'Doñana')
        button = next(w for w in descendants(window) if isinstance(w, self.ctk.CTkButton) and w.cget('text') == 'Aplicar filtros')
        with patch('lynx_api_ui.fetch_api_package') as request:
            button.invoke()
            request.assert_not_called()
        self.assertEqual(page.filters['year'], '2025')
        self.assertEqual(page.filters['site'], 'Doñana')
        self.assertFalse(window.winfo_exists())
        self.assertIn('Año de inicio: 2025', page.filter_summary.cget('text'))

    def test_metadata_sources_open_before_valid_preview(self):
        from types import SimpleNamespace
        from lynx_bulk_ui import BulkEditor
        row = {'metadata': {'cameras.camera_model': ['Model X'], 'unmatched.name': []}}
        editor = BulkEditor(SimpleNamespace(root=self.root, lang='es'), lambda task, options: ([row], []))
        self.assertFalse(next(f for f in editor.collect() if f['name'] == 'MarkedIndividual.individualID')['enabled'])
        source = editor.rows[0][2]
        source._open_dropdown_menu()
        self.wait_until(lambda: not self.root.jobs.busy)
        self.assertIn('cameras.camera_model', source.cget('values'))
        self.assertNotIn('unmatched.name', source.cget('values'))
        self.assertIsNone(editor.snapshot)
        self.assertTrue(source._choice_popup.winfo_exists())
        source._choice_popup.destroy()
        editor.destroy()

    def test_api_source_loads_without_forcing_authorization(self):
        from lynx_bulk_workflow import BulkWorkflow
        from lynx_camtrap import read_package
        from lynx_tasks import TaskContext
        from camtrap_factory import make_package
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            package = read_package(TaskContext(), make_package(folder) / 'datapackage.json')
            flow = BulkWorkflow(self.root, 'es')
            for provider in ('Agouti API', 'Trapper API'):
                flow.select(provider + ' (alpha)')
                page = flow.pages[provider + ' (alpha)']
                if provider == 'Trapper API':
                    page.server.insert(0, 'https://trapper.example')
                page.project.insert(0, '123')
                self.assertIsNone(page.download_auth)
                with patch('lynx_api_ui.dialogs.askdirectory', return_value=folder), patch('lynx_api_ui.fetch_api_package', return_value=package) as fetch:
                    page.load()
                    self.wait_until(lambda: not self.root.jobs.busy)
                self.assertIsNone(fetch.call_args.args[5])
                self.assertIn('Lynx pardinus', page.checks)
                page.checks['Lynx pardinus'].select()
                page.review()
                self.wait_until(lambda: not self.root.jobs.busy)
                self.assertTrue(page.records)
                self.assertEqual(page.download.cget('state'), 'normal')
                page.access()
                self.assertEqual(page._access_dialog.method.get(), 'Agouti API key' if provider == 'Agouti API' else 'Trapper token')
                page._access_dialog.close()

    def test_download_access_and_mode_selection_share_controls(self):
        from lynx_bulk_workflow import BulkWorkflow
        from lynx_download_auth_ui import DownloadAccess
        flow = BulkWorkflow(self.root, 'es')
        flow.select('Camtrap DP')
        page = flow.pages['Camtrap DP']
        page.records = [{'prepared': True}]
        page.batches = ['previous']
        page.local.set(1)
        page.acquisition_changed()
        self.assertEqual(page.records, [{'prepared': True}])
        self.assertEqual(page.batches, [])
        page.access()
        dialog = page._access_dialog
        self.assertIsInstance(dialog, DownloadAccess)
        dialog.method.set('Trapper token')
        dialog.server.delete(0, 'end')
        dialog.server.insert(0, 'https://trapper.example.org')
        dialog.secret.insert(0, 'test-token')
        self.assertTrue(dialog.secret.cget('show'))
        dialog.apply()
        self.assertEqual(page.download_auth.headers('https://trapper.example.org/image'),
                         {'Authorization': 'Token test-token'})
        self.assertEqual(dialog.secret.get(), '')
        dialog.clear()
        self.assertIsNone(page.download_auth)
        dialog.close()
        flow.select('Wildlife Insights')
        wi = flow.pages['Wildlife Insights']
        wi.access()
        google = wi._access_dialog
        self.assertTrue(google.google)
        with patch('lynx_download_auth_ui.google_login', return_value=True) as login:
            google.login_google()
            self.wait_until(lambda: not self.root.jobs.busy)
        login.assert_called_once()
        google.close()

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
        tab.local.set(1)
        with tempfile.TemporaryDirectory() as folder, patch('lynx_camtrap_ui.filedialog.askdirectory', side_effect=[str(fixture.parent), folder]), patch('lynx_ui_jobs.messagebox.showwarning') as warning:
            tab.obtain()
            self.wait_until(lambda: not self.root.jobs.busy)
            self.assertEqual(len(list(Path(folder).rglob('*.jpg'))), 2)
            warning.assert_called_once()
            self.assertEqual(len(tab.failed), 1)
        tab.select_all(False)
        self.assertIsNone(tab.records)
        self.assertEqual(tab.download.cget('state'), 'disabled')

    def test_editor_layout_and_localized_profile_roundtrip(self):
        from types import SimpleNamespace
        from lynx_bulk_ui import BulkEditor
        from lynx_bulk import default_fields
        for lang in ('es', 'pt', 'en'):
            editor = BulkEditor(SimpleNamespace(root=self.root, lang=lang), lambda task, options: ([], []))
            self.root.update()
            self.assertEqual(editor.collect(), default_fields())
            self.assertIs(editor.group.master, editor.interval.master)
            self.assertEqual(editor.group.pack_info()['side'], 'left')
            self.assertEqual(editor.interval.pack_info()['side'], 'left')
            index = next(i for i, field in enumerate(editor.collect()) if field['name'] == 'Encounter.locationID')
            button = editor.location_buttons[index]
            self.assertIs(button.master, editor.rows[index][3].master)
            self.assertEqual(button.winfo_manager(), 'grid')
            self.assertEqual(sum(bool(button.winfo_manager()) for button in editor.location_buttons), 1)
            editor.rows[index][1]._dropdown_callback('Encounter.country')
            self.root.update()
            self.assertFalse(button.winfo_manager())
            editor.destroy()

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
                self.assertTrue(editor.preview_window.winfo_exists())
                profile.assert_not_called()
                destination = Path(folder) / 'bulk.xlsx'
                with patch('lynx_bulk_ui.dialogs.asksaveasfilename', return_value=str(destination)) as save_dialog, patch('lynx_bulk_ui.messagebox.showinfo'):
                    editor.save()
                    self.wait_until(lambda: not self.root.jobs.busy)
                self.assertTrue(destination.is_file())
                self.assertRegex(save_dialog.call_args.kwargs['initialfile'], r'wildbook_bulk_import_.*_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsx')
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
            page.local.set(0)
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
            editor.add()
            chooser = next(w for w in editor.winfo_children() if isinstance(w, self.ctk.CTkToplevel))
            next(w for w in chooser.winfo_children() if isinstance(w, self.ctk.CTkButton) and w.cget('text') == 'Comentarios con metadatos').invoke()
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
        self.assertEqual(field['location_hierarchy']['Exact-á'], ['parent', 'Exact-á'])
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
        footer = next(w for w in picker.winfo_children() if isinstance(w, self.ctk.CTkFrame) and any(isinstance(child, self.ctk.CTkButton) and child.cget('text') == 'Cerrar' for child in w.winfo_children()))
        next(child for child in footer.winfo_children() if isinstance(child, self.ctk.CTkButton) and child.cget('text') == 'Cerrar').invoke()
        self.assertFalse(picker.winfo_exists())
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
            page.local.set(1)
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
            flow.select(flow.choice.get())
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
        self.assertEqual(len(app.main_tabs.tabs()), 3)
        self.assertEqual(app.main_tabs.tab(0, 'text'), 'Bulk Import')
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
