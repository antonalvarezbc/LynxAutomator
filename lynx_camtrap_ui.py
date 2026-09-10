"""Camtrap DP adapter for the shared media import interface."""
import lynx_dialogs as filedialog
from lynx_camtrap import read_package, select_media, acquire_media
from lynx_ui_jobs import action, start
from lynx_media_ui import MediaImportTab


class CamtrapTab(MediaImportTab):
    @action
    def open_bulk_import(self):
        from lynx_bulk_ui import open_camtrap
        open_camtrap(self)

    @action
    def load(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[('Camtrap DP', ('*.json', '*.zip'))])
        if path:
            self.reset_source()
            start(self, 'Camtrap DP', lambda task: read_package(task, path), self.receive)

    def receive(self, package):
        self.batches = []
        self.package = package
        title = str(package.descriptor.get('title', package.path.name))
        if 'synthetic' in str(package.descriptor.get('keywords', [])).lower() or 'synthetic' in title.lower():
            title += '\n' + self.text[15]
        self.receive_species(package.species(), title)

    @action
    def review(self):
        species = {name for name, check in self.checks.items() if check.get()}
        if not self.package or not species:
            raise ValueError(self.text[8])
        self.invalidate()
        package, events = self.package, bool(self.events.get())
        start(self, 'Camtrap DP', lambda task: select_media(task, package, species, events), self.receive_selection)

    def receive_selection(self, result):
        self.records, issues = result
        remote = sum(r['media']['filePath'].startswith(('https://', 'http://')) for r in self.records)
        private = sum(r['media'].get('filePublic') in ('false', False) for r in self.records)
        summary = f'{len(self.records)} {self.text[11]} | {len(self.records)-remote} {self.text[12]} | {remote} {self.text[13]} | {private} {self.text[14]}'
        self.show(summary + '\n' + self.text[16] + '\n' + '\n'.join(issues[:8]) + '\n' +
                  '\n'.join(r['media'].get('fileName', r['media']['mediaID']) for r in self.records[:30]))
        self.download.configure(state='normal' if self.records else 'disabled')

    @action
    def obtain(self):
        if not self.records:
            return
        destination = filedialog.askdirectory(parent=self)
        if destination:
            package, records, local = self.package, self.records, bool(self.local.get())
            start(self, 'Camtrap DP', lambda task: acquire_media(task, package, records, destination, local), self.receive_download)

    def receive_download(self, result):
        self.batches.append(result.output_directory)
        if result.completed:
            self.export.configure(state='normal')
        self.failed = result.failed_records
        self.retry.configure(state='normal' if self.failed else 'disabled')
        self.show(result.output_directory + '\nmanifest.csv\n' +
                  f'{result.completed} OK; {result.skipped} skipped; {len(result.errors)} errors')

    @action
    def retry_failed(self):
        if self.failed:
            destination = filedialog.askdirectory(parent=self)
            if destination:
                package, records = self.package, list(self.failed)
                start(self, 'Camtrap DP', lambda task: acquire_media(task, package, records, destination), self.receive_download)
