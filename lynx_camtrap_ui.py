"""Camtrap DP adapter for the shared media import interface."""
import lynx_dialogs as filedialog
from lynx_camtrap import read_package, select_media, acquire_media
from lynx_ui_jobs import action, start
from lynx_media_ui import MediaImportTab


class CamtrapTab(MediaImportTab):
    def __init__(self, root, lang='es'):
        super().__init__(root, lang)


    @action
    def open_bulk_import(self):
        from lynx_bulk_ui import open_camtrap
        open_camtrap(self)

    @action
    def load(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[('Camtrap DP ZIP', '*.zip')])
        if path:
            if not path.lower().endswith('.zip'):
                raise ValueError('Selecciona el ZIP completo de Camtrap DP.')
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
        detail = f'{len(self.records)-remote} {self.text[12]} | {remote} {self.text[13]} | {private} {self.text[14]}'
        if self.events.get():
            detail += '\n' + self.text[16]
        self.selection_summary(len(self.records),
            [r['media'].get('fileName', r['media']['mediaID']) for r in self.records], issues, detail)

    @action
    def obtain(self):
        if not self.records:
            return
        local = bool(self.local.get())
        local_folder = None
        if local:
            local_folder = filedialog.askdirectory(parent=self, title={'es': 'Carpeta con las fotografías originales', 'pt': 'Pasta com fotografias originais', 'en': 'Folder containing original photographs'}[self.lang])
            if not local_folder:
                return
        destination = filedialog.askdirectory(parent=self, title={'es': 'Destino de las fotografías preparadas', 'pt': 'Destino das fotografias preparadas', 'en': 'Destination for prepared photographs'}[self.lang])
        if destination:
            self.local_folder = local_folder
            package, records, auth = self.package, self.records, self.download_auth
            start(self, 'Camtrap DP', lambda task: acquire_media(task, package, records, destination,
                  local_only=local, local_folder=local_folder, auth=auth, allow_private=not local), self.receive_download)

    def receive_download(self, result):
        self.batches.append(result.output_directory)
        if result.completed:
            self.export.configure(state='normal')
        self.failed = result.failed_records
        self.retry.configure(state='normal' if self.failed else 'disabled')
        self.show(result.output_directory + '\nmanifest.csv\n' +
                  ({'es': '{} preparadas; {} omitidas; {} fallidas', 'pt': '{} preparadas; {} omitidas; {} falhadas', 'en': '{} prepared; {} skipped; {} failed'}[self.lang].format(result.completed, result.skipped, len(result.errors))) + '\n' + '\n'.join(result.errors[:12]))

    @action
    def retry_failed(self):
        if self.failed:
            destination = filedialog.askdirectory(parent=self)
            if destination:
                package, records = self.package, list(self.failed)
                local_folder, auth = getattr(self, 'local_folder', None), self.download_auth
                start(self, 'Camtrap DP', lambda task: acquire_media(task, package, records, destination,
                      local_only=bool(local_folder), local_folder=local_folder, auth=auth, allow_private=not local_folder), self.receive_download)
