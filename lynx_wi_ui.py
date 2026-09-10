"""Wildlife Insights ZIP adapter for the common photo import interface."""
from pathlib import Path
import shutil
import lynx_dialogs as filedialog
from lynx_bulk import wi_species, wi_rows
from lynx_media_ui import MediaImportTab, LABELS
from lynx_ui_jobs import action, start
from lynx_wi_media import acquire_wi, acquired_rows, sources_for


class WITab(MediaImportTab):
    def __init__(self, root, lang='es'):
        labels = list(LABELS[lang])
        labels[0] = {'es': 'Cargar ZIP de Wildlife Insights', 'pt': 'Carregar ZIP do Wildlife Insights', 'en': 'Load Wildlife Insights ZIP'}[lang]
        labels[5] = {'es': 'Usar fotografías de una carpeta local', 'pt': 'Usar fotografias de uma pasta local', 'en': 'Use photographs from a local folder'}[lang]
        super().__init__(root, lang, labels=labels, events=False)
        self.available = {}
        self.directory = ''
        self.executable = None
        self.gsutil = shutil.which('gsutil')
        if not self.gsutil:
            self.local.select()
        self.hint = {'es': 'ZIP → seleccionar especies → revisar → obtener fotografías → configurar Excel. ',
                     'pt': 'ZIP → selecionar espécies → rever → obter fotografias → configurar Excel. ',
                     'en': 'ZIP → select species → review → obtain photographs → configure Excel. '}[lang]
        self.hint += ({'es': 'gsutil disponible para descargar.', 'pt': 'gsutil disponível para descarregar.', 'en': 'gsutil available for downloads.'}[lang] if self.gsutil else
                      {'es': 'gsutil no está instalado: selecciona una carpeta local.', 'pt': 'gsutil não está instalado: selecione uma pasta local.', 'en': 'gsutil is not installed: select a local folder.'}[lang])
        self.heading.configure(text=self.hint)

    def invalidate(self):
        super().invalidate()
        self.available = {}

    @action
    def load(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[('Wildlife Insights ZIP', '*.zip')])
        if path:
            self.reset_source()
            start(self, 'Wildlife Insights', lambda task: wi_species(task, path),
                  lambda counts: self.receive(path, counts))

    def receive(self, path, counts):
        self.package = path
        self.receive_species(counts, Path(path).name + '\n' + self.hint)

    @action
    def review(self):
        species = {name for name, check in self.checks.items() if check.get()}
        if not self.package or not species:
            raise ValueError(self.text[8])
        self.invalidate()
        path = self.package
        start(self, 'Wildlife Insights', lambda task: wi_rows(task, path, species=species), self.receive_selection)

    @action
    def receive_selection(self, result):
        rows, issues = result
        sources = sources_for(rows)
        self.records = rows
        self.show(f'{len(sources)} {self.text[11]}\n' + '\n'.join(issues[:8]) + '\n' +
                  '\n'.join(sources[:30]))
        self.download.configure(state='normal' if sources else 'disabled')

    @action
    def obtain(self):
        if not self.records:
            return
        executable = None if self.local.get() else shutil.which('gsutil')
        if not self.local.get() and not executable:
            raise ValueError({'es': 'gsutil no está instalado. Selecciona la opción de fotografías locales.',
                              'pt': 'gsutil não está instalado. Selecione fotografias locais.',
                              'en': 'gsutil is not installed. Select the local photographs option.'}[self.lang])
        directory = filedialog.askdirectory(parent=self)
        if directory:
            self.directory, self.executable = directory, executable
            sources = sources_for(self.records)
            start(self, 'Wildlife Insights', lambda task: acquire_wi(task, sources, directory, executable), self.receive_download)

    def receive_download(self, result):
        self.available.update(result.available)
        self.failed = result.failed
        self.retry.configure(state='normal' if self.failed else 'disabled')
        self.export.configure(state='normal' if self.available else 'disabled')
        self.show(f'{len(self.available)} OK; {len(self.failed)} ' +
                  {'es': 'pendientes', 'pt': 'pendentes', 'en': 'pending'}[self.lang] + '\n' +
                  result.manifest + '\n' + '\n'.join(result.errors[:12]))

    @action
    def retry_failed(self):
        if self.failed:
            sources, directory, executable = list(self.failed), self.directory, self.executable
            start(self, 'Wildlife Insights', lambda task: acquire_wi(task, sources, directory, executable), self.receive_download)

    @action
    def open_bulk_import(self):
        if not self.records or not self.available:
            raise ValueError({'es': 'Obtén primero las fotografías de la selección actual.',
                              'pt': 'Obtenha primeiro as fotografias selecionadas.',
                              'en': 'Obtain photographs for the current selection first.'}[self.lang])
        rows, available = list(self.records), dict(self.available)
        loader = lambda task, options: acquired_rows(task, rows, available, options[0], options[1])
        if getattr(self, 'workflow', None):
            self.workflow.edit(loader)
        else:
            from lynx_bulk_ui import BulkEditor
            BulkEditor(self, loader)
