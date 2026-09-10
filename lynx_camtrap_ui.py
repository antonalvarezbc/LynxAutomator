"""Camtrap DP tab: dynamic species checkboxes and photo acquisition."""
import customtkinter as ctk
from tkinter import messagebox
import lynx_dialogs as filedialog
from lynx_camtrap import read_package, select_media, acquire_media
from lynx_ui_jobs import action, start

LABELS = {
    'es': ['Cargar datapackage.json o ZIP', 'Buscar especie', 'Seleccionar visibles', 'Deseleccionar todas',
           'Incluir imágenes asociadas por evento (requieren revisión)', 'Sólo copiar imágenes locales',
           'Revisar selección', 'Obtener fotografías seleccionadas', 'Selecciona al menos una especie.',
           'Paquete', 'observaciones', 'fotografías únicas', 'locales', 'remotas', 'privadas',
           'SINTÉTICO: etiquetas de prueba; las fotos no son de lince.',
           'Los eventos agrupan imágenes por despliegue e intervalo; no confirman la especie en cada foto.'],
    'pt': ['Carregar datapackage.json ou ZIP', 'Pesquisar espécie', 'Selecionar visíveis', 'Desmarcar todas',
           'Incluir imagens associadas por evento (requerem revisão)', 'Copiar apenas imagens locais',
           'Rever seleção', 'Obter fotografias selecionadas', 'Selecione pelo menos uma espécie.',
           'Pacote', 'observações', 'fotografias únicas', 'locais', 'remotas', 'privadas',
           'SINTÉTICO: etiquetas de teste; as fotos não são de lince.',
           'Os eventos associam imagens por instalação e intervalo; não confirmam a espécie em cada foto.'],
    'en': ['Load datapackage.json or ZIP', 'Search species', 'Select visible', 'Clear selection',
           'Include event-associated images (review required)', 'Copy local images only',
           'Review selection', 'Get selected photographs', 'Select at least one species.',
           'Package', 'observations', 'unique photographs', 'local', 'remote', 'private',
           'SYNTHETIC: test labels; the photos are not lynxes.',
           'Events associate images by deployment and interval; they do not confirm the species in each photo.']}


class CamtrapTab(ctk.CTkFrame):
    def __init__(self, root, lang='es'):
        super().__init__(root)
        self.root, self.lang = root, lang
        self.text = LABELS[lang]
        self.package = None
        self.checks = {}
        self.records = None
        self.failed = []
        self.pack(fill='both', expand=True)
        ctk.CTkButton(self, text=self.text[0], command=self.load).pack(padx=10, pady=8, anchor='w')
        self.heading = ctk.CTkLabel(self, text='', wraplength=850, justify='left')
        self.heading.pack(fill='x', padx=10)
        self.search = ctk.CTkEntry(self, placeholder_text=self.text[1])
        self.search.pack(fill='x', padx=10, pady=5)
        self.search.bind('<KeyRelease>', lambda event: self.filter_species())
        tools = ctk.CTkFrame(self)
        tools.pack(fill='x', padx=10)
        ctk.CTkButton(tools, text=self.text[2], command=lambda: self.select_all(True)).pack(side='left', padx=5)
        ctk.CTkButton(tools, text=self.text[3], command=lambda: self.select_all(False)).pack(side='left', padx=5)
        self.list = ctk.CTkScrollableFrame(self, height=150)
        self.list.pack(fill='both', expand=True, padx=10, pady=5)
        self.events = ctk.CTkCheckBox(self, text=self.text[4], command=self.invalidate)
        self.events.pack(anchor='w', padx=15, pady=4)
        self.local = ctk.CTkCheckBox(self, text=self.text[5])
        self.local.pack(anchor='w', padx=15, pady=4)
        self.preview = ctk.CTkTextbox(self, height=120)
        self.preview.pack(fill='x', padx=10)
        self.preview.configure(state='disabled')
        buttons = ctk.CTkFrame(self)
        buttons.pack(fill='x', padx=10, pady=8)
        ctk.CTkButton(buttons, text=self.text[6], command=self.review).pack(side='left', padx=5)
        self.download = ctk.CTkButton(buttons, text=self.text[7], command=self.obtain, state='disabled')
        self.download.pack(side='left', padx=5)
        retry_label = {'es': 'Reintentar fallidas', 'pt': 'Repetir falhadas', 'en': 'Retry failed'}[lang]
        self.retry = ctk.CTkButton(buttons, text=retry_label, command=self.retry_failed, state='disabled')
        self.retry.pack(side='left', padx=5)

    def invalidate(self):
        self.failed = []
        self.retry.configure(state='disabled')
        self.records = None
        self.download.configure(state='disabled')
        self.show('')

    def show(self, text):
        self.preview.configure(state='normal')
        self.preview.delete('1.0', 'end')
        self.preview.insert('1.0', text)
        self.preview.configure(state='disabled')

    @action
    def load(self):
        path = filedialog.askopenfilename(filetypes=[('Camtrap DP', ('*.json', '*.zip'))])
        if path:
            self.invalidate()
            self.package = None
            for check in self.checks.values():
                check.destroy()
            self.checks = {}
            self.heading.configure(text='')
            start(self, 'Camtrap DP', lambda task: read_package(task, path), self.receive)

    def receive(self, package):
        self.package = package
        title = str(package.descriptor.get('title', package.path.name))
        if 'synthetic' in str(package.descriptor.get('keywords', [])).lower() or 'synthetic' in title.lower():
            title += '\n' + self.text[15]
        self.heading.configure(text=title)
        for species, count in sorted(package.species().items()):
            check = ctk.CTkCheckBox(self.list, text=f'{species} — {count} {self.text[10]}', command=self.invalidate)
            self.checks[species] = check
        self.filter_species()

    def filter_species(self):
        query = self.search.get().casefold()
        for species, check in self.checks.items():
            check.pack_forget()
            if query in species.casefold():
                check.pack(anchor='w', padx=8, pady=4)

    def select_all(self, selected):
        for species, check in self.checks.items():
            if not selected:
                check.deselect()
            elif self.search.get().casefold() in species.casefold():
                check.select()
        self.invalidate()

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
        destination = filedialog.askdirectory()
        if destination:
            package, records, local = self.package, self.records, bool(self.local.get())
            start(self, 'Camtrap DP', lambda task: acquire_media(task, package, records, destination, local), self.receive_download)

    def receive_download(self, result):
        self.failed = result.failed_records
        self.retry.configure(state='normal' if self.failed else 'disabled')
        self.show(result.output_directory + '\nmanifest.csv\n' +
                  f'{result.completed} OK; {result.skipped} skipped; {len(result.errors)} errors')

    @action
    def retry_failed(self):
        if self.failed:
            destination = filedialog.askdirectory()
            if destination:
                package, records = self.package, list(self.failed)
                start(self, 'Camtrap DP', lambda task: acquire_media(task, package, records, destination), self.receive_download)
