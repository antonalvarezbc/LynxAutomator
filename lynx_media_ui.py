"""Shared species selection and acquisition controls for WI and Camtrap DP."""
import customtkinter as ctk

LABELS = {
    'es': ['Cargar datapackage.json o ZIP', 'Buscar especie', 'Seleccionar visibles', 'Deseleccionar todas',
           'Incluir todas las fotografías de los eventos', 'Sólo copiar imágenes locales',
           'Revisar selección', 'Obtener fotografías seleccionadas', 'Selecciona al menos una especie.',
           'Paquete', 'observaciones', 'fotografías únicas', 'locales', 'remotas', 'privadas',
           'SINTÉTICO: etiquetas de prueba; las fotos no son de lince.',
           'Se incluyen las fotos del evento, aunque estén vacías. La detección se realizará en Wildbook.'],
    'pt': ['Carregar datapackage.json ou ZIP', 'Pesquisar espécie', 'Selecionar visíveis', 'Desmarcar todas',
           'Incluir todas as fotografias dos eventos', 'Copiar apenas imagens locais',
           'Rever seleção', 'Obter fotografias selecionadas', 'Selecione pelo menos uma espécie.',
           'Pacote', 'observações', 'fotografias únicas', 'locais', 'remotas', 'privadas',
           'SINTÉTICO: etiquetas de teste; as fotos não são de lince.',
           'Incluem-se as fotos do evento, mesmo vazias. A deteção será realizada no Wildbook.'],
    'en': ['Load datapackage.json or ZIP', 'Search species', 'Select visible', 'Clear selection',
           'Include all event photographs', 'Copy local images only',
           'Review selection', 'Get selected photographs', 'Select at least one species.',
           'Package', 'observations', 'unique photographs', 'local', 'remote', 'private',
           'SYNTHETIC: test labels; the photos are not lynxes.',
           'Event photos are included even when empty. Detection will take place in Wildbook.']}


class MediaImportTab(ctk.CTkFrame):
    def __init__(self, root, lang='es', labels=None, events=True):
        super().__init__(root)
        self.root, self.lang = root, lang
        self.text = list(labels or LABELS[lang])
        self.package = None
        self.batches = []
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
        if events:
            self.events.pack(anchor='w', padx=15, pady=4)
        self.events.select()
        self.local = ctk.CTkCheckBox(self, text=self.text[5], command=self.invalidate)
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
        self.export = ctk.CTkButton(buttons, text='Configurar Excel' if lang == 'es' else 'Configurar Excel' if lang == 'pt' else 'Configure Excel', command=self.open_bulk_import, state='disabled')
        self.export.pack(side='left', padx=5)

    def invalidate(self):
        self.failed = []
        self.batches = []
        self.export.configure(state='disabled')
        self.retry.configure(state='disabled')
        self.records = None
        self.download.configure(state='disabled')
        self.show('')

    def show(self, text):
        self.preview.configure(state='normal')
        self.preview.delete('1.0', 'end')
        self.preview.insert('1.0', text)
        self.preview.configure(state='disabled')

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

    def reset_source(self):
        self.invalidate()
        self.package = None
        for check in self.checks.values():
            check.destroy()
        self.checks = {}
        self.heading.configure(text='')

    def receive_species(self, counts, title):
        self.heading.configure(text=title)
        for check in self.checks.values():
            check.destroy()
        self.checks = {}
        for species, count in sorted(counts.items()):
            self.checks[species] = ctk.CTkCheckBox(
                self.list, text=f'{species} — {count} {self.text[10]}', command=self.invalidate)
        self.filter_species()

