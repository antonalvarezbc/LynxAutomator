"""API source forms reuse Camtrap's species, photo and Excel workflow."""
import customtkinter as ctk
import lynx_dialogs as dialogs
from lynx_api_sources import fetch_api_package, api_base, project_key
from lynx_camtrap_ui import CamtrapTab
from lynx_ui_jobs import action, start


class APIImportTab(CamtrapTab):
    def __init__(self, root, provider, lang='es'):
        self.provider = provider
        super().__init__(root, lang)
        labels = {'es': ['Servidor HTTPS', 'ID del proyecto', 'Cargar proyecto', 'Sólo identificaciones aprobadas'],
                  'pt': ['Servidor HTTPS', 'ID do projeto', 'Carregar projeto', 'Só identificações aprovadas'],
                  'en': ['HTTPS server', 'Project ID', 'Load project', 'Approved classifications only']}[lang]
        if provider == 'Trapper API':
            labels[1] = {'es': 'ID del proyecto de clasificación', 'pt': 'ID do projeto de classificação', 'en': 'Classification project ID'}[lang]
        form = ctk.CTkFrame(self)
        form.pack(fill='x', padx=10, pady=5, before=self.source_controls)
        for column in (1, 3):
            form.grid_columnconfigure(column, weight=1)
        ctk.CTkLabel(form, text=labels[0]).grid(row=0, column=0, padx=6)
        self.server = ctk.CTkEntry(form, width=270, placeholder_text='https://…')
        self.server.grid(row=0, column=1, sticky='ew', padx=6, pady=5)
        if provider == 'Agouti API':
            self.server.insert(0, 'https://api.agouti.eu')
        ctk.CTkLabel(form, text=labels[1]).grid(row=0, column=2, padx=6)
        self.project = ctk.CTkEntry(form, width=260)
        self.project.grid(row=0, column=3, sticky='ew', padx=6, pady=5)
        self.approved = ctk.CTkCheckBox(form, text=labels[3])
        self.approved.select()
        if provider == 'Trapper API':
            self.approved.grid(row=1, column=0, columnspan=4, sticky='w', padx=8, pady=5)
        self.load_button.configure(text=labels[2])
        self.heading.configure(text=provider)
        self.server.bind('<KeyRelease>', self.connection_changed)
        self.project.bind('<KeyRelease>', lambda _: self.reset_source())
        self.approved.configure(command=self.reset_source)

    def connection_changed(self, *_):
        self.download_auth = None
        self.reset_source()

    @action
    def load(self):
        server, project = api_base(self.server.get()), project_key(self.project.get())
        destination = dialogs.askdirectory(parent=self, title={
            'es': 'Carpeta para guardar los datos del proyecto',
            'pt': 'Pasta para guardar os dados do projeto',
            'en': 'Folder for the project data'}[self.lang])
        if not destination:
            return
        provider, auth, approved = self.provider, self.download_auth, bool(self.approved.get())
        self.reset_source()
        start(self, provider, lambda task: fetch_api_package(task, provider, server, project, destination, auth, approved), self.receive)
