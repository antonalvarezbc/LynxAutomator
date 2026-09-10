"""API source forms reuse Camtrap's species, photo and Excel workflow."""
import customtkinter as ctk
import lynx_dialogs as dialogs
from lynx_api_sources import fetch_api_package, discover_agouti, api_base, project_key
from lynx_camtrap_ui import CamtrapTab
from lynx_ui_jobs import action, start
from lynx_api_filters import normalize_filters, filter_package
from lynx_windows import foreground


class APIImportTab(CamtrapTab):
    def __init__(self, root, provider, lang='es'):
        self.provider = provider
        self.filters = normalize_filters()
        self.index = None
        self.facet_choices = {}
        self.index_connection = None
        self.index_auth = None
        self.destination = None
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
        self.load_button.configure(text={'es': 'Consultar valores del proyecto', 'pt': 'Consultar valores do projeto', 'en': 'Read project values'}[lang])
        self.heading.configure(text=provider + ' (alpha)')
        self.filters_button = ctk.CTkButton(self.source_controls, text={'es': 'Seleccionar datos…', 'pt': 'Selecionar dados…', 'en': 'Select data…'}[lang], command=self.filter_dialog)
        self.filters_button.pack(side='left', padx=8)
        self.filter_summary = ctk.CTkLabel(self, text='', wraplength=950)
        self.filter_summary.pack(fill='x', padx=10, after=self.source_controls)
        self.update_filter_summary()
        self.server.bind('<KeyRelease>', self.connection_changed)
        self.project.bind('<KeyRelease>', self.clear_index)
        self.approved.configure(command=self.clear_index)

    def update_filter_summary(self):
        if self.index is None:
            text = ({'es': 'Consultar obtiene sólo el descriptor y los despliegues para elegir valores. No descarga fotografías.', 'pt': 'Consultar obtém apenas o descritor e as instalações para escolher valores. Não descarrega fotografias.', 'en': 'Read only the descriptor and deployments to choose values. No photographs are downloaded.'} if self.provider == 'Agouti API' else {'es': 'Consultar descarga el ZIP de metadatos del proyecto para ofrecer valores reales. Las fotografías se descargan después, sólo si lo eliges.', 'pt': 'Consultar descarrega o ZIP de metadados do projeto para oferecer valores reais. As fotografias descarregam-se depois, se o escolher.', 'en': 'Read downloads the project metadata ZIP to offer actual values. Photographs download later, only if you choose to.'})[self.lang]
        else:
            selected = len(self.filters.get('deployment_ids', self.index['deployments']))
            text = f"{selected} / {len(self.index['deployments'])} " + {'es': 'despliegues seleccionados. Revisa las especies antes de obtener fotografías.', 'pt': 'instalações selecionadas. Reveja as espécies antes de obter fotografias.', 'en': 'deployments selected. Review species before obtaining photographs.'}[self.lang]
        self.filter_summary.configure(text=text)

    def connection(self):
        return (api_base(self.server.get()), project_key(self.project.get()), bool(self.approved.get()))

    def clear_index(self, *_):
        window = getattr(self, '_filters_window', None)
        if window is not None and window.winfo_exists():
            window.destroy()
        self.index = None
        self.facet_choices = {}
        self.filters = normalize_filters()
        self.reset_source()
        self.update_filter_summary()

    def connection_changed(self, *_):
        self.download_auth = None
        self.clear_index()

    @action
    def filter_dialog(self):
        if self.index is None or self.index_connection != self.connection() or self.index_auth != self.download_auth:
            self.load()
            return
        previous = getattr(self, '_filters_window', None)
        if previous is not None and previous.winfo_exists():
            previous.lift()
            return
        from lynx_api_filters_ui import DeploymentFilterDialog
        self._filters_window = DeploymentFilterDialog(self)

    @action
    def load(self):
        connection = self.connection()
        server, project, approved = connection
        destination = dialogs.askdirectory(parent=self, title={
            'es': 'Carpeta para guardar los datos seleccionados', 'pt': 'Pasta para guardar os dados selecionados',
            'en': 'Folder for selected data'}[self.lang])
        if not destination:
            return
        provider, auth = self.provider, self.download_auth
        self.clear_index()
        self.destination = destination
        def work(task):
            if provider == 'Agouti API':
                return discover_agouti(task, server, project, auth)
            package = fetch_api_package(task, provider, server, project, destination, auth, approved)
            return {'deployments': package.deployments, 'package': package}
        def receive(index):
            self.index = index
            self.index_connection, self.index_auth = connection, auth
            self.update_filter_summary()
            self.filter_dialog()
        start(self, provider, work, receive)

    @action
    def apply_selection(self, ids):
        if self.index is None or self.index_connection != self.connection() or self.index_auth != self.download_auth:
            raise ValueError('Consulta de nuevo los valores: la conexión o autorización ha cambiado.')
        self.filters = normalize_filters({'deployment_ids': ids})
        filters, index = dict(self.filters), self.index
        self.reset_source()
        self.update_filter_summary()
        server, project, approved = self.index_connection
        destination, auth, provider = self.destination, self.download_auth, self.provider
        if provider == 'Agouti API':
            work = lambda task: fetch_api_package(task, provider, server, project, destination, auth, approved, filters, index)
        else:
            work = lambda task: filter_package(index['package'], filters)
        start(self, provider, work, self.receive)
