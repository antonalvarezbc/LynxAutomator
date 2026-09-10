"""API source forms reuse Camtrap's species, photo and Excel workflow."""
import customtkinter as ctk
import lynx_dialogs as dialogs
from lynx_api_sources import fetch_api_package, api_base, project_key
from lynx_camtrap_ui import CamtrapTab
from lynx_ui_jobs import action, start
from lynx_api_filters import normalize_filters
from lynx_windows import foreground


class APIImportTab(CamtrapTab):
    def __init__(self, root, provider, lang='es'):
        self.provider = provider
        self.filters = normalize_filters()
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
        self.heading.configure(text=provider + ' (alpha)')
        self.filters_button = ctk.CTkButton(self.source_controls, text={'es': 'Filtrar datos…', 'pt': 'Filtrar dados…', 'en': 'Filter data…'}[lang], command=self.filter_dialog)
        self.filters_button.pack(side='left', padx=8)
        self.filter_summary = ctk.CTkLabel(self, text='', wraplength=950)
        self.filter_summary.pack(fill='x', padx=10, after=self.source_controls)
        self.update_filter_summary()
        self.server.bind('<KeyRelease>', self.connection_changed)
        self.project.bind('<KeyRelease>', lambda _: self.reset_source())
        self.approved.configure(command=self.reset_source)

    def update_filter_summary(self):
        labels = dict(zip(('year', 'site', 'deployment', 'latest', 'exclude_blank'), {
            'es': ['Año de inicio', 'Sitio contiene', 'ID contiene', 'Últimos', 'Excluir vacías'],
            'pt': ['Ano de início', 'Local contém', 'ID contém', 'Últimas', 'Excluir vazias'],
            'en': ['Start year', 'Site contains', 'ID contains', 'Latest', 'Exclude blank']}[self.lang]))
        active = [labels[key] if value is True else f'{labels[key]}: {value}' for key, value in self.filters.items() if value]
        self.filter_summary.configure(text=' · '.join(active) if active else {
            'es': 'Sin filtros de despliegue: se solicitarán todos los datos del proyecto.',
            'pt': 'Sem filtros de instalação: serão pedidos todos os dados do projeto.',
            'en': 'No deployment filters: all project data will be requested.'}[self.lang])

    @action
    def filter_dialog(self):
        previous = getattr(self, '_filters_window', None)
        if previous is not None and previous.winfo_exists():
            previous.lift()
            return
        window = self._filters_window = ctk.CTkToplevel(self)
        window.title({'es': 'Filtrar datos', 'pt': 'Filtrar dados', 'en': 'Filter data'}[self.lang])
        window.geometry('730x470')
        window.transient(self.winfo_toplevel())
        explanations = {
            'Agouti API': {
                'es': 'Primero se consulta la tabla de despliegues. Después se piden media y observaciones sólo de los despliegues elegidos. Las fotos se preparan en el siguiente paso.',
                'pt': 'Primeiro consulta-se a tabela de instalações. Depois pedem-se media e observações apenas das instalações escolhidas. As fotos preparam-se no passo seguinte.',
                'en': 'First read the deployments table, then request media and observations only for selected deployments. Photographs are prepared in the next step.'},
            'Trapper API': {
                'es': 'ID y excluir vacías se envían al servidor. Año, sitio y últimos se aplican después de descargar el ZIP: reducen las fotos posteriores, pero no la descarga inicial. El año corresponde al inicio del despliegue.',
                'pt': 'ID e excluir vazias são enviados ao servidor. Ano, local e últimas aplicam-se após descarregar o ZIP: reduzem as fotos posteriores, não a descarga inicial. O ano refere-se ao início da instalação.',
                'en': 'ID and exclude blank are sent to the server. Year, site and latest are applied after downloading the ZIP: they reduce later photos, not the initial download. Year means deployment start year.'}}
        ctk.CTkLabel(window, text=explanations[self.provider][self.lang], wraplength=680, justify='left').pack(padx=12, pady=12)
        labels = {'es': ['Año de inicio del despliegue', 'El nombre del sitio contiene', 'El ID del despliegue contiene (distingue mayúsculas)', 'Últimos N despliegues por fecha de inicio'],
                  'pt': ['Ano de início da instalação', 'O nome do local contém', 'O ID da instalação contém (distingue maiúsculas)', 'Últimas N instalações por data de início'],
                  'en': ['Deployment start year', 'Site name contains', 'Deployment ID contains (case sensitive)', 'Latest N deployments by start date']}[self.lang]
        form = ctk.CTkFrame(window)
        form.pack(fill='x', padx=12)
        entries = {}
        for index, (key, label) in enumerate(zip(('year', 'site', 'deployment', 'latest'), labels)):
            ctk.CTkLabel(form, text=label).grid(row=index, column=0, sticky='w', padx=8, pady=6)
            entry = entries[key] = ctk.CTkEntry(form, width=230)
            entry.grid(row=index, column=1, padx=8, pady=6)
            entry.insert(0, self.filters[key])
        blank = ctk.CTkCheckBox(window, text={'es': 'Excluir observaciones vacías en el servidor', 'pt': 'Excluir observações vazias no servidor', 'en': 'Exclude blank observations on the server'}[self.lang])
        if self.filters['exclude_blank']:
            blank.select()
        if self.provider == 'Trapper API':
            blank.pack(pady=6)
        @action
        def apply(owner):
            owner.filters = normalize_filters(dict({key: entry.get() for key, entry in entries.items()}, exclude_blank=bool(blank.get())))
            owner.reset_source()
            owner.update_filter_summary()
            window.destroy()
        def clear():
            for entry in entries.values():
                entry.delete(0, 'end')
            blank.deselect()
        controls = ctk.CTkFrame(window)
        controls.pack(pady=12)
        for label, command in [({'es': 'Aplicar filtros', 'pt': 'Aplicar filtros', 'en': 'Apply filters'}[self.lang], lambda: apply(self)),
                               ({'es': 'Limpiar', 'pt': 'Limpar', 'en': 'Clear'}[self.lang], clear),
                               ({'es': 'Cancelar', 'pt': 'Cancelar', 'en': 'Cancel'}[self.lang], window.destroy)]:
            ctk.CTkButton(controls, text=label, command=command).pack(side='left', padx=5)
        foreground(window, self.winfo_toplevel())

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
        filters = dict(self.filters)
        self.reset_source()
        start(self, provider, lambda task: fetch_api_package(task, provider, server, project, destination, auth, approved, filters), self.receive)
