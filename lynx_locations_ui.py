from lynx_choices import StableComboBox, StableOptionMenu
"""Location picker: optional catalog and common/per-deployment assignments."""
import customtkinter as ctk
from tkinter import ttk
import lynx_dialogs as dialogs
from lynx_ui_jobs import action, start
from lynx_locations import CATALOGS, read_catalog, flatten, deployment_key, encounter_key, github_branches, branch_url
from lynx_bulk import settings
from lynx_windows import foreground


class LocationPicker(ctk.CTkToplevel):
    def __init__(self, editor, rows):
        super().__init__(editor)
        self.root, self.editor, self.lang = editor.root, editor, editor.lang
        self.title('Wildbook · Encounter.locationID')
        self.geometry('1000x780')
        self.transient(editor)
        foreground(self, editor, modal=True)
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.data = None
        self.entries = []
        self.source_rows = rows
        scopes_frame = ctk.CTkFrame(self)
        scopes_frame.pack(fill='x', padx=10)
        location_fields = sorted({key for row in rows for key in row.get('metadata', {}) if any(word in key.lower() for word in ['location', 'locality', 'placename', 'country', 'site'])})
        self.scope_labels = dict(zip(['Location / verbatimLocality', 'Coordinates', 'Deployment', 'Encounter'], {'es': ['Localidad de origen', 'Coordenadas', 'Despliegue de cámara', 'Encuentro'], 'pt': ['Localidade de origem', 'Coordenadas', 'Instalação da câmara', 'Encontro'], 'en': ['Source locality', 'Coordinates', 'Camera deployment', 'Encounter']}[self.lang]))
        self.scope_mode = StableOptionMenu(scopes_frame, values=list(self.scope_labels.values())[:2] + location_fields + list(self.scope_labels.values())[2:], command=lambda _: self.populate_scopes())
        self.scope_mode.pack(anchor='w', pady=4)
        self.scope = ttk.Treeview(scopes_frame, columns=('label',), show='headings', selectmode='extended', height=5)
        self.scope.heading('label', text={'es': 'Localidades de origen (Ctrl / Shift para seleccionar varias)', 'pt': 'Localizações de origem (Ctrl / Shift para selecionar várias)', 'en': 'Source locations (Ctrl / Shift for multiple selection)'}[editor.lang])
        self.scope_items = {}
        self.populate_scopes()
        scroll = ttk.Scrollbar(scopes_frame, command=self.scope.yview)
        scroll.pack(side='right', fill='y')
        self.scope.configure(yscrollcommand=scroll.set)
        self.scope.pack(fill='x', expand=True)

        bar = ctk.CTkFrame(self)
        bar.pack(fill='x', padx=10)
        self.book = StableComboBox(bar, values=list(CATALOGS), width=240, command=self.select_book)
        self.book.pack(side='left', padx=5)
        self.url = ctk.CTkEntry(self, width=850)
        # Internal source storage: URLs are never shown in the normal flow.
        advanced = ctk.CTkFrame(self)
        toggle = ctk.CTkCheckBox(self, text={'es': 'Opciones avanzadas', 'pt': 'Opções avançadas', 'en': 'Advanced options'}[editor.lang], command=lambda: advanced.pack(fill='x', padx=10) if toggle.get() else advanced.pack_forget())
        # Advanced controls are packed below the location application button.
        import json
        branch_cache = settings / 'locations' / 'branches.json'
        try:
            branch_names = json.loads(branch_cache.read_text(encoding='utf-8')) if branch_cache.exists() else []
        except (ValueError, OSError):
            branch_names = []
        self.branch = StableComboBox(advanced, values=branch_names, width=320, command=self.select_branch)
        self.branch.pack(side='left', padx=4)
        ctk.CTkButton(advanced, text={'es': 'Consultar ramas GitHub', 'pt': 'Consultar ramos GitHub', 'en': 'Fetch GitHub branches'}[editor.lang], command=self.branches).pack(side='left', padx=4)
        ctk.CTkButton(advanced, text={'es': 'JSON local', 'pt': 'JSON local', 'en': 'Local JSON'}[editor.lang], command=self.local).pack(side='left', padx=4)
        labels = {'es': ['Cargar catálogo', 'Actualizar catálogo', 'Abrir JSON local', 'Buscar ubicación', 'Asignar locationID'],
                  'pt': ['Carregar catálogo', 'Atualizar catálogo', 'Abrir JSON local', 'Pesquisar localização', 'Atribuir locationID'],
                  'en': ['Load catalog', 'Refresh catalog', 'Open local JSON', 'Search location', 'Assign locationID']}[editor.lang]
        for text, command in [(labels[0], lambda: self.load(False)), (labels[1], lambda: self.load(True))]:
            ctk.CTkButton(bar, text=text, command=command).pack(side='left', padx=3)
        self.search = ctk.CTkEntry(self, placeholder_text=labels[3], width=850)
        self.search.pack(pady=5)
        self.search.bind('<KeyRelease>', lambda _: self.filter())
        frame = ctk.CTkFrame(self)
        frame.pack(fill='both', expand=True, padx=10)
        self.tree = ttk.Treeview(frame, columns=('route', 'id'), show='headings')
        self.tree.heading('route', text={'es': 'Ubicación en Wildbook', 'pt': 'Localização no Wildbook', 'en': 'Wildbook location'}[self.lang])
        self.tree.heading('id', text='ID')
        self.tree.column('route', width=570)
        self.tree.column('id', width=220)
        scroll = ttk.Scrollbar(frame, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)
        self.info = ctk.CTkLabel(self, text='', wraplength=850)
        self.info.pack(padx=10, pady=5)
        ctk.CTkButton(self, text=labels[4], command=self.apply).pack(pady=8)
        current = next((f for f in editor.collect() if f['name'] == 'Encounter.locationID'), {})
        toggle.pack(anchor='w', padx=10, pady=4)
        self.book.set(current.get('catalog_name', 'Lynx'))
        self.url.insert(0, current.get('catalog_source', CATALOGS['Lynx']))

    def populate_scopes(self):
        self.scope.delete(*self.scope.get_children())
        self.scope.heading('label', text=self.scope_mode.get() + {'es': ' (Ctrl / Shift para seleccionar varias filas)', 'pt': ' (Ctrl / Shift para selecionar várias linhas)', 'en': ' (Ctrl / Shift to select multiple rows)'}[self.lang])
        self.scope_items = {'all': [None]}
        self.scope.insert('', 'end', iid='all', values=({'es': 'Todas las filas', 'pt': 'Todas as linhas', 'en': 'All rows'}[self.lang],))
        grouped = {}
        value = self.scope_mode.get()
        mode = next((key for key, label in self.scope_labels.items() if value == label), value)
        for index, row in enumerate(self.source_rows):
            metadata = row.get('metadata', {})
            deployment = deployment_key(row)
            if mode == 'Encounter':
                label = f"{self.scope_labels['Encounter']} {index+1}: {row.get('species', '')}"
                key = encounter_key(row)
            elif mode == 'Coordinates':
                label = f"{row.get('latitude', '')}, {row.get('longitude', '')}"
                key = deployment
            elif mode not in ('Location / verbatimLocality', 'Deployment'):
                label = ' | '.join(metadata.get(mode, [])) or {'es': '(Sin valor)', 'pt': '(Sem valor)', 'en': '(No value)'}[self.lang]
                key = encounter_key(row)
            elif mode == 'Deployment':
                identifier = (metadata.get('deployment.deploymentID') or metadata.get('deployment.deployment_id') or ['—'])[0]
                project = (metadata.get('deployment.project_id') or [''])[0]
                locality = row.get('locality') or (metadata.get('deployment.locationName') or metadata.get('deployment.placename') or [''])[0]
                label = ' · '.join(str(value) for value in (project, identifier, locality) if value)
                key = deployment
            else:
                label = str(row.get('locality') or (metadata.get('deployment.locationName') or metadata.get('deployment.placename') or [''])[0] or {'es': '(Sin localidad)', 'pt': '(Sem localidade)', 'en': '(No locality)'}[self.lang])
                key = deployment
            grouped.setdefault(label, set()).add(key)
        for index, (label, keys) in enumerate(sorted(grouped.items())):
            iid = 's' + str(index)
            self.scope_items[iid] = sorted(keys)
            count = sum(1 for row in self.source_rows if deployment_key(row) in keys or encounter_key(row) in keys)
            self.scope.insert('', 'end', iid=iid, values=(f"{label} · {count} " + {'es': 'filas', 'pt': 'linhas', 'en': 'rows'}[self.lang],))
        self.scope.selection_set('all')

    @action
    def branches(self):
        start(self, 'GitHub branches', lambda task: github_branches(task, settings / 'locations', True),
              lambda names: self.branch.configure(values=names))

    def select_branch(self, name):
        self.select_book(name)
        self.book.set(name)
        self.url.delete(0, 'end')
        self.url.insert(0, branch_url(name))

    def select_book(self, name):
        self.url.delete(0, 'end')
        self.url.insert(0, CATALOGS.get(name, ''))
        self.data = None
        self.entries = []
        self.filter()

    def close(self):
        if not self.root.winfo_toplevel().jobs.busy:
            self.destroy()

    @action
    def load(self, refresh=False):
        source = self.url.get().strip()
        start(self, 'Location catalog', lambda task: read_catalog(task, source, settings / 'locations', refresh), self.receive)

    def receive(self, data):
        self.data = data
        self.entries = flatten(data['catalog'])
        self.filter()
        self.info.configure(text=self.book.get() + ' · ' + data['updated'])

    @action
    def local(self):
        path = dialogs.askopenfilename(parent=self, filetypes=[('JSON', '*.json')])
        if path:
            self.book.set('Custom')
            self.url.delete(0, 'end')
            self.url.insert(0, path)
            self.load(True)

    def filter(self):
        self.tree.delete(*self.tree.get_children())
        query = self.search.get().casefold()
        for index, entry in enumerate(self.entries):
            if query in (entry['label'] + ' ' + entry['id']).casefold():
                self.tree.insert('', 'end', iid=str(index), values=(entry['label'], entry['id']))

    @action
    def apply(self):
        selected = self.tree.selection()
        if not selected or not self.data:
            raise ValueError('Selecciona una ubicación del catálogo cargado.')
        if self.url.get().strip() != self.data['source']:
            from lynx_locations import normalize_url
            if not self.url.get().startswith('https://') or normalize_url(self.url.get().strip()) != self.data['source']:
                raise ValueError('Carga el catálogo de la URL actual antes de aplicar.')
        entry = self.entries[int(selected[0])]
        fields = self.editor.collect()
        field = next((f for f in fields if f['name'] == 'Encounter.locationID'), None)
        if field is None:
            field = dict(name='Encounter.locationID', value='')
            fields.append(field)
        previous = field.get('catalog_source')
        if previous and previous != self.data['source']:
            field['locations'] = {}
            field['value'] = ''
        field.update(enabled=True, type='text', catalog_name=self.book.get(), catalog_source=self.data['source'])
        scopes = list(dict.fromkeys(key for iid in self.scope.selection() for key in self.scope_items[iid]))
        if not scopes:
            raise ValueError({'es': 'Selecciona las filas a las que quieres asignar el locationID.', 'pt': 'Selecione as linhas às quais pretende atribuir o locationID.', 'en': 'Select the rows to assign this locationID to.'}[self.lang])
        if None in scopes:
            field.update(source='fixed', value=entry['id'], locations={})
        else:
            if field.get('source') not in ('fixed', 'location_map'):
                field['value'] = ''
            field['source'] = 'location_map'
            for scope in scopes:
                field.setdefault('locations', {})[scope] = entry['id']
        self.editor.render(fields)
        count = sum(1 for row in self.source_rows if None in scopes or deployment_key(row) in scopes or encounter_key(row) in scopes)
        self.info.configure(text={'es': f'✓ Aplicado: {entry["id"]} → {count} filas. Guarda el perfil para reutilizarlo.', 'pt': f'✓ Aplicado: {entry["id"]} → {count} linhas.', 'en': f'✓ Applied: {entry["id"]} → {count} rows.'}[self.lang])
        for iid in self.scope.selection():
            label = self.scope.item(iid, 'values')[0].split('  ✓')[0]
            self.scope.item(iid, values=(label + '  ✓ ' + entry['id'],))
