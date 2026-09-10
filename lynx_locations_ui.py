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
        self.scopes = {'*': None}
        for row in rows:
            key = deployment_key(row)
            import json
            project, deployment = json.loads(key)
            metadata = row.get('metadata', {})
            locality = (metadata.get('deployment.locationName') or metadata.get('deployment.placename') or [''])[0]
            label = ' / '.join(v for v in [project, deployment] if v)
            if locality:
                label += ' — ' + locality
            self.scopes[label or key] = key
        scopes_frame = ctk.CTkFrame(self)
        scopes_frame.pack(fill='x', padx=10)
        self.scope = ttk.Treeview(scopes_frame, columns=('label',), show='headings', selectmode='extended', height=5)
        self.scope.heading('label', text={'es': 'Aplicar a: selecciona varios con Ctrl / Shift', 'pt': 'Aplicar a: selecione vários com Ctrl / Shift', 'en': 'Apply to: select multiple with Ctrl / Shift'}[editor.lang])
        self.scope_items = {}
        for i, (label, key) in enumerate(self.scopes.items()):
            iid = 'd' + str(i)
            self.scope.insert('', 'end', iid=iid, values=(label,))
            self.scope_items[iid] = key
        for i, row in enumerate(rows):
            iid = 'e' + str(i)
            self.scope.insert('', 'end', iid=iid, values=(f"Encounter {i+1}: {row.get('species', '')} · {row.get('year', '')}-{row.get('month', '')}-{row.get('day', '')} · {len(row.get('media', []))} photos",))
            self.scope_items[iid] = encounter_key(row)
        scroll = ttk.Scrollbar(scopes_frame, command=self.scope.yview)
        scroll.pack(side='right', fill='y')
        self.scope.configure(yscrollcommand=scroll.set)
        self.scope.pack(fill='x', expand=True)
        self.scope.selection_set('d0')
        bar = ctk.CTkFrame(self)
        bar.pack(fill='x', padx=10)
        self.book = ctk.CTkComboBox(bar, values=list(CATALOGS), width=240, command=self.select_book)
        self.book.pack(side='left', padx=5)
        self.url = ctk.CTkEntry(self, width=850)
        # Internal source storage: URLs are never shown in the normal flow.
        advanced = ctk.CTkFrame(self)
        toggle = ctk.CTkCheckBox(self, text={'es': 'Opciones avanzadas', 'pt': 'Opções avançadas', 'en': 'Advanced options'}[editor.lang], command=lambda: advanced.pack(fill='x', padx=10) if toggle.get() else advanced.pack_forget())
        toggle.pack(anchor='w', padx=10, pady=4)
        import json
        branch_cache = settings / 'locations' / 'branches.json'
        try:
            branch_names = json.loads(branch_cache.read_text(encoding='utf-8')) if branch_cache.exists() else []
        except (ValueError, OSError):
            branch_names = []
        self.branch = ctk.CTkComboBox(advanced, values=branch_names, width=320, command=self.select_branch)
        self.branch.pack(side='left', padx=4)
        ctk.CTkButton(advanced, text={'es': 'Consultar ramas GitHub', 'pt': 'Consultar ramos GitHub', 'en': 'Fetch GitHub branches'}[editor.lang], command=self.branches).pack(side='left', padx=4)
        ctk.CTkButton(advanced, text={'es': 'JSON local', 'pt': 'JSON local', 'en': 'Local JSON'}[editor.lang], command=self.local).pack(side='left', padx=4)
        labels = {'es': ['Cargar / usar caché', 'Actualizar catálogo', 'Abrir JSON local', 'Buscar ubicación', 'Aplicar ubicación'],
                  'pt': ['Carregar / usar cache', 'Atualizar catálogo', 'Abrir JSON local', 'Pesquisar localização', 'Aplicar localização'],
                  'en': ['Load / use cache', 'Refresh catalog', 'Open local JSON', 'Search location', 'Apply location']}[editor.lang]
        for text, command in [(labels[0], lambda: self.load(False)), (labels[1], lambda: self.load(True))]:
            ctk.CTkButton(bar, text=text, command=command).pack(side='left', padx=3)
        self.search = ctk.CTkEntry(self, placeholder_text=labels[3], width=850)
        self.search.pack(pady=5)
        self.search.bind('<KeyRelease>', lambda _: self.filter())
        frame = ctk.CTkFrame(self)
        frame.pack(fill='both', expand=True, padx=10)
        self.tree = ttk.Treeview(frame, columns=('route', 'id'), show='headings')
        self.tree.heading('route', text='Location')
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
        self.book.set(current.get('catalog_name', 'Lynx'))
        self.url.insert(0, current.get('catalog_source', CATALOGS['Lynx']))

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
        scopes = [self.scope_items[iid] for iid in self.scope.selection()]
        if not scopes:
            raise ValueError('Selecciona al menos un despliegue o encuentro.')
        if None in scopes:
            field.update(source='fixed', value=entry['id'], locations={})
        else:
            if field.get('source') not in ('fixed', 'location_map'):
                field['value'] = ''
            field['source'] = 'location_map'
            for scope in scopes:
                field.setdefault('locations', {})[scope] = entry['id']
        self.editor.render(fields)
        self.info.configure(text=entry['label'] + ' → ' + entry['id'])
