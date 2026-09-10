from lynx_choices import StableComboBox, StableOptionMenu
"""Reusable field editor and preview for both input formats."""
import customtkinter as ctk
from tkinter import ttk, messagebox
import lynx_dialogs as dialogs
from lynx_bulk import (SOURCES, load_profile, save_profile, read_profiles, default_fields, build_table,
                       write_excel, wi_rows, camtrap_rows, wi_species)
from lynx_ui_jobs import action, start
from lynx_wildbook_fields import FIELDS
from lynx_windows import foreground

TEXT = {
'es': ['Añadir campo', 'Restablecer campos', 'Vista previa / validar', 'Guardar Excel',
       'Nombre de columna', 'Origen', 'Valor fijo', 'Tipo', 'Agrupar fotografías',
       'Filas: {}. Fotos omitidas: {}.', 'Selecciona la carpeta con las fotografías de Wildlife Insights'],
'pt': ['Adicionar campo', 'Repor campos', 'Pré-visualizar / validar', 'Guardar Excel',
       'Nome da coluna', 'Origem', 'Valor fixo', 'Tipo', 'Agrupar fotografias',
       'Linhas: {}. Fotos omitidas: {}.', 'Selecione a pasta com as fotografias do Wildlife Insights'],
'en': ['Add field', 'Reset fields', 'Preview / validate', 'Save Excel',
       'Column name', 'Source', 'Fixed value', 'Type', 'Group photographs',
       'Rows: {}. Omitted photos: {}.', 'Select the Wildlife Insights photographs folder']}


class BulkEditor(ctk.CTkToplevel):
    def __init__(self, owner, loader):
        super().__init__(owner.root)
        self.root = owner.root
        self.lang = getattr(owner, 'lang', 'es')
        self.words = TEXT.get(self.lang, TEXT['es'])
        self.loader = loader
        self.rows = []
        self.source_names = list(SOURCES)
        self.title('Wildbook Bulk Import')
        self.geometry('1120x740')
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.group = ctk.CTkCheckBox(self, text=self.words[8], command=self.invalidate)
        self.group.pack(anchor='w', padx=12, pady=8)
        self.interval = ctk.CTkEntry(self, width=140)
        self.interval.insert(0, '3')
        self.interval.pack(anchor='w', padx=12)
        self.interval.bind('<KeyRelease>', self.invalidate)
        ctk.CTkLabel(self, text={'es': 'Intervalo máximo entre fotos (segundos). Los eventos explícitos de Camtrap DP se conservan.', 'pt': 'Intervalo máximo entre fotos (segundos). Eventos explícitos de Camtrap DP são preservados.', 'en': 'Maximum gap between photos (seconds). Explicit Camtrap DP events are preserved.'}[self.lang]).pack(anchor='w', padx=12)
        heading = ctk.CTkFrame(self)
        heading.pack(fill='x', padx=10)
        for label in self.words[4:8]:
            ctk.CTkLabel(heading, text=label, width=220).pack(side='left')
        self.fields_frame = ctk.CTkScrollableFrame(self, height=280)
        self.fields_frame.pack(fill='both', expand=True, padx=10)
        profile_bar = ctk.CTkFrame(self)
        profile_bar.pack(fill='x', padx=10, before=heading)
        labels = {'es': ['Perfil local (opcional)', 'Cargar perfil', 'Guardar perfil', 'Nombre: Doñana…'],
                  'pt': ['Perfil local (opcional)', 'Carregar perfil', 'Guardar perfil', 'Nome: Doñana…'],
                  'en': ['Local profile (optional)', 'Load profile', 'Save profile', 'Name: Doñana…']}[self.lang]
        ctk.CTkLabel(profile_bar, text=labels[0]).pack(side='left', padx=4)
        try:
            names = sorted(read_profiles())
        except (ValueError, OSError, KeyError) as exc:
            messagebox.showwarning('Bulk Import', str(exc), parent=self)
            names = []
        self.profile_choice = StableOptionMenu(profile_bar, values=names or ['—'])
        self.profile_choice.pack(side='left', padx=4)
        ctk.CTkButton(profile_bar, text=labels[1], command=self.use_profile).pack(side='left', padx=4)
        self.profile_name = ctk.CTkEntry(profile_bar, placeholder_text=labels[3])
        self.profile_name.pack(side='left', padx=4)
        ctk.CTkButton(profile_bar, text=labels[2], command=self.store_profile).pack(side='left', padx=4)
        ctk.CTkButton(self, text='Wildbook / locationID', command=self.choose_location).pack(before=heading, pady=4)
        fields = default_fields()
        self.render(fields)
        buttons = ctk.CTkFrame(self)
        buttons.pack(fill='x', padx=10, pady=8)
        for text, command in [(self.words[0], self.add), (self.words[1], lambda: self.render(default_fields())),
                              (self.words[2], self.preview), (self.words[3], self.save),
                              ({'es': 'Comentarios con metadatos', 'pt': 'Comentários com metadados', 'en': 'Metadata comments'}[self.lang], self.metadata_comments)]:
            ctk.CTkButton(buttons, text=text, command=command).pack(side='left', padx=4)
        self.status = ctk.CTkLabel(self, text='', wraplength=1050)
        self.status.pack(fill='x', padx=10)
        ctk.CTkLabel(self, text={'es': 'Obligatorios: ubicación · año · foto (automática) · género y especie. template: texto con {deployment.cameraID}.', 'pt': 'Obrigatórios: localização · ano · foto (automática) · género e espécie. template: texto com {deployment.cameraID}.', 'en': 'Required: location · year · photo (automatic) · genus and species. template: text with {deployment.cameraID}.'}[self.lang], wraplength=1050).pack(fill='x')
        frame = ctk.CTkFrame(self)
        frame.pack(fill='both', expand=True, padx=10, pady=8)
        self.tree = ttk.Treeview(frame, show='headings', height=7)
        self.tree.grid(row=0, column=0, sticky='nsew')
        x = ttk.Scrollbar(frame, orient='horizontal', command=self.tree.xview)
        y = ttk.Scrollbar(frame, orient='vertical', command=self.tree.yview)
        x.grid(row=1, column=0, sticky='ew')
        y.grid(row=0, column=1, sticky='ns')
        self.tree.configure(xscrollcommand=x.set, yscrollcommand=y.set)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        foreground(self, self.root)

    def group_options(self):
        threshold = float(self.interval.get())
        import math
        if not math.isfinite(threshold) or threshold < 0:
            raise ValueError('El intervalo debe ser un número finito mayor o igual a cero.')
        return bool(self.group.get()), threshold

    @action
    def choose_location(self):
        from lynx_locations_ui import LocationPicker
        group = self.group_options()
        start(self, 'Locations', lambda task: self.loader(task, group),
              lambda result: LocationPicker(self, result[0]))

    @action
    def use_profile(self):
        name = self.profile_choice.get()
        if name == '—':
            return
        self.render(load_profile(name=name))
        self.profile_name.delete(0, 'end')
        self.profile_name.insert(0, name)

    @action
    def store_profile(self):
        name = self.profile_name.get().strip()
        save_profile(self.collect(), name=name)
        self.profile_choice.configure(values=sorted(read_profiles()))
        self.profile_choice.set(name)
        messagebox.showinfo('Bulk Import', name, parent=self)

    def close(self):
        if not self.root.winfo_toplevel().jobs.busy:
            self.destroy()

    def invalidate(self, *_):
        self.snapshot = None

    def collect(self):
        return [dict(original, name=n.get(), source=s.get(), value=v.get(), type=t.get(), enabled=bool(e.get()))
                for original, (e, n, s, v, t) in zip(self.field_options, self.rows)]

    def render(self, fields):
        self.field_options = fields
        self.invalidate()
        for child in self.fields_frame.winfo_children():
            child.destroy()
        self.rows = []
        for index, field in enumerate(fields):
            line = ctk.CTkFrame(self.fields_frame)
            line.pack(fill='x', pady=2)
            enabled = ctk.CTkCheckBox(line, text='', width=25, command=self.invalidate)
            enabled.pack(side='left')
            if field.get('enabled', True):
                enabled.select()
            name = StableComboBox(line, width=250, values=FIELDS, command=self.invalidate)
            name.set(field['name'])
            name.pack(side='left', padx=3)
            source = StableComboBox(line, values=['fixed', 'template', 'location_map'] + self.source_names, width=155, command=self.invalidate)
            source.set(field['source'])
            source.pack(side='left', padx=3)
            value = ctk.CTkEntry(line, width=230)
            value.insert(0, field.get('value', ''))
            value.pack(side='left', padx=3)
            kind = StableOptionMenu(line, values=['text', 'integer', 'decimal', 'boolean'], width=110, command=self.invalidate)
            kind.set(field['type'])
            kind.pack(side='left', padx=3)
            for entry in (name, source, value):
                entry.bind('<KeyRelease>', self.invalidate)
            ctk.CTkButton(line, text='↑', width=30, command=lambda i=index: self.move(i)).pack(side='left', padx=2)
            ctk.CTkButton(line, text='×', width=30, command=lambda i=index: self.remove(i)).pack(side='left', padx=2)
            self.rows.append((enabled, name, source, value, kind))

    def move(self, index):
        fields = self.collect()
        if index:
            fields[index - 1], fields[index] = fields[index], fields[index - 1]
            self.render(fields)

    def remove(self, index):
        fields = self.collect()
        del fields[index]
        self.render(fields)

    def add(self):
        self.render(self.collect() + [dict(name='', source='fixed', value='', type='text', enabled=True)])

    @action
    def metadata_comments(self):
        group = self.group_options()
        def receive(result):
            rows, _ = result
            self.source_names = sorted(set(SOURCES) | {key for row in rows for key in row.get('metadata', {})})
            window = ctk.CTkToplevel(self)
            window.title({'es': 'Seleccionar metadatos para comentarios', 'pt': 'Selecionar metadados para comentários', 'en': 'Select metadata for comments'}[self.lang])
            window.geometry('620x540')
            window.transient(self)
            foreground(window, self, modal=True)
            destination = StableComboBox(window, width=400, values=['Sighting.comments', 'Encounter.sightingRemarks', 'Encounter.researcherComments'])
            destination.pack(pady=8)
            destination.set('Sighting.comments')
            search = ctk.CTkEntry(window, placeholder_text='camera / setup / deployment / …', width=450)
            search.pack(pady=5)
            listing = ctk.CTkScrollableFrame(window)
            listing.pack(fill='both', expand=True, padx=10)
            checks = {}
            for key in self.source_names:
                if '.' not in key:
                    continue
                check = ctk.CTkCheckBox(listing, text=key)
                check.pack(anchor='w', pady=3)
                if key in ['deployment.cameraID', 'deployment.cameraModel', 'deployment.setupBy']:
                    check.select()
                checks[key] = check
            def filter_keys(_):
                for key, check in checks.items():
                    check.pack_forget()
                    if search.get().casefold() in key.casefold():
                        check.pack(anchor='w', pady=3)
            search.bind('<KeyRelease>', filter_keys)
            def apply():
                keys = [key for key, check in checks.items() if check.get()]
                if not keys:
                    return
                fields = self.collect()
                name = destination.get().strip()
                template = '; '.join(key + ': {' + key + '}' for key in keys)
                existing = next((f for f in fields if f['name'] == name), None)
                if existing:
                    # Preserve existing fixed/template notes rather than replacing them.
                    previous = existing.get('value', '') if existing['source'] in ['fixed', 'template'] else '{' + existing['source'] + '}'
                    existing.update(source='template', value='; '.join(v for v in [previous, template] if v), enabled=True, type='text')
                else:
                    fields.append(dict(name=name, source='template', value=template, type='text', enabled=True))
                if name == 'Sighting.comments' and not any(f.get('enabled', True) and f['name'] in ['Encounter.sightingID', 'Sighting.sightingID'] for f in fields):
                    fields.append(dict(name='Encounter.sightingID', source='eventID', value='', type='text', enabled=True))
                self.render(fields)
                window.destroy()
            ctk.CTkButton(window, text={'es': 'Añadir a comentarios', 'pt': 'Adicionar aos comentários', 'en': 'Add to comments'}[self.lang], command=apply).pack(pady=10)
        start(self, 'Metadata', lambda task: self.loader(task, group), receive)

    @action
    def preview(self):
        fields, group = self.collect(), self.group_options()
        self.invalidate()
        def work(task):
            rows, missing = self.loader(task, group)
            table = build_table(rows, fields, task)
            return table, missing, rows
        def receive(result):
            table, missing, source_rows = result
            self.source_names = sorted(set(SOURCES) | {key for row in source_rows for key in row.get('metadata', {})})
            for _, _, source, _, _ in self.rows:
                source.configure(values=['fixed', 'template', 'location_map'] + self.source_names)
            self.snapshot = (fields, group, source_rows)
            headers = list(dict.fromkeys(k for row in table for k in row))
            self.tree.delete(*self.tree.get_children())
            self.tree.configure(columns=headers)
            for name in headers:
                self.tree.heading(name, text=name)
                self.tree.column(name, width=180)
            for row in table[:100]:
                self.tree.insert('', 'end', values=[row.get(k, '') for k in headers])
            unmatched = sorted({name for row in source_rows for name in row.get('_unmatched_tables', [])})
            warning = ({'es': 'Tablas sin correspondencia en algunas filas: ', 'pt': 'Tabelas sem correspondência em algumas linhas: ', 'en': 'Tables without matches in some rows: '}[self.lang] + ', '.join(unmatched)) if unmatched else ''
            self.status.configure(text=self.words[9].format(len(table), len(missing)) + '\n' + ', '.join(missing[:10]) + '\n' + warning)
        start(self, 'Bulk Import', work, receive)

    @action
    def save(self):
        if not self.snapshot or self.snapshot[:2] != (self.collect(), self.group_options()):
            raise ValueError(self.words[2])
        path = dialogs.asksaveasfilename(parent=self, initialfile='wildbook_bulk_import.xlsx', defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if path:
            fields, _, rows = self.snapshot
            start(self, 'Bulk Import', lambda task: write_excel(task, build_table(rows, fields, task), path),
                  lambda _: messagebox.showinfo('Bulk Import', path, parent=self))


class WIInput(ctk.CTkToplevel):
    """CSV/ZIP input independent of the legacy template form."""
    def __init__(self, owner):
        super().__init__(owner.root)
        self.root, self.lang = owner.root, owner.lang
        self.images = getattr(owner, 'images_csv_path', '')
        self.deployments = getattr(owner, 'deployments_csv_path', '')
        self.archive = ''
        self.extra_paths = []
        self.title('Wildlife Insights → Bulk Import')
        self.geometry('800x680')
        self.checks = {}
        words = {
            'es': ['Carga el ZIP de Wildlife Insights o sus dos CSV. No necesitas plantilla Excel ni fotos locales.',
                   'Cargar ZIP', 'Seleccionar images*.csv', 'Seleccionar deployments.csv', 'Configurar Bulk Import', 'Intervalo de agrupación (segundos)'],
            'pt': ['Carregue o ZIP do Wildlife Insights ou os dois CSV. Não precisa de modelo Excel nem fotos locais.',
                   'Carregar ZIP', 'Selecionar images*.csv', 'Selecionar deployments.csv', 'Configurar Bulk Import', 'Intervalo de agrupamento (segundos)'],
            'en': ['Load the Wildlife Insights ZIP or its two CSVs. No Excel template or local photos are needed.',
                   'Load ZIP', 'Select images*.csv', 'Select deployments.csv', 'Configure Bulk Import', 'Grouping interval (seconds)']}[self.lang]
        ctk.CTkLabel(self, text=words[0], wraplength=700).pack(padx=12, pady=10)
        for label, kind in zip(words[1:4], ['archive', 'images', 'deployments']):
            ctk.CTkButton(self, text=label, command=lambda k=kind: self.choose(k)).pack(pady=4)
        self.selection = ctk.CTkLabel(self, text='', wraplength=700)
        self.selection.pack(padx=10, pady=6)
        ctk.CTkButton(self, text={'es': 'Añadir projects.csv / otro CSV', 'pt': 'Adicionar projects.csv / outro CSV', 'en': 'Add projects.csv / other CSV'}[self.lang], command=self.add_extra).pack(pady=4)
        ctk.CTkButton(self, text={'es': 'Leer especies', 'pt': 'Ler espécies', 'en': 'Read species'}[self.lang], command=self.load_species).pack(pady=5)
        self.species_search = ctk.CTkEntry(self, placeholder_text={'es': 'Buscar especie', 'pt': 'Pesquisar espécie', 'en': 'Search species'}[self.lang])
        self.species_search.pack(fill='x', padx=10)
        self.species_search.bind('<KeyRelease>', lambda _: self.filter_species())
        self.species_list = ctk.CTkScrollableFrame(self, height=150)
        self.species_list.pack(fill='both', expand=True, padx=10)
        ctk.CTkButton(self, text={'es': 'Seleccionar visibles', 'pt': 'Selecionar visíveis', 'en': 'Select visible'}[self.lang], command=self.select_species).pack(pady=4)
        ctk.CTkButton(self, text=words[4], command=self.configure_bulk).pack(pady=10)
        self.show_selection()
        foreground(self, self.root)

    def show_selection(self):
        self.selection.configure(text=(self.archive or '\n'.join(str(p or '—') for p in (self.images, self.deployments))) + '\n' + ', '.join(self.extra_paths))

    @action
    def choose(self, kind):
        path = dialogs.askopenfilename(parent=self, filetypes=[('ZIP', '*.zip')] if kind == 'archive' else [('CSV', '*.csv')])
        if path:
            setattr(self, kind, path)
            for check in self.checks.values():
                check.destroy()
            self.checks = {}
            if kind != 'archive':
                self.archive = ''
            self.show_selection()

    @action
    def load_species(self):
        images, deployments = (self.archive, None) if self.archive else (self.images, self.deployments)
        if not images:
            raise ValueError('Selecciona un ZIP o los CSV.')
        def receive(counts):
            for check in self.checks.values():
                check.destroy()
            self.checks = {}
            for species, count in sorted(counts.items()):
                check = ctk.CTkCheckBox(self.species_list, text=f'{species} — {count}')
                self.checks[species] = check
            self.filter_species()
        start(self, 'Wildlife Insights: species', lambda task: wi_species(task, images, deployments), receive)

    def filter_species(self):
        for name, check in self.checks.items():
            check.pack_forget()
            if self.species_search.get().casefold() in name.casefold():
                check.pack(anchor='w', padx=6, pady=3)

    def select_species(self):
        for name, check in self.checks.items():
            if self.species_search.get().casefold() in name.casefold():
                check.select()

    @action
    def add_extra(self):
        path = dialogs.askopenfilename(parent=self, filetypes=[('CSV', '*.csv')])
        if path and path not in self.extra_paths:
            self.extra_paths.append(path)
            self.show_selection()

    @action
    def configure_bulk(self):
        images, deployments = (self.archive, None) if self.archive else (self.images, self.deployments)
        if not images or (not self.archive and not deployments):
            raise ValueError('Selecciona images.csv y deployments.csv, o un ZIP con ambos.')
        species = {name for name, check in self.checks.items() if check.get()}
        if not species:
            raise ValueError('Pulsa Leer especies y selecciona al menos una especie.')
        extra_paths = tuple(self.extra_paths)
        BulkEditor(self, lambda task, options: wi_rows(task, images, deployments, group=options[0], threshold=options[1], extra_paths=extra_paths, species=species))
        self.destroy()


def open_wi(owner):
    WIInput(owner)


def open_camtrap(owner):
    if not owner.records or not owner.batches:
        raise ValueError('Obtén primero las fotografías de la selección actual.')
    package, records, batches = owner.package, list(owner.records), list(owner.batches)
    BulkEditor(owner, lambda task, options: camtrap_rows(task, package, records, batches, options[0], options[1]))
