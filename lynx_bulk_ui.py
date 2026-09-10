from lynx_choices import StableComboBox, StableOptionMenu
"""Reusable field editor and preview for both input formats."""
import customtkinter as ctk
from tkinter import ttk, messagebox
import lynx_dialogs as dialogs
from lynx_bulk import (SOURCES, load_profile, save_profile, read_profiles, default_fields, build_table,
                       write_excel, camtrap_rows)
from lynx_ui_jobs import action, start
from lynx_wildbook_fields import FIELDS
from lynx_windows import foreground, WorkflowPanel

TEXT = {
'es': ['Añadir campo', 'Restablecer campos', 'Previsualizar', 'Guardar Excel',
       'Nombre de columna', 'Obtener valor de', 'Valor o formato', 'Tipo de dato', 'Agrupar fotografías',
       'Filas: {}. Fotos omitidas: {}.', 'Selecciona la carpeta con las fotografías de Wildlife Insights'],
'pt': ['Adicionar campo', 'Repor campos', 'Pré-visualizar', 'Guardar Excel',
       'Nome da coluna', 'Obter valor de', 'Valor ou formato', 'Tipo de dado', 'Agrupar fotografias',
       'Linhas: {}. Fotos omitidas: {}.', 'Selecione a pasta com as fotografias do Wildlife Insights'],
'en': ['Add field', 'Reset fields', 'Preview', 'Save Excel',
       'Column name', 'Value source', 'Value or format', 'Data type', 'Group photographs',
       'Rows: {}. Omitted photos: {}.', 'Select the Wildlife Insights photographs folder']}


class MetadataSourceBox(StableComboBox):
    def _open_dropdown_menu(self):
        self.editor.ensure_sources(lambda: super(MetadataSourceBox, self)._open_dropdown_menu())


class BulkEditor(WorkflowPanel):
    def __init__(self, owner, loader):
        super().__init__(owner)
        self.root = owner.root
        self.lang = getattr(owner, 'lang', 'es')
        self.words = TEXT.get(self.lang, TEXT['es'])
        self.loader = loader
        self.rows = []
        self.source_names = list(SOURCES)
        self.title('Wildbook Bulk Import')
        self.geometry('1120x740')
        self.protocol('WM_DELETE_WINDOW', self.close)
        grouping = ctk.CTkFrame(self)
        grouping.pack(fill='x', padx=10, pady=8)
        self.group = ctk.CTkCheckBox(grouping, text=self.words[8], command=self.invalidate)
        self.group.pack(side='left', padx=8)
        ctk.CTkLabel(grouping, text={'es': 'Intervalo máximo entre fotos:', 'pt': 'Intervalo máximo entre fotos:', 'en': 'Maximum gap between photos:'}[self.lang]).pack(side='left', padx=(16, 4))
        self.interval = ctk.CTkEntry(grouping, width=75)
        self.interval.insert(0, '3')
        self.interval.pack(side='left', padx=4)
        self.interval.bind('<KeyRelease>', self.invalidate)
        ctk.CTkLabel(grouping, text={'es': 'segundos', 'pt': 'segundos', 'en': 'seconds'}[self.lang]).pack(side='left', padx=4)
        ctk.CTkLabel(grouping, text={'es': 'Se respetan los eventos definidos en Camtrap DP.', 'pt': 'Respeitam-se os eventos definidos no Camtrap DP.', 'en': 'Defined Camtrap DP events are preserved.'}[self.lang]).pack(side='left', padx=12)
        self.source_labels = dict(zip(['fixed', 'template', 'location_map'], {
            'es': ['Valor fijo', 'Texto con metadatos', 'Ubicaciones asignadas'],
            'pt': ['Valor fixo', 'Texto com metadados', 'Localizações atribuídas'],
            'en': ['Fixed value', 'Text with metadata', 'Assigned locations']}[self.lang]))
        self.source_labels.update(dict(zip(SOURCES, {
            'es': ['Género', 'Epíteto específico', 'Latitud', 'Longitud', 'Localidad', 'Año', 'Mes', 'Día', 'Hora', 'Minutos', 'ID de evento', 'ID de individuo', 'Nombre científico'],
            'pt': ['Género', 'Epíteto específico', 'Latitude', 'Longitude', 'Localidade', 'Ano', 'Mês', 'Dia', 'Hora', 'Minutos', 'ID de evento', 'ID de indivíduo', 'Nome científico'],
            'en': ['Genus', 'Specific epithet', 'Latitude', 'Longitude', 'Locality', 'Year', 'Month', 'Day', 'Hour', 'Minutes', 'Event ID', 'Individual ID', 'Scientific name']}[self.lang])))
        self.type_labels = dict(zip(['text', 'integer', 'decimal', 'boolean'], {
            'es': ['Texto', 'Entero', 'Decimal', 'Lógico'],
            'pt': ['Texto', 'Inteiro', 'Decimal', 'Lógico'],
            'en': ['Text', 'Integer', 'Decimal', 'Boolean']}[self.lang]))
        heading = ctk.CTkFrame(self)
        heading.pack(fill='x', padx=(10, 26))
        self.field_grid(heading)
        for column, label in enumerate(self.words[4:8], 1):
            ctk.CTkLabel(heading, text=label, anchor='w').grid(row=0, column=column, sticky='ew', padx=3)
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
        fields = default_fields()
        self.render(fields)
        buttons = ctk.CTkFrame(self)
        buttons.pack(fill='x', padx=10, pady=8)
        for text, command in [(self.words[0], self.add), (self.words[1], lambda: self.render(default_fields())),
                              (self.words[2], self.preview), (self.words[3], self.save)]:
            ctk.CTkButton(buttons, text=text, command=command).pack(side='left', padx=4)
        self.status = ctk.CTkLabel(self, text='', wraplength=1050)
        self.status.pack(fill='x', padx=10)
        ctk.CTkLabel(self, text={'es': 'Obligatorios: ubicación · año (opcional en catálogo) · foto (automática) · nombre científico (género y epíteto específico).', 'pt': 'Obrigatórios: localização · ano (opcional no catálogo) · foto (automática) · nome científico (género e epíteto específico).', 'en': 'Required: location · year (optional for catalog) · photo (automatic) · scientific name (genus and specific epithet).'}[self.lang], wraplength=1050).pack(fill='x')
        self.preview_window = None
        self.sources_group = None
        foreground(self, self.root)

    def preview_popup(self):
        window = self.preview_window
        if window is not None and window.winfo_exists():
            window.destroy()
        window = self.preview_window = ctk.CTkToplevel(self)
        window.title(self.words[2] + ' · Bulk Import')
        window.geometry('1100x480')
        window.transient(self.winfo_toplevel())
        def close_preview():
            if not self.root.winfo_toplevel().jobs.busy:
                window.destroy()
        window.protocol('WM_DELETE_WINDOW', close_preview)
        ctk.CTkLabel(window, text={'es': 'Se muestran hasta 100 filas. El Excel incluye todas las filas.', 'pt': 'Mostram-se até 100 linhas. O Excel inclui todas as linhas.', 'en': 'Showing up to 100 rows. Excel includes all rows.'}[self.lang]).pack()
        frame = ctk.CTkFrame(window)
        frame.pack(fill='both', expand=True, padx=10, pady=8)
        self.tree = ttk.Treeview(frame, show='headings', height=12)
        self.tree.grid(row=0, column=0, sticky='nsew')
        x = ttk.Scrollbar(frame, orient='horizontal', command=self.tree.xview)
        y = ttk.Scrollbar(frame, orient='vertical', command=self.tree.yview)
        x.grid(row=1, column=0, sticky='ew')
        y.grid(row=0, column=1, sticky='ns')
        self.tree.configure(xscrollcommand=x.set, yscrollcommand=y.set)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        ctk.CTkButton(window, text={'es': 'Cerrar', 'pt': 'Fechar', 'en': 'Close'}[self.lang], command=close_preview).pack(pady=6)
        foreground(window, self.winfo_toplevel())

    def refresh_sources(self, rows, group):
        self.source_names = sorted(set(SOURCES) | {key for row in rows for key, values in row.get('metadata', {}).items() if values})
        self.sources_group = group
        for _, _, source, _, _ in self.rows:
            source.configure(values=self.source_choices())

    @action
    def ensure_sources(self, callback):
        group = self.group_options()
        if self.sources_group == group:
            callback()
            return
        def receive(result):
            self.refresh_sources(result[0], group)
            callback()
        start(self, 'Metadatos / Metadata', lambda task: self.loader(task, group), receive)

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
        def receive(result):
            self.refresh_sources(result[0], group)
            LocationPicker(self, result[0])
        start(self, {'es': 'Ubicaciones', 'pt': 'Localizações', 'en': 'Locations'}[self.lang], lambda task: self.loader(task, group), receive)

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
        window = getattr(self, 'preview_window', None)
        if window is not None and window.winfo_exists():
            window.destroy()

    def collect(self):
        return [dict(original, name=n.get(), source=self.internal_label(s.get(), self.source_labels), value=v.get(), type=self.internal_label(t.get(), self.type_labels), enabled=bool(e.get()))
                for original, (e, n, s, v, t) in zip(self.field_options, self.rows)]

    @staticmethod
    def internal_label(value, labels):
        return next((key for key, label in labels.items() if label == value), value)

    def source_choices(self):
        return list(self.source_labels.values()) + [name for name in self.source_names if name not in self.source_labels]

    @staticmethod
    def field_grid(widget):
        for column, (minimum, weight) in enumerate([(28, 0), (245, 2), (165, 1), (280, 2), (110, 0), (74, 0)]):
            widget.grid_columnconfigure(column, minsize=minimum, weight=weight)

    def render(self, fields):
        self.field_options = fields
        self.invalidate()
        for child in self.fields_frame.winfo_children():
            child.destroy()
        self.rows = []
        self.location_buttons = []
        for index, field in enumerate(fields):
            line = ctk.CTkFrame(self.fields_frame)
            line.pack(fill='x', pady=2)
            self.field_grid(line)
            enabled = ctk.CTkCheckBox(line, text='', width=25, command=self.invalidate)
            enabled.grid(row=0, column=0)
            if field.get('enabled', True):
                enabled.select()
            name = StableComboBox(line, width=235, values=FIELDS)
            name.set(field['name'])
            name.grid(row=0, column=1, sticky='ew', padx=3)
            source = MetadataSourceBox(line, values=self.source_choices(), width=155, command=self.invalidate)
            source.editor = self
            source.set(self.source_labels.get(field['source'], field['source']))
            source.grid(row=0, column=2, sticky='ew', padx=3)
            value_frame = ctk.CTkFrame(line, fg_color='transparent')
            value_frame.grid(row=0, column=3, sticky='ew', padx=3)
            value_frame.grid_columnconfigure(0, weight=1)
            value = ctk.CTkEntry(value_frame, width=140)
            value.insert(0, field.get('value', ''))
            value.grid(row=0, column=0, sticky='ew')
            locations = ctk.CTkButton(value_frame, width=120, text={
                'es': 'Elegir locationID', 'pt': 'Escolher locationID', 'en': 'Choose locationID'}[self.lang], command=self.choose_location)
            self.location_buttons.append(locations)
            def changed(*_, name=name, button=locations):
                if name.get().strip() == 'Encounter.locationID':
                    button.grid(row=0, column=1, padx=(4, 0))
                else:
                    button.grid_remove()
                self.invalidate()
            name.configure(command=changed)
            name.bind('<KeyRelease>', changed)
            changed()
            kind = StableOptionMenu(line, values=list(self.type_labels.values()), width=100, command=self.invalidate)
            kind.set(self.type_labels.get(field['type'], field['type']))
            kind.grid(row=0, column=4, sticky='ew', padx=3)
            def kind_changed(*_, kind=kind, value=value):
                value.configure(placeholder_text='true / false' if self.internal_label(kind.get(), self.type_labels) == 'boolean' else '')
                self.invalidate()
            kind.configure(command=kind_changed)
            kind_changed()
            for entry in (source, value):
                entry.bind('<KeyRelease>', self.invalidate)
            actions = ctk.CTkFrame(line, fg_color='transparent')
            actions.grid(row=0, column=5)
            ctk.CTkButton(actions, text='↑', width=30, command=lambda i=index: self.move(i)).pack(side='left', padx=2)
            ctk.CTkButton(actions, text='×', width=30, command=lambda i=index: self.remove(i)).pack(side='left', padx=2)
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

    def add_standard(self):
        self.render(self.collect() + [dict(name='', source='fixed', value='', type='text', enabled=True)])

    def add(self):
        window = ctk.CTkToplevel(self)
        window.title(self.words[0])
        window.geometry('440x170')
        window.transient(self.winfo_toplevel())
        def choose(command):
            window.destroy()
            command()
        for label, command in [({'es': 'Campo del Excel', 'pt': 'Campo do Excel', 'en': 'Excel field'}[self.lang], self.add_standard),
                               ({'es': 'Comentarios con metadatos', 'pt': 'Comentários com metadados', 'en': 'Metadata comments'}[self.lang], self.metadata_comments)]:
            ctk.CTkButton(window, text=label, width=300, command=lambda cmd=command: choose(cmd)).pack(pady=14)
        foreground(window, self.winfo_toplevel())

    @action
    def metadata_comments(self):
        group = self.group_options()
        def receive(result):
            rows, _ = result
            self.refresh_sources(rows, group)
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
            selection = ctk.CTkFrame(window)
            selection.pack(fill='x', padx=10)
            def select_visible(selected):
                for key, check in checks.items():
                    if not selected:
                        check.deselect()
                    elif search.get().casefold() in key.casefold():
                        check.select()
            for label, value in [({'es': 'Seleccionar visibles', 'pt': 'Selecionar visíveis', 'en': 'Select visible'}[self.lang], True),
                                 ({'es': 'Deseleccionar todos', 'pt': 'Desmarcar todos', 'en': 'Clear selection'}[self.lang], False)]:
                ctk.CTkButton(selection, text=label, command=lambda v=value: select_visible(v)).pack(side='left', padx=4)

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
        start(self, {'es': 'Metadatos', 'pt': 'Metadados', 'en': 'Metadata'}[self.lang], lambda task: self.loader(task, group), receive)

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
            self.refresh_sources(source_rows, group)
            self.preview_table = table
            self.preview_popup()
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
        from lynx_locations import export_filename
        path = dialogs.asksaveasfilename(parent=self, initialfile=export_filename(self.preview_table, self.snapshot[0]), defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if path:
            fields, _, rows = self.snapshot
            start(self, 'Bulk Import', lambda task: write_excel(task, build_table(rows, fields, task), path),
                  lambda _: messagebox.showinfo('Bulk Import', path, parent=self))


class WIInput(WorkflowPanel):
    """Standalone wrapper for the same WI page used by Bulk Import."""
    def __init__(self, owner):
        super().__init__(owner)
        self.root, self.lang = owner.root, owner.lang
        self.title('Wildlife Insights → Bulk Import')
        self.geometry('1000x720')
        from lynx_wi_ui import WITab
        self.page = WITab(self, self.lang)
        self.page.root = self.root
        self.page.workflow = self.workflow
        foreground(self, self.root)


def open_wi(owner):
    WIInput(owner)


def open_camtrap(owner):
    if not owner.records or not owner.batches:
        raise ValueError('Obtén primero las fotografías de la selección actual.')
    package, records, batches = owner.package, list(owner.records), list(owner.batches)
    loader = lambda task, options: camtrap_rows(task, package, records, batches, options[0], options[1])
    if getattr(owner, 'workflow', None):
        owner.workflow.edit(loader)
    else:
        BulkEditor(owner, loader)
