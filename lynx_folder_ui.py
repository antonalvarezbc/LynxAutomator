"""Configure folder/catalog inputs; export uses the same editor as WI and DP."""
import customtkinter as ctk
import lynx_dialogs as dialogs
from lynx_choices import StableOptionMenu
from lynx_windows import foreground, WorkflowPanel
from lynx_ui_jobs import action, start
from lynx_bulk_ui import BulkEditor
from lynx_folder_bulk import folder_rows, scan_layout, validate_stations


class FolderInput(WorkflowPanel):
    def __init__(self, owner, catalog=False):
        super().__init__(owner)
        self.root, self.lang = owner.root, owner.lang
        self.catalog = catalog
        self.folder = getattr(owner, 'folder_path', '') or ''
        self.title('Wildbook Bulk Import · ' + {'es': ['Crear desde carpeta', 'Catálogo'], 'pt': ['Criar a partir de pasta', 'Catálogo'], 'en': ['Create from folder', 'Catalog']}[self.lang][int(catalog)])
        self.geometry('920x720')
        self.stations = {}
        self.layout_signature = None
        labels = {
            'es': ['Seleccionar carpeta', 'Nombre científico (ejemplo: Lynx pardinus)', 'Incluir subcarpetas',
                   'Identificación del individuo (si consta en los nombres)', 'Año para fotos sin fecha EXIF (opcional)',
                   'Configurar Excel', 'Las fechas salen del EXIF. No se usa la fecha de copia del archivo. Sin EXIF ni año, la foto se informa como omitida.'],
            'pt': ['Selecionar pasta', 'Nome científico (exemplo: Lynx pardinus)', 'Incluir subpastas',
                   'Identificação do indivíduo (se constar nos nomes)', 'Ano para fotos sem data EXIF (opcional)',
                   'Configurar Excel', 'Datas obtidas do EXIF. A data de cópia não é utilizada. Sem EXIF nem ano, a foto é indicada como omitida.'],
            'en': ['Choose folder', 'Scientific name (example: Lynx pardinus)', 'Include subfolders',
                   'Individual identification (if recorded in names)', 'Year for photos without EXIF date (optional)',
                   'Configure Excel', 'Dates come from EXIF, never file copy time. Photos with neither EXIF nor a supplied year are reported as omitted.']}[self.lang]
        ctk.CTkButton(self, text=labels[0], command=self.choose).pack(pady=10)
        self.path_label = ctk.CTkLabel(self, text=self.folder, wraplength=720)
        self.path_label.pack()
        ctk.CTkLabel(self, text=labels[1]).pack()
        self.species = ctk.CTkEntry(self, width=330)
        self.species.pack()
        self.recursive = ctk.CTkCheckBox(self, text=labels[2])
        self.recursive.pack(pady=8)
        if catalog:
            self.recursive.select()
        ctk.CTkLabel(self, text=labels[3]).pack()
        self.identity_labels = dict(zip({'es': ['No asignar individuo', 'Primera palabra del archivo', 'Nombre de la subcarpeta'], 'pt': ['Não atribuir indivíduo', 'Primeira palavra do ficheiro', 'Nome da subpasta'], 'en': ['Do not assign an individual', 'First word of filename', 'Subfolder name']}[self.lang], ['none', 'filename', 'folder']))
        self.identity = StableOptionMenu(self, values=list(self.identity_labels), width=300)
        self.identity.pack()
        ctk.CTkLabel(self, text=labels[4]).pack()
        self.year = ctk.CTkEntry(self)
        self.year.pack()
        ctk.CTkLabel(self, text=({'es': 'Catálogo: las fotos sin fecha se exportan con los campos temporales vacíos.', 'pt': 'Catálogo: fotos sem data são exportadas com campos temporais vazios.', 'en': 'Catalog: undated photos are exported with empty date and time fields.'}[self.lang] if catalog else labels[6]), wraplength=700).pack(pady=10)
        self.structured = ctk.CTkCheckBox(self, text={'es': 'Interpretar los niveles de las subcarpetas', 'pt': 'Interpretar os níveis das subpastas', 'en': 'Map subfolder levels'}[self.lang], command=self.toggle_structure)
        self.structured.pack(pady=5)
        self.structure_frame = ctk.CTkFrame(self)
        self.levels = {}
        descriptions = {'es': ['Localidad (verbatimLocality)', 'Estación de fototrampeo', 'Nombre científico', 'Individuo'],
                        'pt': ['Localidade (verbatimLocality)', 'Estação de armadilhagem', 'Nome científico', 'Indivíduo'],
                        'en': ['Locality (verbatimLocality)', 'Camera trap station', 'Scientific name', 'Individual']}[self.lang]
        ctk.CTkLabel(self.structure_frame, text={'es': 'Nivel 1 = primera subcarpeta bajo la raíz. — = no usar. Ejemplo: Localidad/Estación/Individuo/foto.jpg', 'pt': 'Nível 1 = primeira subpasta na raiz. — = não usar. Exemplo: Localidade/Estação/Indivíduo/foto.jpg', 'en': 'Level 1 = first subfolder below the root. — = unused. Example: Locality/Station/Individual/photo.jpg'}[self.lang], wraplength=850).grid(row=0, column=0, columnspan=4, padx=6)
        for index, (key, label) in enumerate(zip(('locality', 'station', 'species', 'individual'), descriptions)):
            ctk.CTkLabel(self.structure_frame, text=label).grid(row=1, column=index, padx=6)
            menu = self.levels[key] = StableOptionMenu(self.structure_frame, values=['—'] + [str(i) for i in range(1, 21)], width=145, command=lambda _: self.invalidate_layout())
            menu.set('1' if key in ('locality', 'station') else '—')
            menu.grid(row=2, column=index, padx=6, pady=5)
        self.station_button = ctk.CTkButton(self.structure_frame, text={'es': 'Revisar subcarpetas y estaciones', 'pt': 'Rever subpastas e estações', 'en': 'Review subfolders and stations'}[self.lang], command=self.review_stations)
        self.station_button.grid(row=3, column=0, columnspan=4, pady=6)
        self.layout_status = ctk.CTkLabel(self, text='', wraplength=850)
        self.layout_status.pack()
        self.configure_button = ctk.CTkButton(self, text=labels[5], command=self.configure_bulk)
        self.configure_button.pack(pady=10)
        foreground(self, self.root)

    @action
    def choose(self):
        path = dialogs.askdirectory(parent=self)
        if path:
            self.folder = path
            self.stations = {}
            self.invalidate_layout()
            self.path_label.configure(text=path)

    @action
    def configure_bulk(self):
        if not self.folder:
            raise ValueError('Selecciona una carpeta.')
        species = self.species.get().strip()
        levels = self.level_values() if self.structured.get() else None
        if not (levels and levels['species']) and len(species.split()) < 2:
            raise ValueError('Escribe el nombre científico: género y epíteto específico.')
        if levels and self.layout_signature != (self.folder, levels):
            raise ValueError('Revisa y guarda la tabla de estaciones con los niveles actuales antes de configurar el Excel.')
        year = int(self.year.get()) if self.year.get().strip() else None
        if year is not None and not 1 <= year <= 9999:
            raise ValueError('Año inválido.')
        folder, recursive = self.folder, bool(self.recursive.get())
        identity = self.identity_labels[self.identity.get()]
        stations, catalog = dict(self.stations), self.catalog
        self.open_editor(lambda task, options: folder_rows(task, folder, species, recursive, identity, year, options[0], options[1], allow_undated=catalog, levels=levels, stations=stations))

    def level_values(self):
        return {key: int(menu.get()) if menu.get() != '—' else 0 for key, menu in self.levels.items()}

    def invalidate_layout(self):
        self.layout_signature = None
        self.layout_status.configure(text={'es': 'Revisa las estaciones después de cambiar niveles o carpeta.', 'pt': 'Reveja as estações após alterar níveis ou pasta.', 'en': 'Review stations after changing levels or folder.'}[self.lang])

    def toggle_structure(self):
        if self.structured.get():
            self.recursive.select()
            self.structure_frame.pack(fill='x', padx=10, before=self.layout_status)
        else:
            self.structure_frame.pack_forget()

    @action
    def review_stations(self):
        if not self.folder:
            raise ValueError('Selecciona una carpeta.')
        folder, levels = self.folder, self.level_values()
        def receive(result):
            if not result['count']:
                from tkinter import messagebox
                messagebox.showerror('Subcarpetas', 'No hay fotografías en la carpeta.', parent=self.winfo_toplevel())
                return
            if result['errors']:
                from tkinter import messagebox
                messagebox.showerror('Subcarpetas', '\n'.join(result['errors'][:10]), parent=self.winfo_toplevel())
                return
            self.station_dialog(result, folder, levels)
        start(self, 'Estaciones / Stations', lambda task: scan_layout(task, folder, levels), receive)

    def station_dialog(self, result, folder, levels):
        window = self._station_window = ctk.CTkToplevel(self)
        window.title({'es': 'Estaciones y coordenadas', 'pt': 'Estações e coordenadas', 'en': 'Stations and coordinates'}[self.lang])
        window.geometry('960x600')
        ctk.CTkLabel(window, text={'es': 'Coordenadas en grados decimales WGS84. Completa ambas o deja ambas vacías. La localidad se exporta como verbatimLocality.', 'pt': 'Coordenadas em graus decimais WGS84. Preencha ambas ou deixe ambas vazias. A localidade é exportada como verbatimLocality.', 'en': 'WGS84 decimal degrees. Fill both coordinates or leave both blank. Locality exports as verbatimLocality.'}[self.lang], wraplength=900).pack(padx=10, pady=6)
        ctk.CTkLabel(window, text='\n'.join(result['examples'][:4]), wraplength=900, justify='left').pack()
        table = ctk.CTkScrollableFrame(window)
        table.pack(fill='both', expand=True, padx=10, pady=5)
        for col, label in enumerate(('stationPath', 'verbatimLocality', 'latitude', 'longitude', 'Fotos')):
            ctk.CTkLabel(table, text=label).grid(row=0, column=col, padx=6)
        entries = self.station_entries = {}
        for index, (key, data) in enumerate(sorted(result['stations'].items()), 1):
            previous = self.stations.get(key, {}) if self.layout_signature == (folder, levels) else {}
            values = dict(data, **previous)
            ctk.CTkLabel(table, text=key, width=250, wraplength=250).grid(row=index, column=0, padx=5, pady=4)
            controls = {}
            for col, name in enumerate(('locality', 'latitude', 'longitude'), 1):
                entry = controls[name] = ctk.CTkEntry(table, width=220 if col == 1 else 130)
                entry.insert(0, str(values.get(name, '')))
                entry.grid(row=index, column=col, padx=5, pady=4)
            ctk.CTkLabel(table, text=str(data['count'])).grid(row=index, column=4, padx=5)
            entries[key] = controls
        @action
        def apply(owner):
            if owner.folder != folder or owner.level_values() != levels:
                raise ValueError('La carpeta o los niveles han cambiado. Vuelve a revisar las estaciones.')
            owner.stations = validate_stations({key: dict(result['stations'][key], **{name: entry.get() for name, entry in controls.items()}) for key, controls in entries.items()})
            owner.layout_signature = (folder, levels)
            owner.layout_status.configure(text=f"{len(entries)} " + {'es': 'estaciones configuradas', 'pt': 'estações configuradas', 'en': 'stations configured'}[owner.lang])
            window.destroy()
        ctk.CTkButton(window, text={'es': 'Guardar estaciones', 'pt': 'Guardar estações', 'en': 'Save stations'}[self.lang], command=lambda: apply(self)).pack(pady=8)
        foreground(window, self.winfo_toplevel())
