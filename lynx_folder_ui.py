"""Configure folder/catalog inputs; export uses the same editor as WI and DP."""
import customtkinter as ctk
import lynx_dialogs as dialogs
from lynx_choices import StableOptionMenu
from lynx_windows import foreground, WorkflowPanel
from lynx_ui_jobs import action
from lynx_bulk_ui import BulkEditor
from lynx_folder_bulk import folder_rows


class FolderInput(WorkflowPanel):
    def __init__(self, owner, catalog=False):
        super().__init__(owner)
        self.root, self.lang = owner.root, owner.lang
        self.catalog = catalog
        self.folder = getattr(owner, 'folder_path', '') or ''
        self.title('Wildbook Bulk Import · ' + {'es': ['Crear desde carpeta', 'Catálogo'], 'pt': ['Criar a partir de pasta', 'Catálogo'], 'en': ['Create from folder', 'Catalog']}[self.lang][int(catalog)])
        self.geometry('760x490')
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
        ctk.CTkButton(self, text=labels[5], command=self.configure_bulk).pack(pady=10)
        foreground(self, self.root)

    @action
    def choose(self):
        path = dialogs.askdirectory(parent=self)
        if path:
            self.folder = path
            self.path_label.configure(text=path)

    @action
    def configure_bulk(self):
        if not self.folder:
            raise ValueError('Selecciona una carpeta.')
        species = self.species.get().strip()
        if len(species.split()) < 2:
            raise ValueError('Escribe género y especie.')
        year = int(self.year.get()) if self.year.get().strip() else None
        if year is not None and not 1 <= year <= 9999:
            raise ValueError('Año inválido.')
        folder, recursive = self.folder, bool(self.recursive.get())
        identity = self.identity_labels[self.identity.get()]
        self.open_editor(lambda task, options: folder_rows(task, folder, species, recursive, identity, year, options[0], options[1], allow_undated=self.catalog))
