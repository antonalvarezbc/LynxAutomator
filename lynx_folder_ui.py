"""Configure folder/catalog inputs; export uses the same editor as WI and DP."""
import customtkinter as ctk
import lynx_dialogs as dialogs
from lynx_choices import StableOptionMenu
from lynx_windows import foreground
from lynx_ui_jobs import action
from lynx_bulk_ui import BulkEditor
from lynx_folder_bulk import folder_rows


class FolderInput(ctk.CTkToplevel):
    def __init__(self, owner, catalog=False):
        super().__init__(owner.root)
        self.root, self.lang = owner.root, owner.lang
        self.folder = getattr(owner, 'folder_path', '') or ''
        self.title('Wildbook Bulk Import · ' + ('Catalog' if catalog else 'Folder'))
        self.geometry('760x490')
        labels = {
            'es': ['Seleccionar carpeta', 'Especie científica (ejemplo: Lynx pardinus)', 'Incluir subcarpetas',
                   'Identidad (sólo si los nombres identifican animales)', 'Año para fotos sin fecha EXIF (opcional)',
                   'Configurar Bulk Import', 'Las fechas salen del EXIF. No se usa la fecha de copia del archivo. Sin EXIF ni año, la foto se informa como omitida.'],
            'pt': ['Selecionar pasta', 'Espécie científica (exemplo: Lynx pardinus)', 'Incluir subpastas',
                   'Identidade (apenas se os nomes identificam animais)', 'Ano para fotos sem data EXIF (opcional)',
                   'Configurar Bulk Import', 'Datas obtidas do EXIF. A data de cópia não é utilizada. Sem EXIF nem ano, a foto é indicada como omitida.'],
            'en': ['Choose folder', 'Scientific species (example: Lynx pardinus)', 'Include subfolders',
                   'Identity (only if names identify animals)', 'Year for photos without EXIF date (optional)',
                   'Configure Bulk Import', 'Dates come from EXIF, never file copy time. Photos with neither EXIF nor a supplied year are reported as omitted.']}[self.lang]
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
        self.identity = StableOptionMenu(self, values=['none', 'filename: first word', 'folder name'], width=270)
        self.identity.pack()
        ctk.CTkLabel(self, text=labels[4]).pack()
        self.year = ctk.CTkEntry(self)
        self.year.pack()
        ctk.CTkLabel(self, text=labels[6], wraplength=700).pack(pady=10)
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
        identity = {'none': 'none', 'filename: first word': 'filename', 'folder name': 'folder'}[self.identity.get()]
        BulkEditor(self, lambda task, options: folder_rows(task, folder, species, recursive, identity, year, options[0], options[1]))
        self.destroy()
