"""Single entry point for all four Bulk Import sources."""
import customtkinter as ctk
from lynx_choices import StableOptionMenu
from lynx_ui_jobs import action


class BulkWorkflow(ctk.CTkFrame):
    def __init__(self, root, lang='es'):
        super().__init__(root)
        self.root, self.lang, self.workflow = root, lang, self
        self.pack(fill='both', expand=True)
        labels = {'es': ['Carpeta', 'Catálogo', 'Wildlife Insights', 'Camtrap DP'],
                  'pt': ['Pasta', 'Catálogo', 'Wildlife Insights', 'Camtrap DP'],
                  'en': ['Folder', 'Catalog', 'Wildlife Insights', 'Camtrap DP']}[lang]
        self.sources = labels
        bar = ctk.CTkFrame(self)
        bar.pack(fill='x', padx=10, pady=8)
        ctk.CTkLabel(bar, text={'es': 'Origen de los datos', 'pt': 'Origem dos dados', 'en': 'Data source'}[lang]).pack(side='left', padx=8)
        self.choice = StableOptionMenu(bar, values=labels, command=self.select)
        self.choice.pack(side='left', padx=8)
        ctk.CTkButton(bar, text={'es': 'Volver al origen', 'pt': 'Voltar à origem', 'en': 'Back to source'}[lang], command=self.back).pack(side='left', padx=8)
        self.body = ctk.CTkFrame(self)
        self.body.pack(fill='both', expand=True)
        self.pages = {}
        self.editor = None
        self.select(labels[0])

    def hide_pages(self):
        for child in self.body.winfo_children():
            child.pack_forget()

    @action
    def select(self, value):
        self.hide_pages()
        if self.editor:
            self.editor.destroy()
            self.editor = None
        if value not in self.pages:
            index = self.sources.index(value)
            if index < 2:
                from lynx_folder_ui import FolderInput
                page = FolderInput(self, catalog=index == 1)
            elif index == 2:
                from lynx_wi_ui import WITab
                page = WITab(self.body, self.lang)
                page.workflow = self
            else:
                from lynx_camtrap_ui import CamtrapTab
                page = CamtrapTab(self.body, self.lang)
                page.workflow = self
            self.pages[value] = page
        self.pages[value].pack(fill='both', expand=True)

    @action
    def back(self):
        self.select(self.choice.get())

    def edit(self, loader):
        from lynx_bulk_ui import BulkEditor
        self.hide_pages()
        if self.editor:
            self.editor.destroy()
        self.editor = BulkEditor(self, loader)
