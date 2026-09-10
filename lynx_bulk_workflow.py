"""Single entry point for all Bulk Import sources."""
import customtkinter as ctk
from lynx_choices import StableOptionMenu
from lynx_ui_jobs import action


class BulkWorkflow(ctk.CTkFrame):
    def __init__(self, root, lang='es'):
        super().__init__(root)
        self.root, self.lang, self.workflow = root, lang, self
        self.pack(fill='both', expand=True)
        labels = {'es': ['Crear desde carpeta', 'Catálogo', 'Wildlife Insights', 'Camtrap DP', 'Agouti API (alpha)', 'Trapper API (alpha)'],
                  'pt': ['Criar a partir de pasta', 'Catálogo', 'Wildlife Insights', 'Camtrap DP', 'Agouti API (alpha)', 'Trapper API (alpha)'],
                  'en': ['Create from folder', 'Catalog', 'Wildlife Insights', 'Camtrap DP', 'Agouti API (alpha)', 'Trapper API (alpha)']}[lang]
        self.sources = labels
        bar = ctk.CTkFrame(self)
        bar.pack(fill='x', padx=10, pady=8)
        ctk.CTkLabel(bar, text={'es': 'Origen de los datos', 'pt': 'Origem dos dados', 'en': 'Data source'}[lang]).pack(side='left', padx=8)
        self.choice = StableOptionMenu(bar, values=labels, command=self.select, width=240)
        self.choice.pack(side='left', padx=8)
        self.body = ctk.CTkFrame(self)
        self.body.pack(fill='both', expand=True)
        self.pages = {}
        self.editor = None
        self.choice.set('Wildlife Insights')
        self.select('Wildlife Insights')

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
            elif index == 3:
                from lynx_camtrap_ui import CamtrapTab
                page = CamtrapTab(self.body, self.lang)
                page.workflow = self
            else:
                from lynx_api_ui import APIImportTab
                page = APIImportTab(self.body, ('Agouti API', 'Trapper API')[index - 4], self.lang)
                page.workflow = self
            self.pages[value] = page
        self.pages[value].pack(fill='both', expand=True)

    def edit(self, loader):
        from lynx_bulk_ui import BulkEditor
        self.hide_pages()
        if self.editor:
            self.editor.destroy()
        self.editor = BulkEditor(self, loader)
