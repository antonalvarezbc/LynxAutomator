"""Persistent searchable choices; avoids transient Tk.Menu focus/grab conflicts."""
import customtkinter as ctk
from tkinter import ttk


class ChoicePopup(ctk.CTkToplevel):
    def __init__(self, owner):
        super().__init__(owner.winfo_toplevel())
        self.owner = owner
        self.title('Seleccionar / Select')
        self.transient(owner.winfo_toplevel())
        self.geometry('520x380')
        self.search = ctk.CTkEntry(self, placeholder_text='Buscar / Search')
        self.search.pack(fill='x', padx=8, pady=8)
        self.search.bind('<KeyRelease>', lambda _: self.filter())
        frame = ctk.CTkFrame(self)
        frame.pack(fill='both', expand=True, padx=8)
        self.list = ttk.Treeview(frame, columns=('value',), show='tree', height=10, selectmode='browse')
        scroll = ttk.Scrollbar(frame, command=self.list.yview)
        scroll.pack(side='right', fill='y')
        self.list.configure(yscrollcommand=scroll.set)
        self.list.pack(fill='both', expand=True)
        self.list.bind('<ButtonRelease-1>', self.choose)
        self.list.bind('<Return>', self.choose)
        self.bind('<Escape>', lambda _: self.destroy())
        self.filter()
        self.lift()

    def filter(self):
        self.list.delete(*self.list.get_children())
        for i, value in enumerate(self.owner.cget('values')):
            if self.search.get().casefold() in value.casefold():
                self.list.insert('', 'end', iid=str(i), text=value)

    def choose(self, event=None):
        if not self.owner.winfo_exists():
            self.destroy()
            return
        if self.owner.cget('state') == 'disabled':
            return
        selected = self.list.selection()
        if selected:
            value = self.owner.cget('values')[int(selected[0])]
            owner = self.owner
            self.destroy()
            owner._dropdown_callback(value)


class StableChoices:
    def _open_dropdown_menu(self):
        popup = getattr(self, '_choice_popup', None)
        if popup is not None and popup.winfo_exists():
            popup.lift()
        else:
            self._choice_popup = ChoicePopup(self)


class StableComboBox(StableChoices, ctk.CTkComboBox):
    pass


class StableOptionMenu(StableChoices, ctk.CTkOptionMenu):
    pass
