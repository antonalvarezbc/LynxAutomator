"""Own dialogs without delayed focus stealing or competing modal grabs."""
def foreground(window, parent, modal=False):
    window.transient(parent.winfo_toplevel())
    window.lift()


import customtkinter as ctk


class WorkflowPanel(ctk.CTkFrame):
    """A workflow page, embedded when a host is supplied, otherwise in a dialog."""
    def __init__(self, owner, force_window=False):
        self.workflow = None if force_window else getattr(owner, 'workflow', None)
        self.window = None if self.workflow else ctk.CTkToplevel(owner.root)
        super().__init__(self.workflow.body if self.workflow else self.window)
        self.pack(fill='both', expand=True)

    def title(self, text):
        if self.window:
            self.window.title(text)

    def geometry(self, value):
        if self.window:
            self.window.geometry(value)

    def protocol(self, name, callback):
        if self.window:
            self.window.protocol(name, callback)

    def transient(self, parent=None):
        if self.window:
            return self.window.transient(parent)

    def destroy(self):
        if self.window and self.window.winfo_exists():
            window, self.window = self.window, None
            window.destroy()
        else:
            super().destroy()

    def open_editor(self, loader):
        from lynx_bulk_ui import BulkEditor
        if self.workflow:
            self.workflow.edit(loader)
        else:
            BulkEditor(self, loader)
