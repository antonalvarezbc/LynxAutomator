"""Own dialogs without delayed focus stealing or competing modal grabs."""
def foreground(window, parent, modal=False):
    window.transient(parent.winfo_toplevel())
    window.lift()
