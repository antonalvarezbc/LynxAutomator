"""Keep task dialogs above their owning window, without permanent topmost state."""
def foreground(window, parent, modal=False):
    window.transient(parent.winfo_toplevel())
    def show():
        if window.winfo_exists():
            window.lift()
            window.focus_force()
            if modal:
                window.grab_set()
    window.after(180, show)
