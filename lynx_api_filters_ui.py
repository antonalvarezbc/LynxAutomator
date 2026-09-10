"""Select actual deployment metadata values, never guess server-side filters."""
import customtkinter as ctk
from lynx_choices import StableOptionMenu
from lynx_api_filters import deployment_facets, match_facets
from lynx_windows import foreground
from lynx_ui_jobs import action


class DeploymentFilterDialog(ctk.CTkToplevel):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner, self.root, self.lang = owner, owner.root, owner.lang
        self.title({'es': 'Seleccionar datos del proyecto', 'pt': 'Selecionar dados do projeto', 'en': 'Select project data'}[self.lang])
        self.geometry('900x640')
        self.rows = owner.index['deployments']
        self.facets = deployment_facets(self.rows)
        self.choices = {key: set(values) for key, values in owner.facet_choices.items() if key in self.facets}
        self.checks = {}
        text = ({'es': 'Agouti: sólo se han consultado los despliegues. Al continuar se pedirán media y observaciones de la selección.', 'pt': 'Agouti: só foram consultadas as instalações. Ao continuar serão pedidos media e observações da seleção.', 'en': 'Agouti: only deployments have been read. Continuing requests media and observations for the selection.'} if owner.provider == 'Agouti API' else {'es': 'Trapper: ya se descargó el ZIP de metadatos para conocer los valores. Estos filtros se aplican localmente y no vuelven a descargarlo.', 'pt': 'Trapper: o ZIP de metadados já foi descarregado para conhecer os valores. Estes filtros aplicam-se localmente sem o descarregar novamente.', 'en': 'Trapper: the metadata ZIP has been downloaded to discover values. These filters apply locally without downloading it again.'})[self.lang]
        ctk.CTkLabel(self, text=text + '\n' + {'es': 'Todavía no se descargan fotografías. Años = inicio del despliegue.', 'pt': 'Ainda não se descarregam fotografias. Anos = início da instalação.', 'en': 'No photographs are downloaded yet. Years refer to deployment start.'}[self.lang], wraplength=850, justify='left').pack(padx=12, pady=8)
        ctk.CTkLabel(self, text={'es': 'Marca varios valores por variable. Sin marcas = cualquier valor. Se combinan las variables elegidas.', 'pt': 'Marque vários valores por variável. Sem marcas = qualquer valor. Combinam-se as variáveis escolhidas.', 'en': 'Select multiple values per variable. No selection = any value. Selected variables are combined.'}[self.lang], wraplength=850).pack()
        bar = ctk.CTkFrame(self)
        bar.pack(fill='x', padx=12, pady=5)
        self.variable = StableOptionMenu(bar, values=list(self.facets) or ['deploymentID'], width=240, command=self.show_variable)
        self.variable.pack(side='left', padx=5)
        self.search = ctk.CTkEntry(bar, placeholder_text={'es': 'Buscar valores', 'pt': 'Pesquisar valores', 'en': 'Search values'}[self.lang], width=320)
        self.search.pack(side='left', padx=5)
        self.search.bind('<KeyRelease>', lambda _: self.filter_values())
        self.listing = ctk.CTkScrollableFrame(self)
        self.listing.pack(fill='both', expand=True, padx=12)
        controls = ctk.CTkFrame(self)
        controls.pack(fill='x', padx=12, pady=5)
        for label, command in [({'es': 'Marcar visibles', 'pt': 'Marcar visíveis', 'en': 'Select visible'}[self.lang], self.select_visible),
                               ({'es': 'Cualquier valor', 'pt': 'Qualquer valor', 'en': 'Any value'}[self.lang], self.clear_variable),
                               ({'es': 'Limpiar filtros', 'pt': 'Limpar filtros', 'en': 'Clear filters'}[self.lang], self.clear_all)]:
            ctk.CTkButton(controls, text=label, command=command).pack(side='left', padx=5)
        self.summary = ctk.CTkLabel(self, text='', wraplength=850)
        self.summary.pack(fill='x', padx=12, pady=6)
        self.apply_button = ctk.CTkButton(self, text={'es': 'Cargar selección de datos', 'pt': 'Carregar seleção de dados', 'en': 'Load selected data'}[self.lang], command=self.apply)
        self.apply_button.pack(pady=8)
        self.show_variable(self.variable.get())
        foreground(self, owner.winfo_toplevel())

    def show_variable(self, key):
        for widget in self.listing.winfo_children():
            widget.destroy()
        self.checks = {}
        self.search.delete(0, 'end')
        for value, count in self.facets.get(key, {}).items():
            check = self.checks[value] = ctk.CTkCheckBox(self.listing, text=f'{value} ({count})', command=self.changed)
            if value in self.choices.get(key, set()):
                check.select()
            check.pack(anchor='w', padx=6, pady=3)
        self.update_summary()

    def filter_values(self):
        query = self.search.get().casefold()
        for value, check in self.checks.items():
            check.pack_forget()
            if query in value.casefold():
                check.pack(anchor='w', padx=6, pady=3)

    def changed(self):
        self.choices[self.variable.get()] = {value for value, check in self.checks.items() if check.get()}
        self.update_summary()

    def update_summary(self):
        selected = match_facets(self.rows, self.choices)
        labels = [f'{key}: ' + ', '.join(sorted(values)) for key, values in self.choices.items() if values]
        self.summary.configure(text=f'{len(selected)} / {len(self.rows)} ' + {'es': 'despliegues', 'pt': 'instalações', 'en': 'deployments'}[self.lang] + '\n' + ' · '.join(labels))
        self.apply_button.configure(state='normal' if selected else 'disabled')

    def select_visible(self):
        for value, check in self.checks.items():
            if self.search.get().casefold() in value.casefold():
                check.select()
        self.changed()

    def clear_variable(self):
        for check in self.checks.values():
            check.deselect()
        self.changed()

    def clear_all(self):
        self.choices = {}
        self.show_variable(self.variable.get())

    @action
    def apply(self):
        ids = [row['deploymentID'] for row in match_facets(self.rows, self.choices)]
        if not ids:
            raise ValueError('La selección no contiene despliegues.')
        self.owner.facet_choices = {key: set(values) for key, values in self.choices.items()}
        owner = self.owner
        self.destroy()
        owner.apply_selection(ids)
