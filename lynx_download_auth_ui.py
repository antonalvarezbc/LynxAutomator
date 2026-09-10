"""Download access dialog shared by both import sources."""
import webbrowser
import customtkinter as ctk
from lynx_choices import StableOptionMenu
from lynx_download_auth import DownloadAuth, google_login, origin
from lynx_ui_jobs import action, start
from lynx_windows import foreground

WORDS = {
    'es': ['Autorización (opcional)', 'Servidor HTTPS', 'Clave API / token', 'Usar para esta sesión', 'Quitar acceso guardado en memoria', 'Abrir web / ayuda', 'Iniciar sesión con Google', 'Acceso configurado para esta sesión. Los permisos se comprobarán al descargar.', 'Google ha completado el inicio de sesión. Puedes reintentar la descarga.'],
    'pt': ['Autorização (opcional)', 'Servidor HTTPS', 'Chave API / token', 'Usar nesta sessão', 'Remover acesso da memória', 'Abrir web / ajuda', 'Iniciar sessão com Google', 'Acesso configurado nesta sessão. As permissões serão verificadas na descarga.', 'Google concluiu o início de sessão. Pode repetir a descarga.'],
    'en': ['Authorization (optional)', 'HTTPS server', 'API key / token', 'Use for this session', 'Clear access from memory', 'Open website / help', 'Sign in with Google', 'Access configured for this session. Permissions will be checked when downloading.', 'Google sign-in completed. You can retry the download.']}


class DownloadAccess(ctk.CTkToplevel):
    def __init__(self, owner, google=False):
        super().__init__(owner.root)
        self.root, self.lang, self.owner = owner.root, owner.lang, owner
        self.words = WORDS[self.lang]
        self.title(self.words[0])
        self.geometry('680x480')
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.google = google
        if google:
            text = {'es': 'Inicia sesión en Google desde el navegador. Google Cloud CLI administra las credenciales que usa gsutil; necesitas una cuenta con permiso para las fotos de WI.',
                    'pt': 'Inicie sessão no navegador. Google Cloud CLI gere as credenciais usadas por gsutil; a conta precisa de acesso às fotos WI.',
                    'en': 'Sign in through your browser. Google Cloud CLI manages the credentials used by gsutil; your account needs access to the WI photos.'}[self.lang]
            ctk.CTkLabel(self, text=text, wraplength=620).pack(pady=20)
            ctk.CTkButton(self, text=self.words[6], command=self.login_google).pack(pady=10)
        else:
            self.method = StableOptionMenu(self, values=['Agouti API key', 'Agouti Bearer', 'Trapper token'], command=self.changed)
            self.method.pack(pady=12)
            ctk.CTkLabel(self, text=self.words[1]).pack()
            self.server = ctk.CTkEntry(self, width=580, placeholder_text='https://…')
            self.server.pack(pady=6)
            ctk.CTkLabel(self, text=self.words[2]).pack()
            self.secret = ctk.CTkEntry(self, width=580, show='•')
            self.secret.pack(pady=6)
            text = {'es': 'Agouti: usa una API key o un token Bearer. Trapper: inicia sesión en tu servidor y genera un token en tu perfil. La clave sólo se envía al servidor indicado y permanece en memoria hasta cerrar la aplicación o quitar el acceso.',
                    'pt': 'Agouti: use uma API key ou token Bearer. Trapper: inicie sessão no servidor e gere um token no perfil. A chave só é enviada ao servidor indicado e permanece apenas na memória.',
                    'en': 'Agouti: use an API key or Bearer token. Trapper: sign in on your server and generate a token in your profile. The credential is sent only to the specified server and stays in memory until the app closes or access is cleared.'}[self.lang]
            ctk.CTkLabel(self, text=text, wraplength=620).pack(pady=10)
            ctk.CTkButton(self, text=self.words[3], command=self.apply).pack(pady=6)
            ctk.CTkButton(self, text=self.words[4], command=self.clear).pack(pady=6)
            self.changed(self.method.get())
            if owner.download_auth:
                self.method.set(owner.download_auth.method)
                self.server.delete(0, 'end')
                self.server.insert(0, owner.download_auth.server)
        ctk.CTkButton(self, text=self.words[5], command=self.help).pack(pady=10)
        self.status = ctk.CTkLabel(self, text='', wraplength=620)
        self.status.pack(pady=8)
        provider = getattr(owner, 'provider', '')
        if not google and provider:
            self.method.configure(values=['Agouti API key', 'Agouti Bearer'] if provider == 'Agouti API' else ['Trapper token'])
            if not owner.download_auth:
                self.method.set('Agouti API key' if provider == 'Agouti API' else 'Trapper token')
                self.server.delete(0, 'end')
                self.server.insert(0, owner.server.get())
        foreground(self, self.root)

    def changed(self, method):
        self.secret.delete(0, 'end')
        self.server.delete(0, 'end')
        if getattr(self.owner, 'provider', ''):
            self.server.insert(0, self.owner.server.get())
        elif method.startswith('Agouti'):
            self.server.insert(0, 'https://api.agouti.eu')

    @action
    def apply(self):
        credential = DownloadAuth(self.server.get().strip(), self.method.get(), self.secret.get().strip())
        self.owner.download_auth = credential
        self.secret.delete(0, 'end')
        self.status.configure(text=self.words[7])

    @action
    def clear(self):
        self.owner.download_auth = None
        self.secret.delete(0, 'end')
        self.status.configure(text=self.words[4])

    @action
    def help(self):
        if self.google:
            url = 'https://cloud.google.com/sdk/docs/authorizing'
        elif self.method.get().startswith('Agouti'):
            url = 'https://docs.agouti.eu/api/endpoints.html'
        else:
            url = self.server.get().strip()
            if origin(url)[0] != 'https':
                raise ValueError(self.words[1])
        webbrowser.open(url)

    @action
    def login_google(self):
        start(self, 'Google', google_login, lambda _: self.status.configure(text=self.words[8]))

    def close(self):
        if not self.root.winfo_toplevel().jobs.busy:
            self.destroy()
