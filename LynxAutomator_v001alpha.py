from lynx_ui_jobs import (JobPanel, FolderJobs, CatalogJobs, WIJobs, LynxJobs, VideoJobs, DateJobs, RenamerJobs, DownloadJobs)
import customtkinter as ctk
from tkinter import filedialog, messagebox, IntVar, StringVar, ttk, BooleanVar
import pandas as pd
import os
from PIL import Image, ImageTk, ImageFile
from PIL.ExifTags import TAGS
from datetime import datetime
import tempfile
import cv2
import platform
from lynx_core import file_timestamp, set_file_timestamp, shift_file_date, frame_step, unique_path, merge_deployments
import tkinter as tk
import threading
import queue
import math
import subprocess
import shutil
import piexif
import webbrowser
import re
import shutil
from pathlib import Path
import customtkinter as ctk  
import tkinter as tk
from tkinter import ttk
import sys


class BaseApp:
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")
        ctk.set_default_color_theme("green")  # Color theme ("blue", "green", "dark-blue")
        
        self.translations = {
            "es": {
                "title": "LynxAutomator",
                "description": "Esta aplicación te permite automatizar ciertos procesos en el monitoreo del Lince Ibérico y otras especies",
                "about": "Acerca de",
                "wildbook": "Wildbook",
                "wildlife_insights": "Wildlife Insights",
                "wi_wildbook": "Wildlife Insight - Wildbook",
                "iberian_lynx": "Lince Ibérico",
                "functionalities": "Funcionalidades",
                "presentation": "Presentación",
                "date_changer": "Cambiador de Fecha",
                "video_frame_extractor": "Extractor de Fotogramas de Video",
                "images_renamer": "Renombrador de Imágenes",
                "wiwbe_folder": "BIWbE desde Carpeta",
                "wiwbe_catalog": "Catálogo BIWbE",
                "wi_downloader": "Descargador WI",
                "lynx_feature_1": "Función Lince Ibérico",
                "wi_csvs_to_biwbe": "WI CSVs a BIWbE"
            },
            "pt": {
                "title": "LynxAutomator",
                "description": "Este aplicativo permite automatizar certos processos no monitoramento do Lince Ibérico e outras espécies",
                "about": "Sobre",
                "wildbook": "Wildbook",
                "wildlife_insights": "Wildlife Insights",
                "wi_wildbook": "Wildlife Insight - Wildbook",
                "iberian_lynx": "Lince Ibérico",
                "functionalities": "Funcionalidades",
                "presentation": "Apresentação",
                "date_changer": "Mudança de Data",
                "video_frame_extractor": "Extractor de Quadros de Vídeo",
                "images_renamer": "Renomeador de Imagens",
                "wiwbe_folder": "BIWbE da Pasta",
                "wiwbe_catalog": "Catálogo BIWbE",
                "wi_downloader": "Downloader WI",
                "lynx_feature_1": "Recurso Lince Ibérico",
                "wi_csvs_to_biwbe": "WI CSVs para BIWbE"
            },
            "en": {
                "title": "LynxAutomator",
                "description": "This application allows you to automate certain processes in the monitoring of the Iberian Lynx and other species",
                "about": "About",
                "wildbook": "Wildbook",
                "wildlife_insights": "Wildlife Insights",
                "wi_wildbook": "Wildlife Insight - Wildbook",
                "iberian_lynx": "Iberian Lynx",
                "functionalities": "Functionalities",
                "presentation": "Presentation",
                "date_changer": "Date Changer",
                "video_frame_extractor": "Video Frame Extractor",
                "images_renamer": "Images Renamer",
                "wiwbe_folder": "BIWbE from Folder",
                "wiwbe_catalog": "BIWbE Catalog",
                "wi_downloader": "WI Downloader",
                "lynx_feature_1": "Iberian Lynx Feature",
                "wi_csvs_to_biwbe": "WI CSVs to BIWbE"
            }
        }

        # Variable para almacenar el idioma seleccionado
        self.language = StringVar(value="Español")  # Valor predeterminado: Español

        # Configurar la interfaz inicial
        self.setup_ui()

    def setup_ui(self):
        # Obtener las traducciones según el idioma seleccionado
        lang = "es" if self.language.get() == "Español" else "pt" if self.language.get() == "Português" else "en"
        tr = self.translations[lang]

        self.root.jobs = JobPanel(self.root, lang)

        # Actualizar el título de la ventana
        self.root.title(tr["title"])

        # Crear el encabezado con el texto grande en el centro
        self.header_frame = ctk.CTkFrame(self.root)
        self.header_frame.pack(pady=10, fill='x')

        # Mover el menú de selección de idioma a la esquina superior derecha
        self.language_menu = ctk.CTkOptionMenu(
            self.header_frame, 
            variable=self.language, 
            values=["Español", "Português", "English"], 
            command=self.change_language,
            width=120,  # Ancho reducido para hacerlo más pequeño
            height=30   # Altura reducida para hacerlo más pequeño
        )
        self.language_menu.pack(pady=10, side="right", padx=10, anchor="ne")  # Se ubica en la esquina superior derecha

        # Etiqueta de encabezado con texto ligeramente más pequeño
        self.header_label = ctk.CTkLabel(self.header_frame, text=tr["title"], font=("Helvetica", 24, "bold"))
        self.header_label.pack(pady=5, padx=5)

        # Etiqueta de descripción en el encabezado
        self.description_label = ctk.CTkLabel(
            self.header_frame,
            text=tr["description"],
            font=("Helvetica", 12),
            wraplength=800,
            justify="left"
        )
        self.description_label.pack(pady=10, padx=10)

        # Crear el panel de pestañas principal (pestañas de nivel superior)
        self.main_tabs = ttk.Notebook(self.root)
        self.main_tabs.pack(pady=20, padx=20, fill='both', expand=True)

        # Crear marcos para cada pestaña principal
        self.about_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.wildbook_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.wildlife_insights_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.iberian_lynx_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.camtrap_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.wildlife_insights_wildbook_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)  # Nueva pestaña

        # Agregar pestañas al notebook principal
        self.main_tabs.add(self.wildbook_frame, text=tr["wildbook"])
        self.main_tabs.add(self.wildlife_insights_frame, text=tr["wildlife_insights"])
        self.main_tabs.add(self.wildlife_insights_wildbook_frame, text=tr["wi_wildbook"])
        self.main_tabs.add(self.iberian_lynx_frame, text=tr["iberian_lynx"])
        self.main_tabs.add(self.camtrap_frame, text=tr["functionalities"])
        self.main_tabs.add(self.about_frame, text=tr["about"])


        # Estilo para las pestañas
        style = ttk.Style()
        style.configure(
            'TNotebook.Tab',
            font=('Helvetica', 14, 'bold'),  # Fuente, tamaño y estilo
            padding=[10, 5],  # Espaciado alrededor del texto
            relief='flat'  # Sin borde alrededor de las pestañas
        )

        # About TabView
        self.about_tabs = ctk.CTkTabview(self.about_frame, width=600, height=400)
        self.about_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.presentation_tab = self.about_tabs.add(tr["presentation"])

        # Camtrap Functionalities TabView
        self.camtrap_tabs = ctk.CTkTabview(self.camtrap_frame, width=600, height=400)
        self.camtrap_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.date_changer_tab = self.camtrap_tabs.add(tr["date_changer"])
        self.video_frame_extractor_tab = self.camtrap_tabs.add(tr["video_frame_extractor"])
        self.images_renamer_tab = self.camtrap_tabs.add(tr["images_renamer"])

        # Wildbook TabView
        self.wildbook_tabs = ctk.CTkTabview(self.wildbook_frame, width=600, height=400)
        self.wildbook_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.wiwbe_folder_tab = self.wildbook_tabs.add(tr["wiwbe_folder"])
        self.wbcatalog_tab = self.wildbook_tabs.add(tr["wiwbe_catalog"])

        # Wildlife Insights TabView
        self.wildlife_insights_tabs = ctk.CTkTabview(self.wildlife_insights_frame, width=600, height=400)
        self.wildlife_insights_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.wi_downloader_tab = self.wildlife_insights_tabs.add(tr["wi_downloader"])

        # Iberian Lynx TabView
        self.iberian_lynx_tabs = ctk.CTkTabview(self.iberian_lynx_frame, width=600, height=400)
        self.iberian_lynx_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.lynx_feature_1_tab = self.iberian_lynx_tabs.add(tr["lynx_feature_1"])

        # Wildlife Insight - Wildbook TabView (Nuevo)
        self.wildlife_insights_wildbook_tabs = ctk.CTkTabview(self.wildlife_insights_wildbook_frame, width=600, height=400)
        self.wildlife_insights_wildbook_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.excel_combiner_tab = self.wildlife_insights_wildbook_tabs.add(tr["wi_csvs_to_biwbe"])

        # Integrar las clases en la BaseApp
        self.photo_date_app = WBFolderApp(self.wiwbe_folder_tab, lang=lang)
        self.gcs_downloader_app = GCSDownloaderAndRenamer(self.wi_downloader_tab, lang=lang)
        self.excel_combiner_app = ExcelCombinerApp(self.excel_combiner_tab, lang=lang)
        self.data_changer_app = DateChangerApp(self.date_changer_tab, lang=lang)
        self.frame_extractor_app = FrameExtractorApp(self.video_frame_extractor_tab, lang=lang)
        self.wbcatalog_app = WBCatalogApp(self.wbcatalog_tab, lang=lang)
        self.images_renamer_app = ImagesRenamer(self.images_renamer_tab, lang=lang)
        self.presentation_app = Presentation(self.presentation_tab, lang=lang)
        self.lynxone_app = LynxOne(self.lynx_feature_1_tab, lang=lang)

    def change_language(self, *args):
        if self.root.jobs.busy:
            return
        for widget in self.root.winfo_children():
            widget.destroy()
        self.setup_ui()

        # self.language_menu = ctk.CTkOptionMenu(
        #     self.header_frame, 
        #     variable=self.language, 
        #     values=["Español", "Português", "English"], 
        #     command=self.change_language,
        #     width=120,  # Ancho reducido para hacerlo más pequeño
        #     height=30   # Altura reducida para hacerlo más pequeño
        # )
        # self.language_menu.pack(pady=10, side="right", padx=10, anchor="ne") 

class Presentation(ctk.CTkFrame):
    def __init__(self, root, lang="es"):
        super().__init__(root)
        self.lang = lang

        self.translations = {
            "es": {
                "title": "LynxAutomator",
                "description": "LynxAutomator ha sido desarrollada por WWF España\n"
                               "en el ámbito del proyecto LIFE LynxConnect 19NAT/ES/001055\n"
                               "en la acción A8 Nuevas técnicas complementarias para el seguimiento de las poblaciones de lince"
            },
            "pt": {
                "title": "LynxAutomator",
                "description": "LynxAutomator foi desenvolvida pela WWF Espanha\n"
                               "no âmbito do projeto LIFE LynxConnect 19NAT/ES/001055\n"
                               "na ação A8 Novas técnicas complementares para o monitoramento das populações de lince"
            },
            "en": {
                "title": "LynxAutomator",
                "description": "LynxAutomator has been developed by WWF Spain\n"
                               "within the framework of the LIFE LynxConnect project 19NAT/ES/001055\n"
                               "under Action A8 New complementary techniques for monitoring lynx populations"
            }
        }

        self.setup_ui()

    def resource_path(self, relative_path):
        """Obtiene la ruta al recurso, compatible con PyInstaller."""
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.abspath(os.path.dirname(__file__))
        return os.path.join(base_path, relative_path)

    def setup_ui(self):
        tr = self.translations[self.lang]
        self.pack(fill="both", expand=True)

        # Cargar imagen del logo
        logo_path = self.resource_path("logo.png")
        
        try:
            # 1. Carga la imagen con PIL (como antes)
            pil_image = Image.open(logo_path)
            
            # 2. Convierte la imagen de PIL a CTkImage
            # Puedes especificar un tamaño si quieres controlarlo,
            # pero para mantener el tamaño original, solo pasas la imagen PIL.
            self.logo = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(pil_image.width, pil_image.height))
            # Si solo tienes un modo de color, puedes usar solo 'light_image' o 'dark_image'.
            # light_image es para el tema claro, dark_image para el tema oscuro.
            # Si son iguales, puedes pasar la misma imagen a ambos.

        except FileNotFoundError:
            print(f"Error: No se encontró el archivo de imagen en {logo_path}")
            # Si no se encuentra el logo, puedes optar por:
            # - No mostrar el label de imagen.
            # - Usar una imagen de marcador de posición.
            # - Salir o levantar una excepción.
            self.logo = None # Asegurarse de que self.logo es None o una imagen por defecto
            
        if self.logo: # Solo crea el label si la imagen fue cargada
            # 3. Usa CTkImage en el CTkLabel
            logo_label = ctk.CTkLabel(self, image=self.logo, text="") # text="" para que no muestre texto adicional
            logo_label.pack(pady=20)
        else:
            # Opcional: mostrar un mensaje si no se encuentra la imagen
            error_label = ctk.CTkLabel(self, text="Logo no encontrado", font=("Helvetica", 14), text_color="red")
            error_label.pack(pady=20)


        # Texto de descripción
        text_label = ctk.CTkLabel(self, text=tr["description"], font=("Helvetica", 16))
        text_label.pack(pady=10)

class AppHelp:
    def __init__(self, parent):
        self.description_parts = {
            "WIWbE from Folder": (
                "This app helps to fulfill a Bulk Import Wildbook Excel (WIWbE). To generate the WIWbE, you will need a folder containing images and an Excel file. The Excel file should be a WIWbE template with only the first row partially filled out. The app extracts the filenames (corresponding to Encounter.mediaAsset0) and the EXIF dates from the images in the selected folder. It then uses this extracted data to populate the relevant columns in the Excel file, completing the WIWbE with the necessary information. The rest of the variables in the WIWbE file are filled in based on the data from the first row of the template, ensuring consistency across the entries."
            ),
            "Lorem": (
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua\n"
            ),
            "ipsum": (
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. \n"
            ),
            "dolor": (
                "- Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua\n"
                "- Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua\n\n"
            ),
            "sit": (
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua:\n"
                "Example Files"
            )
        }
        self.create_widgets(parent)

    def create_widgets(self, parent):
        index_frame = ctk.CTkFrame(parent)
        index_frame.pack(side="left", fill="y", padx=10, pady=10)

        content_frame = ctk.CTkFrame(parent)
        content_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.scrollbar = tk.Scrollbar(content_frame)
        self.scrollbar.pack(side="right", fill="y")

        self.text_widget = tk.Text(content_frame, wrap="word", yscrollcommand=self.scrollbar.set)
        self.text_widget.pack(fill="both", expand=True)
        self.scrollbar.configure(command=self.text_widget.yview)

        for section in self.description_parts:
            button = ctk.CTkButton(index_frame, text=section, command=lambda s=section: self.scroll_to_section(s))
            button.pack(pady=5, fill="x")

        self.populate_text()

    def populate_text(self):
        self.text_widget.configure(state="normal")
        self.text_widget.delete("1.0", tk.END)
        for section, content in self.description_parts.items():
            self.text_widget.insert(tk.END, f"{section}\n", ("section",))
            self.insert_with_links(content)
            self.text_widget.insert(tk.END, "\n\n", ("content",))
        
        # Insert the "Example Files" link at the end
        example_files_text = "Example Files"
        self.text_widget.insert(tk.END, example_files_text, ("link",))
        self.text_widget.tag_bind("link", "<Button-1>", lambda e: self.open_url("https://docs.google.com/spreadsheets/d/e/2PACX-1vR4va7tsFS2NgGxADc6U_HK7zyO5511yh-r7l5S2mSAmdreguKXQNOd8L-RlDtQXzUOniiINhz3fw9Y/pubhtml?gid=0&single=true"))

        self.text_widget.tag_configure("section", font=("Helvetica", 14, "bold"))
        self.text_widget.tag_configure("content", font=("Helvetica", 12))
        self.text_widget.tag_configure("link", foreground="blue", underline=True)
        self.text_widget.configure(state="disabled")

    def insert_with_links(self, content):
        """Insert text with clickable links."""
        words = content.split()
        for word in words:
            if word.startswith("[") and word.endswith(")"):
                # Extract link text and URL
                link_text = word[word.find("[") + 1:word.find("]")]
                url = word[word.find("(") + 1:word.find(")")]
                # Insert the link text
                self.text_widget.insert(tk.END, link_text, ("link",))
                self.text_widget.tag_bind("link", "<Button-1>", lambda e, url=url: self.open_url(url))
                self.text_widget.insert(tk.END, " ")
            else:
                self.text_widget.insert(tk.END, word + " ", ("content",))

    def open_url(self, url):
        webbrowser.open(url)

    def scroll_to_section(self, section):
        self.text_widget.configure(state="normal")
        index = self.text_widget.search(section, "1.0", tk.END)
        if index:
            self.text_widget.see(index)
            self.text_widget.tag_remove("highlight", "1.0", tk.END)
            end_index = f"{index}+{len(section)}c"
            self.text_widget.tag_add("highlight", index, end_index)
            self.text_widget.tag_configure("highlight", background="yellow")
        self.text_widget.configure(state="disabled")    


class WBFolderApp(FolderJobs):
    def __init__(self, root, lang="es"):
        self.root = root
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "select_folder": "Selecciona la carpeta con las imágenes. El nombre y la fecha se tomarán de estas imágenes.",
                "browse_folder": "Buscar Carpeta",
                "no_folder_selected": "No se ha seleccionado ninguna carpeta",
                "select_excel": "Selecciona el archivo Excel inicial.",
                "browse_file": "Buscar Archivo",
                "no_file_selected": "No se ha seleccionado ningún archivo",
                "process": "Procesar",
                "download_excel": "Descargar Excel Actualizado",
                "success_message": "¡Archivo Excel procesado con éxito!",
                "group_by_time": "Agrupar por tiempo"
            },
            "pt": {
                "select_folder": "Selecione a pasta com as imagens. O nome e a data serão retirados dessas imagens.",
                "browse_folder": "Procurar Pasta",
                "no_folder_selected": "Nenhuma pasta selecionada",
                "select_excel": "Selecione o arquivo Excel inicial.",
                "browse_file": "Procurar Arquivo",
                "no_file_selected": "Nenhum arquivo selecionado",
                "process": "Processar",
                "download_excel": "Baixar Excel Atualizado",
                "success_message": "Arquivo Excel processado com sucesso!",
                "group_by_time": "Agrupar por tempo"
            },
            "en": {
                "select_folder": "Select the folder with the images. The name and the date will be taken from these images.",
                "browse_folder": "Browse Folder",
                "no_folder_selected": "No folder selected",
                "select_excel": "Select the Initial Excel file.",
                "browse_file": "Browse File",
                "no_file_selected": "No file selected",
                "process": "Process",
                "download_excel": "Download Updated Excel",
                "success_message": "Excel file processed successfully!",
                "group_by_time": "Group by time"
            }
        }

        # Resto de la configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtain translations based on the selected language
        tr = self.translations[self.lang]

        # Main frame (single large square)
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Label asking the user to select a folder
        self.label = ctk.CTkLabel(self.main_frame, text=tr["select_folder"], anchor="w")
        self.label.pack(pady=10)

        # Frame for the folder selection widgets
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select a folder
        self.select_folder_btn = ctk.CTkButton(self.folder_frame, text=tr["browse_folder"], command=self.select_folder)
        self.select_folder_btn.pack(side="left", padx=5)

        # Label to display the selected folder's name
        self.folder_label = ctk.CTkLabel(self.folder_frame, text=tr["no_folder_selected"], anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Message between the two buttons
        self.message_label = ctk.CTkLabel(self.main_frame, text=tr["select_excel"])
        self.message_label.pack(pady=10)

        # Frame for the file selection widgets
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select an Excel file
        self.select_file_btn = ctk.CTkButton(self.file_frame, text=tr["browse_file"], command=self.select_file)
        self.select_file_btn.pack(side="left", padx=5)

        # Label to display the selected file's name
        self.file_label = ctk.CTkLabel(self.file_frame, text=tr["no_file_selected"], anchor="w")
        self.file_label.pack(side="left", padx=5)

        # Frame for Checkbox and Time Threshold
        self.options_frame = ctk.CTkFrame(self.main_frame)
        self.options_frame.pack(pady=10, padx=10, fill="x")

        # Checkbox for multiple images (group by time)
        self.multiple_images_var = tk.BooleanVar()
        self.multiple_images_check = ctk.CTkCheckBox(self.options_frame, text=tr["group_by_time"], variable=self.multiple_images_var)
        self.multiple_images_check.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Label for time threshold
        self.time_threshold_label = ctk.CTkLabel(self.options_frame, text="Time Threshold (seconds):", anchor="w")
        self.time_threshold_label.grid(row=0, column=1, padx=(20, 5), pady=5, sticky="w")

        # Entry for time threshold
        self.time_threshold_entry = ctk.CTkEntry(self.options_frame)
        self.time_threshold_entry.grid(row=0, column=2, padx=5, pady=5)

        # Frame for Process and Download Buttons
        self.buttons_frame = ctk.CTkFrame(self.main_frame)
        self.buttons_frame.pack(pady=10)

        # Process files button
        self.process_btn = ctk.CTkButton(self.buttons_frame, text=tr["process"], command=self.process_files, state=ctk.DISABLED)
        self.process_btn.pack(side="left", padx=5)

        # Download Excel button
        self.download_btn = ctk.CTkButton(self.buttons_frame, text=tr["download_excel"], command=self.download_file, state=ctk.DISABLED)
        self.download_btn.pack(side="left", padx=5)

        # Variable to store the path of the temporary file

    def select_folder(self):
        # Open a dialog to select a folder
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.folder_label.configure(text=os.path.basename(self.folder_path))
            self.check_ready_to_process()

    def select_file(self):
        # Open a dialog to select an Excel file
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.file_label.configure(text=os.path.basename(self.file_path))
            self.check_ready_to_process()

    def check_ready_to_process(self):
        # Enable the process button if both a folder and a file have been selected
        if hasattr(self, 'folder_path') and hasattr(self, 'file_path'):
            self.process_btn.configure(state=ctk.NORMAL)


class WBCatalogApp(CatalogJobs):
    def __init__(self, root, lang="es"):
        self.root = root
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "select_directory": "Selecciona un directorio para analizar. La primera palabra del nombre de archivo se utilizará como el ID individual.",
                "browse_folder": "Buscar Carpeta",
                "no_folder_selected": "No se ha seleccionado ninguna carpeta",
                "select_excel": "Selecciona un archivo Excel para procesar junto con las fotos.",
                "browse_file": "Buscar Archivo",
                "no_file_selected": "No se ha seleccionado ningún archivo",
                "collapse_rows": "Colapsar filas con el mismo ID individual",
                "capitalize_id": "Capitalizar el ID individual",
                "process": "Procesar",
                "download_excel": "Descargar Excel Actualizado",
                "success_message": "¡Archivo Excel procesado con éxito!"
            },
            "pt": {
                "select_directory": "Selecione um diretório para analisar. A primeira palavra do nome do arquivo será usada como o ID individual.",
                "browse_folder": "Procurar Pasta",
                "no_folder_selected": "Nenhuma pasta selecionada",
                "select_excel": "Selecione um arquivo Excel para processar junto com as fotos.",
                "browse_file": "Procurar Arquivo",
                "no_file_selected": "Nenhum arquivo selecionado",
                "collapse_rows": "Colapsar linhas com o mesmo ID individual",
                "capitalize_id": "Capitalizar o ID individual",
                "process": "Processar",
                "download_excel": "Baixar Excel Atualizado",
                "success_message": "Arquivo Excel processado com sucesso!"
            },
            "en": {
                "select_directory": "Select a directory to analyze. The first word of the filename will be used as the individual ID.",
                "browse_folder": "Browse Folder",
                "no_folder_selected": "No folder selected",
                "select_excel": "Select an Excel file to process along with the photos.",
                "browse_file": "Browse File",
                "no_file_selected": "No file selected",
                "collapse_rows": "Collapse rows with the same individual ID",
                "capitalize_id": "Capitalize individual ID",
                "process": "Process",
                "download_excel": "Download Updated Excel",
                "success_message": "Excel file processed successfully!"
            }
        }

        # Configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtener las traducciones según el idioma seleccionado
        tr = self.translations[self.lang]

        # Frame principal
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Etiqueta para selección de directorio
        self.label1 = ctk.CTkLabel(self.main_frame, text=tr["select_directory"], anchor="w")
        self.label1.pack(pady=10)

        # Frame para la selección de directorio
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        # Botón para buscar y seleccionar una carpeta
        self.select_folder_btn = ctk.CTkButton(self.folder_frame, text=tr["browse_folder"], command=self.select_folder)
        self.select_folder_btn.pack(side="left", padx=5)

        # Etiqueta para mostrar la carpeta seleccionada
        self.folder_label = ctk.CTkLabel(self.folder_frame, text=tr["no_folder_selected"], anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Etiqueta para selección de archivo Excel
        self.label2 = ctk.CTkLabel(self.main_frame, text=tr["select_excel"], anchor="w")
        self.label2.pack(pady=10)

        # Frame para la selección de archivo Excel
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(pady=5, padx=10, fill="x")

        # Botón para buscar y seleccionar un archivo
        self.select_file_btn = ctk.CTkButton(self.file_frame, text=tr["browse_file"], command=self.select_file)
        self.select_file_btn.pack(side="left", padx=5)

        # Etiqueta para mostrar el archivo seleccionado
        self.file_label = ctk.CTkLabel(self.file_frame, text=tr["no_file_selected"], anchor="w")
        self.file_label.pack(side="left", padx=5)

        # Frame para las opciones de checkboxes
        self.checkbox_frame = ctk.CTkFrame(self.main_frame)
        self.checkbox_frame.pack(pady=5, padx=10, fill="x")

        # Checkbox para la opción de colapsar filas
        self.collapse_var = ctk.BooleanVar()
        self.collapse_check = ctk.CTkCheckBox(self.checkbox_frame, text=tr["collapse_rows"], variable=self.collapse_var)
        self.collapse_check.pack(side="left", padx=5)

        # Checkbox para la opción de capitalizar el ID individual
        self.capitalize_var = ctk.BooleanVar(value=True)
        self.capitalize_check = ctk.CTkCheckBox(self.checkbox_frame, text=tr["capitalize_id"], variable=self.capitalize_var)
        self.capitalize_check.pack(side="left", padx=5)

        # Frame para los botones de procesar y descargar
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(pady=10)

        # Botón para procesar
        self.process_btn = ctk.CTkButton(self.button_frame, text=tr["process"], command=self.process_files)
        self.process_btn.pack(side="left", padx=5)

        # Botón para descargar el Excel
        self.download_btn = ctk.CTkButton(self.button_frame, text=tr["download_excel"], command=self.download_file, state=ctk.DISABLED)
        self.download_btn.pack(side="left", padx=5)
        
    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.folder_label.configure(text=self.folder_path)
            messagebox.showinfo("Information", f"Selected folder: {self.folder_path}")

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.file_label.configure(text=os.path.basename(self.file_path))
            messagebox.showinfo("Information", f"Selected file: {self.file_path}")


class FrameExtractorApp(VideoJobs):
    def __init__(self, root, lang="es"):
        self.root = root
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "select_videos": "Selecciona la carpeta que contiene los videos.",
                "browse_folder": "Buscar Carpeta",
                "no_folder_selected": "No se ha seleccionado ninguna carpeta",
                "interval_between_frames": "Intervalo entre fotogramas (segundos):",
                "extract_frames": "Extraer fotogramas",
                "processing": "Procesando...",
                "success_message": "Fotogramas extraídos y guardados con éxito.",
                "error_message": "Error: "
            },
            "pt": {
                "select_videos": "Selecione a pasta que contém os vídeos.",
                "browse_folder": "Procurar Pasta",
                "no_folder_selected": "Nenhuma pasta selecionada",
                "interval_between_frames": "Intervalo entre quadros (segundos):",
                "extract_frames": "Extrair quadros",
                "processing": "Processando...",
                "success_message": "Quadros extraídos e salvos com sucesso.",
                "error_message": "Erro: "
            },
            "en": {
                "select_videos": "Select the folder containing the videos.",
                "browse_folder": "Browse Folder",
                "no_folder_selected": "No folder selected",
                "interval_between_frames": "Interval between frames (seconds):",
                "extract_frames": "Extract frames",
                "processing": "Processing...",
                "success_message": "Frames extracted and saved successfully.",
                "error_message": "Error: "
            }
        }

        # Configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtener las traducciones según el idioma seleccionado
        tr = self.translations[self.lang]

        # Frame principal
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Etiqueta para seleccionar la carpeta de videos
        self.label = ctk.CTkLabel(self.main_frame, text=tr["select_videos"], anchor="w")
        self.label.pack(pady=10)

        # Frame para la selección de la carpeta de videos
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        # Botón para buscar y seleccionar la carpeta
        self.select_videos_btn = ctk.CTkButton(self.folder_frame, text=tr["browse_folder"], command=self.select_videos_folder)
        self.select_videos_btn.pack(side="left", padx=5)

        # Etiqueta para mostrar la carpeta seleccionada
        self.folder_label = ctk.CTkLabel(self.folder_frame, text=tr["no_folder_selected"], anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Frame para el intervalo entre cuadros
        self.interval_frame = ctk.CTkFrame(self.main_frame)
        self.interval_frame.pack(pady=10, padx=10, fill="x")

        self.interval_label = ctk.CTkLabel(self.interval_frame, text=tr["interval_between_frames"])
        self.interval_label.pack(side="left", padx=5)

        self.interval_var = ctk.DoubleVar(value=1.0)  # Valor por defecto de 1 segundo
        self.interval_entry = ctk.CTkEntry(self.interval_frame, textvariable=self.interval_var, width=100)
        self.interval_entry.pack(side="left", padx=5)

        # Botón para extraer cuadros
        self.process_btn = ctk.CTkButton(self.main_frame, text=tr["extract_frames"], command=self.start_extraction)
        self.process_btn.pack(pady=10)

        # Etiqueta de estado para mostrar mensajes de procesamiento
        self.status_label = ctk.CTkLabel(self.main_frame, text="", anchor="center", justify="center")
        self.status_label.pack(pady=10, padx=10, fill="x")
        
    def select_videos_folder(self):
        self.folder_path = filedialog.askdirectory(title="Select the folder with videos")
        if self.folder_path:
            self.folder_label.configure(text=os.path.basename(self.folder_path))
            messagebox.showinfo("Information", f"Selected folder: {self.folder_path}")
        else:
            self.status_label.configure(text="Please select a folder with videos.")


class GCSDownloaderAndRenamer(DownloadJobs, ctk.CTkFrame):
    def __init__(self, root, lang="es"):
        super().__init__(root)  # Llamada correcta al constructor de la clase base
        self.root = root
        self.lang = lang  # Almacena el idioma actual
        self.pack(fill="both", expand=True)  # Asegura que el marco se ajuste al tamaño de la ventana
        ctk.set_appearance_mode("System")  # Modo de apariencia ("System", "Dark", "Light")
        
        # Diccionario de traducciones
        self.translations = {
            "es": {
                "select_csv": "Selecciona el archivo CSV con URLs de Wildlife Insights",
                "browse_csv": "Buscar CSV",
                "no_csv_selected": "No se ha seleccionado un archivo CSV",
                "multiple_folders": "Guardar en carpetas separadas por deployment_id",
                "download_images": "Descargar Imágenes",
                "stop": "Detener",
                "downloading": "Descargando...",
                "stopped": "Descarga detenida"
            },
            "pt": {
                "select_csv": "Selecione o arquivo CSV com URLs do Wildlife Insights",
                "browse_csv": "Procurar CSV",
                "no_csv_selected": "Nenhum arquivo CSV selecionado",
                "multiple_folders": "Salvar em pastas separadas por deployment_id",
                "download_images": "Baixar Imagens",
                "stop": "Parar",
                "downloading": "Baixando...",
                "stopped": "Download interrompido"
            },
            "en": {
                "select_csv": "Select the images CSV file with URLs from Wildlife Insights",
                "browse_csv": "Browse CSV",
                "no_csv_selected": "No CSV file selected",
                "multiple_folders": "Save in separate folders by deployment_id",
                "download_images": "Download Images",
                "stop": "Stop",
                "downloading": "Downloading...",
                "stopped": "Download stopped"
            }
        }

        # Marco principal para contener todos los widgets
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Etiqueta para la selección de archivos CSV
        self.label1 = ctk.CTkLabel(self.main_frame, text=self.translations[self.lang]["select_csv"], anchor="w")
        self.label1.pack(pady=10)

        # Marco para contener los widgets de selección de CSV
        self.selection_frame = ctk.CTkFrame(self.main_frame)
        self.selection_frame.pack(pady=5, padx=10, fill="x")

        # Botón para buscar y seleccionar el archivo CSV
        self.select_csv_btn = ctk.CTkButton(self.selection_frame, text=self.translations[self.lang]["browse_csv"], command=self.select_csv)
        self.select_csv_btn.pack(side="left", padx=5)

        # Etiqueta para mostrar el nombre del archivo CSV seleccionado
        self.csv_label = ctk.CTkLabel(self.selection_frame, text=self.translations[self.lang]["no_csv_selected"], anchor="w")
        self.csv_label.pack(side="left", padx=5)

        # Añadir opción para carpetas simples o múltiples
        self.use_multiple_folders = IntVar()
        self.multiple_folders_checkbtn = ctk.CTkCheckBox(self.main_frame, text=self.translations[self.lang]["multiple_folders"], variable=self.use_multiple_folders)
        self.multiple_folders_checkbtn.pack(pady=10)

        # Etiqueta de estado para mostrar el progreso de la descarga
        self.status_var = StringVar()
        self.status_label = ctk.CTkLabel(self.main_frame, textvariable=self.status_var)
        self.status_label.pack(pady=10)

        # Botones para el procesamiento
        self.buttons_frame = ctk.CTkFrame(self.main_frame)
        self.buttons_frame.pack(pady=10)

        self.download_btn = ctk.CTkButton(self.buttons_frame, text=self.translations[self.lang]["download_images"], command=self.start_download, state=ctk.DISABLED)
        self.download_btn.pack(side="left", padx=5)


    def select_csv(self):
        self.csv_path = filedialog.askopenfilename(title="Select CSV File", filetypes=[("CSV files", "*.csv")])
        if self.csv_path:
            self.csv_label.configure(text=f"Selected CSV File: {os.path.basename(self.csv_path)}")
            self.update_download_button_state()

    def update_download_button_state(self):
        if hasattr(self, 'csv_path'):
            self.download_btn.configure(state=ctk.NORMAL)
        else:
            self.download_btn.configure(state=ctk.DISABLED)


class ExcelCombinerApp(WIJobs):
    """
    Clase principal de la aplicación para combinar datos de Excel y CSV,
    procesar imágenes y generar un nuevo archivo Excel.
    """
    def __init__(self, root, lang="es"):
        """
        Constructor de la clase.
        Inicializa la ventana principal de la aplicación y las variables de estado.

        Args:
            root (tk.Tk o ctk.CTk): La ventana raíz de la aplicación.
            lang (str): El idioma de la interfaz ("es", "pt", "en").
        """
        self.root = root # Ventana raíz de la aplicación
        self.lang = lang  # Idioma actual de la interfaz

        # Rutas de los archivos seleccionados por el usuario
        self.initial_excel_path = None
        self.images_csv_path = None
        self.deployments_csv_path = None

        # Diccionario de traducciones para la interfaz de usuario
        self.translations = {
            "es": {
                "select_initial_excel": "Selecciona el archivo Excel inicial",
                "browse_excel": "Buscar Excel",
                "no_excel_selected": "No se ha seleccionado ningún archivo Excel",
                "select_images_csv": "Selecciona el archivo CSV de imágenes",
                "browse_csv": "Buscar CSV",
                "no_csv_selected": "No se ha seleccionado ningún archivo CSV",
                "select_deployments_csv": "Selecciona el archivo CSV de deployments",
                "process_multiple_images": "Procesar múltiples imágenes  ",
                "time_threshold": "Umbral de tiempo (segundos):",
                "process_files": "Procesar Archivos",
                "download_excel": "Descargar Excel Actualizado",
                "process_completed": "Archivos procesados con éxito",
                "error_message": "Se produjo un error: ",
                "separate_objects_gt_1": "Separar si objetos > 1", # Nuevo texto para el checkbox
                "missing_columns_error": "Columnas 'project_id' o 'deployment_id' faltan en los archivos CSV de imágenes o despliegues. Asegúrate de que ambos archivos contengan estas columnas.",
                "error_column_missing_in_images_csv": "La columna '{col}' falta en el archivo CSV de imágenes.",
                "error_column_missing_in_deployments_csv": "La columna '{col}' falta en el archivo CSV de deployments.",
                "error_no_initial_excel_selected": "No se ha seleccionado un archivo de Excel inicial.",
                "file_saved_successfully": "Archivo guardado con éxito en {path}",
                "error_threshold_value": "El umbral de tiempo debe ser un número entero.",
                "warning_empty_final_df": "El DataFrame final está vacío, pero había datos para procesar. Verifique los filtros y umbrales.",
                "info_no_valid_data": "No había datos válidos para procesar después de la carga inicial y combinación.",
                "error_saving_file": "No se pudo guardar el archivo: ",
                "warning_download_empty": "No hay datos procesados para descargar. El Excel resultante estaría vacío.",
                "error_download_not_available": "Los archivos aún no se han procesado o el procesamiento no generó datos."
            },
            "pt": {
                "select_initial_excel": "Selecione o arquivo Excel inicial",
                "browse_excel": "Procurar Excel",
                "no_excel_selected": "Nenhum arquivo Excel selecionado",
                "select_images_csv": "Selecione o arquivo CSV de imagens",
                "browse_csv": "Procurar CSV",
                "no_csv_selected": "Nenhum arquivo CSV selecionado",
                "select_deployments_csv": "Selecione o arquivo CSV de deployments",
                "process_multiple_images": "Processar múltiplas imagens  ",
                "time_threshold": "Limite de tempo (segundos):",
                "process_files": "Processar Arquivos",
                "download_excel": "Baixar Excel Atualizado",
                "process_completed": "Arquivos processados com sucesso",
                "error_message": "Ocorreu um erro: ",
                "separate_objects_gt_1": "Separar se objetos > 1",
                "missing_columns_error": "Colunas 'project_id' ou 'deployment_id' estão faltando nos arquivos CSV de imagens ou implantações. Certifique-se de que ambos os arquivos contenham essas colunas.",
                "error_column_missing_in_images_csv": "A coluna '{col}' está faltando no arquivo CSV de imagens.",
                "error_column_missing_in_deployments_csv": "A coluna '{col}' está faltando no arquivo CSV de implantações.",
                "error_no_initial_excel_selected": "Nenhum arquivo Excel inicial foi selecionado.",
                "file_saved_successfully": "Arquivo salvo com sucesso em {path}",
                "error_threshold_value": "O limite de tempo deve ser um número inteiro.",
                "warning_empty_final_df": "O DataFrame final está vazio, mas havia dados para processar. Verifique os filtros e limites.",
                "info_no_valid_data": "Não havia dados válidos para processar após o carregamento inicial e a combinação.",
                "error_saving_file": "Não foi possível salvar o arquivo: ",
                "warning_download_empty": "Não há dados processados para download. O Excel resultante estaria vazio.",
                "error_download_not_available": "Os arquivos ainda não foram processados ou o processamento não gerou dados."
            },
            "en": {
                "select_initial_excel": "Select the Initial Excel file",
                "browse_excel": "Browse Excel",
                "no_excel_selected": "No Excel file selected",
                "select_images_csv": "Select the Images CSV file",
                "browse_csv": "Browse CSV",
                "no_csv_selected": "No CSV file selected",
                "select_deployments_csv": "Select the Deployments CSV file",
                "process_multiple_images": "Process multiple images  ",
                "time_threshold": "Time threshold (seconds):",
                "process_files": "Process Files",
                "download_excel": "Download Updated Excel",
                "process_completed": "Files processed successfully",
                "error_message": "An error occurred: ",
                "separate_objects_gt_1": "Separate if objects > 1",
                "missing_columns_error": "Columns 'project_id' or 'deployment_id' are missing in the images or deployments CSV files. Ensure both files contain these columns.",
                "error_column_missing_in_images_csv": "Column '{col}' is missing in the images CSV file.",
                "error_column_missing_in_deployments_csv": "Column '{col}' is missing in the deployments CSV file.",
                "error_no_initial_excel_selected": "No initial Excel file has been selected.",
                "file_saved_successfully": "File saved successfully to {path}",
                "error_threshold_value": "Time threshold must be an integer.",
                "warning_empty_final_df": "The final DataFrame is empty, but there was data to process. Check filters and thresholds.",
                "info_no_valid_data": "No valid data was found to process after initial load and merge.",
                "error_saving_file": "Could not save file: ",
                "warning_download_empty": "No processed data available for download. The resulting Excel would be empty.",
                "error_download_not_available": "Files have not been processed yet or processing resulted in no data."
            }
        }
        # Configura la interfaz de usuario
        self.setup_ui()

    def setup_ui(self):
        """
        Configura todos los elementos de la interfaz de usuario (widgets)
        utilizando CustomTkinter.
        """
        tr = self.translations[self.lang] # Obtiene las traducciones para el idioma actual

        # Marco principal que contendrá todos los demás elementos
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Sección para la selección del archivo Excel inicial
        self.label1 = ctk.CTkLabel(self.main_frame, text=tr["select_initial_excel"], anchor="w")
        self.label1.pack(pady=10, fill="x")

        self.excel_frame = ctk.CTkFrame(self.main_frame)
        self.excel_frame.pack(pady=5, padx=10, fill="x")
        self.select_excel_btn = ctk.CTkButton(self.excel_frame, text=tr["browse_excel"], command=self.select_initial_excel)
        self.select_excel_btn.pack(side="left", padx=5)
        self.excel_label = ctk.CTkLabel(self.excel_frame, text=tr["no_excel_selected"], anchor="w")
        self.excel_label.pack(side="left", padx=5, expand=True, fill="x")

        # Sección para la selección del archivo CSV de imágenes
        self.label2 = ctk.CTkLabel(self.main_frame, text=tr["select_images_csv"], anchor="w")
        self.label2.pack(pady=10, fill="x")

        self.images_frame = ctk.CTkFrame(self.main_frame)
        self.images_frame.pack(pady=5, padx=10, fill="x")
        self.select_images_btn = ctk.CTkButton(self.images_frame, text=tr["browse_csv"], command=self.select_images_csv)
        self.select_images_btn.pack(side="left", padx=5)
        self.images_label = ctk.CTkLabel(self.images_frame, text=tr["no_csv_selected"], anchor="w")
        self.images_label.pack(side="left", padx=5, expand=True, fill="x")

        # Sección para la selección del archivo CSV de despliegues
        self.label3 = ctk.CTkLabel(self.main_frame, text=tr["select_deployments_csv"], anchor="w")
        self.label3.pack(pady=10, fill="x")

        self.deployments_frame = ctk.CTkFrame(self.main_frame)
        self.deployments_frame.pack(pady=5, padx=10, fill="x")
        self.select_deployments_btn = ctk.CTkButton(self.deployments_frame, text=tr["browse_csv"], command=self.select_deployments_csv)
        self.select_deployments_btn.pack(side="left", padx=5)
        self.deployments_label = ctk.CTkLabel(self.deployments_frame, text=tr["no_csv_selected"], anchor="w")
        self.deployments_label.pack(side="left", padx=5, expand=True, fill="x")

        # Marco para las opciones de procesamiento (checkboxes y umbral de tiempo)
        self.options_frame = ctk.CTkFrame(self.main_frame)
        self.options_frame.pack(pady=10, padx=10, fill="x")

        # Checkbox para procesar múltiples imágenes (agrupar por tiempo)
        self.multiple_images_var = tk.BooleanVar()
        self.multiple_images_check = ctk.CTkCheckBox(self.options_frame, text=tr["process_multiple_images"], variable=self.multiple_images_var)
        self.multiple_images_check.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Etiqueta y campo de entrada para el umbral de tiempo
        self.time_threshold_label = ctk.CTkLabel(self.options_frame, text=tr["time_threshold"], anchor="w")
        self.time_threshold_label.grid(row=0, column=1, padx=(20, 5), pady=5, sticky="w")
        self.time_threshold_entry = ctk.CTkEntry(self.options_frame, width=100) # Ancho ajustado
        self.time_threshold_entry.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.time_threshold_entry.insert(0, "3") # Valor predeterminado de 3 segundos

        # Checkbox para separar imágenes si 'number_of_objects' es mayor que 1
        self.separate_large_groups_var = tk.BooleanVar(value=False) # Valor inicial False
        self.separate_large_groups_check = ctk.CTkCheckBox(
            self.options_frame,
            text=tr["separate_objects_gt_1"], # Texto actualizado desde las traducciones
            variable=self.separate_large_groups_var
        )
        self.separate_large_groups_check.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        
        # Configuración de las columnas del marco de opciones para un diseño flexible
        self.options_frame.columnconfigure(0, weight=1) # Permite que el texto del checkbox se expanda
        self.options_frame.columnconfigure(1, weight=0) # Etiqueta de umbral de tiempo fija
        self.options_frame.columnconfigure(2, weight=0) # Campo de entrada de umbral de tiempo fijo

        # Marco para los botones de acción
        self.buttons_frame = ctk.CTkFrame(self.main_frame)
        self.buttons_frame.pack(pady=20) # Mayor padding vertical

        # Botón para iniciar el procesamiento de archivos
        self.process_btn = ctk.CTkButton(self.buttons_frame, text=tr["process_files"], command=self.process_files, state=ctk.DISABLED)
        self.process_btn.pack(side="left", padx=10) # Mayor padding horizontal

        # Botón para descargar el archivo Excel actualizado
        self.download_btn = ctk.CTkButton(self.buttons_frame, text=tr["download_excel"], command=self.save_file, state=ctk.DISABLED)
        self.download_btn.pack(side="left", padx=10) # Mayor padding horizontal
        
    def select_initial_excel(self):
        """
        Abre un diálogo para seleccionar el archivo Excel inicial (.xlsx o .xls).
        Actualiza la etiqueta de la interfaz y verifica si todos los archivos están seleccionados.
        """
        self.initial_excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.initial_excel_path:
            self.excel_label.configure(text=os.path.basename(self.initial_excel_path))
        else:
            self.excel_label.configure(text=self.translations[self.lang]["no_excel_selected"])
        self.check_all_selected()

    def select_images_csv(self):
        """
        Abre un diálogo para seleccionar el archivo CSV de imágenes.
        Actualiza la etiqueta de la interfaz y verifica si todos los archivos están seleccionados.
        """
        self.images_csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.images_csv_path:
            self.images_label.configure(text=os.path.basename(self.images_csv_path))
        else:
            self.images_label.configure(text=self.translations[self.lang]["no_csv_selected"])
        self.check_all_selected()

    def select_deployments_csv(self):
        """
        Abre un diálogo para seleccionar el archivo CSV de despliegues.
        Actualiza la etiqueta de la interfaz y verifica si todos los archivos están seleccionados.
        """
        self.deployments_csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.deployments_csv_path:
            self.deployments_label.configure(text=os.path.basename(self.deployments_csv_path))
        else:
            self.deployments_label.configure(text=self.translations[self.lang]["no_csv_selected"])
        self.check_all_selected()

    def check_all_selected(self):
        """
        Verifica si los tres archivos (Excel inicial, CSV de imágenes, CSV de despliegues)
        han sido seleccionados. Habilita o deshabilita el botón de "Procesar Archivos"
        en consecuencia.
        """
        if self.initial_excel_path and self.images_csv_path and self.deployments_csv_path:
            self.process_btn.configure(state=ctk.NORMAL)
        else:
            self.process_btn.configure(state=ctk.DISABLED)


class DateChangerApp(DateJobs):
    def __init__(self, root, lang="es"):
        self.root = root
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "select_folder": "Selecciona la carpeta que contiene los archivos",
                "browse_folder": "Buscar Carpeta",
                "no_folder_selected": "No se ha seleccionado ninguna carpeta",
                "set_real_date": "Establece la Fecha Real (AAAA-MM-DD HH:MM:SS)",
                "use_newest_date": "Usar la fecha más reciente",
                "use_oldest_date": "Usar la fecha más antigua",
                "use_custom_date": "Usar fecha personalizada",
                "rewrite_dates": "Reescribir fechas de los archivos",
                "copy_to_folder": "Copiar a la Carpeta",
                "success_message": "¡Fechas de los archivos actualizadas con éxito!",
            },
            "pt": {
                "select_folder": "Selecione a pasta que contém os arquivos",
                "browse_folder": "Procurar Pasta",
                "no_folder_selected": "Nenhuma pasta selecionada",
                "set_real_date": "Defina a Data Real (AAAA-MM-DD HH:MM:SS)",
                "use_newest_date": "Usar a data mais recente",
                "use_oldest_date": "Usar a data mais antiga",
                "use_custom_date": "Usar data personalizada",
                "rewrite_dates": "Reescrever datas dos arquivos",
                "copy_to_folder": "Copiar para a Pasta",
                "success_message": "Datas dos arquivos atualizadas com sucesso!",
            },
            "en": {
                "select_folder": "Select the folder containing files",
                "browse_folder": "Browse Folder",
                "no_folder_selected": "No folder selected",
                "set_real_date": "Set the Real Date (YYYY-MM-DD HH:MM:SS)",
                "use_newest_date": "Use newest date",
                "use_oldest_date": "Use oldest date",
                "use_custom_date": "Use custom date",
                "rewrite_dates": "Rewrite files Change Dates",
                "copy_to_folder": "Copy to Folder",
                "success_message": "File dates updated successfully!",
            }
        }

        # Configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtener las traducciones según el idioma seleccionado
        tr = self.translations[self.lang]

        # Frame principal
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Label para selección de carpeta
        self.label1 = ctk.CTkLabel(self.main_frame, text=tr["select_folder"], anchor="w")
        self.label1.pack(pady=10)

        # Frame para la selección de carpeta
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        # Botón para seleccionar carpeta
        self.select_folder_btn = ctk.CTkButton(self.folder_frame, text=tr["browse_folder"], command=self.select_folder)
        self.select_folder_btn.pack(side="left", padx=5)

        # Label para mostrar la carpeta seleccionada
        self.folder_label = ctk.CTkLabel(self.folder_frame, text=tr["no_folder_selected"], anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Frame para las opciones de fecha
        self.date_frame = ctk.CTkFrame(self.main_frame)
        self.date_frame.pack(pady=10, padx=10, fill="x")

        # Label y entrada para la fecha real
        self.label2 = ctk.CTkLabel(self.date_frame, text=tr["set_real_date"], anchor="w")
        self.label2.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.real_date_entry = ctk.CTkEntry(self.date_frame)
        self.real_date_entry.grid(row=0, column=1, padx=10, pady=5, sticky="we")

        # Radio buttons y etiquetas para opciones de fecha
        self.date_option = ctk.IntVar()
        self.date_option.set(1)

        self.newest_date_radio = ctk.CTkRadioButton(self.date_frame, text=tr["use_newest_date"], variable=self.date_option, value=1)
        self.newest_date_radio.grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.oldest_date_radio = ctk.CTkRadioButton(self.date_frame, text=tr["use_oldest_date"], variable=self.date_option, value=2)
        self.oldest_date_radio.grid(row=2, column=0, padx=5, pady=2, sticky="w")

        self.custom_date_radio = ctk.CTkRadioButton(self.date_frame, text=tr["use_custom_date"], variable=self.date_option, value=3)
        self.custom_date_radio.grid(row=3, column=0, padx=5, pady=2, sticky="w")

        self.custom_date_entry = ctk.CTkEntry(self.date_frame)
        self.custom_date_entry.grid(row=3, column=1, padx=10, pady=2, sticky="we")

        # Frame para los botones de cambiar fechas y copiar
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(pady=10)

        # Botón para cambiar las fechas
        self.change_dates_btn = ctk.CTkButton(self.button_frame, text=tr["rewrite_dates"], command=self.change_dates)
        self.change_dates_btn.pack(side="left", padx=10)

        # Botón para copiar a otra carpeta
        self.copy_btn = ctk.CTkButton(self.button_frame, text=tr["copy_to_folder"], command=self.copy_to_folder)
        self.copy_btn.pack(side="left", padx=10)
        
    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            self.selected_folder = folder
            self.folder_label.configure(text=folder)
            self.get_file_dates(folder)
        else:
            self.folder_label.configure(text="No folder selected")

    def get_exif_date(self, filepath):
        try:
            exif_data = piexif.load(filepath)
            date_time_original = exif_data['Exif'].get(piexif.ExifIFD.DateTimeOriginal)
            if date_time_original:
                return datetime.strptime(date_time_original.decode('utf-8'), "%Y:%m:%d %H:%M:%S")
            return None
        except Exception as e:
            return None


    def get_date_difference(self):
        if not getattr(self, "selected_folder", None):
            raise ValueError("Please select a folder first.")
        real_date = datetime.strptime(self.real_date_entry.get(), "%Y-%m-%d %H:%M:%S")
        options = {1: self.newest_date, 2: self.oldest_date, 3: self.custom_date_entry}
        reference = datetime.strptime(options[self.date_option.get()].get(), "%Y-%m-%d %H:%M:%S")
        return real_date - reference


class ImagesRenamer(RenamerJobs):
    def __init__(self, root, lang="es"):
        self.root = root
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "select_source_directory": "Selecciona el directorio de origen:",
                "browse_source": "Buscar Origen",
                "no_source_folder_selected": "No se ha seleccionado ninguna carpeta de origen",
                "settings": "Configuración",
                "use_folder_name": "Usar nombre de carpeta",
                "keep_original_filename": "Mantener nombre de archivo original",
                "add_exif_date": "Agregar fecha EXIF",
                "replace_spaces": "Reemplazar espacios",
                "use_custom_text": "Usar texto personalizado",
                "copy_photos": "Copiar fotos a la carpeta de destino",
                "select_destination_directory": "Selecciona el directorio de destino si deseas mover las fotos:",
                "browse_destination": "Buscar Destino",
                "no_destination_folder_selected": "No se ha seleccionado ninguna carpeta de destino",
                "processing": "Procesando...",
                "rename": "Renombrar",
                "success_message": "Las fotos han sido renombradas y procesadas con éxito.",
                "information": "Información"
            },
            "pt": {
                "select_source_directory": "Selecione o diretório de origem:",
                "browse_source": "Procurar Origem",
                "no_source_folder_selected": "Nenhuma pasta de origem selecionada",
                "settings": "Configurações",
                "use_folder_name": "Usar nome da pasta",
                "keep_original_filename": "Manter nome de arquivo original",
                "add_exif_date": "Adicionar data EXIF",
                "replace_spaces": "Substituir espaços",
                "use_custom_text": "Usar texto personalizado",
                "copy_photos": "Copiar fotos para a pasta de destino",
                "select_destination_directory": "Selecione o diretório de destino se deseja mover as fotos:",
                "browse_destination": "Procurar Destino",
                "no_destination_folder_selected": "Nenhuma pasta de destino selecionada",
                "processing": "Processando...",
                "rename": "Renomear",
                "success_message": "As fotos foram renomeadas e processadas com sucesso.",
                "information": "Informação"
            },
            "en": {
                "select_source_directory": "Select the source directory:",
                "browse_source": "Browse Source",
                "no_source_folder_selected": "No source folder selected",
                "settings": "Settings",
                "use_folder_name": "Use folder name",
                "keep_original_filename": "Keep original filename",
                "add_exif_date": "Add EXIF date",
                "replace_spaces": "Replace spaces",
                "use_custom_text": "Use custom text",
                "copy_photos": "Copy photos to destination folder",
                "select_destination_directory": "Select the destination directory if you want to move the photos:",
                "browse_destination": "Browse Destination",
                "no_destination_folder_selected": "No destination folder selected",
                "processing": "Processing...",
                "rename": "Rename",
                "success_message": "Photos have been renamed and processed successfully.",
                "information": "Information"
            }
        }

        # Configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtener las traducciones según el idioma seleccionado
        tr = self.translations[self.lang]

        # Frame principal
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Sección: Directorio de origen
        self.label1 = ctk.CTkLabel(self.main_frame, text=tr["select_source_directory"], anchor="w")
        self.label1.pack(pady=10)

        self.source_folder_frame = ctk.CTkFrame(self.main_frame)
        self.source_folder_frame.pack(pady=5, padx=10, fill="x")

        self.select_source_folder_btn = ctk.CTkButton(self.source_folder_frame, text=tr["browse_source"], command=self.select_source_folder)
        self.select_source_folder_btn.pack(side="left", padx=5)

        self.source_folder_label = ctk.CTkLabel(self.source_folder_frame, text=tr["no_source_folder_selected"], anchor="w")
        self.source_folder_label.pack(side="left", padx=5)

        # Sección: Configuración
        self.label2 = ctk.CTkLabel(self.main_frame, text=tr["settings"], anchor="w")
        self.label2.pack(pady=10)

        self.checkbox_frame = ctk.CTkFrame(self.main_frame)
        self.checkbox_frame.pack(pady=10, padx=10, fill="x")

        self.use_folder_name_var = BooleanVar(value=True)
        self.checkbox_use_folder_name = ctk.CTkCheckBox(self.checkbox_frame, text=tr["use_folder_name"], variable=self.use_folder_name_var)
        self.checkbox_use_folder_name.pack(side="left", padx=5)

        self.keep_original_name_var = BooleanVar(value=False)
        self.checkbox_keep_original_name = ctk.CTkCheckBox(self.checkbox_frame, text=tr["keep_original_filename"], variable=self.keep_original_name_var)
        self.checkbox_keep_original_name.pack(side="left", padx=5)

        self.add_exif_date_var = BooleanVar(value=False)
        self.checkbox_add_exif_date = ctk.CTkCheckBox(self.checkbox_frame, text=tr["add_exif_date"], variable=self.add_exif_date_var)
        self.checkbox_add_exif_date.pack(side="left", padx=5)

        self.replace_spaces_var = BooleanVar(value=False)
        self.checkbox_replace_spaces = ctk.CTkCheckBox(self.checkbox_frame, text=tr["replace_spaces"], variable=self.replace_spaces_var)
        self.checkbox_replace_spaces.pack(side="left", padx=5)

        self.custom_text_frame = ctk.CTkFrame(self.main_frame)
        self.custom_text_frame.pack(pady=10, padx=10, fill="x")

        self.use_custom_text_var = BooleanVar(value=False)
        self.checkbox_use_custom_text = ctk.CTkCheckBox(self.custom_text_frame, text=tr["use_custom_text"], variable=self.use_custom_text_var, command=self.toggle_custom_text_entry)
        self.checkbox_use_custom_text.pack(side="left", padx=5)

        self.custom_text_var = StringVar(value="")
        self.custom_text_entry = ctk.CTkEntry(self.custom_text_frame, textvariable=self.custom_text_var, state="disabled", width=200)
        self.custom_text_entry.pack(side="left", padx=5)

        self.copy_photos_var = BooleanVar(value=True)
        self.checkbox_copy_photos = ctk.CTkCheckBox(self.custom_text_frame, text=tr["copy_photos"], variable=self.copy_photos_var)
        self.checkbox_copy_photos.pack(side="left", padx=5)

        # Sección: Directorio de destino
        self.label3 = ctk.CTkLabel(self.main_frame, text=tr["select_destination_directory"], anchor="w")
        self.label3.pack(pady=10)

        self.dest_folder_frame = ctk.CTkFrame(self.main_frame)
        self.dest_folder_frame.pack(pady=5, padx=10, fill="x")

        self.select_dest_folder_btn = ctk.CTkButton(self.dest_folder_frame, text=tr["browse_destination"], command=self.select_dest_folder)
        self.select_dest_folder_btn.pack(side="left", padx=5)

        self.dest_folder_label = ctk.CTkLabel(self.dest_folder_frame, text=tr["no_destination_folder_selected"], anchor="w")
        self.dest_folder_label.pack(side="left", padx=5)

        # Etiqueta de procesamiento
        self.processing_label = ctk.CTkLabel(self.main_frame, text="", anchor="w")
        self.processing_label.pack(pady=10)

        # Botón para renombrar y copiar
        self.rename_and_copy_btn = ctk.CTkButton(self.main_frame, text=tr["rename"], command=self.rename_and_copy_photos)
        self.rename_and_copy_btn.pack(pady=20)

        self.source_folder = None
        self.dest_folder = None

    def toggle_custom_text_entry(self):
        if self.use_custom_text_var.get():
            self.custom_text_entry.configure(state="normal")
        else:
            self.custom_text_entry.configure(state="disabled")

    def select_source_folder(self):
        self.source_folder = filedialog.askdirectory()
        if self.source_folder:
            self.source_folder_label.configure(text=self.source_folder)
            messagebox.showinfo(self.translations[self.lang]["information"], f"{self.translations[self.lang]['select_source_directory']} {self.source_folder}")

    def select_dest_folder(self):
        self.dest_folder = filedialog.askdirectory()
        if self.dest_folder:
            self.dest_folder_label.configure(text=self.dest_folder)
            messagebox.showinfo(self.translations[self.lang]["information"], f"{self.translations[self.lang]['select_destination_directory']} {self.dest_folder}")


    def get_exif_date(self, file_path):
        try:
            image = Image.open(file_path)
            exif_data = image._getexif()
            if exif_data:
                for tag, value in exif_data.items():
                    tag_name = TAGS.get(tag, tag)
                    if tag_name == "DateTimeOriginal":
                        # Format the EXIF date
                        exif_date = value.replace(":", "-").replace(" ", "_")
                        return exif_date
        except Exception as e:
            print(f"Error getting EXIF data from {file_path}: {e}")
        return None


class LynxOne(LynxJobs):
    def __init__(self, root, lang="es"):
        self.root = root
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "source_directory": "Selecciona el directorio de origen.",
                "browse_source": "Buscar Origen",
                "no_source_folder_selected": "No se ha seleccionado ninguna carpeta de origen",
                "settings": "Configuración para agrupar imágenes.",
                "group_by_minutes": "Agrupar por minutos:",
                "linces_folder_exists": "¿Existe la carpeta Lince/Linces?",
                "revision_folder_exists": "¿Existe la carpeta Revisión?",
                "optional_files": "Opcional: Selecciona archivos Excel adicionales para unir.",
                "estaciones_file": "No se ha seleccionado ningún archivo de Estaciones",
                "browse_estaciones": "Buscar Estaciones",
                "individuos_file": "No se ha seleccionado ningún archivo de Individuos",
                "browse_individuos": "Buscar Individuos",
                "generate_excel": "Generar Excel",
                "download_excel": "Descargar Excel",
                "warning": "Advertencia",
                "info": "Información",
                "success_message": "Archivo Excel guardado con éxito en {file_path}",
                "error_message": "Error: ",
                "no_source_folder": "Por favor, selecciona una carpeta de origen primero."
            },
            "pt": {
                "source_directory": "Selecione o diretório de origem.",
                "browse_source": "Procurar Origem",
                "no_source_folder_selected": "Nenhuma pasta de origem selecionada",
                "settings": "Configurações para agrupar imagens.",
                "group_by_minutes": "Agrupar por minutos:",
                "linces_folder_exists": "A pasta Lince/Linces existe?",
                "revision_folder_exists": "A pasta Revisão existe?",
                "optional_files": "Opcional: Selecione arquivos Excel adicionais para unir.",
                "estaciones_file": "Nenhum arquivo Estaciones selecionado",
                "browse_estaciones": "Procurar Estaciones",
                "individuos_file": "Nenhum arquivo Individuos selecionado",
                "browse_individuos": "Procurar Individuos",
                "generate_excel": "Gerar Excel",
                "download_excel": "Baixar Excel",
                "warning": "Aviso",
                "info": "Informação",
                "success_message": "Arquivo Excel salvo com sucesso em {file_path}",
                "error_message": "Erro: ",
                "no_source_folder": "Por favor, selecione uma pasta de origem primeiro."
            },
            "en": {
                "source_directory": "Select the source directory.",
                "browse_source": "Browse Source",
                "no_source_folder_selected": "No source folder selected",
                "settings": "Settings for grouping images.",
                "group_by_minutes": "Group by minutes:",
                "linces_folder_exists": "Does the lynx/lynxes folder exist?",
                "revision_folder_exists": "Does the develoment folder exist?",
                "optional_files": "Optional: Select additional Excel files for joining.",
                "estaciones_file": "No Station file selected",
                "browse_estaciones": "Browse Station",
                "individuos_file": "No Individuals file selected",
                "browse_individuos": "Browse Individuals",
                "generate_excel": "Generate Excel",
                "download_excel": "Download Excel",
                "warning": "Warning",
                "info": "Information",
                "success_message": "Excel file successfully saved at {file_path}",
                "error_message": "Error: ",
                "no_source_folder": "Please select a source folder first."
            }
        }

        # Configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtener las traducciones según el idioma seleccionado
        tr = self.translations[self.lang]

        # Crear el main_frame
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True)

        # Crear el panel de pestañas
        self.tabview = ctk.CTkTabview(self.main_frame, width=400, height=300)
        self.tabview.pack(pady=20, padx=20, fill="both", expand=True)
        self.tabview.pack_propagate(False)

        # Agregar pestañas
        self.source_tab = self.tabview.add("Source")
        self.settings_tab = self.tabview.add("Settings")
        self.optional_tab = self.tabview.add("Optional Files")

        # Contenido de la pestaña "Source"
        self.label1 = ctk.CTkLabel(self.source_tab, text=tr["source_directory"], anchor="w")
        self.label1.pack(pady=10)

        self.source_folder_frame = ctk.CTkFrame(self.source_tab)
        self.source_folder_frame.pack(pady=5, padx=10, fill="x")

        self.select_source_folder_btn = ctk.CTkButton(self.source_folder_frame, text=tr["browse_source"], command=self.select_source_folder)
        self.select_source_folder_btn.pack(side="left", padx=5)

        self.source_folder_label = ctk.CTkLabel(self.source_folder_frame, text=tr["no_source_folder_selected"], anchor="w")
        self.source_folder_label.pack(side="left", padx=5)

        # Contenido de la pestaña "Settings"
        self.label2 = ctk.CTkLabel(self.settings_tab, text=tr["settings"], anchor="w")
        self.label2.pack(pady=10)

        self.minutes_frame = ctk.CTkFrame(self.settings_tab)
        self.minutes_frame.pack(pady=5, padx=10, fill="x")

        self.minutes_label = ctk.CTkLabel(self.minutes_frame, text=tr["group_by_minutes"])
        self.minutes_label.pack(side="left", padx=5)

        self.minutes_entry = ctk.CTkEntry(self.minutes_frame, placeholder_text="Minutes (0 to disable)")
        self.minutes_entry.pack(side="left", padx=5)

        self.lince_checkbox_var = ctk.BooleanVar(value=True)
        self.checkbox_lince_exists = ctk.CTkCheckBox(self.settings_tab, text=tr["linces_folder_exists"], variable=self.lince_checkbox_var)
        self.checkbox_lince_exists.pack(pady=10)

        self.revision_checkbox_var = ctk.BooleanVar(value=False)
        self.checkbox_revision_exists = ctk.CTkCheckBox(self.settings_tab, text=tr["revision_folder_exists"], variable=self.revision_checkbox_var)
        self.checkbox_revision_exists.pack(pady=10)

        # Contenido de la pestaña "Optional Files"
        self.label3 = ctk.CTkLabel(self.optional_tab, text=tr["optional_files"], anchor="w")
        self.label3.pack(pady=10)

        self.estaciones_frame = ctk.CTkFrame(self.optional_tab)
        self.estaciones_frame.pack(pady=5, padx=10, fill="x")

        self.estaciones_file_label = ctk.CTkLabel(self.estaciones_frame, text=tr["estaciones_file"], anchor="w")
        self.estaciones_file_label.pack(side="left", padx=5)

        self.select_estaciones_file_btn = ctk.CTkButton(self.estaciones_frame, text=tr["browse_estaciones"], command=self.select_estaciones_file)
        self.select_estaciones_file_btn.pack(side="left", padx=5)

        self.individuos_frame = ctk.CTkFrame(self.optional_tab)
        self.individuos_frame.pack(pady=5, padx=10, fill="x")

        self.individuos_file_label = ctk.CTkLabel(self.individuos_frame, text=tr["individuos_file"], anchor="w")
        self.individuos_file_label.pack(side="left", padx=5)

        self.select_individuos_file_btn = ctk.CTkButton(self.individuos_frame, text=tr["browse_individuos"], command=self.select_individuos_file)
        self.select_individuos_file_btn.pack(side="left", padx=5)

        self.source_folder = None
        self.estaciones_file = None
        self.individuos_file = None
        self.excel_data = None

        # Frame for Process and Download Buttons
        self.buttons_frame = ctk.CTkFrame(self.main_frame)
        self.buttons_frame.pack(pady=10)

        # Botones para generar y descargar archivos Excel
        self.generate_button = ctk.CTkButton(self.buttons_frame, text=tr["generate_excel"], command=self.generate_excel)
        self.generate_button.pack(side="left", padx=5)

        self.download_button = ctk.CTkButton(self.buttons_frame, text=tr["download_excel"], command=self.save_excel)
        self.download_button.pack(side="left", padx=5)

    def select_source_folder(self):
        self.source_folder = filedialog.askdirectory()
        if self.source_folder:
            self.source_folder_label.configure(text=self.source_folder)
            messagebox.showinfo(self.translations[self.lang]["info"], f"{self.translations[self.lang]['source_directory']} {self.source_folder}")

    def select_estaciones_file(self):
        self.estaciones_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.estaciones_file:
            self.estaciones_file_label.configure(text=self.estaciones_file)
            messagebox.showinfo(self.translations[self.lang]["info"], f"{self.translations[self.lang]['estaciones_file']} {self.estaciones_file}")

    def select_individuos_file(self):
        self.individuos_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.individuos_file:
            self.individuos_file_label.configure(text=self.individuos_file)
            messagebox.showinfo(self.translations[self.lang]["info"], f"{self.translations[self.lang]['individuos_file']} {self.individuos_file}")


   

        
if __name__ == "__main__":
    root = ctk.CTk()
    app = BaseApp(root)
    smoke_marker = os.environ.get("LYNXAUTOMATOR_SMOKE_TEST")
    if smoke_marker:
        def finish_smoke_test():
            if app.presentation_app.logo is None:
                root.destroy()
                return
            Path(smoke_marker).write_text("started", encoding="utf-8")
            root.destroy()
        root.after(500, finish_smoke_test)
    root.mainloop()
