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
import win32file
import pywintypes
import tkinter as tk
import threading
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
                "wiwbe_folder": "WIWbE desde Carpeta",
                "wiwbe_catalog": "Catálogo WIWbE",
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
                "wiwbe_folder": "WIWbE da Pasta",
                "wiwbe_catalog": "Catálogo WIWbE",
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
                "wiwbe_folder": "WIWbE from Folder",
                "wiwbe_catalog": "WIWbE Catalog",
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
        self.main_tabs.add(self.about_frame, text=tr["about"])
        self.main_tabs.add(self.wildbook_frame, text=tr["wildbook"])
        self.main_tabs.add(self.wildlife_insights_frame, text=tr["wildlife_insights"])
        self.main_tabs.add(self.wildlife_insights_wildbook_frame, text=tr["wi_wildbook"])
        self.main_tabs.add(self.iberian_lynx_frame, text=tr["iberian_lynx"])
        self.main_tabs.add(self.camtrap_frame, text=tr["functionalities"])

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
        # Limpiar y reconstruir la interfaz cuando se cambia el idioma
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
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "title": "LynxAutomator",
                "description": "Esta aplicación forma parte del proyecto LIFE+ Lynxconnect"
            },
            "pt": {
                "title": "LynxAutomator",
                "description": "Este aplicativo faz parte do projeto LIFE+ Lynxconnect"
            },
            "en": {
                "title": "LynxAutomator",
                "description": "This app is part of the LIFE+ Lynxconnect project"
            }
        }

        # Configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtener las traducciones según el idioma seleccionado
        tr = self.translations[self.lang]

        self.pack(fill="both", expand=True)

        # # Crear widgets
        # logo_path = "ruta/a/tu/logo.png"  # Asegúrate de actualizar esto con la ruta correcta
        # logo_image = Image.open(logo_path)

        # self.logo = ImageTk.PhotoImage(logo_image)
        # logo_label = tk.Label(self, image=self.logo)
        # logo_label.pack(pady=20)

        # text_label = ctk.CTkLabel(self, text=tr["description"], font=("Helvetica", 16))
        # text_label.pack(pady=10)


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


class WBFolderApp:
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
                "success_message": "¡Archivo Excel procesado con éxito!"
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
                "success_message": "Arquivo Excel processado com sucesso!"
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
                "success_message": "Excel file processed successfully!"
            }
        }

        # Resto de la configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtener traducciones según el idioma seleccionado
        tr = self.translations[self.lang]

        # Marco principal (cuadrado grande único)
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Etiqueta para solicitar acción al usuario
        self.label = ctk.CTkLabel(self.main_frame, text=tr["select_folder"], anchor="w")
        self.label.pack(pady=10)

        # Marco para contener los widgets de selección de carpeta
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        # Botón para buscar y seleccionar una carpeta
        self.select_folder_btn = ctk.CTkButton(self.folder_frame, text=tr["browse_folder"], command=self.select_folder)
        self.select_folder_btn.pack(side="left", padx=5)

        # Etiqueta para mostrar el nombre de la carpeta seleccionada
        self.folder_label = ctk.CTkLabel(self.folder_frame, text=tr["no_folder_selected"], anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Mensaje entre los dos botones
        self.message_label = ctk.CTkLabel(self.main_frame, text=tr["select_excel"])
        self.message_label.pack(pady=10)

        # Marco para contener los widgets de selección de archivo
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(pady=5, padx=10, fill="x")

        # Botón para buscar y seleccionar un archivo Excel
        self.select_file_btn = ctk.CTkButton(self.file_frame, text=tr["browse_file"], command=self.select_file)
        self.select_file_btn.pack(side="left", padx=5)

        # Etiqueta para mostrar el nombre del archivo seleccionado
        self.file_label = ctk.CTkLabel(self.file_frame, text=tr["no_file_selected"], anchor="w")
        self.file_label.pack(side="left", padx=5)

        # Botón de procesar
        self.process_btn = ctk.CTkButton(self.main_frame, text=tr["process"], command=self.process_files)
        self.process_btn.pack(pady=10)

        # Botón de descargar
        self.download_btn = ctk.CTkButton(self.main_frame, text=tr["download_excel"], command=self.download_file, state=ctk.DISABLED)
        self.download_btn.pack(pady=10)

        # Variable para almacenar la ruta del archivo temporal
        self.temp_file_path = None

    def select_folder(self):
        # Open a dialog to select a folder
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.folder_label.configure(text=os.path.basename(self.folder_path))
            messagebox.showinfo("Information", f"Selected folder: {self.folder_path}")

    def select_file(self):
        # Open a dialog to select an Excel file
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.file_label.configure(text=os.path.basename(self.file_path))
            messagebox.showinfo("Information", f"Selected file: {self.file_path}")

    def get_exif_data(self, image_path):
        # Retrieve EXIF data from an image
        image = Image.open(image_path)
        exif_data = image._getexif()
        if not exif_data:
            return None
        exif = {}
        for tag, value in exif_data.items():
            decoded = TAGS.get(tag, tag)
            exif[decoded] = value
        return exif

    def get_date_taken(self, exif_data):
        # Extract the date when the photo was taken from EXIF data
        date_taken = exif_data.get("DateTimeOriginal")
        if date_taken:
            return datetime.strptime(date_taken, '%Y:%m:%d %H:%M:%S')
        return None

    def process_files(self):
        # Ensure both folder and file are selected
        if not hasattr(self, 'folder_path') or not hasattr(self, 'file_path'):
            messagebox.showerror("Error", "Please select both a folder and a file.")
            return

        photo_data = []

        # Iterate over files in the selected folder
        for filename in os.listdir(self.folder_path):
            if filename.lower().endswith(('png', 'jpg', 'jpeg')):
                file_path = os.path.join(self.folder_path, filename)
                exif_data = self.get_exif_data(file_path)
                if exif_data:
                    date_taken = self.get_date_taken(exif_data)
                    if date_taken:
                        photo_data.append([filename, date_taken])

        if not photo_data:
            messagebox.showinfo("Information", "No photos with date information found.")
            return

        df = pd.read_excel(self.file_path)
        
        # Ensure the 'Encounter.mediaAsset0' column exists, create it if it doesn't
        if 'Encounter.mediaAsset0' not in df.columns:
            df['Encounter.mediaAsset0'] = ""

        df['Encounter.mediaAsset0'] = df['Encounter.mediaAsset0'].astype(str)  # Ensure column is string type

        new_data = pd.DataFrame(photo_data, columns=['Encounter.mediaAsset0', 'Date'])
        new_data['Encounter.mediaAsset0'] = new_data['Encounter.mediaAsset0'].astype(str) 
        
        # Convert 'Date' column to datetime and sort data by date
        new_data['Date'] = pd.to_datetime(new_data['Date'])
        new_data.sort_values(by='Date', inplace=True)
        
        # Determine time threshold for grouping
        time_threshold = int(self.time_threshold_entry.get())
        new_data['TimeDiff'] = new_data['Date'].diff().dt.total_seconds().fillna(0)
        
        grouped_data = []
        current_group = []
        for index, row in new_data.iterrows():
            if current_group and row['TimeDiff'] > time_threshold:
                grouped_data.append(current_group)
                current_group = []
            current_group.append(row)
        if current_group:
            grouped_data.append(current_group)
        
        # Create a new DataFrame to store the grouped data
        final_data = []
        for group in grouped_data:
            base_row = group[0].copy()
            for i, additional_row in enumerate(group[1:], start=1):
                base_row[f'Encounter.mediaAsset{i}'] = additional_row['Encounter.mediaAsset0']
            final_data.append(base_row)
        
        final_df = pd.DataFrame(final_data).drop(columns=['TimeDiff'])

        # Add date-related columns
        final_df['Encounter.year'] = final_df['Date'].dt.year
        final_df['Encounter.month'] = final_df['Date'].dt.month
        final_df['Encounter.day'] = final_df['Date'].dt.day
        final_df['Encounter.hour'] = final_df['Date'].dt.hour
        final_df['Encounter.minutes'] = final_df['Date'].dt.minute
        final_df = final_df.drop('Date', axis=1)

        # Merge with original DataFrame
        df_merged = pd.merge(df, final_df, on='Encounter.mediaAsset0', how='right')

        # Fill missing data in original DataFrame columns with the first row value
        for column in df.columns:
            if column in df_merged.columns:
                first_row_value = df[column].iloc[0]
                df_merged[column] = df_merged[column].fillna(first_row_value)

        # Ensure only one set of date-related columns appears
        for time_unit in ['year', 'month', 'day', 'hour', 'minutes']:  
            column_x = f'Encounter.{time_unit}_x'
            column_y = f'Encounter.{time_unit}_y'
            column = f'Encounter.{time_unit}'
            if column_x in df_merged.columns and column_y in df_merged.columns:
                # Remove empty entries before combining
                df_merged[column_x].replace('', pd.NA, inplace=True)
                df_merged[column_y].replace('', pd.NA, inplace=True)
                df_merged[column] = df_merged[column_x].combine_first(df_merged[column_y])
                df_merged.drop([column_x, column_y], axis=1, inplace=True)
            elif column_y in df_merged.columns:
                df_merged.rename(columns={column_y: column}, inplace=True)

        # Create a temporary file to save the updated Excel data
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            self.temp_file_path = tmp.name
            df_merged.to_excel(self.temp_file_path, index=False)

        messagebox.showinfo("Information", "Excel file processed successfully!")
        self.download_btn.configure(state=ctk.NORMAL)

    def download_file(self):
        # Open a dialog to save the processed Excel file
        if self.temp_file_path:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                os.rename(self.temp_file_path, save_path)
                messagebox.showinfo("Information", f"Excel file saved successfully at {save_path}!")
                self.download_btn.configure(state=ctk.DISABLED)
                self.temp_file_path = None



class WBCatalogApp:
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

    def process_files(self):
        # Ensure both folder and file are selected
        if not hasattr(self, 'folder_path') or not hasattr(self, 'file_path'):
            messagebox.showerror("Error", "Please select both a folder and a file.")
            return

        photo_data = []

        # Iterate over files in the selected folder
        for root, dirs, files in os.walk(self.folder_path):
            # Use the first word of the filename without the extension as the individual name
            individual_name = {
                file: os.path.splitext(file)[0].split()[0] for file in files if file.lower().endswith(('png', 'jpg', 'jpeg'))
            }

            # Capitalize individualID if the checkbox is checked
            if self.capitalize_var.get():
                individual_name = {file: name.capitalize() for file, name in individual_name.items()}

            for file in files:
                if file.lower().endswith(('png', 'jpg', 'jpeg')):
                    file_path = os.path.join(root, file)
                    relative_path = os.path.relpath(file_path, self.folder_path)  # Get the relative path
                    relative_path = relative_path.replace("\\", "/")  # Replace backslashes with forward slashes
                    name_used = individual_name[file]
                    photo_data.append([relative_path, name_used])

        if not photo_data:
            messagebox.showinfo("Information", "No photos found.")
            return

        try:
            df_original = pd.read_excel(self.file_path)
            original_columns = df_original.columns.tolist()  # Save the original column order
            first_row_data = df_original.iloc[0]  # Get the first row to propagate values
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
            return

        # Check if the 'Encounter.mediaAsset0' column exists
        if 'Encounter.mediaAsset0' not in df_original.columns:
            df_original['Encounter.mediaAsset0'] = pd.NA

        df_original['Encounter.mediaAsset0'] = df_original['Encounter.mediaAsset0'].astype(str)
        
        new_data = pd.DataFrame(photo_data, columns=['Encounter.mediaAsset0', 'MarkedIndividual.individualID'])
        new_data['Encounter.mediaAsset0'] = new_data['Encounter.mediaAsset0'].astype(str)

        # Perform the merge operation
        df_merged = pd.merge(df_original, new_data, on='Encounter.mediaAsset0', how='right')

        # Ensure 'MarkedIndividual.individualID' is correctly placed
        if 'MarkedIndividual.individualID_x' in df_merged.columns and 'MarkedIndividual.individualID_y' in df_merged.columns:
            df_merged['MarkedIndividual.individualID'] = df_merged['MarkedIndividual.individualID_y'].fillna(df_merged['MarkedIndividual.individualID_x'])
            df_merged.drop(['MarkedIndividual.individualID_x', 'MarkedIndividual.individualID_y'], axis=1, inplace=True)
        elif 'MarkedIndividual.individualID_y' in df_merged.columns:
            df_merged.rename(columns={'MarkedIndividual.individualID_y': 'MarkedIndividual.individualID'}, inplace=True)
        elif 'MarkedIndividual.individualID_x' in df_merged.columns:
            df_merged.rename(columns={'MarkedIndividual.individualID_x': 'MarkedIndividual.individualID'}, inplace=True)

        if self.collapse_var.get():
            # Collapse rows with the same 'MarkedIndividual.individualID'
            df_merged['RowNumber'] = df_merged.groupby('MarkedIndividual.individualID').cumcount()
            df_pivot = df_merged.pivot_table(index='MarkedIndividual.individualID', columns='RowNumber', values='Encounter.mediaAsset0', aggfunc='first')
            df_pivot.columns = [f'Encounter.mediaAsset{int(col)}' for col in df_pivot.columns]
            df_merged = pd.merge(df_merged.drop(columns='Encounter.mediaAsset0').drop_duplicates('MarkedIndividual.individualID'), df_pivot, on='MarkedIndividual.individualID')

            # Drop the 'RowNumber' column after use
            df_merged.drop(columns=['RowNumber'], inplace=True)

        # Reorder columns to match the original Excel file order
        new_column_order = [col for col in original_columns if col in df_merged.columns] + \
                        [col for col in df_merged.columns if col not in original_columns]
        df_merged = df_merged[new_column_order]

        # Fill in missing values from the first row where appropriate
        for column in df_merged.columns:
            if df_merged[column].isnull().any():
                if column in first_row_data.index:  # Corrected to check the index
                    df_merged[column] = df_merged[column].fillna(first_row_data[column])

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            self.temp_file_path = tmp.name
            df_merged.to_excel(self.temp_file_path, index=False)

        messagebox.showinfo("Information", "Excel file processed successfully!")
        self.download_btn.configure(state=ctk.NORMAL)

    def download_file(self):
        if self.temp_file_path:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                os.rename(self.temp_file_path, save_path)
                messagebox.showinfo("Information", f"Excel file saved successfully at {save_path}!")
                self.download_btn.configure(state=ctk.DISABLED)
                self.temp_file_path = None



class FrameExtractorApp:
    def __init__(self, root, lang="es"):
        self.root = root
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "select_videos": "Selecciona la carpeta que contiene los videos.",
                "browse_folder": "Buscar Carpeta",
                "no_folder_selected": "No se ha seleccionado ninguna carpeta",
                "interval_between_frames": "Intervalo entre cuadros (segundos):",
                "extract_frames": "Extraer cuadros",
                "processing": "Procesando...",
                "success_message": "Cuadros extraídos y guardados con éxito.",
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

    def start_extraction(self):
        try:
            interval = float(self.interval_var.get())
            if interval <= 0:
                raise ValueError("The interval must be greater than zero.")
            
            if not self.folder_path:
                self.status_label.configure(text="Please select a folder with videos.")
                return
            
            # Ask the user to select the output folder
            self.output_folder = filedialog.askdirectory(title="Select the output folder")
            if not self.output_folder:
                self.status_label.configure(text="Please select an output folder.")
                return

            # Mostrar mensaje de "Processing..." antes de comenzar
            self.status_label.configure(text="Processing...")
            self.status_label.update_idletasks()  # Forzar actualización de la UI

            total_frames = 0
            for filename in os.listdir(self.folder_path):
                if filename.lower().endswith(('.mp4', '.avi', '.mov', '.mkv', '.flv')):
                    video_path = os.path.join(self.folder_path, filename)
                    vidcap = cv2.VideoCapture(video_path)
                    if vidcap.isOpened():
                        total_frames += int(vidcap.get(cv2.CAP_PROP_FRAME_COUNT))
                        vidcap.release()

            self.status_label.configure(text="")

            for filename in os.listdir(self.folder_path):
                if filename.lower().endswith(('.mp4', '.avi', '.mov', '.mkv', '.flv')):
                    video_path = os.path.join(self.folder_path, filename)
                    self.extract_frames(video_path, interval)

            # Actualizar el mensaje una vez que el procesamiento haya terminado
            self.status_label.configure(text="Frames extracted and saved successfully.")

        except ValueError as e:
            self.status_label.configure(text=f"Error: {e}")


    def extract_frames(self, video_path, interval):
        vidcap = cv2.VideoCapture(video_path)
        if not vidcap.isOpened():
            print(f"Error opening video {video_path}")
            return
        fps = vidcap.get(cv2.CAP_PROP_FPS)
        success, image = vidcap.read()
        count = 0
        frame_number = 0
        creation_time = datetime.fromtimestamp(os.path.getctime(video_path))

        total_frames = int(vidcap.get(cv2.CAP_PROP_FRAME_COUNT))
        self.status_label.configure(text=f"Processing {os.path.basename(video_path)}")

        while success:
            if count % int(fps * interval) == 0:
                frame_path = os.path.join(self.output_folder, f"{os.path.basename(video_path)}_frame{frame_number}.jpg")
                cv2.imwrite(frame_path, image, [cv2.IMWRITE_JPEG_QUALITY, 100])
                self.change_file_dates(frame_path, creation_time)
                self.set_exif_date(frame_path, creation_time)
                frame_number += 1
            
            success, image = vidcap.read()
            count += 1

        vidcap.release()
        self.status_label.configure(text="Frames extracted and saved successfully.")

    def set_exif_date(self, image_path, creation_time):
        exif_date = creation_time.strftime("%Y:%m:%d %H:%M:%S")
        exif_dict = piexif.load(image_path)
        exif_dict['Exif'][piexif.ExifIFD.DateTimeOriginal] = exif_date
        exif_dict['Exif'][piexif.ExifIFD.DateTimeDigitized] = exif_date
        exif_bytes = piexif.dump(exif_dict)
        piexif.insert(exif_bytes, image_path)

    def change_file_dates(self, file_path, creation_time):
        try:
            os.utime(file_path, (creation_time.timestamp(), creation_time.timestamp()))
            if platform.system() == 'Windows':
                wintime = pywintypes.Time(creation_time.timestamp())
                fh = win32file.CreateFile(file_path, win32file.GENERIC_WRITE, 0, None, win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
                win32file.SetFileTime(fh, wintime, wintime, wintime)
                fh.close()
        except Exception as e:
            print(f"Error setting file dates: {e}")


class GCSDownloaderAndRenamer(ctk.CTkFrame):
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

        self.stop_btn = ctk.CTkButton(self.buttons_frame, text=self.translations[self.lang]["stop"], command=self.stop_download, state=ctk.DISABLED)
        self.stop_btn.pack(side="left", padx=5)

        self.stop_flag = False

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

    def start_download(self):
        # Open file dialog to select destination folder
        self.folder_path = filedialog.askdirectory(title="Select Destination Folder")
        if not self.folder_path:
            messagebox.showerror("Error", "No folder selected.")
            return

        # Ensure the destination folder exists
        if self.use_multiple_folders.get():
            deployment_folders = self.get_deployment_folders()
            for folder in deployment_folders:
                os.makedirs(folder, exist_ok=True)
        else:
            os.makedirs(self.folder_path, exist_ok=True)
        
        self.download_btn.configure(state=ctk.DISABLED)
        self.stop_btn.configure(state=ctk.NORMAL)
        self.stop_flag = False
        threading.Thread(target=self.download_and_rename_files).start()

    def stop_download(self):
        self.stop_flag = True
        self.download_btn.configure(state=ctk.NORMAL)
        self.stop_btn.configure(state=ctk.DISABLED)

    def clean_filename(self, name):
        # Replace any character that is not alphanumeric, dot, underscore, hyphen, space, or parentheses with an underscore
        return re.sub(r'[^a-zA-Z0-9._Ññ\-\(\) ]', '_', name)

    def clean_deployment_id(self, deployment_id):
        # Clean the deployment_id only
        return self.clean_filename(deployment_id)

    def get_deployment_folders(self):
        # Get unique deployment_ids from CSV and return their corresponding folder paths
        df = pd.read_csv(self.csv_path)
        deployment_ids = df['deployment_id'].unique()
        return [os.path.join(self.folder_path, self.clean_deployment_id(str(deployment_id))) for deployment_id in deployment_ids]

    def download_and_rename_files(self):
        if not hasattr(self, 'csv_path') or not self.csv_path:
            messagebox.showerror("Error", "Please select a CSV file.")
            self.download_btn.configure(state=ctk.NORMAL)
            self.stop_btn.configure(state=ctk.DISABLED)
            return

        try:
            df = pd.read_csv(self.csv_path)
            if 'location' not in df.columns or 'deployment_id' not in df.columns:
                messagebox.showerror("Error", "The CSV file must contain columns named 'location' and 'deployment_id'.")
                self.download_btn.configure(state=ctk.NORMAL)
                self.stop_btn.configure(state=ctk.DISABLED)
                return

            urls = df['location'].tolist()
            deployment_ids = df['deployment_id'].tolist()
            total_urls = len(urls)

            for i, (url, deployment_id) in enumerate(zip(urls, deployment_ids)):
                if self.stop_flag:
                    break

                try:
                    if self.use_multiple_folders.get():
                        # Create folder for deployment_id if it doesn't exist
                        clean_deployment_id = self.clean_deployment_id(str(deployment_id))
                        deployment_folder = os.path.join(self.folder_path, clean_deployment_id)
                        os.makedirs(deployment_folder, exist_ok=True)
                    else:
                        deployment_folder = self.folder_path

                    # Ensure the deployment folder exists
                    os.makedirs(deployment_folder, exist_ok=True)

                    filename = os.path.basename(url)
                    clean_filename = self.clean_filename(filename)
                    new_name = clean_filename.rsplit('.', 1)[0] + '.JPG'
                    new_file = os.path.join(deployment_folder, new_name)

                    # Check if the file already exists
                    if os.path.exists(new_file):
                        # If the file exists, skip downloading
                        self.status_var.set(f"File already exists, skipping... {i + 1} of {total_urls}")
                        continue

                    # Download the file
                    command = f"gsutil -m cp {url} \"{deployment_folder}\""
                    result = subprocess.run(command, check=True, shell=True, capture_output=True, text=True)
                    if result.returncode != 0:
                        raise subprocess.CalledProcessError(result.returncode, command, result.stdout, result.stderr)

                    old_file = os.path.join(deployment_folder, filename)
                    shutil.move(old_file, new_file)

                    # Update the status label
                    self.status_var.set(f"Downloading... {i + 1} of {total_urls} files")
                except subprocess.CalledProcessError as e:
                    error_message = f"Failed to download file {url} to {deployment_folder}\n{e.stderr}"
                    messagebox.showerror("Error", error_message)

            if not self.stop_flag:
                messagebox.showinfo("Success", "All files have been downloaded.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process the CSV file: {e}")
        finally:
            self.download_btn.configure(state=ctk.NORMAL)
            self.stop_btn.configure(state=ctk.DISABLED)
            # Reset the status label
            self.status_var.set("Download process completed.")

       
class ExcelCombinerApp:
    def __init__(self, root, lang="es"):
        self.root = root
        self.lang = lang  # Guardar el idioma actual

        # Diccionario de traducciones
        self.translations = {
            "es": {
                "select_initial_excel": "Selecciona el archivo Excel inicial",
                "browse_excel": "Buscar Excel",
                "no_excel_selected": "No se ha seleccionado ningún archivo Excel",
                "select_images_csv": "Selecciona el archivo CSV de imágenes",
                "browse_csv": "Buscar CSV",
                "no_csv_selected": "No se ha seleccionado ningún archivo CSV",
                "select_deployments_csv": "Selecciona el archivo CSV de despliegues",
                "process_files": "Procesar Archivos",
                "download_excel": "Descargar Excel Actualizado",
                "process_completed": "Archivos procesados con éxito",
                "error_message": "Se produjo un error: "
            },
            "pt": {
                "select_initial_excel": "Selecione o arquivo Excel inicial",
                "browse_excel": "Procurar Excel",
                "no_excel_selected": "Nenhum arquivo Excel selecionado",
                "select_images_csv": "Selecione o arquivo CSV de imagens",
                "browse_csv": "Procurar CSV",
                "no_csv_selected": "Nenhum arquivo CSV selecionado",
                "select_deployments_csv": "Selecione o arquivo CSV de implantações",
                "process_files": "Processar Arquivos",
                "download_excel": "Baixar Excel Atualizado",
                "process_completed": "Arquivos processados com sucesso",
                "error_message": "Ocorreu um erro: "
            },
            "en": {
                "select_initial_excel": "Select the Initial Excel file",
                "browse_excel": "Browse Excel",
                "no_excel_selected": "No Excel file selected",
                "select_images_csv": "Select the Images CSV file",
                "browse_csv": "Browse CSV",
                "no_csv_selected": "No CSV file selected",
                "select_deployments_csv": "Select the Deployments CSV file",
                "process_files": "Process Files",
                "download_excel": "Download Updated Excel",
                "process_completed": "Files processed successfully",
                "error_message": "An error occurred: "
            }
        }

        # Resto de la configuración de la interfaz
        self.setup_ui()

    def setup_ui(self):
        # Obtener traducciones según el idioma seleccionado
        tr = self.translations[self.lang]

        # Marco principal
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Etiqueta para el archivo Excel inicial
        self.label1 = ctk.CTkLabel(self.main_frame, text=tr["select_initial_excel"], anchor="w")
        self.label1.pack(pady=10)

        # Marco para la selección del archivo Excel inicial
        self.excel_frame = ctk.CTkFrame(self.main_frame)
        self.excel_frame.pack(pady=5, padx=10, fill="x")

        self.select_excel_btn = ctk.CTkButton(self.excel_frame, text=tr["browse_excel"], command=self.select_initial_excel)
        self.select_excel_btn.pack(side="left", padx=5)

        self.excel_label = ctk.CTkLabel(self.excel_frame, text=tr["no_excel_selected"], anchor="w")
        self.excel_label.pack(side="left", padx=5)

        # Etiqueta para el archivo CSV de imágenes
        self.label2 = ctk.CTkLabel(self.main_frame, text=tr["select_images_csv"], anchor="w")
        self.label2.pack(pady=10)

        # Marco para la selección del archivo CSV de imágenes
        self.images_frame = ctk.CTkFrame(self.main_frame)
        self.images_frame.pack(pady=5, padx=10, fill="x")

        self.select_images_btn = ctk.CTkButton(self.images_frame, text=tr["browse_csv"], command=self.select_images_csv)
        self.select_images_btn.pack(side="left", padx=5)

        self.images_label = ctk.CTkLabel(self.images_frame, text=tr["no_csv_selected"], anchor="w")
        self.images_label.pack(side="left", padx=5)

        # Etiqueta para el archivo CSV de despliegues
        self.label3 = ctk.CTkLabel(self.main_frame, text=tr["select_deployments_csv"], anchor="w")
        self.label3.pack(pady=10)

        # Marco para la selección del archivo CSV de despliegues
        self.deployments_frame = ctk.CTkFrame(self.main_frame)
        self.deployments_frame.pack(pady=5, padx=10, fill="x")

        self.select_deployments_btn = ctk.CTkButton(self.deployments_frame, text=tr["browse_csv"], command=self.select_deployments_csv)
        self.select_deployments_btn.pack(side="left", padx=5)

        self.deployments_label = ctk.CTkLabel(self.deployments_frame, text=tr["no_csv_selected"], anchor="w")
        self.deployments_label.pack(side="left", padx=5)

        # Botón de procesar archivos
        self.process_btn = ctk.CTkButton(self.main_frame, text=tr["process_files"], command=self.process_files, state=ctk.DISABLED)
        self.process_btn.pack(pady=10)

        # Botón de descargar Excel
        self.download_btn = ctk.CTkButton(self.main_frame, text=tr["download_excel"], command=self.save_file, state=ctk.DISABLED)
        self.download_btn.pack(pady=10)
        
    def select_initial_excel(self):
        self.initial_excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.initial_excel_path:
            self.excel_label.configure(text=os.path.basename(self.initial_excel_path))
        self.check_all_selected()

    def select_images_csv(self):
        self.images_csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.images_csv_path:
            self.images_label.configure(text=os.path.basename(self.images_csv_path))
        self.check_all_selected()

    def select_deployments_csv(self):
        self.deployments_csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.deployments_csv_path:
            self.deployments_label.configure(text=os.path.basename(self.deployments_csv_path))
        self.check_all_selected()

    def check_all_selected(self):
        if self.initial_excel_path and self.images_csv_path and self.deployments_csv_path:
            self.process_btn.configure(state=ctk.NORMAL)

    def process_files(self):
        try:
            # Load the provided files with dtype=str to avoid DtypeWarning
            images_df = pd.read_csv(self.images_csv_path, dtype=str, low_memory=False)
            deployments_df = pd.read_csv(self.deployments_csv_path, dtype=str, low_memory=False)
            
            # Load the initial Excel file
            initial_excel_path = self.initial_excel_path
            if initial_excel_path is None:
                raise ValueError("No se ha seleccionado un archivo de Excel inicial.")
            
            initial_df = pd.read_excel(initial_excel_path, sheet_name=None)
            sheet_names = initial_df.keys()
            first_sheet_name = list(sheet_names)[0]
            initial_df = initial_df[first_sheet_name]

            # Merge the dataframes on project_id and deployment_id
            merged_df = images_df.merge(deployments_df, on=['project_id', 'deployment_id'], suffixes=('_image', '_deployment'))

            # Select the relevant columns
            result_df = merged_df[['latitude', 'longitude', 'placename', 'location', 'timestamp', 'project_id', 'deployment_id', 'subproject_name']]

            # Convert 'timestamp' column to datetime
            result_df['timestamp'] = pd.to_datetime(result_df['timestamp'], format='%Y-%m-%d %H:%M:%S')

            # Check if the "Multiple Images" checkbox is selected
            if self.multiple_images_var.get():
                self.process_multiple_images(result_df, initial_df)
            else:
                # Process as single image per row, no grouping
                combined_df = pd.DataFrame()
                combined_df['Encounter.decimalLatitude'] = result_df['latitude']
                combined_df['Encounter.decimalLongitude'] = result_df['longitude']
                combined_df['Encounter.verbatimLocality'] = result_df['placename']
                combined_df['Encounter.mediaAsset0'] = result_df['location'].apply(lambda x: x.split('/')[-1] if pd.notna(x) else x)

                # Generate Occurrence.occurrenceID
                combined_df['Occurrence.occurrenceID'] = result_df.apply(self.generate_occurrence_id, axis=1)

                # Ensure the file extension is .JPG
                combined_df['Encounter.mediaAsset0'] = combined_df['Encounter.mediaAsset0'].apply(self.ensure_jpg_extension)

                # Add time-related columns
                combined_df['Encounter.year'] = result_df['timestamp'].dt.year
                combined_df['Encounter.month'] = result_df['timestamp'].dt.month
                combined_df['Encounter.day'] = result_df['timestamp'].dt.day
                combined_df['Encounter.hour'] = result_df['timestamp'].dt.hour
                combined_df['Encounter.minutes'] = result_df['timestamp'].dt.minute

                # Fill missing columns from initial_df to combined_df with default values from initial_df
                for column in initial_df.columns:
                    if column not in combined_df.columns:
                        combined_df[column] = initial_df[column].iloc[0]

                # Ensure the columns are in the same order as initial_df and include the new Occurrence.occurrenceID column
                final_columns = ['Occurrence.occurrenceID'] + [col for col in initial_df.columns if col != 'Occurrence.occurrenceID']
                combined_df = combined_df[final_columns]

                # Store the final DataFrame in the class variable
                self.final_df = combined_df

                messagebox.showinfo("Process Completed", f"File processed successfully.")
                self.save_file_btn.configure(state=ctk.NORMAL)
        
        except Exception as e:
            messagebox.showerror("Error", f"Se produjo un error: {e}")
            print(e)

    def process_multiple_images(self, result_df, initial_df):
        try:
            # Sort the DataFrame by deployment_id and timestamp
            result_df = result_df.sort_values(by=['deployment_id', 'timestamp'])

            # Get the time threshold from user input
            time_threshold = int(self.time_threshold_entry.get())

            # Group images by deployment and time difference
            combined_images = []
            for deployment_id, group in result_df.groupby('deployment_id'):
                group['time_diff'] = group['timestamp'].diff().dt.total_seconds().fillna(time_threshold + 1)
                group_images = []
                for _, row in group.iterrows():
                    if group_images and row['time_diff'] > time_threshold:
                        combined_images.append(group_images)
                        group_images = []
                    group_images.append(row)
                if group_images:
                    combined_images.append(group_images)

            # Create the new combined DataFrame
            rows_list = []
            max_assets = 0
            for idx, images_group in enumerate(combined_images):
                if isinstance(images_group, list) or isinstance(images_group, pd.DataFrame):
                    base_row = images_group[0]
                    new_row = {
                        'Encounter.decimalLatitude': base_row['latitude'],
                        'Encounter.decimalLongitude': base_row['longitude'],
                        'Encounter.verbatimLocality': base_row['placename'],
                        'Occurrence.occurrenceID': self.generate_occurrence_id(base_row),  # Generate Occurrence ID
                        'Encounter.year': base_row['timestamp'].year,
                        'Encounter.month': base_row['timestamp'].month,
                        'Encounter.day': base_row['timestamp'].day,
                        'Encounter.hour': base_row['timestamp'].hour,
                        'Encounter.minutes': base_row['timestamp'].minute
                    }
                    for i, image in enumerate(images_group):
                        image_location = image['location'].split('/')[-1]
                        new_row[f'Encounter.mediaAsset{i}'] = self.ensure_jpg_extension(image_location)
                    rows_list.append(new_row)
                    max_assets = max(max_assets, len(images_group))

            # Convert the list of rows into a DataFrame
            combined_df = pd.DataFrame(rows_list)

            # Ensure all rows have columns Encounter.mediaAsset0 to Encounter.mediaAsset{max_assets-1}
            for i in range(max_assets):
                if f'Encounter.mediaAsset{i}' not in combined_df.columns:
                    combined_df[f'Encounter.mediaAsset{i}'] = None

            # Fill missing columns from initial_df to combined_df with default values from initial_df
            for column in initial_df.columns:
                if column not in combined_df.columns:
                    combined_df[column] = initial_df[column].iloc[0]

            # Ensure the columns are in the same order as initial_df and include the new Occurrence.occurrenceID column
            final_columns = ['Occurrence.occurrenceID'] + [col for col in initial_df.columns if col != 'Occurrence.occurrenceID']
            combined_df = combined_df[final_columns + [col for col in combined_df.columns if col.startswith('Encounter.mediaAsset')]]

            # Store the final DataFrame in the class variable
            self.final_df = combined_df

            messagebox.showinfo("Process Completed", f"Files processed successfully with multiple images handling.")
            self.save_file_btn.configure(state=ctk.NORMAL)
        
        except Exception as e:
            messagebox.showerror("Error", f"Se produjo un error: {e}")
            print(e)

    # Function to sanitize and generate Occurrence.occurrenceID
    def generate_occurrence_id(self, row):
        # Convert to string and handle NaN by replacing with an empty string
        sanitized_project_id = re.sub(r'[^a-zA-Z0-9-_]', '_', str(row['project_id']) if pd.notna(row['project_id']) else '')
        sanitized_subproject_name = re.sub(r'[^a-zA-Z0-9-_]', '_', str(row['subproject_name']) if pd.notna(row['subproject_name']) else '')
        sanitized_deployment_id = re.sub(r'[^a-zA-Z0-9-_]', '_', str(row['deployment_id']) if pd.notna(row['deployment_id']) else '')
        return f"{sanitized_project_id}-{sanitized_subproject_name}-{sanitized_deployment_id}"

    # Ensure the file extension is .JPG
    def ensure_jpg_extension(self, location):
        if pd.isna(location):
            return location
        parts = location.split('.')
        if len(parts) > 1 and parts[-1].lower() != 'jpg':
            return '.'.join(parts[:-1]) + '.JPG'
        return location


    def save_file(self):
        if self.final_df is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                self.final_df.to_excel(save_path, index=False)
                messagebox.showinfo("File Saved", f"File saved successfully to {save_path}")


class DateChangerApp:
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

    def get_file_dates(self, folder):
        files = [os.path.join(folder, f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        dates = []
        
        for file in files:
            exif_date = self.get_exif_date(file)
            if exif_date:
                dates.append(exif_date.timestamp())
            else:
                dates.append(os.path.getctime(file))
        
        if dates:
            self.newest_date.set(datetime.fromtimestamp(max(dates)).strftime("%Y-%m-%d %H:%M:%S"))
            self.oldest_date.set(datetime.fromtimestamp(min(dates)).strftime("%Y-%m-%d %H:%M:%S"))
        else:
            self.newest_date.set("")
            self.oldest_date.set("")

    def correct_exif_format(self, exif_data):
        for ifd in ("0th", "Exif", "GPS", "1st"):
            for tag in exif_data[ifd]:
                if isinstance(exif_data[ifd][tag], int):
                    exif_data[ifd][tag] = (exif_data[ifd][tag], 1)
        return exif_data

    def change_dates(self):
        real_date = datetime.strptime(self.real_date_entry.get(), "%Y-%m-%d %H:%M:%S")
        
        if self.date_option.get() == 1:
            reference_date = datetime.strptime(self.newest_date.get(), "%Y-%m-%d %H:%M:%S")
        elif self.date_option.get() == 2:
            reference_date = datetime.strptime(self.oldest_date.get(), "%Y-%m-%d %H:%M:%S")
        else:
            reference_date = datetime.strptime(self.custom_date_entry.get(), "%Y-%m-%d %H:%M:%S")

        difference = real_date - reference_date

        folder = self.selected_folder
        files = [os.path.join(folder, f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]

        for file in files:
            # Update EXIF dates if it's an image
            exif_data = piexif.load(file)
            date_time_original = exif_data['Exif'].get(piexif.ExifIFD.DateTimeOriginal)
            if date_time_original:
                new_exif_date = datetime.strptime(date_time_original.decode('utf-8'), "%Y:%m:%d %H:%M:%S") + difference
                new_date_str = new_exif_date.strftime("%Y:%m:%d %H:%M:%S")
                exif_data['Exif'][piexif.ExifIFD.DateTimeOriginal] = new_date_str.encode('utf-8')
                exif_data['Exif'][piexif.ExifIFD.DateTimeDigitized] = new_date_str.encode('utf-8')
                exif_data['0th'][piexif.ImageIFD.DateTime] = new_date_str.encode('utf-8')
                exif_data = self.correct_exif_format(exif_data)
                exif_bytes = piexif.dump(exif_data)
                piexif.insert(exif_bytes, file)

            # Update file creation dates
            handle = win32file.CreateFile(
                file, 
                win32file.GENERIC_WRITE, 
                0, 
                None, 
                win32file.OPEN_EXISTING, 
                win32file.FILE_ATTRIBUTE_NORMAL, 
                None
            )

            new_creation_date = datetime.fromtimestamp(os.path.getctime(file)) + difference
            new_creation_date_filetime = pywintypes.Time(new_creation_date)

            (creation, access, modification) = win32file.GetFileTime(handle)

            win32file.SetFileTime(handle, new_creation_date_filetime, access, modification)
            handle.close()
        
        messagebox.showinfo("Success", "File dates updated successfully!")

    def copy_to_folder(self):
        destination_folder = filedialog.askdirectory(title="Select Destination Folder")
        if destination_folder:
            folder = self.selected_folder
            files = [os.path.join(folder, f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]

            for file in files:
                file_name = os.path.basename(file)
                destination_path = os.path.join(destination_folder, file_name)

                if os.path.exists(destination_path):
                    base_name, extension = os.path.splitext(file_name)
                    destination_path = os.path.join(destination_folder, f"{base_name}_datechanged{extension}")

                # Apply date change only to the file being copied
                self.apply_date_change_to_copy(file, destination_path)

            messagebox.showinfo("Success", "Files copied successfully!")

    def apply_date_change_to_copy(self, source_file, destination_file):
        # Copy the file first
        shutil.copy(source_file, destination_file)

        # Then apply the date changes to the copied file
        real_date = datetime.strptime(self.real_date_entry.get(), "%Y-%m-%d %H:%M:%S")
        
        if self.date_option.get() == 1:
            reference_date = datetime.strptime(self.newest_date.get(), "%Y-%m-%d %H:%M:%S")
        elif self.date_option.get() == 2:
            reference_date = datetime.strptime(self.oldest_date.get(), "%Y-%m-%d %H:%M:%S")
        else:
            reference_date = datetime.strptime(self.custom_date_entry.get(), "%Y-%m-%d %H:%M:%S")

        difference = real_date - reference_date

        # Update EXIF dates if it's an image
        try:
            exif_data = piexif.load(destination_file)
            date_time_original = exif_data['Exif'].get(piexif.ExifIFD.DateTimeOriginal)
            if date_time_original:
                new_exif_date = datetime.strptime(date_time_original.decode('utf-8'), "%Y:%m:%d %H:%M:%S") + difference
                new_date_str = new_exif_date.strftime("%Y:%m:%d %H:%M:%S")
                exif_data['Exif'][piexif.ExifIFD.DateTimeOriginal] = new_date_str.encode('utf-8')
                exif_data['Exif'][piexif.ExifIFD.DateTimeDigitized] = new_date_str.encode('utf-8')
                exif_data['0th'][piexif.ImageIFD.DateTime] = new_date_str.encode('utf-8')
                exif_data = self.correct_exif_format(exif_data)
                exif_bytes = piexif.dump(exif_data)
                piexif.insert(exif_bytes, destination_file)
        except:
            pass  # Ignore if not an image

        # Update file creation dates
        handle = win32file.CreateFile(
            destination_file, 
            win32file.GENERIC_WRITE, 
            0, 
            None, 
            win32file.OPEN_EXISTING, 
            win32file.FILE_ATTRIBUTE_NORMAL, 
            None
        )

        new_creation_date = datetime.fromtimestamp(os.path.getctime(destination_file)) + difference
        new_creation_date_filetime = pywintypes.Time(new_creation_date)

        (creation, access, modification) = win32file.GetFileTime(handle)

        win32file.SetFileTime(handle, new_creation_date_filetime, access, modification)
        handle.close()


class ImagesRenamer:
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

    def rename_and_copy_photos(self):
        if not self.source_folder:
            messagebox.showerror("Error", "Please select the source folder.")
            return

        if self.copy_photos_var.get() and not self.dest_folder:
            messagebox.showerror("Error", "Please select the destination folder.")
            return

        # Show processing label
        self.processing_label.configure(text="Processing...")
        self.root.update_idletasks()  # Ensure the label updates immediately

        source_path = Path(self.source_folder)
        dest_path = Path(self.dest_folder) if self.copy_photos_var.get() else None

        if not source_path.is_dir() or (dest_path and not dest_path.is_dir()):
            messagebox.showerror("Error", "Invalid directory selected.")
            self.processing_label.configure(text="")  # Hide processing label on error
            return

        custom_text = self.custom_text_var.get().strip()

        for root, dirs, files in os.walk(source_path):
            root_path = Path(root)
            folder_name = root_path.name.split()[0].title()

            for file_name in files:
                file = root_path / file_name
                if file.is_file() and file.suffix.lower() in ['.jpg', '.jpeg', '.png', '.gif']:
                    exif_date = f" {self.get_exif_date(file)}" if self.add_exif_date_var.get() else ""
                    original_name = file.stem.title()

                    # Determine the base name for the new file
                    base_name = ""
                    if self.use_folder_name_var.get():
                        base_name += folder_name
                    if self.use_custom_text_var.get():
                        base_name += f" {custom_text}"

                    new_name = f"{base_name} {exif_date}{original_name}{file.suffix}" if self.keep_original_name_var.get() else f"{base_name} {exif_date}{file.suffix}"

                    # Replace spaces with underscores if the option is selected
                    if self.replace_spaces_var.get():
                        new_name = new_name.replace(" ", "_")

                    new_path = (dest_path / new_name) if self.copy_photos_var.get() else (root_path / new_name)

                    # Check if the new file name already exists and modify it if necessary
                    counter = 1
                    original_new_path = new_path
                    while new_path.exists():
                        new_name = original_new_path.stem + f" {counter}" + original_new_path.suffix
                        new_path = new_path.parent / new_name
                        counter += 1

                    if self.copy_photos_var.get():
                        shutil.copy2(file, new_path)
                    else:
                        file.rename(new_path)

        # Hide processing label and show success message
        self.processing_label.configure(text="Processing completed.")
        messagebox.showinfo("Information", "Photos have been renamed and processed successfully.")
        self.processing_label.after(3000, lambda: self.processing_label.configure(text=""))  # Hide label after 3 seconds

class LynxOne:
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
                "linces_folder_exists": "Does the Lince/Linces folder exist?",
                "revision_folder_exists": "Does the Revision folder exist?",
                "optional_files": "Optional: Select additional Excel files for joining.",
                "estaciones_file": "No Estaciones file selected",
                "browse_estaciones": "Browse Estaciones",
                "individuos_file": "No Individuos file selected",
                "browse_individuos": "Browse Individuos",
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

        # Crear el panel de pestañas
        self.tabview = ctk.CTkTabview(self.root, width=400, height=300)
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

        # Botones para generar y descargar archivos Excel
        self.generate_button = ctk.CTkButton(self.root, text=tr["generate_excel"], command=self.generate_excel)
        self.generate_button.pack(pady=10, side="left", padx=20)

        self.download_button = ctk.CTkButton(self.root, text=tr["download_excel"], command=self.save_excel)
        self.download_button.pack(pady=10, side="left", padx=20)

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


    def get_exif_data(self, image_path):
        # Retrieve EXIF data from an image
        image = Image.open(image_path)
        exif_data = image._getexif()
        if not exif_data:
            return None
        exif = {}
        for tag, value in exif_data.items():
            decoded = TAGS.get(tag, tag)
            exif[decoded] = value
        return exif

    def get_date_taken(self, exif_data):
        # Extract the date when the photo was taken from EXIF data
        date_taken = exif_data.get("DateTimeOriginal")
        if date_taken:
            return datetime.strptime(date_taken, '%Y:%m:%d %H:%M:%S')
        return None

    def generate_excel(self):
        if not self.source_folder:
            messagebox.showwarning(self.translations[self.lang]["warning"], self.translations[self.lang]["no_source_folder"])
            return

        # File extensions to look for
        valid_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', 
                            '.mp4', '.avi', '.mov', '.mkv', '.wmv')

        # Determine if the "Lince/linces" and "Revision" folders exist
        lince_exists = self.lince_checkbox_var.get()
        revision_exists = self.revision_checkbox_var.get()

        # Get the number of minutes to group by
        try:
            minutes_to_group = int(self.minutes_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for minutes.")
            return

        # Traverse the directory and extract information
        data = []
        for root, dirs, files in os.walk(self.source_folder):
            for file in files:
                if file.lower().endswith(valid_extensions):
                    # Convert the path to a standard format with Path
                    root_path = Path(root)

                    # Extract parts of the path depending on whether Lince and Revision exist
                    parts = root_path.parts
                    if lince_exists and revision_exists and len(parts) >= 6:
                        finca = parts[-5]  # Fifth folder from the end
                        estacion = parts[-4]  # Fourth folder from the end
                        revision = parts[-3]  # Third folder from the end
                        lince = parts[-1]  # Current folder name
                    elif lince_exists and not revision_exists and len(parts) >= 5:
                        finca = parts[-4]  # Fourth folder from the end
                        estacion = parts[-3]  # Third folder from the end
                        revision = "N/A"  # No Revision folder
                        lince = parts[-1]  # Current folder name
                    elif not lince_exists and revision_exists and len(parts) >= 5:
                        finca = parts[-4]  # Fourth folder from the end
                        estacion = parts[-3]  # Third folder from the end
                        revision = parts[-2]  # Second folder from the end
                        lince = parts[-1]  # Current folder name
                    elif not lince_exists and not revision_exists and len(parts) >= 4:
                        finca = parts[-3]  # Third folder from the end
                        estacion = parts[-2]  # Second folder from the end
                        revision = "N/A"  # No Revision folder
                        lince = parts[-1]  # Current folder name
                    else:
                        continue  # Skip if the expected structure is not met

                    # Get the Capture Date (EXIF DateTimeOriginal) from the image
                    file_name = root_path / file  # Combine path and file
                    exif_data = self.get_exif_data(str(file_name))
                    if exif_data:
                        capture_date = self.get_date_taken(exif_data)
                    else:
                        capture_date = None

                    # Handle multiple linces
                    # Replace both " y " and " Y " with a common separator
                    lince = re.sub(r' y | Y ', ' y ', lince)
                    lince_names = lince.split(" y ")

                    # Add individual rows for each lince
                    for individual_lince in lince_names:
                        individual_lince_row = {
                            "Finca": finca,
                            "Estación": estacion,
                            "Revisión": revision,
                            "Linces": lince,  # Combined names
                            "Lince": individual_lince.strip(),  # Individual lince name
                            "Archivo": str(file_name),
                            "Fecha de Captura": capture_date
                        }
                        data.append(individual_lince_row)

        if not data:
            messagebox.showinfo("Info", "No image or video files were found in the selected directory.")
            return

        # Convert to DataFrame
        df = pd.DataFrame(data, columns=["Finca", "Estación", "Revisión", "Linces", "Lince", "Archivo", "Fecha de Captura"])

        if minutes_to_group > 0:
            # Group by individual and collapse rows based on time difference
            grouped_data = []
            for name, group in df.groupby(["Finca", "Estación", "Revisión", "Lince"]):
                group = group.sort_values(by="Fecha de Captura")
                collapsed_files = []
                last_time = None
                for _, row in group.iterrows():
                    if last_time and row["Fecha de Captura"] and (row["Fecha de Captura"] - last_time).total_seconds() / 60 <= minutes_to_group:
                        collapsed_files[-1]["Archivo"] += ";" + row["Archivo"]
                    else:
                        collapsed_files.append(row.to_dict())
                    last_time = row["Fecha de Captura"]
                grouped_data.extend(collapsed_files)
            df = pd.DataFrame(grouped_data)

        # Perform join with Estaciones and Individuos if the files are provided
        if self.estaciones_file:
            estaciones_df = pd.read_excel(self.estaciones_file)
            df = df.merge(estaciones_df, how='left', left_on='Estación', right_on='Estacion')

        if self.individuos_file:
            individuos_df = pd.read_excel(self.individuos_file)
            df = df.merge(individuos_df, how='left', left_on='Lince', right_on='Lince')

        # Store the DataFrame to use it later for saving
        self.excel_data = df

    def save_excel(self):
        if not hasattr(self, 'excel_data') or self.excel_data is None:
            messagebox.showerror("Error", "No Excel data to save. Please generate the Excel file first.")
            return

        # Open save file dialog to choose where to save the file
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 title="Save Excel File")
        if file_path:
            self.excel_data.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Excel file successfully saved at {file_path}")

        
if __name__ == "__main__":
    root = ctk.CTk()
    app = BaseApp(root)
    root.mainloop()
