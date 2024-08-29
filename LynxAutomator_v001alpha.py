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
        self.root.title("LynxAutomator")
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")
        ctk.set_default_color_theme("green")  # Color theme ("blue", "green", "dark-blue")

        # Create the header with large text in the middle
        self.header_frame = ctk.CTkFrame(root)
        self.header_frame.pack(pady=10, fill='x')

        # Header label with slightly smaller text
        self.header_label = ctk.CTkLabel(self.header_frame, text="LynxAutomator", font=("Helvetica", 24, "bold"))
        self.header_label.pack(pady=5, padx=5)

        # Description label in the header
        self.description_label = ctk.CTkLabel(self.header_frame, text="This application allows you to automate certain processes in the monitoring of the Iberian Lynx and other species", font=("Helvetica", 12), wraplength=800, justify="left")
        self.description_label.pack(pady=10, padx=10)

        # Create the main tab panel (top-level tabs)
        self.main_tabs = ttk.Notebook(root)
        self.main_tabs.pack(pady=20, padx=20, fill='both', expand=True)

        # Create frames for each main tab
        self.camtrap_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.wildbook_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.wildlife_insights_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.iberian_lynx_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.help_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)
        self.about_frame = ctk.CTkFrame(self.main_tabs, width=600, height=400)

        self.main_tabs.add(self.about_frame, text="About")
        self.main_tabs.add(self.wildbook_frame, text="Wildbook")
        self.main_tabs.add(self.wildlife_insights_frame, text="Wildlife Insights")
        self.main_tabs.add(self.iberian_lynx_frame, text="Iberian Lynx")
        self.main_tabs.add(self.camtrap_frame, text="Functionalities")
        self.main_tabs.add(self.help_frame, text="Help")

        # Style for the tabs
        style = ttk.Style()
        style.configure('TNotebook.Tab',
                        font=('Helvetica', 14, 'bold'),  # Font, size, and style
                        padding=[10, 5],  # Padding around the text
                        relief='flat')  # No border around tabs

        # About TabView
        self.about_tabs = ctk.CTkTabview(self.about_frame, width=600, height=400)
        self.about_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.presentation_tab = self.about_tabs.add("Presentation")

        # Camtrap Functionalities TabView
        self.camtrap_tabs = ctk.CTkTabview(self.camtrap_frame, width=600, height=400)
        self.camtrap_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.date_changer_tab = self.camtrap_tabs.add("Date Changer")
        self.video_frame_extractor_tab = self.camtrap_tabs.add("Video Frame Extractor")
        self.images_renamer_tab = self.camtrap_tabs.add("Images Renamer") 


        # Wildbook TabView
        self.wildbook_tabs = ctk.CTkTabview(self.wildbook_frame, width=600, height=400)
        self.wildbook_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.wiwbe_folder_tab = self.wildbook_tabs.add("WIWbE from Folder")
        self.wbcatalog_tab = self.wildbook_tabs.add("WIWbE Catalog") 

        # Wildlife Insights TabView
        self.wildlife_insights_tabs = ctk.CTkTabview(self.wildlife_insights_frame, width=600, height=400)
        self.wildlife_insights_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.wi_images_csv_filter_tab = self.wildlife_insights_tabs.add("WI Images CSV Filter")
        self.wi_downloader_tab = self.wildlife_insights_tabs.add("WI Downloader")
        self.wi_to_wiwbe_tab = self.wildlife_insights_tabs.add("WI to WIWbE")

        # Iberian Lynx TabView
        self.iberian_lynx_tabs = ctk.CTkTabview(self.iberian_lynx_frame, width=600, height=400)
        self.iberian_lynx_tabs.pack(pady=20, padx=20, fill='both', expand=True)
        
        self.lynx_feature_1_tab = self.iberian_lynx_tabs.add("Iberian Lynx Feature")

        # Help TabView
        self.help_tab = ctk.CTkTabview(self.help_frame, width=600, height=400)
        self.help_tab.pack(pady=20, padx=20, fill='both', expand=True)

        self.help_content_tab = self.help_tab.add("Help")

        # Integrate the class in the BaseApp
        self.help_app = AppHelp(self.help_content_tab)
        self.photo_date_app = WBFolderApp(self.wiwbe_folder_tab)
        self.csv_filter_app = CSVFilterApp(self.wi_images_csv_filter_tab)
        self.gcs_downloader_app = GCSDownloaderAndRenamer(self.wi_downloader_tab)
        self.excel_combiner_app = ExcelCombinerApp(self.wi_to_wiwbe_tab)
        self.data_changer_app = DateChangerApp(self.date_changer_tab)
        self.frame_extractor_app = FrameExtractorApp(self.video_frame_extractor_tab)
        self.wbcatalog_app = WBCatalogApp(self.wbcatalog_tab)
        self.images_renamer_app = ImagesRenamer(self.images_renamer_tab) 
        self.presentation_app = Presentation(self.presentation_tab)
        self.lynxone_app = LynxOne(self.lynx_feature_1_tab)


        
        # Placeholder methods for future functionality
        self.temp_file_path = None

class Presentation(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(fill="both", expand=True)
        self.create_widgets()

    def create_widgets(self):
        self.title = "Mi Aplicación"

        logo_path = "C:/Users/WWF-POR113/Desktop/PythonApps/BIWBapp/logo.png"
        logo_image = Image.open(logo_path)

        self.logo = ImageTk.PhotoImage(logo_image)
        logo_label = tk.Label(self, image=self.logo)
        logo_label.pack(pady=20)

        text_label = ctk.CTkLabel(self, text="This app is part of the LIFE+ Lynxconnect project", font=("Helvetica", 16))
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



class WBFolderApp:

    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")

        # Main frame (single large square)
        self.main_frame = ctk.CTkFrame(root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Label to prompt user action
        self.label = ctk.CTkLabel(self.main_frame, text="Select the folder with the images, tha name and the date will be take from these images.", anchor="w")
        self.label.pack(pady=10)

        # Frame to hold folder selection widgets
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select a folder
        self.select_folder_btn = ctk.CTkButton(self.folder_frame, text="Browse Folder", command=self.select_folder)
        self.select_folder_btn.pack(side="left", padx=5)

        # Label to display the selected folder name
        self.folder_label = ctk.CTkLabel(self.folder_frame, text="No folder selected", anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Message between the two buttons
        self.message_label = ctk.CTkLabel(self.main_frame, text="Select the Initial Excel file.")
        self.message_label.pack(pady=10)

        # Frame to hold file selection widgets
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select an Excel file
        self.select_file_btn = ctk.CTkButton(self.file_frame, text="Browse File", command=self.select_file)
        self.select_file_btn.pack(side="left", padx=5)

        # Label to display the selected file name
        self.file_label = ctk.CTkLabel(self.file_frame, text="No file selected", anchor="w")
        self.file_label.pack(side="left", padx=5)

        # Section for options in a single line
        self.options_frame = ctk.CTkFrame(self.main_frame)
        self.options_frame.pack(pady=10, padx=10, fill='x')

        # Checkbox for multiple images processing
        self.multiple_images_checkbox = ctk.CTkCheckBox(self.options_frame, text="Process with Multiple Images")
        self.multiple_images_checkbox.pack(side="left", padx=5)

        # Label for time threshold
        self.time_threshold_label = ctk.CTkLabel(self.options_frame, text="Time threshold (seconds):", anchor="w")
        self.time_threshold_label.pack(side="left", padx=5)

        # Entry for time threshold
        self.time_threshold_entry = ctk.CTkEntry(self.options_frame, width=100)
        self.time_threshold_entry.pack(side="left", padx=5)
        self.time_threshold_entry.insert(0, "3")  # Default value

        # Frame to hold process and download buttons side by side
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(pady=10)

        # Process button
        self.process_btn = ctk.CTkButton(self.button_frame, text="Process", command=self.process_files)
        self.process_btn.pack(side="left", padx=5)

        # Download button
        self.download_btn = ctk.CTkButton(self.button_frame, text="Download Updated Excel", command=self.download_file, state=ctk.DISABLED)
        self.download_btn.pack(side="left", padx=5)

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
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")

        # Frame to hold all widgets
        self.main_frame = ctk.CTkFrame(root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Label for folder selection
        self.label1 = ctk.CTkLabel(self.main_frame, text="Select a directory to analyze. The first word of the filename's image files will be used as the individualID.", anchor="w")
        self.label1.pack(pady=10)

        # Frame to hold folder selection widgets
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select a folder
        self.select_folder_btn = ctk.CTkButton(self.folder_frame, text="Browse Folder", command=self.select_folder)
        self.select_folder_btn.pack(side="left", padx=5)

        # Label to display the selected folder name
        self.folder_label = ctk.CTkLabel(self.folder_frame, text="No folder selected", anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Label for file selection
        self.label2 = ctk.CTkLabel(self.main_frame, text="Select an Excel file to process along with the photos.", anchor="w")
        self.label2.pack(pady=10)

        # Frame to hold file selection widgets
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select an Excel file
        self.select_file_btn = ctk.CTkButton(self.file_frame, text="Browse File", command=self.select_file)
        self.select_file_btn.pack(side="left", padx=5)

        # Label to display the selected file name
        self.file_label = ctk.CTkLabel(self.file_frame, text="No file selected", anchor="w")
        self.file_label.pack(side="left", padx=5)

        # Frame to hold the checkboxes side by side
        self.checkbox_frame = ctk.CTkFrame(self.main_frame)
        self.checkbox_frame.pack(pady=5, padx=10, fill="x")
    
        # Checkbox for collapsing option
        self.collapse_var = ctk.BooleanVar()
        self.collapse_check = ctk.CTkCheckBox(self.checkbox_frame, text="Collapse rows with the same individualID", variable=self.collapse_var)
        self.collapse_check.pack(side="left", padx=5)

        # Checkbox for capitalizing individualID
        self.capitalize_var = ctk.BooleanVar(value=True)  # Default to True
        self.capitalize_check = ctk.CTkCheckBox(self.checkbox_frame, text="Capitalize individualID", variable=self.capitalize_var)
        self.capitalize_check.pack(side="left", padx=5)

        # Frame to hold process and download buttons side by side
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(pady=10)

        # Process button
        self.process_btn = ctk.CTkButton(self.button_frame, text="Process", command=self.process_files)
        self.process_btn.pack(side="left", padx=5)

        # Download button
        self.download_btn = ctk.CTkButton(self.button_frame, text="Download Updated Excel", command=self.download_file, state=ctk.DISABLED)
        self.download_btn.pack(side="left", padx=5)

        self.temp_file_path = None

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
    def __init__(self, root):
        self.root = root
        
        self.folder_path = None
        self.output_folder = None

        # Create the tab panel
        self.tabview = ctk.CTkTabview(root, width=200, height=150)
        self.tabview.pack(pady=20, padx=20)

        # Add tabs
        self.tab1 = self.tabview.add("Videos Folder")
        self.tab2 = self.tabview.add("Output Folder")

        # Content of the "Videos Folder" tab
        self.label_videos = ctk.CTkLabel(self.tab1, text="Select the folder containing the videos:", anchor="w")
        self.label_videos.pack(pady=10)

        self.videos_frame = ctk.CTkFrame(self.tab1)
        self.videos_frame.pack(pady=5, padx=10, fill="x")

        self.select_videos_btn = ctk.CTkButton(self.videos_frame, text="Browse", command=self.select_videos_folder)
        self.select_videos_btn.pack(side="left", padx=5)

        self.videos_folder_label = ctk.CTkLabel(self.videos_frame, text="No folder selected", anchor="w")
        self.videos_folder_label.pack(side="left", padx=5)

        # Content of the "Output Folder" tab
        self.label_output = ctk.CTkLabel(self.tab2, text="Select the output folder for frames:", anchor="w")
        self.label_output.pack(pady=10)

        self.output_frame = ctk.CTkFrame(self.tab2)
        self.output_frame.pack(pady=5, padx=10, fill="x")

        self.select_output_btn = ctk.CTkButton(self.output_frame, text="Browse", command=self.select_output_folder)
        self.select_output_btn.pack(side="left", padx=5)

        self.output_folder_label = ctk.CTkLabel(self.output_frame, text="No folder selected", anchor="w")
        self.output_folder_label.pack(side="left", padx=5)

        # Add the label and entry for the interval
        ctk.CTkLabel(root, text="Interval between frames (seconds):").pack(pady=10)
        self.interval_var = ctk.DoubleVar(value=1.0)  # Default value of 1 second
        ctk.CTkEntry(root, textvariable=self.interval_var).pack(pady=10)

        # Add the process button
        self.process_btn = ctk.CTkButton(root, text="Extract frames", command=self.start_extraction)
        self.process_btn.pack(pady=10)

        # Add status label
        self.status_label = ctk.CTkLabel(root, text="")
        self.status_label.pack(pady=10)

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

    def select_videos_folder(self):
        self.folder_path = filedialog.askdirectory(title="Select the folder with videos")
        if self.folder_path:
            self.videos_folder_label.configure(text=os.path.basename(self.folder_path))
            messagebox.showinfo("Information", f"Selected folder: {self.folder_path}")
        else:
            self.status_label.configure(text="Please select a folder with videos.")
    
    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory(title="Select the output folder")
        if self.output_folder:
            self.output_folder_label.configure(text=os.path.basename(self.output_folder))
            messagebox.showinfo("Information", f"Selected folder: {self.output_folder}")
        else:
            self.status_label.configure(text="Please select an output folder.")

    def start_extraction(self):
        try:
            interval = float(self.interval_var.get())
            if interval <= 0:
                raise ValueError("The interval must be greater than zero.")
            
            if not self.folder_path or not self.output_folder:
                self.status_label.configure(text="Please select both input and output folders.")
                return

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

        except ValueError as e:
            self.status_label.configure(text=f"Error: {e}")


class CSVFilterApp:
    def __init__(self, root):
        self.root = root
        self.frames = {}

        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

        # Create frames for each step
        for F in (self.StartPage, self.DeploymentPage, self.CommonNamePage, self.ResultPage):
            page_name = F.__name__
            frame = F(parent=root, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")

        self.df = None
        self.filtered_df = None
        self.selected_deployments = []
        self.selected_common_names = []

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        self.df = pd.read_csv(file_path)
        required_columns = ['deployment_id', 'common_name']

        if not all(col in self.df.columns for col in required_columns):
            messagebox.showerror("Error", f"CSV file must contain the following columns: {', '.join(required_columns)}")
            self.df = None
            return

        unique_deployments = self.df['deployment_id'].unique()
        unique_common_names = self.df['common_name'].unique()

        self.frames["DeploymentPage"].update_listbox(unique_deployments)
        self.frames["CommonNamePage"].update_listbox(unique_common_names)

    def set_selected_deployments(self, selected_deployments):
        self.selected_deployments = selected_deployments

    def set_selected_common_names(self, selected_common_names):
        self.selected_common_names = selected_common_names

    def filter_csv(self):
        if not self.selected_deployments or not self.selected_common_names:
            messagebox.showwarning("Warning", "Please select at least one item from each list")
            return

        self.filtered_df = self.df[
            (self.df['deployment_id'].isin(self.selected_deployments)) &
            (self.df['common_name'].isin(self.selected_common_names))
        ]

        self.frames["ResultPage"].update_result_text(f"Filtered {len(self.filtered_df)} rows")

    def save_csv(self):
        if self.filtered_df is None:
            messagebox.showwarning("Warning", "No filtered data to save")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        self.filtered_df.to_csv(file_path, index=False)
        messagebox.showinfo("Info", "Filtered CSV saved successfully")

    class StartPage(ctk.CTkFrame):
        def __init__(self, parent, controller):
            ctk.CTkFrame.__init__(self, parent)
            self.controller = controller

            self.grid_rowconfigure(0, weight=1)
            self.grid_columnconfigure(0, weight=1)

            left_frame = ctk.CTkFrame(self)
            left_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

            instructions_frame = ctk.CTkFrame(self)
            instructions_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

            instructions_label = ctk.CTkLabel(instructions_frame, text="Instructions:\n\n1. Load a CSV file by clicking 'Load CSV'.\n2. Click 'Next' to proceed to select Deployment IDs.\n3. Select the IDs and click 'Next' again to select Common Names.\n4. Finally, filter and save the CSV file.", justify="left")
            instructions_label.pack(pady=10, padx=10)

            load_button = ctk.CTkButton(left_frame, text="Load CSV", command=self.controller.load_csv)
            load_button.pack(pady=20)

            next_button = ctk.CTkButton(left_frame, text="Next", command=self.next_page)
            next_button.pack()

        def next_page(self):
            if self.controller.df is None:
                messagebox.showwarning("Warning", "Please load a CSV file before proceeding.")
                return
            self.controller.show_frame("DeploymentPage")

    class DeploymentPage(ctk.CTkFrame):
        def __init__(self, parent, controller):
            ctk.CTkFrame.__init__(self, parent)
            self.controller = controller

            self.grid_rowconfigure(0, weight=1)
            self.grid_columnconfigure(0, weight=1)

            left_frame = ctk.CTkFrame(self)
            left_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

            instructions_frame = ctk.CTkFrame(self)
            instructions_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

            instructions_label = ctk.CTkLabel(instructions_frame, text="Instructions:\n\n1. Select Deployment IDs from the list.\n2. Use 'Select All' if needed.\n3. Click 'Next' to proceed.", justify="left")
            instructions_label.pack(pady=10, padx=10)

            label = ctk.CTkLabel(left_frame, text="Select Deployment IDs")
            label.pack(pady=10)

            self.deployment_listbox = tk.Listbox(left_frame, selectmode=tk.EXTENDED, height=15, width=50)
            self.deployment_listbox.pack(pady=5)

            select_all_button = ctk.CTkButton(left_frame, text="Select All Deployment IDs", command=self.select_all)
            select_all_button.pack(pady=5)

            back_button = ctk.CTkButton(left_frame, text="Back", command=lambda: controller.show_frame("StartPage"))
            back_button.pack(pady=5)

            next_button = ctk.CTkButton(left_frame, text="Next", command=self.next_page)
            next_button.pack()

        def update_listbox(self, items):
            self.deployment_listbox.delete(0, tk.END)
            for item in items:
                self.deployment_listbox.insert(tk.END, item)

        def select_all(self):
            self.deployment_listbox.select_set(0, tk.END)

        def next_page(self):
            selected_items = [self.deployment_listbox.get(i) for i in self.deployment_listbox.curselection()]
            if not selected_items:
                messagebox.showwarning("Warning", "Please select at least one deployment ID")
                return
            self.controller.set_selected_deployments(selected_items)
            self.controller.show_frame("CommonNamePage")

    class CommonNamePage(ctk.CTkFrame):
        def __init__(self, parent, controller):
            ctk.CTkFrame.__init__(self, parent)
            self.controller = controller

            self.grid_rowconfigure(0, weight=1)
            self.grid_columnconfigure(0, weight=1)

            left_frame = ctk.CTkFrame(self)
            left_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

            instructions_frame = ctk.CTkFrame(self)
            instructions_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

            instructions_label = ctk.CTkLabel(instructions_frame, text="Instructions:\n\n1. Select Common Names from the list.\n2. Use 'Select All' if needed.\n3. Click 'Next' to filter the CSV.", justify="left")
            instructions_label.pack(pady=10, padx=10)

            label = ctk.CTkLabel(left_frame, text="Select Common Names")
            label.pack(pady=10)

            self.common_name_listbox = tk.Listbox(left_frame, selectmode=tk.EXTENDED, height=15, width=50)
            self.common_name_listbox.pack(pady=5)

            select_all_button = ctk.CTkButton(left_frame, text="Select All Common Names", command=self.select_all)
            select_all_button.pack(pady=5)

            back_button = ctk.CTkButton(left_frame, text="Back", command=lambda: controller.show_frame("DeploymentPage"))
            back_button.pack(pady=5)

            next_button = ctk.CTkButton(left_frame, text="Next", command=self.next_page)
            next_button.pack()

        def update_listbox(self, items):
            self.common_name_listbox.delete(0, tk.END)
            for item in items:
                self.common_name_listbox.insert(tk.END, item)

        def select_all(self):
            self.common_name_listbox.select_set(0, tk.END)

        def next_page(self):
            selected_items = [self.common_name_listbox.get(i) for i in self.common_name_listbox.curselection()]
            if not selected_items:
                messagebox.showwarning("Warning", "Please select at least one common name")
                return
            self.controller.set_selected_common_names(selected_items)
            self.controller.filter_csv()
            self.controller.show_frame("ResultPage")

    class ResultPage(ctk.CTkFrame):
        def __init__(self, parent, controller):
            ctk.CTkFrame.__init__(self, parent)
            self.controller = controller

            self.grid_rowconfigure(0, weight=1)
            self.grid_columnconfigure(0, weight=1)

            left_frame = ctk.CTkFrame(self)
            left_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

            instructions_frame = ctk.CTkFrame(self)
            instructions_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

            instructions_label = ctk.CTkLabel(instructions_frame, text="Instructions:\n\n1. Review the number of filtered rows.\n2. Click 'Save Filtered CSV' to save the file.", justify="left")
            instructions_label.pack(pady=10, padx=10)

            self.result_label = ctk.CTkLabel(left_frame, text="")
            self.result_label.pack(pady=10)

            back_button = ctk.CTkButton(left_frame, text="Back", command=lambda: controller.show_frame("CommonNamePage"))
            back_button.pack(pady=5)

            save_button = ctk.CTkButton(left_frame, text="Save Filtered CSV", command=self.controller.save_csv)
            save_button.pack(pady=5)

        def update_result_text(self, text):
            self.result_label.configure(text=text)


class GCSDownloaderAndRenamer(ctk.CTkFrame):
    def __init__(self, root):
        super().__init__(root)  # Correct call to the base class constructor
        self.root = root
        self.pack(fill="both", expand=True)  # Ensure the frame adjusts to the window size
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")
        
        # Main frame to hold all widgets
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Label for CSV file selection
        self.label1 = ctk.CTkLabel(self.main_frame, text="Select the images CSV file with URLs from Wildlife Insights", anchor="w")
        self.label1.pack(pady=10)

        # Frame to hold CSV selection widgets
        self.selection_frame = ctk.CTkFrame(self.main_frame)
        self.selection_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select CSV file
        self.select_csv_btn = ctk.CTkButton(self.selection_frame, text="Browse CSV", command=self.select_csv)
        self.select_csv_btn.pack(side="left", padx=5)

        # Label to display the selected CSV file name
        self.csv_label = ctk.CTkLabel(self.selection_frame, text="No CSV file selected", anchor="w")
        self.csv_label.pack(side="left", padx=5)

        # Add option for single or multiple folders
        self.use_multiple_folders = IntVar()
        self.multiple_folders_checkbtn = ctk.CTkCheckBox(self.main_frame, text="Save in separate folders by deployment_id", variable=self.use_multiple_folders)
        self.multiple_folders_checkbtn.pack(pady=10)

        # Status label for showing download progress
        self.status_var = StringVar()
        self.status_label = ctk.CTkLabel(self.main_frame, textvariable=self.status_var)
        self.status_label.pack(pady=10)

        # Buttons for processing
        self.buttons_frame = ctk.CTkFrame(self.main_frame)
        self.buttons_frame.pack(pady=10)

        self.download_btn = ctk.CTkButton(self.buttons_frame, text="Download Images", command=self.start_download, state=ctk.DISABLED)
        self.download_btn.pack(side="left", padx=5)

        self.stop_btn = ctk.CTkButton(self.buttons_frame, text="Stop", command=self.stop_download, state=ctk.DISABLED)
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
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")
        
        # Create the tab panel
        self.tabview = ctk.CTkTabview(root)
        self.tabview.pack(pady=20, padx=20, fill='both', expand=True)

        # Add tabs
        self.tab1 = self.tabview.add("Initial Excel")
        self.tab2 = self.tabview.add("Images CSV")
        self.tab3 = self.tabview.add("Deployments CSV")
        self.tab4 = self.tabview.add("Destination Folder")

        # Content of the "Initial Excel" tab
        self.label1 = ctk.CTkLabel(self.tab1, text="Select the Initial Excel file", anchor="w")
        self.label1.pack(pady=10)

        self.excel_frame = ctk.CTkFrame(self.tab1)
        self.excel_frame.pack(pady=5, padx=10, fill="x")

        self.select_excel_btn = ctk.CTkButton(self.excel_frame, text="Browse", command=self.select_initial_excel)
        self.select_excel_btn.pack(side="left", padx=5)

        self.excel_label = ctk.CTkLabel(self.excel_frame, text="No Excel file selected", anchor="w")
        self.excel_label.pack(side="left", padx=5)

        # Content of the "Images CSV" tab
        self.label2 = ctk.CTkLabel(self.tab2, text="Select the images CSV file", anchor="w")
        self.label2.pack(pady=10)

        self.images_frame = ctk.CTkFrame(self.tab2)
        self.images_frame.pack(pady=5, padx=10, fill="x")

        self.select_images_btn = ctk.CTkButton(self.images_frame, text="Browse", command=self.select_images_csv)
        self.select_images_btn.pack(side="left", padx=5)

        self.images_label = ctk.CTkLabel(self.images_frame, text="No images CSV file selected", anchor="w")
        self.images_label.pack(side="left", padx=5)

        # Content of the "Deployments CSV" tab
        self.label3 = ctk.CTkLabel(self.tab3, text="Select the deployments CSV file", anchor="w")
        self.label3.pack(pady=10)

        self.deployments_frame = ctk.CTkFrame(self.tab3)
        self.deployments_frame.pack(pady=5, padx=10, fill="x")

        self.select_deployments_btn = ctk.CTkButton(self.deployments_frame, text="Browse", command=self.select_deployments_csv)
        self.select_deployments_btn.pack(side="left", padx=5)

        self.deployments_label = ctk.CTkLabel(self.deployments_frame, text="No deployments CSV file selected", anchor="w")
        self.deployments_label.pack(side="left", padx=5)

        # Content of the "Destination Folder" tab
        self.label4 = ctk.CTkLabel(self.tab4, text="Select the destination folder for the final Excel file", anchor="w")
        self.label4.pack(pady=10)

        self.folder_frame = ctk.CTkFrame(self.tab4)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        self.select_folder_btn = ctk.CTkButton(self.folder_frame, text="Browse", command=self.select_folder)
        self.select_folder_btn.pack(side="left", padx=5)

        self.folder_label = ctk.CTkLabel(self.folder_frame, text="No folder selected", anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Section for additional options
        self.options_frame = ctk.CTkFrame(root)
        self.options_frame.pack(pady=10, padx=10, fill='x')

        # Checkbox for multiple images processing
        self.multiple_images_var = ctk.BooleanVar(value=False)
        self.multiple_images_checkbox = ctk.CTkCheckBox(self.options_frame, text="Process with Multiple Images", variable=self.multiple_images_var)
        self.multiple_images_checkbox.pack(side="left", padx=5)

        # Label and entry for time threshold
        self.time_threshold_label = ctk.CTkLabel(self.options_frame, text="Time threshold (seconds):", anchor="w")
        self.time_threshold_label.pack(side="left", padx=5)

        self.time_threshold_entry = ctk.CTkEntry(self.options_frame)
        self.time_threshold_entry.pack(side="left", padx=5)
        self.time_threshold_entry.insert(0, "3")  # Default value

        # Buttons for processing
        self.buttons_frame = ctk.CTkFrame(root)
        self.buttons_frame.pack(pady=10)

        self.process_btn = ctk.CTkButton(self.buttons_frame, text="Process Files", command=self.process_files, state=ctk.DISABLED)
        self.process_btn.pack(side="left", padx=5)

        self.save_file_btn = ctk.CTkButton(self.buttons_frame, text="Save File", command=self.save_file, state=ctk.DISABLED)
        self.save_file_btn.pack(side="left", padx=5)

        self.initial_excel_path = None
        self.images_csv_path = None
        self.deployments_csv_path = None
        self.destination_folder = None
        self.final_df = None  # To store the final DataFrame

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

    def select_folder(self):
        self.destination_folder = filedialog.askdirectory()
        if self.destination_folder:
            self.folder_label.configure(text=self.destination_folder)
        self.check_all_selected()

    def check_all_selected(self):
        if self.initial_excel_path and self.images_csv_path and self.deployments_csv_path and self.destination_folder:
            self.process_btn.configure(state=ctk.NORMAL)

    def process_files(self):
        if self.multiple_images_var.get():
            self.process_multiple_images()
        else:
            self.process_single_image()

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
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")

        # Frame to hold all widgets
        self.main_frame = ctk.CTkFrame(root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Label for folder selection
        self.label1 = ctk.CTkLabel(self.main_frame, text="Select the folder containing files", anchor="w")
        self.label1.pack(pady=10)

        # Frame to hold folder selection widgets
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select a folder
        self.select_folder_btn = ctk.CTkButton(self.folder_frame, text="Browse Folder", command=self.select_folder)
        self.select_folder_btn.pack(side="left", padx=5)

        # Label to display the selected folder name
        self.folder_label = ctk.CTkLabel(self.folder_frame, text="No folder selected", anchor="w")
        self.folder_label.pack(side="left", padx=5)

        # Frame to hold date options
        self.date_frame = ctk.CTkFrame(self.main_frame)
        self.date_frame.pack(pady=10, padx=10, fill="x")

        # Label and entry for setting the real date
        self.label2 = ctk.CTkLabel(self.date_frame, text="Set the Real Date (YYYY-MM-DD HH:MM:SS)", anchor="w")
        self.label2.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.real_date_entry = ctk.CTkEntry(self.date_frame)
        self.real_date_entry.grid(row=0, column=1, padx=10, pady=5, sticky="we")

        # Radio buttons and labels for date options
        self.date_option = ctk.IntVar()
        self.date_option.set(1)

        self.newest_date_radio = ctk.CTkRadioButton(self.date_frame, text="Use newest date", variable=self.date_option, value=1)
        self.newest_date_radio.grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.newest_date = ctk.StringVar()
        self.newest_date_label = ctk.CTkLabel(self.date_frame, textvariable=self.newest_date, anchor="w")
        self.newest_date_label.grid(row=1, column=1, padx=10, pady=2, sticky="w")

        self.oldest_date_radio = ctk.CTkRadioButton(self.date_frame, text="Use oldest date", variable=self.date_option, value=2)
        self.oldest_date_radio.grid(row=2, column=0, padx=5, pady=2, sticky="w")

        self.oldest_date = ctk.StringVar()
        self.oldest_date_label = ctk.CTkLabel(self.date_frame, textvariable=self.oldest_date, anchor="w")
        self.oldest_date_label.grid(row=2, column=1, padx=10, pady=2, sticky="w")

        self.custom_date_radio = ctk.CTkRadioButton(self.date_frame, text="Use custom date", variable=self.date_option, value=3)
        self.custom_date_radio.grid(row=3, column=0, padx=5, pady=2, sticky="w")

        self.custom_date_entry = ctk.CTkEntry(self.date_frame)
        self.custom_date_entry.grid(row=3, column=1, padx=10, pady=2, sticky="we")

        # Frame to hold the Change Dates and Copy buttons
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(pady=10)

        # Change Dates button
        self.change_dates_btn = ctk.CTkButton(self.button_frame, text="Change Dates", command=self.change_dates)
        self.change_dates_btn.pack(side="left", padx=10)

        # Copy to Folder button
        self.copy_btn = ctk.CTkButton(self.button_frame, text="Copy to Folder", command=self.copy_to_folder)
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
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")

        # Create the tab panel
        self.tabview = ctk.CTkTabview(root, width=400, height=300)
        self.tabview.pack(pady=20, padx=20)

        # Add tabs
        self.source_tab = self.tabview.add("Source")
        self.settings_tab = self.tabview.add("Settings")
        self.dest_tab = self.tabview.add("Destination")

        # Content of the "Source" tab
        self.label1 = ctk.CTkLabel(self.source_tab, text="Select the source directory:", anchor="w")
        self.label1.pack(pady=10)

        # Frame to hold source folder selection widgets
        self.source_folder_frame = ctk.CTkFrame(self.source_tab)
        self.source_folder_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select the source folder
        self.select_source_folder_btn = ctk.CTkButton(self.source_folder_frame, text="Browse Source", command=self.select_source_folder)
        self.select_source_folder_btn.pack(side="left", padx=5)

        # Label to display the selected source folder name
        self.source_folder_label = ctk.CTkLabel(self.source_folder_frame, text="No source folder selected", anchor="w")
        self.source_folder_label.pack(side="left", padx=5)

        # Content of the "Settings" tab
        self.label2 = ctk.CTkLabel(self.settings_tab, text="Settings", anchor="w")
        self.label2.pack(pady=10)

        # Frame to hold checkboxes in parallel
        self.checkbox_frame = ctk.CTkFrame(self.settings_tab)
        self.checkbox_frame.pack(pady=10, padx=10, fill="x")

        # Checkbox to choose whether to use the folder name for renaming
        self.use_folder_name_var = BooleanVar(value=True)
        self.checkbox_use_folder_name = ctk.CTkCheckBox(self.checkbox_frame, text="Use folder name", variable=self.use_folder_name_var)
        self.checkbox_use_folder_name.pack(side="left", padx=5)

        # Checkbox to choose whether to keep the original filename
        self.keep_original_name_var = BooleanVar(value=False)
        self.checkbox_keep_original_name = ctk.CTkCheckBox(self.checkbox_frame, text="Keep original filename", variable=self.keep_original_name_var)
        self.checkbox_keep_original_name.pack(side="left", padx=5)

        # Checkbox to choose whether to add EXIF date to the filename
        self.add_exif_date_var = BooleanVar(value=False)
        self.checkbox_add_exif_date = ctk.CTkCheckBox(self.checkbox_frame, text="Add EXIF date", variable=self.add_exif_date_var)
        self.checkbox_add_exif_date.pack(side="left", padx=5)

        # Checkbox to choose whether to replace spaces with underscores
        self.replace_spaces_var = BooleanVar(value=False)
        self.checkbox_replace_spaces = ctk.CTkCheckBox(self.checkbox_frame, text="Replace spaces", variable=self.replace_spaces_var)
        self.checkbox_replace_spaces.pack(side="left", padx=5)

        # Frame to hold the custom text input option and its entry field
        self.custom_text_frame = ctk.CTkFrame(self.settings_tab)
        self.custom_text_frame.pack(pady=10, padx=10, fill="x")

        # Checkbox to select whether to use custom text for renaming
        self.use_custom_text_var = BooleanVar(value=False)
        self.checkbox_use_custom_text = ctk.CTkCheckBox(self.custom_text_frame, text="Use custom text", variable=self.use_custom_text_var, command=self.toggle_custom_text_entry)
        self.checkbox_use_custom_text.pack(side="left", padx=5)

        # Entry field for custom text input
        self.custom_text_var = StringVar(value="")
        self.custom_text_entry = ctk.CTkEntry(self.custom_text_frame, textvariable=self.custom_text_var, state="disabled", width=200)
        self.custom_text_entry.pack(side="left", padx=5)

        # Button to rename and overwrite photos in Settings tab
        self.rename_and_copy_btn_source = ctk.CTkButton(self.settings_tab, text="Rename and Overwrite", command=self.rename_and_copy_photos)
        self.rename_and_copy_btn_source.pack(pady=20)

        # Content of the "Destination" tab
        self.label3 = ctk.CTkLabel(self.dest_tab, text="Select the destination directory if you want to move the photos:", anchor="w")
        self.label3.pack(pady=10)

        # Frame to hold destination folder selection widgets
        self.dest_folder_frame = ctk.CTkFrame(self.dest_tab)
        self.dest_folder_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select the destination folder
        self.select_dest_folder_btn = ctk.CTkButton(self.dest_folder_frame, text="Browse Destination", command=self.select_dest_folder)
        self.select_dest_folder_btn.pack(side="left", padx=5)

        # Label to display the selected destination folder name
        self.dest_folder_label = ctk.CTkLabel(self.dest_folder_frame, text="No destination folder selected", anchor="w")
        self.dest_folder_label.pack(side="left", padx=5)

        # Checkbox to choose whether to copy photos to the destination folder
        self.copy_photos_var = BooleanVar(value=False)  # Added this line
        self.checkbox_copy_photos = ctk.CTkCheckBox(self.dest_tab, text="Copy photos to destination folder", variable=self.copy_photos_var)
        self.checkbox_copy_photos.pack(pady=10)

        # Button to rename and move photos in Destination tab
        self.rename_and_copy_btn_dest = ctk.CTkButton(self.dest_tab, text="Rename and Move", command=self.rename_and_copy_photos)
        self.rename_and_copy_btn_dest.pack(pady=20)

        self.source_folder = None
        self.dest_folder = None

    def toggle_custom_text_entry(self):
        """Toggle the state of the custom text entry field based on the checkbox."""
        if self.use_custom_text_var.get():
            self.custom_text_entry.configure(state="normal")
        else:
            self.custom_text_entry.configure(state="disabled")

    def select_source_folder(self):
        self.source_folder = filedialog.askdirectory()
        if self.source_folder:
            self.source_folder_label.configure(text=self.source_folder)
            messagebox.showinfo("Information", f"Selected source folder: {self.source_folder}")

    def select_dest_folder(self):
        self.dest_folder = filedialog.askdirectory()
        if self.dest_folder:
            self.dest_folder_label.configure(text=self.dest_folder)
            messagebox.showinfo("Information", f"Selected destination folder: {self.dest_folder}")

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

        source_path = Path(self.source_folder)
        dest_path = Path(self.dest_folder) if self.copy_photos_var.get() else None

        if not source_path.is_dir() or (dest_path and not dest_path.is_dir()):
            messagebox.showerror("Error", "Invalid directory selected.")
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

                    new_name = f"{base_name}{exif_date} {original_name}{file.suffix}" if self.keep_original_name_var.get() else f"{base_name}{exif_date}{file.suffix}"

                    # Replace spaces with underscores if the option is selected
                    if self.replace_spaces_var.get():
                        new_name = new_name.replace(" ", "_")

                    new_path = (dest_path / new_name) if self.copy_photos_var.get() else (root_path / new_name)

                    # Check if the new file name already exists and modify it if necessary
                    counter = 1
                    original_new_path = new_path
                    while new_path.exists():
                        new_name = original_new_path.stem + f"_{counter}" + original_new_path.suffix
                        new_path = new_path.parent / new_name
                        counter += 1

                    if self.copy_photos_var.get():
                        shutil.copy2(file, new_path)
                    else:
                        file.rename(new_path)

        if self.copy_photos_var.get():
            messagebox.showinfo("Information", f"Photos have been renamed and moved to '{self.dest_folder}' successfully.")
        else:
            messagebox.showinfo("Information", "Photos have been renamed and overwritten successfully.")


class LynxOne:
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("System")  # Appearance mode ("System", "Dark", "Light")

        # Create the tab panel
        self.tabview = ctk.CTkTabview(root, width=400, height=300)
        self.tabview.pack(pady=20, padx=20, fill="both", expand=True)
        self.tabview.pack_propagate(False)

        # Add tabs
        self.source_tab = self.tabview.add("Source")
        self.settings_tab = self.tabview.add("Settings")
        self.optional_tab = self.tabview.add("Optional Files")  # New tab for optional files

        # Content of the "Source" tab
        self.label1 = ctk.CTkLabel(self.source_tab, text="Select the source directory.", anchor="w")
        self.label1.pack(pady=10)

        # Frame to hold source folder selection widgets
        self.source_folder_frame = ctk.CTkFrame(self.source_tab)
        self.source_folder_frame.pack(pady=5, padx=10, fill="x")

        # Button to browse and select the source folder
        self.select_source_folder_btn = ctk.CTkButton(self.source_folder_frame, text="Browse Source", command=self.select_source_folder)
        self.select_source_folder_btn.pack(side="left", padx=5)

        # Label to display the selected source folder name
        self.source_folder_label = ctk.CTkLabel(self.source_folder_frame, text="No source folder selected", anchor="w")
        self.source_folder_label.pack(side="left", padx=5)

        # Content of the "Settings" tab
        self.label2 = ctk.CTkLabel(self.settings_tab, text="Settings for grouping images.", anchor="w")
        self.label2.pack(pady=10)

        # Frame for the minutes input
        self.minutes_frame = ctk.CTkFrame(self.settings_tab)
        self.minutes_frame.pack(pady=5, padx=10, fill="x")

        self.minutes_label = ctk.CTkLabel(self.minutes_frame, text="Group by minutes:")
        self.minutes_label.pack(side="left", padx=5)

        self.minutes_entry = ctk.CTkEntry(self.minutes_frame, placeholder_text="Minutes (0 to disable)")
        self.minutes_entry.pack(side="left", padx=5)

        # Checkbox to indicate if the Lince/linces folder exists
        self.lince_checkbox_var = ctk.BooleanVar(value=True)
        self.checkbox_lince_exists = ctk.CTkCheckBox(self.settings_tab, text="Does the Lince/Linces folder exist?", variable=self.lince_checkbox_var)
        self.checkbox_lince_exists.pack(pady=10)

        # Checkbox to indicate if the Revision folder exists
        self.revision_checkbox_var = ctk.BooleanVar(value=False)
        self.checkbox_revision_exists = ctk.CTkCheckBox(self.settings_tab, text="Does the Revision folder exist?", variable=self.revision_checkbox_var)
        self.checkbox_revision_exists.pack(pady=10)

        # Content of the "Optional Files" tab
        self.label3 = ctk.CTkLabel(self.optional_tab, text="Optional: Select additional Excel files for joining.", anchor="w")
        self.label3.pack(pady=10)

        # Frame for Estaciones file
        self.estaciones_frame = ctk.CTkFrame(self.optional_tab)
        self.estaciones_frame.pack(pady=5, padx=10, fill="x")
        self.estaciones_file_label = ctk.CTkLabel(self.estaciones_frame, text="No Estaciones file selected", anchor="w")
        self.estaciones_file_label.pack(side="left", padx=5)
        self.select_estaciones_file_btn = ctk.CTkButton(self.estaciones_frame, text="Browse Estaciones", command=self.select_estaciones_file)
        self.select_estaciones_file_btn.pack(side="left", padx=5)

        # Frame for Individuos file
        self.individuos_frame = ctk.CTkFrame(self.optional_tab)
        self.individuos_frame.pack(pady=5, padx=10, fill="x")
        self.individuos_file_label = ctk.CTkLabel(self.individuos_frame, text="No Individuos file selected", anchor="w")
        self.individuos_file_label.pack(side="left", padx=5)
        self.select_individuos_file_btn = ctk.CTkButton(self.individuos_frame, text="Browse Individuos", command=self.select_individuos_file)
        self.select_individuos_file_btn.pack(side="left", padx=5)

        self.source_folder = None
        self.estaciones_file = None
        self.individuos_file = None
        self.excel_data = None

        # Buttons for generating and downloading Excel files
        self.generate_button = ctk.CTkButton(root, text="Generate Excel", command=self.generate_excel)
        self.generate_button.pack(pady=10, side="left", padx=20)

        self.download_button = ctk.CTkButton(root, text="Download Excel", command=self.save_excel)
        self.download_button.pack(pady=10, side="left", padx=20)

    def select_source_folder(self):
        self.source_folder = filedialog.askdirectory()
        if self.source_folder:
            self.source_folder_label.configure(text=self.source_folder)
            messagebox.showinfo("Information", f"Selected source folder: {self.source_folder}")

    def select_estaciones_file(self):
        self.estaciones_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.estaciones_file:
            self.estaciones_file_label.configure(text=self.estaciones_file)
            messagebox.showinfo("Information", f"Selected Estaciones file: {self.estaciones_file}")

    def select_individuos_file(self):
        self.individuos_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.individuos_file:
            self.individuos_file_label.configure(text=self.individuos_file)
            messagebox.showinfo("Information", f"Selected Individuos file: {self.individuos_file}")

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
            messagebox.showwarning("Warning", "Please select a source folder first.")
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
