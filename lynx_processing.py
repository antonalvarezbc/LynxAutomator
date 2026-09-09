"""Spreadsheet transformations executed by background workers, without Tk calls.

The existing transformations live here so both desktop editions use the same
processing code. Inputs are plain values captured before the task starts.
"""
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import os
import re
import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS
from lynx_core import merge_deployments
from lynx_tasks import TaskContext


def read_exif(path):
    try:
        with Image.open(path) as image:
            data = image._getexif() or {}
            return {TAGS.get(tag, tag): value for tag, value in data.items()}
    except (OSError, ValueError, AttributeError):
        return {}


def date_taken(data):
    value = data.get('DateTimeOriginal')
    if not value:
        return None
    try:
        return datetime.strptime(value, '%Y:%m:%d %H:%M:%S')
    except (TypeError, ValueError):
        return None

@dataclass
class FolderProcessor:
    task: TaskContext
    folder_path: str
    file_path: str
    group_images: bool
    threshold: int
    get_exif_data = staticmethod(read_exif)
    get_date_taken = staticmethod(date_taken)

    def run(self):
        self.task.report('Reading image dates')
        photo_data = []
        for filename in os.listdir(self.folder_path):
            self.task.checkpoint()
            if filename.lower().endswith(('png', 'jpg', 'jpeg')):
                file_path = os.path.join(self.folder_path, filename)
                exif_data = self.get_exif_data(file_path)
                if exif_data:
                    date_taken = self.get_date_taken(exif_data)
                    if date_taken:
                        photo_data.append([filename, date_taken])
        if not photo_data:
            raise ValueError('No photos with capture dates found.')
        self.task.report('Reading Excel')
        df = pd.read_excel(self.file_path)
        if df.empty:
            raise ValueError('The template needs an initial row.')
        if 'Encounter.mediaAsset0' not in df.columns:
            df['Encounter.mediaAsset0'] = ''
        df['Encounter.mediaAsset0'] = df['Encounter.mediaAsset0'].astype(str)
        new_data = pd.DataFrame(photo_data, columns=['Encounter.mediaAsset0', 'Date'])
        new_data['Encounter.mediaAsset0'] = new_data['Encounter.mediaAsset0'].astype(str)
        new_data['Date'] = pd.to_datetime(new_data['Date'])
        new_data.sort_values(by='Date', inplace=True)
        if self.group_images:
            time_threshold = self.threshold
            new_data['TimeDiff'] = new_data['Date'].diff().dt.total_seconds().fillna(0)
            grouped_data = []
            current_group = []
            for index, row in new_data.iterrows():
                self.task.checkpoint()
                if current_group and row['TimeDiff'] > time_threshold:
                    grouped_data.append(current_group)
                    current_group = []
                current_group.append(row)
            if current_group:
                grouped_data.append(current_group)
            final_data = []
            for group in grouped_data:
                self.task.checkpoint()
                base_row = group[0].copy()
                for i, additional_row in enumerate(group[1:], start=1):
                    self.task.checkpoint()
                    base_row[f'Encounter.mediaAsset{i}'] = additional_row['Encounter.mediaAsset0']
                final_data.append(base_row)
            final_df = pd.DataFrame(final_data).drop(columns=['TimeDiff'])
        else:
            final_df = new_data.copy()
        final_df['Encounter.year'] = final_df['Date'].dt.year
        final_df['Encounter.month'] = final_df['Date'].dt.month
        final_df['Encounter.day'] = final_df['Date'].dt.day
        final_df['Encounter.hour'] = final_df['Date'].dt.hour
        final_df['Encounter.minutes'] = final_df['Date'].dt.minute
        final_df = final_df.drop('Date', axis=1)
        df_merged = pd.merge(df, final_df, on='Encounter.mediaAsset0', how='right')
        for column in df.columns:
            self.task.checkpoint()
            if column in df_merged.columns:
                first_row_value = df[column].iloc[0]
                df_merged[column] = df_merged[column].fillna(first_row_value)
        for time_unit in ['year', 'month', 'day', 'hour', 'minutes']:
            self.task.checkpoint()
            column_x = f'Encounter.{time_unit}_x'
            column_y = f'Encounter.{time_unit}_y'
            column = f'Encounter.{time_unit}'
            if column_x in df_merged.columns and column_y in df_merged.columns:
                df_merged[column_x].replace('', pd.NA, inplace=True)
                df_merged[column_y].replace('', pd.NA, inplace=True)
                df_merged[column] = df_merged[column_x].combine_first(df_merged[column_y])
                df_merged.drop([column_x, column_y], axis=1, inplace=True)
            elif column_y in df_merged.columns:
                df_merged.rename(columns={column_y: column}, inplace=True)
        self.task.checkpoint()
        return df_merged


@dataclass
class CatalogProcessor:
    task: TaskContext
    folder_path: str
    file_path: str
    capitalize_ids: bool
    collapse: bool
    get_exif_data = staticmethod(read_exif)
    get_date_taken = staticmethod(date_taken)

    def run(self):
        self.task.report('Reading catalog')
        photo_data = []
        for root, dirs, files in os.walk(self.folder_path):
            self.task.checkpoint()
            individual_name = {file: os.path.splitext(file)[0].split()[0] for file in files if file.lower().endswith(('png', 'jpg', 'jpeg'))}
            if self.capitalize_ids:
                individual_name = {file: name.capitalize() for file, name in individual_name.items()}
            for file in files:
                self.task.checkpoint()
                if file.lower().endswith(('png', 'jpg', 'jpeg')):
                    file_path = os.path.join(root, file)
                    relative_path = os.path.relpath(file_path, self.folder_path)
                    relative_path = relative_path.replace('\\', '/')
                    name_used = individual_name[file]
                    photo_data.append([relative_path, name_used])
        if not photo_data:
            raise ValueError('No photos found.')
        self.task.report('Reading Excel')
        df_original = pd.read_excel(self.file_path)
        if df_original.empty:
            raise ValueError('The template needs an initial row.')
        original_columns = df_original.columns.tolist()
        first_row_data = df_original.iloc[0]
        if 'Encounter.mediaAsset0' not in df_original.columns:
            df_original['Encounter.mediaAsset0'] = pd.NA
        df_original['Encounter.mediaAsset0'] = df_original['Encounter.mediaAsset0'].astype(str)
        new_data = pd.DataFrame(photo_data, columns=['Encounter.mediaAsset0', 'MarkedIndividual.individualID'])
        new_data['Encounter.mediaAsset0'] = new_data['Encounter.mediaAsset0'].astype(str)
        df_merged = pd.merge(df_original, new_data, on='Encounter.mediaAsset0', how='right')
        if 'MarkedIndividual.individualID_x' in df_merged.columns and 'MarkedIndividual.individualID_y' in df_merged.columns:
            df_merged['MarkedIndividual.individualID'] = df_merged['MarkedIndividual.individualID_y'].fillna(df_merged['MarkedIndividual.individualID_x'])
            df_merged.drop(['MarkedIndividual.individualID_x', 'MarkedIndividual.individualID_y'], axis=1, inplace=True)
        elif 'MarkedIndividual.individualID_y' in df_merged.columns:
            df_merged.rename(columns={'MarkedIndividual.individualID_y': 'MarkedIndividual.individualID'}, inplace=True)
        elif 'MarkedIndividual.individualID_x' in df_merged.columns:
            df_merged.rename(columns={'MarkedIndividual.individualID_x': 'MarkedIndividual.individualID'}, inplace=True)
        if self.collapse:
            df_merged['RowNumber'] = df_merged.groupby('MarkedIndividual.individualID').cumcount()
            df_pivot = df_merged.pivot_table(index='MarkedIndividual.individualID', columns='RowNumber', values='Encounter.mediaAsset0', aggfunc='first')
            df_pivot.columns = [f'Encounter.mediaAsset{int(col)}' for col in df_pivot.columns]
            df_merged = pd.merge(df_merged.drop(columns='Encounter.mediaAsset0').drop_duplicates('MarkedIndividual.individualID'), df_pivot, on='MarkedIndividual.individualID')
            df_merged.drop(columns=['RowNumber'], inplace=True)
        new_column_order = [col for col in original_columns if col in df_merged.columns] + [col for col in df_merged.columns if col not in original_columns]
        df_merged = df_merged[new_column_order]
        for column in df_merged.columns:
            self.task.checkpoint()
            if df_merged[column].isnull().any():
                if column in first_row_data.index:
                    df_merged[column] = df_merged[column].fillna(first_row_data[column])
        self.task.checkpoint()
        return df_merged


@dataclass
class WIProcessor:
    task: TaskContext
    initial_excel_path: str
    images_csv_path: str
    deployments_csv_path: str
    group_images: bool
    separate_objects: bool
    threshold: int
    get_exif_data = staticmethod(read_exif)
    get_date_taken = staticmethod(date_taken)

    def run(self):
        self.task.report('Reading images CSV')
        images_df = pd.read_csv(self.images_csv_path, dtype=str, low_memory=False)
        self.task.checkpoint()
        self.task.report('Reading deployments CSV')
        deployments_df = pd.read_csv(self.deployments_csv_path, dtype=str, low_memory=False)
        required_merge_cols = ['project_id', 'deployment_id']
        for col in required_merge_cols:
            self.task.checkpoint()
            if col not in images_df.columns:
                raise ValueError(f'Missing image CSV column: {col}')
            if col not in deployments_df.columns:
                raise ValueError(f'Missing deployment CSV column: {col}')
        self.task.checkpoint()
        self.task.report('Reading Excel template')
        initial_df_dict = pd.read_excel(self.initial_excel_path, sheet_name=None)
        if not initial_df_dict:
            raise ValueError('El archivo Excel inicial está vacío o no se pudo leer.')
        first_sheet_name = list(initial_df_dict.keys())[0]
        initial_df = initial_df_dict[first_sheet_name].reset_index(drop=True)
        self.task.checkpoint()
        self.task.report('Combining data')
        merged_df = merge_deployments(images_df, deployments_df)
        merged_df = merged_df.reset_index(drop=True)
        required_cols_for_result_df = ['latitude', 'longitude', 'placename', 'location', 'timestamp', 'project_id', 'deployment_id', 'subproject_name']
        if 'number_of_objects' not in merged_df.columns:
            merged_df['number_of_objects'] = '1'
        missing_in_merged = [col for col in required_cols_for_result_df if col not in merged_df.columns]
        if missing_in_merged:
            if 'number_of_objects' not in missing_in_merged and 'number_of_objects' in required_cols_for_result_df:
                pass
            else:
                raise ValueError(f"Columnas requeridas faltantes en los datos combinados: {', '.join(missing_in_merged)}")
        if 'number_of_objects' not in required_cols_for_result_df:
            required_cols_for_result_df.append('number_of_objects')
        result_df = merged_df[required_cols_for_result_df].copy()
        result_df['timestamp'] = pd.to_datetime(result_df['timestamp'], errors='coerce')
        invalid_dates = result_df['timestamp'].isna().sum()
        if invalid_dates:
            raise ValueError(f'{invalid_dates} images have invalid timestamps. Correct the CSV before exporting.')
        self.final_df = pd.DataFrame()
        if self.separate_objects:
            if 'number_of_objects' not in result_df.columns:
                result_df['number_of_objects'] = '1'
            result_df['number_of_objects'] = pd.to_numeric(result_df['number_of_objects'], errors='coerce').fillna(0)
            large_objects_df = result_df[result_df['number_of_objects'] > 1].copy().reset_index(drop=True)
            other_objects_df = result_df[result_df['number_of_objects'] <= 1].copy().reset_index(drop=True)
            processed_dfs = []
            if not large_objects_df.empty:
                processed_large_df = self.process_single_image_per_row(large_objects_df, initial_df)
                processed_dfs.append(processed_large_df)
            if not other_objects_df.empty:
                if self.group_images:
                    processed_other_df = self.process_multiple_images(other_objects_df, initial_df)
                else:
                    processed_other_df = self.process_single_image_per_row(other_objects_df, initial_df)
                processed_dfs.append(processed_other_df)
            if processed_dfs:
                self.final_df = pd.concat(processed_dfs, ignore_index=True)
        elif self.group_images:
            self.final_df = self.process_multiple_images(result_df.copy(), initial_df)
        else:
            self.final_df = self.process_single_image_per_row(result_df.copy(), initial_df)
        self.task.checkpoint()
        return self.final_df

    def process_single_image_per_row(self, result_df_input, initial_df_template):
        """
            Procesa un DataFrame para asegurar que cada imagen o entrada esté en una fila separada.
            Ideal para cuando no se desea agrupar imágenes por tiempo.
    
            Args:
                result_df_input (pd.DataFrame): DataFrame con los datos de imágenes a procesar.
                initial_df_template (pd.DataFrame): DataFrame de plantilla para rellenar columnas.
    
            Returns:
                pd.DataFrame: DataFrame con una imagen por fila y columnas estandarizadas.
            """
        result_df = result_df_input.copy().reset_index(drop=True)
        initial_df = initial_df_template.copy()
        combined_rows = []
        for _, row in result_df.iterrows():
            self.task.checkpoint()
            new_row_dict = {}
            new_row_dict['Encounter.decimalLatitude'] = row['latitude']
            new_row_dict['Encounter.decimalLongitude'] = row['longitude']
            new_row_dict['Encounter.verbatimLocality'] = row['placename']
            media_asset = row['location']
            if pd.notna(media_asset):
                new_row_dict['Encounter.mediaAsset0'] = self.ensure_jpg_extension(media_asset.split('/')[-1])
            else:
                new_row_dict['Encounter.mediaAsset0'] = pd.NA
            new_row_dict['Occurrence.occurrenceID'] = self.generate_occurrence_id(row)
            ts = row['timestamp']
            new_row_dict['Encounter.year'] = ts.year
            new_row_dict['Encounter.month'] = ts.month
            new_row_dict['Encounter.day'] = ts.day
            new_row_dict['Encounter.hour'] = ts.hour
            new_row_dict['Encounter.minutes'] = ts.minute
            for col_template in initial_df.columns:
                self.task.checkpoint()
                if col_template not in new_row_dict:
                    new_row_dict[col_template] = initial_df[col_template].iloc[0] if not initial_df.empty and col_template in initial_df else pd.NA
            combined_rows.append(new_row_dict)
        if not combined_rows:
            temp_final_cols = list(initial_df.columns)
            default_cols_to_ensure = ['Occurrence.occurrenceID', 'Encounter.decimalLatitude', 'Encounter.decimalLongitude', 'Encounter.verbatimLocality', 'Encounter.mediaAsset0', 'Encounter.year', 'Encounter.month', 'Encounter.day', 'Encounter.hour', 'Encounter.minutes']
            for c in default_cols_to_ensure:
                self.task.checkpoint()
                if c not in temp_final_cols:
                    temp_final_cols.insert(0, c)
            return pd.DataFrame(columns=list(dict.fromkeys(temp_final_cols)))
        combined_df = pd.DataFrame(combined_rows)
        final_ordered_columns = ['Occurrence.occurrenceID', 'Encounter.decimalLatitude', 'Encounter.decimalLongitude', 'Encounter.verbatimLocality', 'Encounter.mediaAsset0', 'Encounter.year', 'Encounter.month', 'Encounter.day', 'Encounter.hour', 'Encounter.minutes']
        for col in initial_df.columns:
            self.task.checkpoint()
            if col not in final_ordered_columns and col in combined_df.columns:
                final_ordered_columns.append(col)
        for col in combined_df.columns:
            self.task.checkpoint()
            if col not in final_ordered_columns:
                final_ordered_columns.append(col)
        combined_df = combined_df.reindex(columns=final_ordered_columns, fill_value=pd.NA)
        return combined_df

    def process_multiple_images(self, result_df_input, initial_df_template):
        """
            Procesa un DataFrame para agrupar imágenes en una sola fila si están dentro
            de un umbral de tiempo específico para el mismo despliegue.
    
            Args:
                result_df_input (pd.DataFrame): DataFrame con los datos de imágenes a procesar.
                initial_df_template (pd.DataFrame): DataFrame de plantilla para rellenar columnas.
    
            Returns:
                pd.DataFrame: DataFrame con imágenes agrupadas por tiempo en una sola fila.
            """
        result_df = result_df_input.copy().reset_index(drop=True)
        initial_df = initial_df_template.copy()
        time_threshold = self.threshold
        if not pd.api.types.is_datetime64_any_dtype(result_df['timestamp']):
            result_df['timestamp'] = pd.to_datetime(result_df['timestamp'], errors='coerce')
            invalid_dates = result_df['timestamp'].isna().sum()
            if invalid_dates:
                raise ValueError(f'{invalid_dates} images have invalid timestamps. Correct the CSV before exporting.')
        if result_df.empty:
            temp_final_cols = list(initial_df.columns)
            default_cols_to_ensure = ['Occurrence.occurrenceID', 'Encounter.decimalLatitude', 'Encounter.decimalLongitude', 'Encounter.verbatimLocality', 'Encounter.year', 'Encounter.month', 'Encounter.day', 'Encounter.hour', 'Encounter.minutes']
            for i in range(5):
                self.task.checkpoint()
                default_cols_to_ensure.append(f'Encounter.mediaAsset{i}')
            for c in default_cols_to_ensure:
                self.task.checkpoint()
                if c not in temp_final_cols:
                    temp_final_cols.insert(0, c)
            return pd.DataFrame(columns=list(dict.fromkeys(temp_final_cols)))
        result_df = result_df.sort_values(by=['project_id', 'deployment_id', 'timestamp'])
        all_processed_rows = []
        max_assets_in_any_group = 0
        for deployment_id, group_df in result_df.groupby(['project_id', 'deployment_id']):
            self.task.checkpoint()
            current_group_processed = group_df.copy().reset_index(drop=True)
            current_group_processed['time_diff'] = current_group_processed['timestamp'].diff().dt.total_seconds().fillna(time_threshold + 1)
            image_event_accumulator = []
            for _, image_row in current_group_processed.iterrows():
                self.task.checkpoint()
                if image_event_accumulator and image_row['time_diff'] > time_threshold:
                    if image_event_accumulator:
                        base_event_row_data = image_event_accumulator[0]
                        new_combined_row = {'Encounter.decimalLatitude': base_event_row_data['latitude'], 'Encounter.decimalLongitude': base_event_row_data['longitude'], 'Encounter.verbatimLocality': base_event_row_data['placename'], 'Occurrence.occurrenceID': self.generate_occurrence_id(base_event_row_data), 'Encounter.year': base_event_row_data['timestamp'].year, 'Encounter.month': base_event_row_data['timestamp'].month, 'Encounter.day': base_event_row_data['timestamp'].day, 'Encounter.hour': base_event_row_data['timestamp'].hour, 'Encounter.minutes': base_event_row_data['timestamp'].minute}
                        for i, asset_data_row in enumerate(image_event_accumulator):
                            self.task.checkpoint()
                            asset_location = asset_data_row['location']
                            if pd.notna(asset_location):
                                new_combined_row[f'Encounter.mediaAsset{i}'] = self.ensure_jpg_extension(asset_location.split('/')[-1])
                            else:
                                new_combined_row[f'Encounter.mediaAsset{i}'] = pd.NA
                        all_processed_rows.append(new_combined_row)
                        max_assets_in_any_group = max(max_assets_in_any_group, len(image_event_accumulator))
                        image_event_accumulator = []
                image_event_accumulator.append(image_row)
            if image_event_accumulator:
                base_event_row_data = image_event_accumulator[0]
                new_combined_row = {'Encounter.decimalLatitude': base_event_row_data['latitude'], 'Encounter.decimalLongitude': base_event_row_data['longitude'], 'Encounter.verbatimLocality': base_event_row_data['placename'], 'Occurrence.occurrenceID': self.generate_occurrence_id(base_event_row_data), 'Encounter.year': base_event_row_data['timestamp'].year, 'Encounter.month': base_event_row_data['timestamp'].month, 'Encounter.day': base_event_row_data['timestamp'].day, 'Encounter.hour': base_event_row_data['timestamp'].hour, 'Encounter.minutes': base_event_row_data['timestamp'].minute}
                for i, asset_data_row in enumerate(image_event_accumulator):
                    self.task.checkpoint()
                    asset_location = asset_data_row['location']
                    if pd.notna(asset_location):
                        new_combined_row[f'Encounter.mediaAsset{i}'] = self.ensure_jpg_extension(asset_location.split('/')[-1])
                    else:
                        new_combined_row[f'Encounter.mediaAsset{i}'] = pd.NA
                all_processed_rows.append(new_combined_row)
                max_assets_in_any_group = max(max_assets_in_any_group, len(image_event_accumulator))
        if not all_processed_rows:
            temp_final_cols = list(initial_df.columns)
            default_cols_to_ensure = ['Occurrence.occurrenceID', 'Encounter.decimalLatitude', 'Encounter.decimalLongitude', 'Encounter.verbatimLocality', 'Encounter.year', 'Encounter.month', 'Encounter.day', 'Encounter.hour', 'Encounter.minutes']
            for i in range(max_assets_in_any_group if max_assets_in_any_group > 0 else 1):
                self.task.checkpoint()
                default_cols_to_ensure.append(f'Encounter.mediaAsset{i}')
            for c in default_cols_to_ensure:
                self.task.checkpoint()
                if c not in temp_final_cols:
                    temp_final_cols.insert(0, c)
            return pd.DataFrame(columns=list(dict.fromkeys(temp_final_cols)))
        final_combined_df = pd.DataFrame(all_processed_rows)
        for i in range(max_assets_in_any_group):
            self.task.checkpoint()
            col_name = f'Encounter.mediaAsset{i}'
            if col_name not in final_combined_df.columns:
                final_combined_df[col_name] = pd.NA
        for col_template in initial_df.columns:
            self.task.checkpoint()
            if col_template not in final_combined_df.columns:
                final_combined_df[col_template] = initial_df[col_template].iloc[0] if not initial_df.empty and col_template in initial_df else pd.NA
        ordered_cols = ['Occurrence.occurrenceID', 'Encounter.decimalLatitude', 'Encounter.decimalLongitude', 'Encounter.verbatimLocality', 'Encounter.year', 'Encounter.month', 'Encounter.day', 'Encounter.hour', 'Encounter.minutes']
        media_asset_cols_sorted = sorted([col for col in final_combined_df.columns if col.startswith('Encounter.mediaAsset')], key=lambda x: int(x.replace('Encounter.mediaAsset', '')))
        ordered_cols.extend(media_asset_cols_sorted)
        for col in initial_df.columns:
            self.task.checkpoint()
            if col not in ordered_cols and col in final_combined_df.columns:
                ordered_cols.append(col)
        for col in final_combined_df.columns:
            self.task.checkpoint()
            if col not in ordered_cols:
                ordered_cols.append(col)
        final_combined_df = final_combined_df.reindex(columns=ordered_cols, fill_value=pd.NA)
        return final_combined_df

    def generate_occurrence_id(self, row):
        """
            Genera un identificador de ocurrencia único combinando 'project_id' y 'deployment_id'.
            Maneja valores NaN convirtiéndolos a cadenas vacías.
    
            Args:
                row (pd.Series): Una fila del DataFrame de entrada.
    
            Returns:
                str: El ID de ocurrencia generado.
            """
        sanitized_project_id = str(row['project_id']) if pd.notna(row['project_id']) else ''
        sanitized_deployment_id = str(row['deployment_id']) if pd.notna(row['deployment_id']) else ''
        return f'{sanitized_project_id}-{sanitized_deployment_id}'

    def ensure_jpg_extension(self, location):
        """
            Asegura que la ubicación del archivo de imagen tenga la extensión '.JPG'.
            Si ya tiene una extensión, la cambia a '.JPG'. Si no tiene, la añade.
    
            Args:
                location (str): La ruta o nombre del archivo de imagen.
    
            Returns:
                str: La ubicación del archivo con la extensión '.JPG'.
            """
        if pd.isna(location):
            return location
        location_str = str(location)
        parts = location_str.split('.')
        if len(parts) > 1:
            base_name = '.'.join(parts[:-1])
            return base_name + '.JPG'
        return location_str + '.JPG'


@dataclass
class LynxProcessor:
    task: TaskContext
    source_folder: str
    has_linces: bool
    has_revision: bool
    minutes: int
    estaciones_file: str | None
    individuos_file: str | None
    get_exif_data = staticmethod(read_exif)
    get_date_taken = staticmethod(date_taken)

    def run(self):
        self.task.report('Reading lynx folders')
        valid_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.mp4', '.avi', '.mov', '.mkv', '.wmv', '.MOV')
        lince_exists = self.has_linces
        revision_exists = self.has_revision
        minutes_to_group = self.minutes
        data = []
        for root, dirs, files in os.walk(self.source_folder):
            self.task.checkpoint()
            for file in files:
                self.task.checkpoint()
                if file.lower().endswith(valid_extensions):
                    root_path = Path(root)
                    parts = root_path.parts
                    if lince_exists and revision_exists and (len(parts) >= 6):
                        finca = parts[-5]
                        estacion = parts[-4]
                        revision = parts[-3]
                        lince = parts[-1]
                    elif lince_exists and (not revision_exists) and (len(parts) >= 5):
                        finca = parts[-4]
                        estacion = parts[-3]
                        revision = 'N/A'
                        lince = parts[-1]
                    elif not lince_exists and revision_exists and (len(parts) >= 5):
                        finca = parts[-4]
                        estacion = parts[-3]
                        revision = parts[-2]
                        lince = parts[-1]
                    elif not lince_exists and (not revision_exists) and (len(parts) >= 4):
                        finca = parts[-3]
                        estacion = parts[-2]
                        revision = 'N/A'
                        lince = parts[-1]
                    else:
                        continue
                    file_name = root_path / file
                    exif_data = self.get_exif_data(str(file_name))
                    if exif_data:
                        capture_date = self.get_date_taken(exif_data)
                    else:
                        capture_date = None
                    lince = re.sub(' y | Y ', ' y ', lince)
                    lince_names = lince.split(' y ')
                    for individual_lince in lince_names:
                        self.task.checkpoint()
                        individual_lince_row = {'Finca': finca, 'Estación': estacion, 'Revisión': revision, 'Individuos': lince, 'Individuo': individual_lince.strip(), 'Archivo': str(file_name), 'Fecha Foto': capture_date}
                        data.append(individual_lince_row)
        if not data:
            raise ValueError('No files match the selected folder structure.')
        df = pd.DataFrame(data, columns=['Finca', 'Estación', 'Revisión', 'Individuos', 'Individuo', 'Archivo', 'Fecha Foto'])
        if minutes_to_group > 0:
            grouped_data = []
            for name, group in df.groupby(['Finca', 'Estación', 'Revisión']):
                self.task.checkpoint()
                group = group.sort_values(by='Fecha Foto')
                collapsed_files = []
                collapsed_individuos = []
                last_time = None
                for _, row in group.iterrows():
                    self.task.checkpoint()
                    if last_time and row['Fecha Foto'] and ((row['Fecha Foto'] - last_time).total_seconds() / 60 <= minutes_to_group):
                        collapsed_files[-1]['Archivo'] += ';' + row['Archivo']
                        collapsed_individuos[-1].update(row['Individuo'].split(' y '))
                    else:
                        collapsed_files.append(row.to_dict())
                        collapsed_individuos.append(set(row['Individuo'].split(' y ')))
                    last_time = row['Fecha Foto']
                for i, collapsed in enumerate(collapsed_files):
                    self.task.checkpoint()
                    for individual in collapsed_individuos[i]:
                        self.task.checkpoint()
                        new_row = collapsed.copy()
                        new_row['Individuo'] = individual.strip()
                        new_row['Individuos'] = ' y '.join(collapsed_individuos[i])
                        grouped_data.append(new_row)
            df = pd.DataFrame(grouped_data)
        if df['Revisión'].nunique() == 1 and df['Revisión'].iloc[0] == 'N/A':
            df = df.drop(columns=['Revisión'])
        if self.estaciones_file:
            estaciones_df = pd.read_excel(self.estaciones_file)
            df = df.merge(estaciones_df, how='left', left_on='Estación', right_on='Estacion')
        if self.individuos_file:
            individuos_df = pd.read_excel(self.individuos_file)
            df = df.merge(individuos_df, how='left', left_on='Individuo', right_on='Lince')
        df['Número de Fotos'] = df['Archivo'].apply(lambda x: len(x.split(';')))
        df.loc[df['Individuos'] == df['Individuo'], 'Individuos'] = ''
        self.task.checkpoint()
        return df


