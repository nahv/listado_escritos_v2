from tinydb import TinyDB
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from io import StringIO
import webview

class Api:
    def __init__(self):
        self.db = TinyDB('db.json')
        self.data = None

    def read_data(self, file_path):
        # If no file_path, open file dialog
        if not file_path:
            file_types = ["Archivos Excel (*.xlsx)"]  # or ["*.xlsx"]
            file_path = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types
            )
            # Ensure file_path is a string, not tuple/list
            if isinstance(file_path, (tuple, list)):
                file_path = file_path[0] if file_path else None

        if file_path:
            input_data = pd.read_excel(file_path)
            csv_data = StringIO()
            input_data.to_csv(csv_data, index=False)
            csv_data.seek(0)
            self.data = pd.read_csv(csv_data, skiprows=8)

            self.data['Recibido'] = pd.to_datetime(self.data['Recibido'])
            total_records = len(self.data)
            presentaciones_count = len(self.data[self.data['Tipo'].str.contains('escrito', case=False)])
            proyectos_count = len(self.data[self.data['Tipo'].str.contains('proyecto', case=False)])
            oldest_record = min(self.data['Recibido'])
            oldest_record_formatted = oldest_record.strftime("%d/%m/%Y")
            newest_record = max(self.data['Recibido'])
            unique_exptes_count = self.data['Expte'].nunique()
            transferencias_count = len(self.data[self.data['Título'].str.contains('transferencia', case=False)])
            today_date = datetime.now()
            days_difference = (today_date - oldest_record).days
            today_formatted = today_date.strftime("%d/%m/%Y")
            most_titles = self.data['Título'].value_counts().head(10).to_dict()
            most_titles_df = pd.DataFrame(list(most_titles.items()), columns=['Suma', 'Cantidad'])

            info_text = (f'\n'
                         f'     {total_records} presentaciones en un total de {unique_exptes_count} causas.\n'
                         f'\n'
                         f'     Escritos: {presentaciones_count} / Proyectos: {proyectos_count}\n'
                         f'\n'
                         f'     Escrito más antiguo {oldest_record_formatted} - {days_difference} días corridos al {today_formatted}\n'
                         f'\n'
                         f'     {transferencias_count} escritos incluyen la palabra "transferencia".\n'
                         f'\n'
                         f'     Escritos más repetidos:\n{most_titles_df.to_string(index=False)}')
            return info_text
        else:
            return "    No se ha seleccionado ningún archivo."

    def create_listados(self, data):
        today_date = datetime.now()
        data['Recibido'] = pd.to_datetime(data['Recibido'])
        data['DaysDifference'] = (today_date - data['Recibido']).dt.days
        data['Recibido'] = pd.to_datetime(data['Recibido']).dt.strftime('%d/%m/%y')
        listados = []
        for index, row in data.iterrows():
            fifth_column_string = f"{row['DaysDifference']} días al {today_date.strftime('%d/%m')}"
            listados.append((row['Título'], row['Expte'], row['Recibido'], row['Apellido'], fifth_column_string))
        return listados

    def save_listados_to_excel(self, listados, filename):
        wb = Workbook()
        ws = wb.active
        ws.title = "Listados"
        for listado in listados:
            ws.append(listado)
        wb.save(filename)

    def export_excel(self):
        if self.data is not None:
            # Prompt user for save location
            file_types = ["Archivos Excel (*.xlsx)"]
            default_filename = f"listado_{datetime.now().strftime('%d-%m-%Y_%H-%Mhs')}.xlsx"
            save_path = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG, allow_multiple=False, file_types=file_types, save_filename=default_filename
            )
            if isinstance(save_path, (tuple, list)):
                save_path = save_path[0] if save_path else None
            if save_path:
                listados = self.create_listados(self.data)
                self.save_listados_to_excel(listados, save_path)
                return f"\n\nArchivo guardado como '{save_path}'."
            else:
                return "Exportación cancelada por el usuario."
        else:
            return "No hay datos cargados para exportar."
