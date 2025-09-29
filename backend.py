from tinydb import TinyDB
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from io import StringIO
import webview
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

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

    def merge_expedientes(self, records):
        """
        Merge records with the same Expediente, counting occurrences and joining Título.
        Returns a list of merged records.
        """
        merged = {}
        for row in records:
            expte = row[1]
            if expte not in merged:
                merged[expte] = {
                    "Título": row[0],
                    "Expte": expte,
                    "Recibido": row[2],
                    "Presentante": row[3],
                    "Días": row[4],
                    "count": 1
                }
            else:
                merged[expte]["count"] += 1
                merged[expte]["Título"] += f" | {row[0]}"
        result = []
        for v in merged.values():
            titulo = v["Título"]
            if len(titulo) > 42:
                titulo = titulo[:42] + "..."
            if v["count"] > 1:
                otros = v["count"] - 1
                if otros == 1:
                    titulo += " (+1 otro escrito)"
                else:
                    titulo += f" (+{otros} otros escritos)"
            result.append([titulo, v["Expte"], v["Recibido"], v["Presentante"], v["Días"]])
        return result

    def assign_proveyentes(self, records, n_proveyentes):
        """
        Assign merged records to proveyentes in round-robin, keeping all presentaciones for the same Expediente together.
        """
        merged_records = self.merge_expedientes(records)
        groups = [[] for _ in range(n_proveyentes)]
        for idx, record in enumerate(merged_records):
            groups[idx % n_proveyentes].append(record)
        return groups

    def export_pdf(self, n_proveyentes):
        if self.data is None:
            return "No hay datos cargados para exportar."
        try:
            # Prompt user for save location
            file_types = ["Archivos PDF (*.pdf)"]
            default_filename = f"proveyentes_{datetime.now().strftime('%d-%m-%Y_%H-%Mhs')}.pdf"
            save_path = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG, allow_multiple=False, file_types=file_types, save_filename=default_filename
            )
            if isinstance(save_path, (tuple, list)):
                save_path = save_path[0] if save_path else None
            if not save_path:
                return "Exportación cancelada por el usuario."

            # Prepare records for assignment (use same columns as Excel export)
            listados = self.create_listados(self.data)
            # Only take up to n_proveyentes * 15 records (before merging)
            total_needed = n_proveyentes * 15
            selected_records = listados[:total_needed]
            groups = self.assign_proveyentes(selected_records, n_proveyentes)

            # Generate PDF
            doc = SimpleDocTemplate(save_path, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            for i, group in enumerate(groups):
                elements.append(Paragraph(f"Listado {i+1}", styles['Heading2']))
                if not group:
                    elements.append(Paragraph("Sin registros asignados.", styles['Normal']))
                else:
                    data = [["Título", "Expediente", "Fecha", "Presentante", "Días corridos"]]
                    processed_group = []
                    for row in group:
                        titulo = row[0]
                        if len(titulo) > 42:
                            titulo = titulo[:42] + "..."
                        presentante = row[3]
                        processed_group.append([titulo, row[1], row[2], presentante, row[4]])
                    data += processed_group
                    page_width = A4[0]
                    col_widths = [
                        page_width * 0.32,  # Título
                        page_width * 0.17,  # Expte
                        page_width * 0.10,  # Recibido (thinner)
                        page_width * 0.15,  # Presentante
                        page_width * 0.15   # Días corridos
                    ]
                    table = Table(data, repeatRows=1, colWidths=col_widths)
                    # Striped style for rows
                    table_style = [
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                        ('ALIGN', (0,0), (0,-1), 'LEFT'),  # Título column left
                        ('ALIGN', (1,0), (-1,-1), 'CENTER'),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,1), (0,-1), 7),   # Título records smaller font
                        ('FONTSIZE', (1,1), (2,-1), 10),  # Expte and Recibido records bigger font
                        ('FONTSIZE', (3,1), (3,-1), 7),   # Presentante records smaller font
                        ('FONTSIZE', (4,1), (4,-1), 10),  # Días corridos records bigger font
                        ('FONTSIZE', (0,0), (-1,0), 10),  # Header bigger font
                        ('BOTTOMPADDING', (0,0), (-1,0), 8),
                        ('VALIGN', (0,1), (-1,-1), 'TOP'),
                        ('WORDWRAP', (0,1), (0,-1), 'CJK'),  # Only Título column wrapped
                    ]
                    # Add striped background for rows
                    for row_idx in range(1, len(data)):
                        if row_idx % 2 == 1:
                            table_style.append(('BACKGROUND', (0,row_idx), (-1,row_idx), colors.whitesmoke))
                    table.setStyle(TableStyle(table_style))
                    elements.append(table)
                elements.append(Spacer(1, 18))
                if (i + 1) % 2 == 0 and (i + 1) < len(groups):
                    elements.append(PageBreak())
            doc.build(elements)
            return f"PDF guardado como '{save_path}'."
        except Exception as e:
            return f"Error al exportar PDF: {e}"

    def get_window_size(self):
        win = webview.windows[0]
        return {'width': win.width, 'height': win.height}

    def set_window_size(self, width, height):
        win = webview.windows[0]
        win.resize(width, height)
        return True
        win = webview.windows[0]
        screen = win.screen
        if hasattr(screen, 'width') and hasattr(screen, 'height'):
            x = max(0, int((screen.width - width) / 2))
            y = max(0, int((screen.height - height) / 2))
            win.move(x, y)
        return True
        return True
