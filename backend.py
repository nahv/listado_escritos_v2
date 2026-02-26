# from tinydb import TinyDB
import os
import subprocess
import sys
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from io import StringIO
import webview
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

class Api:
    def __init__(self):
        self.data = None
        self.last_split_index = 0
        self.last_assigned_date = None

    def _parse_recibido_column(self, df):
        # Try robust parsing with dayfirst and fallback formats
        if 'Recibido' not in df.columns:
            return df
        parsed = pd.to_datetime(df['Recibido'], dayfirst=True, errors='coerce')
        # Try some common explicit formats for any NaT
        if parsed.isna().any():
            def try_formats(val):
                if pd.isna(val):
                    return pd.NaT
                for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%m/%d/%Y"):
                    try:
                        return datetime.strptime(str(val), fmt)
                    except Exception:
                        continue
                try:
                    return pd.to_datetime(val, dayfirst=True, errors='coerce')
                except:
                    return pd.NaT
            parsed = df['Recibido'].apply(try_formats)
        df['Recibido'] = pd.to_datetime(parsed)
        return df

    def read_data(self, file_path):
        # If no file_path, open file dialog
        if not file_path:
            file_types = ["Archivos Excel (*.xlsx)"]
            file_path = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types
            )
            if isinstance(file_path, (tuple, list)):
                file_path = file_path[0] if file_path else None

        if file_path:
            # Read Excel directly and skip first 8 rows if needed. Use engine to be explicit.
            try:
                df = pd.read_excel(file_path, skiprows=8, engine='openpyxl')
            except Exception:
                # fallback to simple read
                df = pd.read_excel(file_path, engine='openpyxl')

            df = self._parse_recibido_column(df)
            self.data = df.copy()

            # Basic aggregations
            total_records = len(self.data)
            presentaciones_count = len(self.data[self.data['Tipo'].str.contains('escrito', case=False, na=False)])
            proyectos_count = len(self.data[self.data['Tipo'].str.contains('proyecto', case=False, na=False)])
            oldest_record = self.data['Recibido'].min()
            oldest_record_formatted = oldest_record.strftime("%d/%m/%Y") if not pd.isna(oldest_record) else ""
            newest_record = self.data['Recibido'].max()
            unique_exptes_count = self.data['Expte'].nunique()
            transferencias_count = len(self.data[self.data['Título'].str.contains('transferencia', case=False, na=False)])
            today_date = datetime.now()
            days_difference = (today_date - oldest_record).days if not pd.isna(oldest_record) else 0
            today_formatted = today_date.strftime("%d/%m/%Y")
            # Top titles
            most_titles_series = self.data['Título'].value_counts().head(10)
            most_titles = [{"Título": k, "Cantidad": int(v)} for k, v in most_titles_series.items()]

            # Presentaciones por Fecha (group by Recibido date)
            self.data['Recibido_date_str'] = self.data['Recibido'].dt.strftime('%d/%m/%Y')
            grouped = self.data.groupby('Recibido_date_str')
            presentaciones_by_date = []
            for date_str, grp in grouped:
                escritos = int(grp['Tipo'].str.contains('escrito', case=False, na=False).sum())
                proyectos = int(grp['Tipo'].str.contains('proyecto', case=False, na=False).sum())
                total = len(grp)
                presentaciones_by_date.append({"Fecha": date_str, "Escritos": escritos, "Proyectos": proyectos, "Total": int(total)})

            # Compose period
            period = f"{oldest_record_formatted} a {today_formatted}" if oldest_record_formatted else ""

            # Return structured dict (pywebview will send JSON-able object)
            return {
                "status": "ok",
                "total_records": int(total_records),
                "unique_exptes_count": int(unique_exptes_count),
                "presentaciones_count": int(presentaciones_count),
                "proyectos_count": int(proyectos_count),
                "oldest_record": oldest_record_formatted,
                "days_difference": int(days_difference),
                "today_date": today_formatted,
                "transferencias_count": int(transferencias_count),
                "most_titles": most_titles,
                "presentaciones_by_date": presentaciones_by_date,
                "period": period
            }
        else:
            return {"status": "no_file", "message": "No se ha seleccionado ningún archivo."}

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
                return {"status": "ok", "path": save_path}
            else:
                return {"status": "cancelled", "message": "Exportación cancelada por el usuario."}
        else:
            return {"status": "error", "message": "No hay datos cargados para exportar."}

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
            return {"status": "error", "message": "No hay datos cargados para exportar."}
        try:
            file_types = ["Archivos PDF (*.pdf)"]
            default_filename = f"proveyentes_{datetime.now().strftime('%d-%m-%Y_%H-%Mhs')}.pdf"
            save_path = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG, allow_multiple=False, file_types=file_types, save_filename=default_filename
            )
            if isinstance(save_path, (tuple, list)):
                save_path = save_path[0] if save_path else None
            if not save_path:
                return {"status": "cancelled", "message": "Exportación cancelada por el usuario."}

            listados = self.create_listados(self.data)
            total_needed = n_proveyentes * 15
            selected_records = listados[:total_needed]
            # Track last split index and last assigned date
            self.last_split_index = min(total_needed, len(listados))
            # Determine last assigned date among selected records (parse dd/mm/yy)
            try:
                dates = []
                for r in selected_records:
                    # r[2] is Recibido like 'dd/mm/yy' from create_listados
                    try:
                        dates.append(datetime.strptime(r[2], "%d/%m/%y"))
                    except:
                        try:
                            dates.append(datetime.strptime(r[2], "%d/%m/%Y"))
                        except:
                            pass
                if dates:
                    self.last_assigned_date = max(dates)
                else:
                    self.last_assigned_date = datetime.now()
            except Exception:
                self.last_assigned_date = datetime.now()

            groups = self.assign_proveyentes(selected_records, n_proveyentes)

            # Build PDF (Fechas repartidas)
            doc = SimpleDocTemplate(save_path, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            for i, group in enumerate(groups):
                # Title with tag
                elements.append(Paragraph(f"Listado {i+1} <small>(Fechas repartidas)</small>", styles['Heading2']))
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
                        if presentante and len(presentante) > 15:
                            presentante = presentante[:15] + "..."
                        processed_group.append([titulo, row[1], row[2], presentante, row[4]])
                    data += processed_group
                    page_width = A4[0]
                    col_widths = [
                        page_width * 0.32,
                        page_width * 0.17,
                        page_width * 0.10,
                        page_width * 0.15,
                        page_width * 0.15
                    ]
                    table = Table(data, repeatRows=1, colWidths=col_widths)
                    # Striped style for rows
                    table_style = [
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                        ('ALIGN', (0,0), (0,-1), 'LEFT'),
                        ('ALIGN', (1,0), (-1,-1), 'CENTER'),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,1), (0,-1), 7),
                        ('FONTSIZE', (1,1), (2,-1), 10),
                        ('FONTSIZE', (3,1), (3,-1), 7),
                        ('FONTSIZE', (4,1), (4,-1), 10),
                        ('FONTSIZE', (0,0), (-1,0), 10),
                        ('BOTTOMPADDING', (0,0), (-1,0), 8),
                        ('VALIGN', (0,1), (-1,-1), 'TOP'),
                        ('WORDWRAP', (0,1), (0,-1), 'CJK'),
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
            return {"status": "ok", "path": save_path, "last_assigned_date": self.last_assigned_date.strftime('%d/%m/%Y'), "last_index": self.last_split_index}
        except Exception as e:
            return {"status": "error", "message": f"Error al exportar PDF: {e}"}

    def export_pdf_continuous(self, start_date_str, n_listados, per_list=15):
        """
        Export continuous lists starting from start_date_str, taking records after last_split_index,
        grouping per_list per listado, and assigning correlative dates starting at start_date.
        """
        if self.data is None:
            return {"status": "error", "message": "No hay datos cargados para exportar."}
        try:
            # parse start date robustly (dayfirst)
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            except Exception:
                try:
                    start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
                except Exception:
                    start_date = datetime.strptime(start_date_str, "%d/%m/%y")

            # Prepare records from the remaining listados
            listados = self.create_listados(self.data)
            start_idx = getattr(self, 'last_split_index', 0)
            remaining = listados[start_idx:]
            # chunk into n_listados * per_list (or until exhausted)
            total_needed = n_listados * per_list
            selected = remaining[:total_needed]

            # create groups sequentially of size per_list
            groups = [selected[i:i+per_list] for i in range(0, len(selected), per_list)]
            # assign correlative dates across all selected records
            assigned_date = start_date
            processed_groups = []
            today_date = datetime.now()
            for grp in groups:
                processed = []
                for row in grp:
                    # compute new 'Días' relative to today as requested
                    days_diff = (today_date - assigned_date).days
                    rec_str = assigned_date.strftime('%d/%m/%y')
                    fifth_col = f"{days_diff} días al {today_date.strftime('%d/%m')}"
                    processed.append((row[0], row[1], rec_str, row[3] if row[3] else '', fifth_col))
                    # increment date for next record
                    assigned_date = assigned_date + timedelta(days=1)
                processed_groups.append(processed)

            # Save PDF
            file_types = ["Archivos PDF (*.pdf)"]
            default_filename = f"proveyentes_continuo_{datetime.now().strftime('%d-%m-%Y_%H-%Mhs')}.pdf"
            save_path = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG, allow_multiple=False, file_types=file_types, save_filename=default_filename
            )
            if isinstance(save_path, (tuple, list)):
                save_path = save_path[0] if save_path else None
            if not save_path:
                return {"status": "cancelled", "message": "Exportación cancelada por el usuario."}

            # Build PDF (Fechas continuas) using merge_expedientes per group
            doc = SimpleDocTemplate(save_path, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            for i, group in enumerate(processed_groups):
                elements.append(Paragraph(f"Listado {i+1} <small>(Fechas continuas)</small>", styles['Heading2']))
                if not group:
                    elements.append(Paragraph("Sin registros asignados.", styles['Normal']))
                else:
                    # merge expedientes within this small group to keep presentaciones together
                    merged = self.merge_expedientes(group)
                    data = [["Título", "Expediente", "Fecha", "Presentante", "Días corridos"]]
                    processed_group = []
                    for row in merged:
                        titulo = row[0]
                        if len(titulo) > 42:
                            titulo = titulo[:42] + "..."
                        presentante = row[3]
                        if presentante and len(presentante) > 15:
                            presentante = presentante[:15] + "..."
                        processed_group.append([titulo, row[1], row[2], presentante, row[4]])
                    data += processed_group
                    page_width = A4[0]
                    col_widths = [
                        page_width * 0.32,
                        page_width * 0.17,
                        page_width * 0.10,
                        page_width * 0.15,
                        page_width * 0.15
                    ]
                    table = Table(data, repeatRows=1, colWidths=col_widths)
                    table_style = [
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                        ('ALIGN', (0,0), (0,-1), 'LEFT'),
                        ('ALIGN', (1,0), (-1,-1), 'CENTER'),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,1), (0,-1), 7),
                        ('FONTSIZE', (1,1), (2,-1), 10),
                        ('FONTSIZE', (3,1), (3,-1), 7),
                        ('FONTSIZE', (4,1), (4,-1), 10),
                        ('FONTSIZE', (0,0), (-1,0), 10),
                        ('BOTTOMPADDING', (0,0), (-1,0), 8),
                        ('VALIGN', (0,1), (-1,-1), 'TOP'),
                        ('WORDWRAP', (0,1), (0,-1), 'CJK'),
                    ]
                    for row_idx in range(1, len(data)):
                        if row_idx % 2 == 1:
                            table_style.append(('BACKGROUND', (0,row_idx), (-1,row_idx), colors.whitesmoke))
                    table.setStyle(TableStyle(table_style))
                    elements.append(table)
                elements.append(Spacer(1, 18))
                if (i + 1) % 2 == 0 and (i + 1) < len(processed_groups):
                    elements.append(PageBreak())
            doc.build(elements)
            # Update last_split_index to reflect consumed records
            self.last_split_index = start_idx + len(selected)
            return {"status": "ok", "path": save_path}
        except Exception as e:
            return {"status": "error", "message": f"Error al exportar PDF continuo: {e}"}

    def open_file(self, file_path):
        try:
            if os.name == 'nt':
                os.startfile(file_path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', file_path])
            else:
                subprocess.Popen(['xdg-open', file_path])
            return {"status": "ok"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

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
