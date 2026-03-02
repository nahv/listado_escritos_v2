# from tinydb import TinyDB
import os
import subprocess
import sys
import pandas as pd
import base64
import tempfile
from datetime import datetime, timedelta
from openpyxl import Workbook
from io import StringIO, BytesIO
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
        self.raw_records = []  # Store processed records for quick access

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

            return self._process_dataframe(df)
        else:
            return {"status": "no_file", "message": "No se ha seleccionado ningún archivo."}

    def read_file_from_memory(self, filename, base64_content):
        """
        Read an Excel file from base64 content (for drag and drop)
        """
        try:
            # Decode base64 content
            file_content = base64.b64decode(base64_content)
            
            # Create a BytesIO object
            file_bytes = BytesIO(file_content)
            
            # Read Excel from bytes
            try:
                df = pd.read_excel(file_bytes, skiprows=8, engine='openpyxl')
            except Exception:
                file_bytes.seek(0)  # Reset to beginning
                df = pd.read_excel(file_bytes, engine='openpyxl')
            
            return self._process_dataframe(df)
            
        except Exception as e:
            print(f"Error reading file from memory: {e}")
            return {"status": "error", "message": f"Error al leer el archivo: {e}"}

    def _process_dataframe(self, df):
        """
        Common DataFrame processing logic used by both read_data and read_file_from_memory
        """
        df = self._parse_recibido_column(df)
        self.data = df.copy()

        # Store raw records for quick access in detail views
        self._store_raw_records()

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

    def _store_raw_records(self):
        """
        Store raw records in a format ready for API responses
        """
        self.raw_records = []
        for _, row in self.data.iterrows():
            record = {
                "expte": str(row.get('Expte', 'N/A')),
                "titulo": str(row.get('Título', 'N/A')),
                "tipo": str(row.get('Tipo', 'N/A')),
                "presentante": str(row.get('Apellido', 'N/A')),
                "fecha": str(row.get('Recibido_date_str', 'N/A')) if 'Recibido_date_str' in row else 'N/A',
                "recibido": row.get('Recibido', None)
            }
            self.raw_records.append(record)

    # ============================================================================
    # NEW METHODS FOR DETAILED RECORD RETRIEVAL
    # ============================================================================

    def get_records_by_date(self, date_str):
        """
        Get all records for a specific date
        """
        if self.data is None:
            return {"status": "error", "message": "No hay datos cargados", "records": []}
        
        try:
            # Filter records for the given date
            filtered = self.data[self.data['Recibido_date_str'] == date_str]
            
            records = []
            for _, row in filtered.iterrows():
                # Determine record type for badge styling
                tipo = str(row.get('Tipo', 'N/A'))
                is_escrito = 'escrito' in tipo.lower() if tipo != 'N/A' else False
                
                record = {
                    "expte": str(row.get('Expte', 'N/A')),
                    "titulo": str(row.get('Título', 'N/A')),
                    "tipo": tipo,
                    "tipo_class": "escrito" if is_escrito else "proyecto",
                    "presentante": str(row.get('Apellido', 'N/A')),
                    "fecha": str(row.get('Recibido_date_str', 'N/A'))
                }
                records.append(record)
            
            return {
                "status": "ok",
                "records": records,
                "count": len(records)
            }
        except Exception as e:
            print(f"Error getting records by date: {e}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e), "records": []}

    def get_records_by_title(self, title):
        """
        Get all records with a specific title
        """
        if self.data is None:
            return {"status": "error", "message": "No hay datos cargados", "records": []}
        
        try:
            # Filter records with matching title (exact match for the title in most_titles)
            # This handles the exact titles from the most_titles list
            filtered = self.data[self.data['Título'] == title]
            
            records = []
            for _, row in filtered.iterrows():
                # Determine record type for badge styling
                tipo = str(row.get('Tipo', 'N/A'))
                is_escrito = 'escrito' in tipo.lower() if tipo != 'N/A' else False
                
                record = {
                    "expte": str(row.get('Expte', 'N/A')),
                    "titulo": str(row.get('Título', 'N/A')),
                    "tipo": tipo,
                    "tipo_class": "escrito" if is_escrito else "proyecto",
                    "presentante": str(row.get('Apellido', 'N/A')),
                    "fecha": str(row.get('Recibido_date_str', 'N/A'))
                }
                records.append(record)
            
            return {
                "status": "ok",
                "records": records,
                "count": len(records)
            }
        except Exception as e:
            print(f"Error getting records by title: {e}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e), "records": []}

    def search_records(self, query, field=None):
        """
        Search records across multiple fields or a specific field
        """
        if self.data is None:
            return {"status": "error", "message": "No hay datos cargados", "records": []}
        
        try:
            if field and field in self.data.columns:
                # Search in specific field
                mask = self.data[field].astype(str).str.contains(query, case=False, na=False)
            else:
                # Search across multiple fields
                mask = (
                    self.data['Expte'].astype(str).str.contains(query, case=False, na=False) |
                    self.data['Título'].astype(str).str.contains(query, case=False, na=False) |
                    self.data['Apellido'].astype(str).str.contains(query, case=False, na=False) |
                    self.data['Tipo'].astype(str).str.contains(query, case=False, na=False)
                )
            
            filtered = self.data[mask]
            
            records = []
            for _, row in filtered.iterrows():
                tipo = str(row.get('Tipo', 'N/A'))
                is_escrito = 'escrito' in tipo.lower() if tipo != 'N/A' else False
                
                record = {
                    "expte": str(row.get('Expte', 'N/A')),
                    "titulo": str(row.get('Título', 'N/A')),
                    "tipo": tipo,
                    "tipo_class": "escrito" if is_escrito else "proyecto",
                    "presentante": str(row.get('Apellido', 'N/A')),
                    "fecha": str(row.get('Recibido_date_str', 'N/A'))
                }
                records.append(record)
            
            return {
                "status": "ok",
                "records": records,
                "count": len(records)
            }
        except Exception as e:
            print(f"Error searching records: {e}")
            return {"status": "error", "message": str(e), "records": []}

    def get_all_records(self, limit=None, offset=0):
        """
        Get all records with pagination
        """
        if self.data is None:
            return {"status": "error", "message": "No hay datos cargados", "records": []}
        
        try:
            records = []
            for idx, row in self.data.iterrows():
                if offset > 0:
                    offset -= 1
                    continue
                    
                if limit and len(records) >= limit:
                    break
                    
                tipo = str(row.get('Tipo', 'N/A'))
                is_escrito = 'escrito' in tipo.lower() if tipo != 'N/A' else False
                
                record = {
                    "expte": str(row.get('Expte', 'N/A')),
                    "titulo": str(row.get('Título', 'N/A')),
                    "tipo": tipo,
                    "tipo_class": "escrito" if is_escrito else "proyecto",
                    "presentante": str(row.get('Apellido', 'N/A')),
                    "fecha": str(row.get('Recibido_date_str', 'N/A'))
                }
                records.append(record)
            
            return {
                "status": "ok",
                "records": records,
                "total": len(self.data),
                "count": len(records)
            }
        except Exception as e:
            print(f"Error getting all records: {e}")
            return {"status": "error", "message": str(e), "records": []}

    def get_records_summary(self):
        """
        Get a summary of all records (useful for debugging)
        """
        if self.data is None:
            return {"status": "error", "message": "No hay datos cargados"}
        
        try:
            return {
                "status": "ok",
                "total_records": len(self.data),
                "unique_dates": self.data['Recibido_date_str'].nunique() if 'Recibido_date_str' in self.data else 0,
                "unique_titles": self.data['Título'].nunique(),
                "unique_exptes": self.data['Expte'].nunique(),
                "date_range": {
                    "min": self.data['Recibido'].min().strftime('%d/%m/%Y') if not self.data['Recibido'].empty else None,
                    "max": self.data['Recibido'].max().strftime('%d/%m/%Y') if not self.data['Recibido'].empty else None
                }
            }
        except Exception as e:
            print(f"Error getting records summary: {e}")
            return {"status": "error", "message": str(e)}

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

    def export_pdf_continuous(self, start_date_str, n_proveyentes, per_list=15):
        """
        Export continuous lists starting from start_date_str, taking records after last_split_index,
        assigning them sequentially to proveyentes (first 15 to listado 1, next 15 to listado 2, etc.),
        preserving the original receipt dates from the records.
        
        The start_date_str is used for reference in the title but does NOT modify the record dates.
        
        Args:
            start_date_str: Reference date for the title (doesn't modify record dates)
            n_proveyentes: Number of proveyentes (lists to create)
            per_list: Number of records per list BEFORE merging (default 15)
        """
        print(f"export_pdf_continuous called with: start_date={start_date_str}, n_proveyentes={n_proveyentes}")
        
        if self.data is None:
            print("ERROR: No data loaded")
            return {"status": "error", "message": "No hay datos cargados para exportar."}
        
        try:
            # Parse start date for title reference only
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
                print(f"Reference start date (Y-m-d): {start_date}")
            except Exception as e:
                print(f"Failed to parse as Y-m-d: {e}")
                try:
                    start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
                    print(f"Reference start date (d/m/Y): {start_date}")
                except Exception as e:
                    print(f"Failed to parse as d/m/Y: {e}")
                    try:
                        start_date = datetime.strptime(start_date_str, "%d/%m/%y")
                        print(f"Reference start date (d/m/y): {start_date}")
                    except Exception as e:
                        print(f"Failed to parse as d/m/y: {e}")
                        return {"status": "error", "message": f"Formato de fecha inválido: {start_date_str}"}
            
            # IMPORTANT: Use the same listados that export_pdf uses, but continue from last_split_index
            # This ensures we don't reprocess records that were already assigned
            listados = self.create_listados(self.data)
            
            # Get the starting index from where we left off
            start_idx = getattr(self, 'last_split_index', 0)
            print(f"Start index: {start_idx}, total listados: {len(listados)}")
            
            # Get remaining records
            remaining = listados[start_idx:]
            print(f"Remaining records: {len(remaining)}")
            
            if not remaining:
                return {"status": "error", "message": "No hay más registros para exportar."}
            
            # Calculate how many records we need (n_proveyentes * 15)
            total_needed = n_proveyentes * per_list
            print(f"Total needed: {total_needed}")
            
            selected_records = remaining[:total_needed]
            print(f"Selected records: {len(selected_records)}")
            
            # Distribute records sequentially to each listado
            # Group 1: records 0-14, Group 2: records 15-29, etc.
            sequential_groups = []
            for i in range(0, len(selected_records), per_list):
                group = selected_records[i:i + per_list]
                if group:  # Only add non-empty groups
                    sequential_groups.append(group)
            
            print(f"Created {len(sequential_groups)} sequential groups before merging")
            
            # Merge expedientes WITHIN each group
            processed_groups = []
            for group_idx, group in enumerate(sequential_groups):
                print(f"Processing group {group_idx+1} with {len(group)} records before merging")
                merged_group = self.merge_expedientes(group)
                print(f"  After merging: {len(merged_group)} records")
                processed_groups.append(merged_group)
            
            # Save PDF dialog
            file_types = ["Archivos PDF (*.pdf)"]
            default_filename = f"proveyentes_continuo_{datetime.now().strftime('%d-%m-%Y_%H-%Mhs')}.pdf"
            print(f"Opening save dialog with default: {default_filename}")
            
            save_path = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG, allow_multiple=False, file_types=file_types, save_filename=default_filename
            )
            if isinstance(save_path, (tuple, list)):
                save_path = save_path[0] if save_path else None
            
            print(f"Save path: {save_path}")
            
            if not save_path:
                return {"status": "cancelled", "message": "Exportación cancelada por el usuario."}
            
            # Build PDF
            doc = SimpleDocTemplate(save_path, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            
            for i, group in enumerate(processed_groups):
                # Title with tag - indicating these are continuous
                elements.append(Paragraph(f"Listado {i+1} <small>(Fechas continuas)</small>", 
                                         styles['Heading2']))
                
                if not group:
                    elements.append(Paragraph("Sin registros asignados.", styles['Normal']))
                else:
                    data = [["Título", "Expediente", "Fecha", "Presentante", "Días corridos"]]
                    processed_rows = []
                    
                    for row in group:
                        titulo = row[0]
                        if len(titulo) > 42:
                            titulo = titulo[:42] + "..."
                        presentante = row[3]
                        if presentante and len(presentante) > 15:
                            presentante = presentante[:15] + "..."
                        processed_rows.append([titulo, row[1], row[2], presentante, row[4]])
                    
                    data += processed_rows
                    
                    # Same table formatting as repartidas version
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
                    
                    # Striped background
                    for row_idx in range(1, len(data)):
                        if row_idx % 2 == 1:
                            table_style.append(('BACKGROUND', (0,row_idx), (-1,row_idx), colors.whitesmoke))
                    
                    table.setStyle(TableStyle(table_style))
                    elements.append(table)
                
                elements.append(Spacer(1, 18))
                if (i + 1) % 2 == 0 and (i + 1) < len(processed_groups):
                    elements.append(PageBreak())
            
            print(f"Building PDF with {len(elements)} elements")
            doc.build(elements)
            
            # Update last_split_index to reflect consumed records
            self.last_split_index = start_idx + len(selected_records)
            # Get the last date from the selected records for tracking
            try:
                dates = []
                for r in selected_records:
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
            
            print(f"Export successful. Last index: {self.last_split_index}")
            
            return {
                "status": "ok", 
                "path": save_path,
                "last_assigned_date": self.last_assigned_date.strftime('%d/%m/%Y'),
                "last_index": self.last_split_index
            }
            
        except Exception as e:
            print(f"ERROR in export_pdf_continuous: {str(e)}")
            import traceback
            traceback.print_exc()
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