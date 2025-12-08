# 📑 Listados Judiciales App

Aplicación de escritorio desarrollada en **Python** y **PyWebview** que procesa archivos crudos del Excel de escritos recibidos que se extrae de Forum Institucional del Poder Judicial de Corrientes, genera un estadístico y permite exportar listados organizados en **Excel** y **PDF**.  

Fue diseñada para facilitar la repartición del trabajo de los proveyentes en el Juzgado Civil y Comercial 1 del Poder Judicial de la provincia de Corrientes, Argentina.

---

## 🚀 Funcionalidades

- **Importación de archivo de Excel Crudo de Forum**  
  Selección de archivo `.xlsx` mediante cuadro de diálogo nativo.  

- **Procesamiento de datos**  
  - Limpieza y normalización de registros.  
  - Conversión automática de fechas.  
  - Cálculo de días transcurridos desde el escrito más antiguo a la fecha en la que se provee. 

- **Resumen estadístico**  
  Obtención de indicadores clave:  
  - Total de presentaciones y expedientes únicos.  
  - Cantidad de escritos y proyectos pendientes.  
  - Fecha del escrito más antiguo.  
  - Títulos de escritos más repetidos para tener panorama de las solicitudes.

- **Exportación a Excel**  
  Generación de un listado limpio y estructurado en orden cronológico.  

- **Exportación a PDF con repartición de fechas**  
  - Distribuye las fechas indicadas en grupos de 15 escritos para proveer de a varios proveyentes en forma pareja. 
  - Fusiona los escritos que pertenecen al mismo expediente, para evitar asginar escritos de la misma causa a diferentes proveyentes. 

---
<img width="1512" height="1112" alt="Screenshot 2025-09-30 at 2 10 19 PM" src="https://github.com/user-attachments/assets/228e82f9-e9f6-481d-aed4-8e5a9d9c47d7" />



🔗 [Descargar ejecutable de prueba](https://drive.google.com/file/d/1HovopvNSddan1wiDczQ91FN6kogDDWsN/view?usp=sharing)

## 🛠️ Tecnologías utilizadas

- [Python 3.10+](https://www.python.org/)  
- [PyWebview](https://pywebview.flowrl.com/) – interfaz gráfica de escritorio multiplataforma  
- [Pandas](https://pandas.pydata.org/) – manipulación de datos  
- [OpenPyXL](https://openpyxl.readthedocs.io/) – exportación a Excel  
- [ReportLab](https://www.reportlab.com/dev/docs/) – generación de PDFs  
- (Opcional, futuro) [TinyDB](https://tinydb.readthedocs.io/) – base de datos ligera en JSON  

---

## 📂 Estructura del proyecto

```
proyecto/
│── main.py          # Punto de entrada de la aplicación
│── backend.py       # Clase Api con toda la lógica del backend
│── requirements.txt # Dependencias
│── README.md        # Documentación
└── frontend/        # el front
```

---

## 📦 Instalación

1. **Clonar el repositorio**  
   ```bash
   git clone 
   cd listado_escritos
   ```

2. **Crear y activar entorno virtual**  
   ```bash
   python -m venv venv
   source venv/bin/activate   # Linux/Mac
   venv\Scripts\activate      # Windows
   ```

3. **Instalar dependencias**  
   ```bash
   pip install -r requirements.txt
   ```

---

## ▶️ Uso

1. Ejecutar  
   ```bash
   python main.py
   ```

2. Seleccionar un archivo `.xlsx` mediante el cuadro de diálogo.  

3. Opciones disponibles:
   - **Exportar a Excel** – guarda un listado estructurado.  
   - **Exportar a PDF (Proveyentes)** – divide en N grupos de 15 y genera reportes en PDF.  

---

## 📊 Flujo de trabajo típico

- **Entrada:** Archivo Excel con registros judiciales.  
- **Proceso:** Detección de expedientes, escritos, proyectos y profesionales.
- **Salida:**  
  - Excel: `listado_dd-mm-YYYY_HH-MMhs.xlsx`  
  - PDF: `proveyentes_dd-mm-YYYY_HH-MMhs.pdf`  

---

## 🔮 Próximas mejoras

- [ ] Almacenamiento persistente con TinyDB.

---

## 📜 Licencia

Licencia MIT. Se puede clonar y adaptar libremente.
