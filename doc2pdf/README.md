# Docx to Styled PDF Project

Este proyecto proporciona herramientas avanzadas para la inspección de archivos DOCX y su conversión a documentos PDF con estilos personalizados y profesionales.

## Características Principales

### 1. Inspección de Documentos (`inspection/`)
El módulo de inspección permite analizar la estructura interna de cualquier archivo `.docx`.
- **Generación de Reportes**: Crea un informe interactivo en HTML (`report.html`).
- **Extracción de Medios**: Extrae automáticamente todas las imágenes y archivos incrustados a la carpeta `extracted_media/`.
- **Análisis de Metadatos**: Extrae autor, fechas de modificación, revisiones y más.
- **Detalle de Estilos**: Desglosa cada párrafo, su estilo aplicado y el formato individual de sus "runs" (negrita, cursiva, tamaño).

### 2. Generación de PDF Estilizado (`pdfCreation/`)
El motor de conversión transforma archivos `.docx` en PDFs con un diseño moderno y coherentes con la identidad visual corporativa.
- **Diseño Premium**: Paleta de colores basada en **Magenta (#DA1984)**, **Azul Marino (#232D4B)** y **Amarillo (#FCEE21)**.
- **Portada Personalizada**: Fondo de imagen estirado (`portada.jpg`) con títulos centrados en blanco.
- **Detección de Listas**: Identificación robusta de viñetas (bullet points) mediante inspección de metadatos XML (`numPr`).
- **Tablas Sincronizadas**: A diferencia de otros conversores, este motor mantiene el orden exacto de las tablas tal como aparecen entre los párrafos del documento original.
- **Estilos de Título**:
    - **Header 1**: Amarillo Navy.
    - **Header 2**: Magenta Bold.
    - **Header 3**: Azul Marino.

## Estructura del Proyecto

```text
doc2pdf/
├── inspection/
│   ├── docx_inspector.py     # Script de inspección y reporte HTML
│   ├── report.html           # Resultado de la última inspección
│   └── extracted_media/      # Imágenes extraídas del docx
├── pdfCreation/
│   ├── pdf_creator.py        # Generador principal de PDF
│   ├── generated_styled.pdf  # Resultado final en PDF
│   └── portada.jpg           # Imagen de fondo para la portada
└── README.md                 # Documentación del proyecto
```

## Requisitos

- Python 3.10+
- Librerías:
  - `python-docx`: Para la manipulación de archivos Word.
  - `reportlab`: Para la generación de PDFs de alta calidad.

Instalación:
```bash
pip install python-docx reportlab
```

## Uso

### Generar Informe de Inspección
Ejecuta el script desde la carpeta raíz:
```bash
python inspection/docx_inspector.py
```
*Nota: El script está configurado para procesar `1. Artificial Intelligence - Copia.docx` por defecto.*

### Crear PDF Estilizado

Puedes generar el PDF de dos formas:

1. **Modo Interactivo (Recomendado)**:
   Si ejecutas el script sin argumentos, se abrirá una ventana para que selecciones el archivo `.docx` desde tu ordenador:
   ```bash
   python doc2pdf/pdfCreation/pdf_creator.py
   ```

2. **Modo Línea de Comandos**:
   Puedes especificar el archivo directamente:
   ```bash
   python doc2pdf/pdfCreation/pdf_creator.py "tu_archivo.docx"
   ```
   *También puedes usar `-o` para definir el nombre del PDF de salida.*


---
*Desarrollado para la gestión inteligente de documentación técnica.*
