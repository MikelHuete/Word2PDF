import argparse
import os
import subprocess
import tempfile
import shutil
import tkinter as tk
from tkinter import filedialog
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as PlatypusImage
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import json
import re
import time
import urllib.request

from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- COLORES CORPORATIVOS ---
MAGENTA = "#DA1984"
NAVY_TEXT = "#232D4B"
WHITE = "#FFFFFF"
YELLOW = "#FCEE21"

class NotebookLMGenerator:
    """Generador que utiliza el CLI de NotebookLM de forma robusta."""
    def __init__(self):
        self.custom_env = os.environ.copy()
        self.custom_env["PYTHONIOENCODING"] = "utf-8"
        self.custom_env["PYTHONUTF8"] = "1"
        try:
            subprocess.run(["nlm", "--version"], capture_output=True, env=self.custom_env, check=True)
            self.available = True
        except:
            self.available = False
            print("AVISO: No se encontró el CLI 'nlm'. Asegúrate de que esté en el PATH.")

    def run_safe_cmd(self, cmd_list):
        """Ejecuta un comando con protección de codificación en Windows."""
        result = subprocess.run(cmd_list, capture_output=True, env=self.custom_env)
        stdout = result.stdout.decode('utf-8', errors='ignore') if result.stdout else ""
        stderr = result.stderr.decode('utf-8', errors='ignore') if result.stderr else ""
        if result.returncode != 0:
            raise subprocess.CalledProcessError(result.returncode, cmd_list, output=stdout, stderr=stderr)
        return stdout

    def extract_id(self, text):
        """Extrae un ID de la respuesta del CLI."""
        match = re.search(r'([a-zA-Z0-9_-]{20,})', text)
        return match.group(1) if match else None

    def resolve_style_template_path(self, docx_path):
        """Localiza infografias_plantilla.jpg junto al .docx o en la raíz del proyecto."""
        candidates = [
            os.path.join(os.path.dirname(os.path.abspath(docx_path)), "Plantilla.png"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Plantilla.png"),
        ]
        for candidate in candidates:
            candidate_path = os.path.abspath(candidate)
            if os.path.exists(candidate_path):
                return candidate_path
        return None

    def build_infographic_focus_prompt(self, template_path=None):
        if not template_path:
            return "Genera una infografía ejecutiva visual."

        # Aquí especificamos que la imagen NO es contenido, sino ESTRUCTURA
        return (
            f"USA ESTRICTAMENTE la estructura visual de '{os.path.basename(template_path)}'. "
            "Actúa como un diseñador gráfico que debe replicar el layout: "
            "1. Mantén la misma disposición de bloques y cajas que ves en la imagen. "
            "2. Usa la misma jerarquía de fuentes (tamaños y grosores). "
            "3. Respeta la paleta de colores corporativos de la imagen. "
            "4. Distribuye el texto del documento DOCX dentro de los contenedores visuales "
            "exactamente donde la plantilla propone los espacios de información."
        )

    def get_profiles(self):
        """Obtiene la lista de perfiles de autenticación disponibles."""
        try:
            # Ignoramos errores aquí para no bloquear la ejecución si falla el listado
            output = self.run_safe_cmd(["nlm", "login", "profile", "list"])
            profiles = {}
            for line in output.splitlines():
                if ":" in line and "Available profiles" not in line and "(invalid)" not in line:
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        name = parts[0].strip()
                        email = parts[1].strip()
                        if name:
                            profiles[name] = email
            return profiles
        except:
            return {}

    def select_profile(self):
        """Permite al usuario seleccionar qué cuenta de Google usar."""
        if not self.available: return
        
        profiles = self.get_profiles()
        if not profiles:
            return
        
        print("\n" + "="*60)
        print(" SELECCIÓN DE CUENTA DE GOOGLE (NotebookLM)")
        print("="*60)
        profile_names = list(profiles.keys())
        for i, name in enumerate(profile_names):
            print(f" [{i+1}] {name.ljust(15)} -> {profiles[name]}")
        print(f" [{len(profile_names)+1}] USAR OTRA CUENTA / NUEVA SESIÓN")
        print("="*60)
        
        try:
            choice = input(f"Seleccione una opción (1-{len(profile_names)+1}) [1]: ").strip()
            if not choice:
                idx = 0 # Por defecto el primero (default)
            else:
                idx = int(choice) - 1
            
            if 0 <= idx < len(profile_names):
                selected = profile_names[idx]
                print(f"    -> Cambiando al perfil: {selected}")
                self.run_safe_cmd(["nlm", "login", "switch", selected])
            elif idx == len(profile_names):
                print("    -> Preparando inicio de sesión con nueva cuenta...")
                # Usamos --clear para forzar que Chrome pida login de nuevo
                self.run_safe_cmd(["nlm", "login", "--clear"])
            else:
                print("    -> Opción no válida, se mantendrá la sesión actual.")
        except Exception as e:
            print(f"    -> Aviso: No se pudo cambiar de cuenta o se canceló ({e}).")
            print("    -> Continuando con la sesión actual...")

    def _try_download_from_url(self, url, dest_path):
        """Intenta descargar desde una URL con cabeceras de navegador."""
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                with open(dest_path, "wb") as f:
                    f.write(resp.read())
            return os.path.exists(dest_path) and os.path.getsize(dest_path) > 0
        except Exception as e:
            print(f"    -> Fallo al descargar URL ({e})")
            return False

    def _find_url_in_artifact(self, art, raw_json=""):
        """Busca la URL de descarga en múltiples campos posibles del artefacto."""
        url_keys = ["url", "download_url", "downloadUrl", "image_url", "imageUrl",
                    "src", "href", "link", "uri", "downloadUri", "contentUrl"]
        for key in url_keys:
            val = art.get(key, "")
            if val and str(val).startswith("http"):
                return str(val)

        # Búsqueda de URLs en el JSON en bruto que apunten a imágenes
        url_pattern = re.compile(r'https?://[^\s"\'<>\\]+', re.IGNORECASE)
        image_exts = ('.png', '.jpg', '.jpeg', '.webp', '.gif')
        image_keywords = ('image', 'infographic', 'export', 'download', 'media', 'storage')
        for url in url_pattern.findall(raw_json):
            lower = url.lower()
            if any(lower.endswith(ext) for ext in image_exts) or \
               any(kw in lower for kw in image_keywords):
                return url
        return ""

    def _try_cli_download(self, notebook_id, output_path):
        """Intenta descargar la infografía usando el comando correcto del CLI."""
        download_cmds = [
            ["nlm", "download", "infographic", notebook_id, "--output", output_path, "--no-progress"],
            ["nlm", "download", "infographic", notebook_id, "--output", output_path],
        ]
        for cmd in download_cmds:
            try:
                self.run_safe_cmd(cmd)
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    print(f"    -> Descargado vía: {' '.join(cmd[:4])}")
                    return True
            except subprocess.CalledProcessError:
                continue
        return False

    def generate_from_file(self, docx_path, output_path):
        if not self.available: return False
        temp_txt_path = os.path.join(tempfile.gettempdir(), "contexto_puro.txt")
        notebook_id = None
        style_template_path = self.resolve_style_template_path(docx_path)
        
        try:
            # 1. Preparar contenido (Convertir DOCX a TXT)
            print(">>> [1/5] Preparando el contenido del documento...")
            doc = Document(docx_path)
            full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            with open(temp_txt_path, "w", encoding="utf-8") as f:
                f.write(full_text)

            # 0. Selección de cuenta y verificación de autenticación
            print(">>> [0/5] Configurando cuenta de Google...")
            self.select_profile()
            
            try:
                self.run_safe_cmd(["nlm", "login", "--check"])
                print("    -> Sesión activa.")
            except subprocess.CalledProcessError:
                print("    -> Sesión no detectada o expirada, iniciando login...")
                self.run_safe_cmd(["nlm", "login"])

            # 2. Crear Cuaderno
            print(">>> [2/5] Creando Cuaderno en NotebookLM...")
            stdout_nb = self.run_safe_cmd(["nlm", "notebook", "create", "Generador_PDF_Auto"])
            notebook_id = self.extract_id(stdout_nb) or "Generador_PDF_Auto"
            print(f"    -> Cuaderno listo (ID: {notebook_id})")
            print(f"    -> Ver en navegador: https://notebooklm.google.com/notebook/{notebook_id}")

            # 3. Subir fuente de datos (probar variantes del CLI)
            print(">>> [3/5] Subiendo fuente de datos...")
            source_added = False
            source_cmds = [
                ["nlm", "source", "add", notebook_id, "--text", full_text[:8000], "--title", os.path.basename(docx_path)],
                ["nlm", "add", "source", notebook_id, "--path", temp_txt_path],
                ["nlm", "source", "add", notebook_id, "--file", temp_txt_path],
            ]
            for cmd in source_cmds:
                try:
                    self.run_safe_cmd(cmd)
                    source_added = True
                    break
                except subprocess.CalledProcessError:
                    continue
            
            if not source_added:
                print("    -> ERROR: No se pudo subir la fuente de datos.")
                return False

            if style_template_path:
                temp_style_txt = os.path.join(tempfile.gettempdir(), "instrucciones_estilo.txt")
                instrucciones_template = (
                    f"REGLA DE ORO DE DISEÑO: La imagen '{os.path.basename(style_template_path)}' "
                    "es la única plantilla permitida. Tiene un encabezado superior, "
                    "tres columnas de datos y un pie de página con iconos. "
                    "Cualquier infografía generada DEBE encajar en este esquema visual sin excepciones."
                )
                try:
                    with open(temp_style_txt, "w", encoding="utf-8") as f:
                        f.write(instrucciones_template)
                    self.run_safe_cmd(["nlm", "source", "add", notebook_id, "--file", temp_style_txt])
                    print("    -> Instrucciones de diseño vinculadas con éxito.")
                finally:
                    if os.path.exists(temp_style_txt): os.remove(temp_style_txt)

            print("    -> Sincronización completa.")
            time.sleep(25) # Un poco más de tiempo para procesar el texto extraº

            # Subir imagen de referencia de estilo si existe
            if style_template_path:
                try:
                    self.run_safe_cmd(["nlm", "source", "add", notebook_id, "--file", style_template_path])
                    print(f"    -> Referencia visual subida: {os.path.basename(style_template_path)}")
                except subprocess.CalledProcessError:
                    print(f"    -> Aviso: no se pudo subir la referencia visual (se continuará sin ella).")
            else:
                print("    -> Aviso: 'infografias_plantilla.jpg' no encontrada. Se generará sin referencia de estilo.")

            # 4. Generar Infografía (con reintentos y backoff)
            print(">>> [4/5] Generando Infografía en el Studio...")
            infographic_created = False
            infographic_focus_prompt = self.build_infographic_focus_prompt(style_template_path)
            cmd = [
                "nlm", "infographic", "create", notebook_id,
                "--focus", infographic_focus_prompt,
                "--confirm",
            ]
            max_retries = 4
            for attempt in range(1, max_retries + 1):
                try:
                    self.run_safe_cmd(cmd)
                    infographic_created = True
                    break
                except subprocess.CalledProcessError as e:
                    last_err = (getattr(e, 'output', '') or '') + (getattr(e, 'stderr', '') or '')
                    if attempt < max_retries:
                        wait = 20 * attempt
                        print(f"    -> Intento {attempt}/{max_retries} rechazado. Esperando {wait}s...")
                        time.sleep(wait)
                    else:
                        print(f"    -> ERROR tras {max_retries} intentos: {last_err.strip()}")

            if not infographic_created:
                return False

            print("    -> Generación de infografía iniciada.")

            # 5. Esperar a que la infografía esté lista y descargarla
            print(">>> [5/5] Esperando y descargando PNG final...")
            output_path_long = os.path.normpath(os.path.abspath(output_path))

            max_polls = 100
            infographic_completed = False
            for attempt in range(1, max_polls + 1):
                time.sleep(15)
                print(f"    -> Comprobando estado ({attempt}/{max_polls})...")
                try:
                    status_out = self.run_safe_cmd([
                        "nlm", "studio", "status", notebook_id, "--json", "--full"
                    ])
                    artifacts = json.loads(status_out)
                    for art in artifacts:
                        art_type = str(art.get("type", "")).lower()
                        art_status = str(art.get("status", "")).lower()

                        if "infographic" not in art_type:
                            continue

                        if art_status == "failed":
                            print("    -> La generación de la infografía falló en el servidor de NotebookLM.")
                            return False

                        if art_status == "completed":
                            infographic_completed = True
                            print("    -> Infografía lista. Descargando con CLI...")

                            # Estrategia 1: Comando CLI oficial de descarga
                            if self._try_cli_download(notebook_id, output_path_long):
                                print("    -> Descarga completada.")
                                return True

                            # Estrategia 2: URL dentro del artefacto JSON
                            art_url = self._find_url_in_artifact(art, status_out)
                            if art_url:
                                print(f"    -> URL encontrada: {art_url[:80]}...")
                                if self._try_download_from_url(art_url, output_path_long):
                                    print("    -> Descarga completada.")
                                    return True

                            print("    -> Reintentando en el próximo ciclo...")

                        else:
                            print(f"    -> Estado actual: {art_status}...")

                except (json.JSONDecodeError, subprocess.CalledProcessError):
                    pass

                # Si ya sabemos que está completado pero no pudimos descargar, acelerar reintentos
                if infographic_completed:
                    time.sleep(5)

            if infographic_completed:
                print("    -> La infografía se generó pero no fue posible descargarla automáticamente.")
                print(f"       Descárgala manualmente desde: https://notebooklm.google.com/notebook/{notebook_id}")
            else:
                print("    -> Timeout: la infografía no se completó a tiempo.")
            return False

        except Exception as e:
            print(f"ERROR EN NOTEBOOKLM: {e}")
            if hasattr(e, 'output') and e.output:
                print(f"Stdout: {e.output.strip()}")
            if hasattr(e, 'stderr') and e.stderr:
                print(f"Stderr: {e.stderr.strip()}")
            return False
        finally:
            if os.path.exists(temp_txt_path):
                os.remove(temp_txt_path)

class PDFCreator:
    def __init__(self, docx_path, output_path, img1_path=None):
        self.docx_path = os.path.abspath(docx_path)
        self.output_path = output_path
        self.img1_path = img1_path
        self.doc = Document(self.docx_path)
        self.styles = getSampleStyleSheet()
        self.elements = []
        self.temp_dir = tempfile.mkdtemp()
        self.extracted_images = {}
        self.infographic_generator = NotebookLMGenerator()
        self._setup_custom_styles()

    def _setup_custom_styles(self):
        """Configuración de estilos de texto corporativos."""
        self.custom_styles = {
            "TitleCover": ParagraphStyle("TitleCover", parent=self.styles["Title"], fontSize=48, textColor=colors.white, alignment=TA_CENTER, fontName="Helvetica-Bold", leading=56),
            "Title": ParagraphStyle("CustomTitle", parent=self.styles["Title"], fontSize=32, textColor=colors.HexColor(MAGENTA), alignment=TA_CENTER, spaceAfter=20, fontName="Helvetica-Bold"),
            "Heading 1": ParagraphStyle("CustomH1", parent=self.styles["Heading1"], fontSize=20, textColor=colors.HexColor(YELLOW), spaceBefore=20, spaceAfter=12, fontName="Helvetica-Bold"),
            "Heading 2": ParagraphStyle("CustomH2", parent=self.styles["Heading2"], fontSize=18, textColor=colors.HexColor(MAGENTA), spaceBefore=20, spaceAfter=12, fontName="Helvetica-Bold"),
            "Heading 3": ParagraphStyle("CustomH3", parent=self.styles["Heading3"], fontSize=12, textColor=colors.HexColor(NAVY_TEXT), spaceBefore=12, spaceAfter=8, fontName="Helvetica-Bold"),
            "Normal": ParagraphStyle("CustomNormal", parent=self.styles["Normal"], fontSize=11, leading=16, textColor=colors.HexColor(NAVY_TEXT), spaceAfter=10, fontName="Helvetica"),
            "List Bullet": ParagraphStyle("CustomBullet", parent=self.styles["Normal"], fontSize=11, leading=16, leftIndent=30, firstLineIndent=-15, textColor=colors.HexColor(NAVY_TEXT), spaceAfter=6, fontName="Helvetica"),
        }

    def is_bullet_style(self, para):
        style_name = para.style.name.lower()
        if "bullet" in style_name or "list" in style_name:
            return True
        if para._element.xpath('./w:pPr/w:numPr'):
            return True
        return False

    def extract_images(self):
        """Extrae imágenes embebidas del DOCX al directorio temporal."""
        for i, rel in enumerate(self.doc.part.rels.values()):
            if "image" in rel.target_ref:
                img_part = rel.target_part
                img_ext = os.path.splitext(rel.target_ref)[1]
                img_path = os.path.join(self.temp_dir, f"img_{i}{img_ext}")
                with open(img_path, "wb") as f:
                    f.write(img_part.blob)
                self.extracted_images[rel.rId] = img_path

    def cleanup(self):
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)

    def draw_cover(self, canvas, doc):
        """Dibuja la portada del documento."""
        canvas.saveState()
        if self.img1_path and os.path.exists(self.img1_path):
            canvas.drawImage(self.img1_path, 0, 0, width=A4[0], height=A4[1])
        canvas.restoreState()

    def process_table(self, table, pdf_width):
        data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(Paragraph(cell.text, self.custom_styles["Normal"]))
            data.append(row_data)

        t = Table(data, colWidths=[pdf_width / len(data[0])] * len(data[0]))
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.white),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor(NAVY_TEXT)),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor(NAVY_TEXT)),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor(NAVY_TEXT)),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ]))
        return t

    def create_pdf(self):
        """Crea el PDF a partir del archivo Word."""
        print(f"--- Procesando Documento: {os.path.basename(self.docx_path)} ---")

        # Generar infografía global con NotebookLM (sin tocar esta lógica)
        global_info_path = os.path.join(self.temp_dir, "global_info.png")
        has_info = self.infographic_generator.generate_from_file(self.docx_path, global_info_path)

        # Extraer imágenes embebidas del DOCX
        self.extract_images()

        pdf = SimpleDocTemplate(self.output_path, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
        first_title = True

        try:
            for block in self.doc.element.body:
                if block.tag.endswith('p'):
                    para = DocxParagraph(block, self.doc)

                    # Recopilar imágenes embebidas en este párrafo
                    para_images = []
                    for run in para.runs:
                        blips = run._element.xpath('.//a:blip')
                        drawings = run._element.xpath('.//w:drawing')
                        blip_sizes = {}
                        for drawing in drawings:
                            extents = drawing.xpath('.//wp:extent')
                            d_blips = drawing.xpath('.//a:blip')
                            if d_blips and extents:
                                d_rId = d_blips[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                cx = extents[0].get('cx')
                                cy = extents[0].get('cy')
                                d_align = None
                                align_tags = (drawing.xpath('.//wp:positionH/wp:align') or
                                              drawing.xpath('.//wp:inline/wp:align') or
                                              drawing.xpath('.//*[local-name()="align"]'))
                                if align_tags:
                                    d_align = align_tags[0].text.upper()
                                if cx and cy:
                                    blip_sizes[d_rId] = {'size': (int(cx), int(cy)), 'align': d_align}
                        for blip in blips:
                            rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                            if rId in self.extracted_images:
                                img_info = blip_sizes.get(rId, {})
                                para_images.append({
                                    'path': self.extracted_images[rId],
                                    'size': img_info.get('size'),
                                    'align': img_info.get('align'),
                                })

                    text = para.text.strip()
                    style_name = para.style.name

                    # Portada: primer título
                    if style_name == "Title" and first_title:
                        self.elements.extend([Spacer(1, 230), Paragraph(text, self.custom_styles["TitleCover"]), PageBreak()])
                        if has_info:
                            img = PlatypusImage(global_info_path)
                            max_width = pdf.width * 0.95
                            aspect = img.imageHeight / float(img.imageWidth)
                            img.drawWidth = max_width
                            img.drawHeight = max_width * aspect
                            self.elements.extend([
                                Paragraph("Resumen Ejecutivo (NotebookLM)", self.custom_styles["Heading 1"]),
                                Spacer(1, 15),
                                img,
                                PageBreak()
                            ])
                        first_title = False
                        continue

                    # Determinar estilo del párrafo
                    if self.is_bullet_style(para):
                        p_style = self.custom_styles["List Bullet"]
                        if text and not text.startswith(('•', '-', '*')):
                            text = f"• {text}"
                    elif style_name in self.custom_styles:
                        p_style = self.custom_styles[style_name]
                    else:
                        p_style = self.custom_styles["Normal"]

                    # Caso 1: texto con imagen inline (side-by-side)
                    if text and para_images:
                        img_path = para_images[0]['path']
                        try:
                            img = PlatypusImage(img_path)
                            aspect = img.imageHeight / float(img.imageWidth)
                            col_img_width = pdf.width * 0.35
                            col_text_width = pdf.width * 0.62
                            img.drawWidth = col_img_width
                            img.drawHeight = col_img_width * aspect
                            drawing_align = para_images[0].get('align')
                            if drawing_align == 'RIGHT':
                                data = [[Paragraph(text, p_style), img]]
                                col_widths = [col_text_width, col_img_width]
                            elif drawing_align == 'LEFT':
                                data = [[img, Paragraph(text, p_style)]]
                                col_widths = [col_img_width, col_text_width]
                            else:
                                image_is_first = False
                                for run in para.runs:
                                    if run.text.strip():
                                        break
                                    if run._element.xpath('.//a:blip'):
                                        image_is_first = True
                                        break
                                if image_is_first:
                                    data = [[img, Paragraph(text, p_style)]]
                                    col_widths = [col_img_width, col_text_width]
                                else:
                                    data = [[Paragraph(text, p_style), img]]
                                    col_widths = [col_text_width, col_img_width]
                            t = Table(data, colWidths=col_widths)
                            t.setStyle(TableStyle([
                                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                ('LEFTPADDING', (0, 0), (0, 0), 0),
                                ('RIGHTPADDING', (0, 0), (0, 0), 15),
                                ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                            ]))
                            self.elements.append(t)
                        except Exception as e:
                            print(f"Error añadiendo imagen inline: {e}")
                            self.elements.append(Paragraph(text, p_style))

                    # Caso 2: imagen sola (sin texto)
                    elif para_images and not text:
                        for img_data in para_images:
                            img_path = img_data['path']
                            size_emu = img_data['size']
                            try:
                                img = PlatypusImage(img_path)
                                aspect = img.imageHeight / float(img.imageWidth)
                                max_width = pdf.width * 0.9
                                if size_emu:
                                    width_pt = size_emu[0] / 12700.0
                                    height_pt = size_emu[1] / 12700.0
                                    if width_pt > max_width:
                                        img.drawWidth = max_width
                                        img.drawHeight = max_width * aspect
                                    else:
                                        img.drawWidth = width_pt
                                        img.drawHeight = height_pt
                                else:
                                    img.drawWidth = min(img.imageWidth, max_width)
                                    img.drawHeight = img.drawWidth * aspect
                                drawing_align = img_data.get('align')
                                if drawing_align in ['LEFT', 'CENTER', 'RIGHT']:
                                    img.hAlign = drawing_align
                                elif para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                                    img.hAlign = 'CENTER'
                                elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                                    img.hAlign = 'RIGHT'
                                else:
                                    img.hAlign = 'LEFT'
                                self.elements.append(img)
                                self.elements.append(Spacer(1, 15))
                            except Exception as e:
                                print(f"Error añadiendo imagen: {e}")

                    # Caso 3: solo texto
                    elif text:
                        self.elements.append(Paragraph(text, p_style))
                        if "Heading" in style_name:
                            self.elements.append(Spacer(1, 10))

                    # Caso 4: línea vacía
                    else:
                        self.elements.append(Spacer(1, 10))

                elif block.tag.endswith('tbl'):
                    table = DocxTable(block, self.doc)
                    self.elements.append(Spacer(1, 15))
                    self.elements.append(self.process_table(table, pdf.width))
                    self.elements.append(Spacer(1, 15))

            pdf.build(self.elements, onFirstPage=self.draw_cover)
            print(f">>> PDF CREADO EXITOSAMENTE: {self.output_path}")
        finally:
            self.cleanup()

def select_file():
    """Selecciona un archivo Word mediante diálogo."""
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    path = filedialog.askopenfilename(title="Selecciona el documento Word", filetypes=[("Archivos Word", "*.docx")])
    root.destroy()
    return path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convierte un archivo DOCX a un PDF con estilo personalizado.")
    parser.add_argument("input", nargs="?", help="Ruta al archivo .docx de entrada")
    parser.add_argument("-o", "--output", help="Ruta de salida para el PDF (por defecto [input].pdf)")
    parser.add_argument("--cover", help="Ruta a la imagen de portada")
    args = parser.parse_args()

    docx_file = args.input if args.input else select_file()

    if docx_file and os.path.exists(docx_file):
        # Determine the project root (where doc2pdf is)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(os.path.dirname(script_dir))
        
        # Create PDF directory at project root
        output_dir = os.path.join(project_root, "PDF")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        if args.output:
            # If output is just a filename, put it in the PDF folder
            if not os.path.dirname(args.output):
                output = os.path.join(output_dir, args.output)
            else:
                output = args.output
        else:
            filename = os.path.splitext(os.path.basename(docx_file))[0] + ".pdf"
            output = os.path.join(output_dir, filename)

        img1 = args.cover
        if not img1:
            candidates = [
                os.path.join(os.path.dirname(os.path.abspath(docx_file)), "portada.jpg"),
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "portada.jpg"),
                os.path.join("doc2pdf", "pdfCreation", "portada.jpg"),
                os.path.join("pdfCreation", "portada.jpg"),
            ]
            for c in candidates:
                if os.path.exists(c):
                    img1 = c
                    break

        if img1 and not os.path.exists(img1):
            print(f"Aviso: No se encontró la imagen de portada en {img1}. Se generará el PDF sin ella.")
            img1 = None

        PDFCreator(docx_file, output, img1_path=img1).create_pdf()
    else:
        print("Operación cancelada.")
