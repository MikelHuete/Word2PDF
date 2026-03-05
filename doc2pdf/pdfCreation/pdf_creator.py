import argparse
import os
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as PlatypusImage
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import tempfile
import shutil

from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable
from docx.enum.text import WD_ALIGN_PARAGRAPH

# THE BRAND COLORS
MAGENTA = "#DA1984"
NAVY_TEXT = "#232D4B"
WHITE = "#FFFFFF"
YELLOW = "#FCEE21"

class PDFCreator:
    def __init__(self, docx_path, output_path, img1_path=None, img2_path=None):
        self.docx_path = docx_path
        self.output_path = output_path
        self.img1_path = img1_path
        self.img2_path = img2_path
        self.doc = Document(docx_path)
        self.styles = getSampleStyleSheet()
        self.elements = []
        self.temp_dir = None
        self.extracted_images = {}
        
        # Define Custom Styles
        self.custom_styles = {
            "TitleCover": ParagraphStyle(
                "TitleCover",
                parent=self.styles["Title"],
                fontSize=48,
                textColor=colors.white,
                alignment=TA_CENTER,
                fontName="Helvetica-Bold",
                leading=56
            ),
            "Title": ParagraphStyle(
                "CustomTitle",
                parent=self.styles["Title"],
                fontSize=32,
                textColor=colors.HexColor(MAGENTA),
                alignment=TA_CENTER,
                spaceAfter=20,
                fontName="Helvetica-Bold"
            ),
            "Heading 1": ParagraphStyle(
                "CustomH1",
                parent=self.styles["Heading1"],
                fontSize=20,
                textColor=colors.HexColor(YELLOW),
                spaceBefore=20,
                spaceAfter=12,
                fontName="Helvetica-Bold"
            ),
            "Heading 2": ParagraphStyle(
                "CustomH2",
                parent=self.styles["Heading2"],
                fontSize=18,  # Increased size as per reference "6 Azure Function"
                textColor=colors.HexColor(MAGENTA),
                spaceBefore=20,
                spaceAfter=12,
                fontName="Helvetica-Bold"
            ),
            "Heading 3": ParagraphStyle(
                "CustomH3",
                parent=self.styles["Heading3"],
                fontSize=12,
                textColor=colors.HexColor(NAVY_TEXT),
                spaceBefore=12,
                spaceAfter=8,
                fontName="Helvetica-Bold"
            ),
            "Normal": ParagraphStyle(
                "CustomNormal",
                parent=self.styles["Normal"],
                fontSize=11,
                leading=16,
                textColor=colors.HexColor(NAVY_TEXT),
                spaceAfter=10,
                fontName="Helvetica"
            ),
            "List Bullet": ParagraphStyle(
                "CustomBullet",
                parent=self.styles["Normal"],
                fontSize=11,
                leading=16,
                leftIndent=30,
                firstLineIndent=-15,
                textColor=colors.HexColor(NAVY_TEXT),
                spaceAfter=6,
                fontName="Helvetica"
            )
        }

    def is_bullet_style(self, para):
        style_name = para.style.name.lower()
        if "bullet" in style_name or "list" in style_name:
            return True
        if para._element.xpath('./w:pPr/w:numPr'):
            return True
        return False

    def extract_images(self):
        self.temp_dir = tempfile.mkdtemp()
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
        canvas.saveState()
        if self.img1_path and os.path.exists(self.img1_path):
            canvas.drawImage(self.img1_path, 0, 0, width=A4[0], height=A4[1])
        canvas.restoreState()

    def process_table(self, table, pdf_width):
        data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                # Use Normal style for cell content
                row_data.append(Paragraph(cell.text, self.custom_styles["Normal"]))
            data.append(row_data)
        
        t = Table(data, colWidths=[pdf_width/len(data[0])] * len(data[0]))
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.white), # Header background white as per image
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
        self.extract_images()
        pdf = SimpleDocTemplate(self.output_path, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
        
        first_title = True
        
        # Interleave paragraphs and tables by iterating through parent elements
        for block in self.doc.element.body:
            if block.tag.endswith('p'):
                para = DocxParagraph(block, self.doc)
                
                # Collect images in this paragraph
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
                            
                            # Extract drawing alignment
                            d_align = None
                            align_tags = drawing.xpath('.//wp:positionH/wp:align')
                            if not align_tags:
                                align_tags = drawing.xpath('.//wp:inline/wp:align')
                            if not align_tags:
                                # Fallback to any align tag inside drawing
                                align_tags = drawing.xpath('.//*[local-name()="align"]')
                                
                            if align_tags:
                                d_align = align_tags[0].text.upper()
                                
                            if cx and cy:
                                blip_sizes[d_rId] = {
                                    'size': (int(cx), int(cy)),
                                    'align': d_align
                                }

                    for blip in blips:
                        rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if rId in self.extracted_images:
                            img_info = blip_sizes.get(rId, {})
                            para_images.append({
                                'path': self.extracted_images[rId],
                                'size': img_info.get('size'),
                                'align': img_info.get('align')
                            })

                text = para.text.strip()
                style_name = para.style.name
                
                # Determine paragraph style
                if self.is_bullet_style(para):
                    p_style = self.custom_styles["List Bullet"]
                    if text and not text.startswith(('•', '-', '*')):
                        text = f"• {text}"
                elif style_name in self.custom_styles:
                    p_style = self.custom_styles[style_name]
                else:
                    p_style = self.custom_styles["Normal"]

                # Case 1: Title Cover
                if style_name == "Title" and first_title:
                    self.elements.append(Spacer(1, 230)) 
                    self.elements.append(Paragraph(text, self.custom_styles["TitleCover"]))
                    self.elements.append(PageBreak())
                    first_title = False
                    continue

                if text and para_images:
                    img_path = para_images[0]['path']
                    try:
                        img = PlatypusImage(img_path)
                        aspect = img.imageHeight / float(img.imageWidth)
                        
                        col_img_width = pdf.width * 0.35
                        col_text_width = pdf.width * 0.62
                        
                        img.drawWidth = col_img_width
                        img.drawHeight = col_img_width * aspect
                        
                        # Determine if image should be on the LEFT or RIGHT based on XML metadata
                        drawing_align = para_images[0].get('align')
                        
                        if drawing_align == 'RIGHT':
                            # Text on the LEFT, Image on the RIGHT
                            data = [[Paragraph(text, p_style), img]]
                            col_widths = [col_text_width, col_img_width]
                            padding_settings = [
                                ('RIGHTPADDING', (0, 0), (0, 0), 15),
                                ('LEFTPADDING', (1, 0), (1, 0), 0),
                            ]
                        elif drawing_align == 'LEFT':
                            # Image on the LEFT, Text on the RIGHT
                            data = [[img, Paragraph(text, p_style)]]
                            col_widths = [col_img_width, col_text_width]
                            padding_settings = [
                                ('RIGHTPADDING', (0, 0), (0, 0), 15),
                                ('LEFTPADDING', (1, 0), (1, 0), 0),
                            ]
                        else:
                            # Fallback to run order if no explicit XML alignment
                            image_is_first = False
                            for run in para.runs:
                                if run.text.strip():
                                    image_is_first = False
                                    break
                                if run._element.xpath('.//a:blip'):
                                    image_is_first = True
                                    break

                            if image_is_first:
                                # Image on the LEFT, Text on the RIGHT
                                data = [[img, Paragraph(text, p_style)]]
                                col_widths = [col_img_width, col_text_width]
                                padding_settings = [
                                    ('RIGHTPADDING', (0, 0), (0, 0), 15),
                                    ('LEFTPADDING', (1, 0), (1, 0), 0),
                                ]
                            else:
                                # Text on the LEFT, Image on the RIGHT
                                data = [[Paragraph(text, p_style), img]]
                                col_widths = [col_text_width, col_img_width]
                                padding_settings = [
                                    ('RIGHTPADDING', (0, 0), (0, 0), 15),
                                    ('LEFTPADDING', (1, 0), (1, 0), 0),
                                ]

                        t = Table(data, colWidths=col_widths)
                        base_style = [
                            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                            ('LEFTPADDING', (0, 0), (0, 0), 0),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                        ]
                        t.setStyle(TableStyle(base_style + padding_settings))
                        self.elements.append(t)
                    except Exception as e:
                        print(f"Error adding side-by-side image: {e}")
                        self.elements.append(Paragraph(text, p_style))
                
                # Case 3: Standing Alone Image(s)
                elif para_images and not text:
                    for img_data in para_images:
                        img_path = img_data['path']
                        size_emu = img_data['size']
                        try:
                            img = PlatypusImage(img_path)
                            aspect = img.imageHeight / float(img.imageWidth)
                            max_width = pdf.width * 0.9
                            
                            if size_emu:
                                # 1 point = 12700 EMU
                                width_pt = size_emu[0] / 12700.0
                                height_pt = size_emu[1] / 12700.0
                                if width_pt > max_width:
                                    img.drawWidth = max_width
                                    img.drawHeight = max_width * aspect
                                else:
                                    img.drawWidth = width_pt
                                    img.drawHeight = height_pt
                            else:
                                if img.imageWidth > max_width:
                                    img.drawWidth = max_width
                                else:
                                    img.drawWidth = img.imageWidth
                                img.drawHeight = img.drawWidth * aspect
                            
                            # Determine final alignment
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
                            print(f"Error adding standalone image: {e}")

                # Case 4: Text Only
                elif text:
                    self.elements.append(Paragraph(text, p_style))
                    if "Heading" in style_name:
                        self.elements.append(Spacer(1, 10))

                # Case 5: Empty line
                elif not text and not para_images:
                    self.elements.append(Spacer(1, 10))

            elif block.tag.endswith('tbl'):
                table = DocxTable(block, self.doc)
                self.elements.append(Spacer(1, 15))
                self.elements.append(self.process_table(table, pdf.width))
                self.elements.append(Spacer(1, 15))

        try:
            pdf.build(self.elements, onFirstPage=self.draw_cover)
            print(f"PDF Successfully created at: {self.output_path}")
        finally:
            self.cleanup()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convierte un archivo DOCX a un PDF con estilo personalizado.")
    parser.add_argument("input", help="Ruta al archivo .docx de entrada")
    parser.add_argument("-o", "--output", help="Ruta de salida para el PDF (por defecto [input].pdf)")
    parser.add_argument("--cover", help="Ruta a la imagen de portada")
    
    args = parser.parse_args()
    
    input_file = args.input
    if not os.path.exists(input_file):
        print(f"Error: No se encontró el archivo de entrada {input_file}")
        exit(1)
        
    if args.output:
        output_file = args.output
    else:
        output_file = os.path.splitext(input_file)[0] + ".pdf"
        
    # Ensure output directory exists
    output_dir = os.path.dirname(os.path.abspath(output_file))
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    img1 = args.cover
    if not img1:
        # Try to find default cover in common locations
        img1_candidates = [
            os.path.join(os.path.dirname(os.path.abspath(input_file)), "portada.jpg"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "portada.jpg"),
            os.path.join("doc2pdf", "pdfCreation", "portada.jpg"),
            os.path.join("pdfCreation", "portada.jpg"),
        ]
        for candidate in img1_candidates:
            if os.path.exists(candidate):
                img1 = candidate
                break
    
    if img1 and not os.path.exists(img1):
        print(f"Aviso: No se encontró la imagen de portada en {img1}. Se generará el PDF sin ella.")
        img1 = None
        
    creator = PDFCreator(input_file, output_file, img1_path=img1)
    creator.create_pdf()
