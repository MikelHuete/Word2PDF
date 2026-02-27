import os
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as PlatypusImage
from reportlab.lib.enums import TA_CENTER, TA_LEFT

from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable

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

    def draw_cover(self, canvas, doc):
        canvas.saveState()
        if self.img1_path and os.path.exists(self.img1_path):
            canvas.drawImage(self.img1_path, 0, 0, width=800, height=A4[1])
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
        pdf = SimpleDocTemplate(self.output_path, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
        
        first_title = True
        
        # Interleave paragraphs and tables by iterating through parent elements
        for block in self.doc.element.body:
            if block.tag.endswith('p'):
                para = DocxParagraph(block, self.doc)
                text = para.text.strip()
                if not text:
                    self.elements.append(Spacer(1, 10))
                    continue
                    
                style_name = para.style.name
                
                if self.is_bullet_style(para):
                    p_style = self.custom_styles["List Bullet"]
                    if not text.startswith(('•', '-', '*')):
                        text = f"• {text}"
                elif style_name in self.custom_styles:
                    p_style = self.custom_styles[style_name]
                else:
                    p_style = self.custom_styles["Normal"]
                
                if style_name == "Title" and first_title:
                    self.elements.append(Spacer(1, 250)) 
                    self.elements.append(Paragraph(text, self.custom_styles["TitleCover"]))
                    self.elements.append(PageBreak())
                    first_title = False
                    continue

                self.elements.append(Paragraph(text, p_style))
                if "Heading" in style_name:
                    self.elements.append(Spacer(1, 10))

            elif block.tag.endswith('tbl'):
                table = DocxTable(block, self.doc)
                self.elements.append(Spacer(1, 15))
                self.elements.append(self.process_table(table, pdf.width))
                self.elements.append(Spacer(1, 15))

        pdf.build(self.elements, onFirstPage=self.draw_cover)
        print(f"PDF Successfully created at: {self.output_path}")

if __name__ == "__main__":
    input_file = "1. Artificial Intelligence - Copia.docx"
    output_file = os.path.join("pdfCreation", "generated_styled.pdf")
    
    img1 = os.path.join("pdfCreation", "portada.jpg")
    
    if not os.path.exists(input_file):
        input_file = os.path.join("..", input_file)
    if not os.path.exists(img1):
        img1 = os.path.join("..", "extracted_media", "word", "media", "portada.jpg")     
    creator = PDFCreator(input_file, output_file, img1_path=img1)
    creator.create_pdf()
