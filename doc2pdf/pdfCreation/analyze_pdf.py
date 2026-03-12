import fitz  # PyMuPDF
from pathlib import Path

PDF_PATH = r"C:\Users\mikel.huete.ext\Desktop\Word to Pdf\doc2pdf\pdfCreation\generated_styled.pdf"
OUT_DIR = r"C:\Users\mikel.huete.ext\Desktop\Word to Pdf\doc2pdf\pdfCreation"

def rgb_to_hex(r, g, b):
    return "#{:02X}{:02X}{:02X}".format(int(r*255), int(g*255), int(b*255))

def analyze_color(c):
    if c is None:
        return "None"
    if isinstance(c, (int, float)):
        v = int(c * 255)
        return "#{:02X}{:02X}{:02X}".format(v, v, v)
    if isinstance(c, (list, tuple)) and len(c) == 3:
        return rgb_to_hex(c[0], c[1], c[2])
    if isinstance(c, (list, tuple)) and len(c) == 4:
        r = 1 - c[0] - c[3]
        g = 1 - c[1] - c[3]
        b = 1 - c[2] - c[3]
        return rgb_to_hex(max(0,r), max(0,g), max(0,b))
    return str(c)

doc = fitz.open(PDF_PATH)
print(f"=== PDF ANALYSIS ===")
print(f"Total pages: {len(doc)}")
print(f"PDF metadata: {doc.metadata}")
print()

all_fonts = {}

for page_num in range(len(doc)):
    page = doc[page_num]
    rect = page.rect
    print(f"{'='*60}")
    print(f"PAGE {page_num+1}")
    print(f"  Size: {rect.width:.1f} x {rect.height:.1f} pts ({rect.width/72:.2f} x {rect.height/72:.2f} inches)")
    
    try:
        bg = page.get_background_color()
        print(f"  Background: {bg} -> hex: {analyze_color(bg)}")
    except:
        print(f"  Background: could not determine")
    
    drawings = page.get_drawings()
    print(f"\n  DRAWINGS ({len(drawings)} total):")
    for i, d in enumerate(drawings[:20]):
        fill = analyze_color(d.get('fill'))
        stroke = analyze_color(d.get('color'))
        print(f"    [{i+1}] type={d.get('type','?')}, rect={d.get('rect')}, fill={fill}, stroke={stroke}, width={d.get('width')}")
    
    blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
    text_blocks = [b for b in blocks if b['type']==0]
    print(f"\n  TEXT BLOCKS ({len(text_blocks)}):")
    for b in text_blocks:
        bx0, by0, bx1, by1 = b['bbox']
        print(f"    Block bbox=({bx0:.1f},{by0:.1f},{bx1:.1f},{by1:.1f}):")
        for line in b['lines']:
            for span in line['spans']:
                font_name = span['font']
                font_size = span['size']
                color_int = span['color']
                r = ((color_int >> 16) & 0xFF) / 255
                g = ((color_int >> 8) & 0xFF) / 255
                bl = (color_int & 0xFF) / 255
                color_hex = rgb_to_hex(r, g, bl)
                flags = span['flags']
                is_bold = bool(flags & 2**4)
                is_italic = bool(flags & 2**1)
                origin = span['origin']
                text = span['text'][:100]
                print(f"      font='{font_name}' size={font_size:.1f} color={color_hex} bold={is_bold} italic={is_italic} at({origin[0]:.1f},{origin[1]:.1f}) text='{text}'")
                key = f"{font_name}_{font_size:.1f}_{color_hex}"
                if key not in all_fonts:
                    all_fonts[key] = {'font': font_name, 'size': font_size, 'color': color_hex, 'bold': is_bold, 'italic': is_italic, 'count': 0}
                all_fonts[key]['count'] += 1
    
    images = page.get_images(full=True)
    print(f"\n  IMAGES ({len(images)}):")
    for img in images:
        try:
            img_rect = page.get_image_bbox(img)
            print(f"    xref={img[0]}, bbox={img_rect}, size={img[2]}x{img[3]}")
        except:
            print(f"    xref={img[0]}")
    
    if page_num < 3:
        mat = fitz.Matrix(150/72, 150/72)
        pix = page.get_pixmap(matrix=mat)
        out_path = str(Path(OUT_DIR) / f"page{page_num+1}.png")
        pix.save(out_path)
        print(f"\n  -> Saved {out_path}")

print(f"\n{'='*60}")
print(f"ALL UNIQUE FONT STYLES:")
for key, info in sorted(all_fonts.items(), key=lambda x: -x[1]['size']):
    print(f"  font='{info['font']}' size={info['size']:.1f} color={info['color']} bold={info['bold']} italic={info['italic']} count={info['count']}")

doc.close()
print("\nDone!")
