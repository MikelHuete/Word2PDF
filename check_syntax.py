import ast
import sys

try:
    with open('doc2pdf/pdfCreation/pdf_creator.py', encoding='utf-8') as f:
        code = f.read()
    ast.parse(code)
    print('Sintaxis OK')
except SyntaxError as e:
    print(f"SyntaxError: línea {e.lineno}: {e.msg}")
    if e.text:
        print(f"  {e.text.rstrip()}")
        if e.offset:
            print(f"  {' ' * (e.offset - 1)}^")
    sys.exit(1)
except Exception as e:
    print(f"Error: {type(e).__name__}: {e}")
    sys.exit(1)
