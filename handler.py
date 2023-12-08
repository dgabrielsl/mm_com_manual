import os
try: import fitz
except Exception as e: print(e)
try: import pytesseract
except Exception as e: print(e)
from PIL import Image

def makeocr(downloadfrom, scanner, h_crafted_text):
    global h_raw_text, h_text_to_lines
    h_raw_text = ''; h_text_to_lines = ''; required_text = []
    zoom = 4.0
    global doc
    doc = ''
    for df in scanner:
        _df = df.upper()
        if _df.__contains__('PAGAR'): doc = df
    document = fitz.open(f'{downloadfrom}/{doc}')
    mat = fitz.Matrix(zoom, zoom)
    pix = document[0].get_pixmap(matrix=mat)
    new_jpg = f'{downloadfrom}/handler_jpg_for_ocr.jpg'
    pix.save(new_jpg)
    try: h_raw_text = str(((pytesseract.image_to_string(Image.open(new_jpg)))))
    except: pass
    finally:
        try: os.remove(new_jpg)
        except Exception as e: print(e)
    h_text_to_lines = h_raw_text.split('\n')
    for ttl in h_text_to_lines:
        if ttl.__contains__('vencim') and ttl.__contains__('vist') and ttl.__contains__('may'): required_text.append(ttl)
        if ttl.__contains__('-') and ttl.__contains__('presente letra'): required_text.append(ttl)
    search_fullname = required_text[0]
    search_fullname = search_fullname.split(',')
    search_fullname = search_fullname[0]
    search_fullname = search_fullname.split(' ')
    search_fullname = search_fullname[8:]
    search_fullname = ' '.join(search_fullname)
    h_crafted_text.append(search_fullname)
    required_text = required_text[1]
    search_id = ''
    for rt in required_text:
        n = rt.isnumeric()
        if n: search_id += rt
    h_crafted_text.append(search_id)