import os
import sys


def get_base_path():
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.abspath(".")


try:
    import pytesseract
    portable_path = os.path.join(get_base_path(), 'tesseract', 'tesseract.exe')
    if os.path.exists(portable_path):
        pytesseract.pytesseract.tesseract_cmd = portable_path
    else:
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    pytesseract.get_tesseract_version()
    PYTESSERACT_AVAILABLE = True
except Exception:
    PYTESSERACT_AVAILABLE = False
