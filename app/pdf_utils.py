import os


FONT_CANDIDATES = [
    '/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf',
    '/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf',
    '/usr/share/fonts/truetype/amiri/Amiri-Regular.ttf',
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/freefont/FreeSans.ttf',
]

_arabic_font_registered = False


def find_and_register_arabic_font() -> str | None:
    global _arabic_font_registered
    if _arabic_font_registered:
        return 'ArabicFont'

    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    for path in FONT_CANDIDATES:
        try:
            if os.path.exists(path):
                pdfmetrics.registerFont(TTFont('ArabicFont', path))
                _arabic_font_registered = True
                return 'ArabicFont'
        except Exception:
            continue
    return None


def arabic_text(text: str) -> str:
    import arabic_reshaper
    from bidi.algorithm import get_display

    try:
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text
