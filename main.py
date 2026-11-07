"""
Etiketi za apteka - cheте Excel/CSV fail (skladova razpiska/faktura) i generira
PDF s etiketi (kod + tsena v EUR i lv.), goten za pechat na A4.
"""
import csv
import datetime
import json
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth, registerFont
from reportlab.pdfbase.ttfonts import TTFont

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

if os.name == "nt":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

def _bundle_root():
    """Papka s bundlvani resursi (assets) - pri .exe e vremennata PyInstaller
    papka (sys._MEIPASS), pri raboti ot iztochnik e papkata na main.py."""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def _app_dir():
    """Papka do izpalnimiya .exe (ili do main.py pri raboti ot iztochnik) - tam
    se zapazva config.json, za da prezhivyava mezhdu startirania na programata."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(_app_dir(), "config.json")
ASSETS_DIR = os.path.join(_bundle_root(), "assets")
LOGO_PATH = os.path.join(ASSETS_DIR, "logo.png")

SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv", ".pdf")

# Tsvetova palitra: gorna chast (logo/zaglavie) - cherno kato loгото; osnovnata
# rabotna oblast (pod tabovete) - toplo bezhovo, za po-lesno chetene na formite.
HEADER_BG = "#020203"
TAB_INACTIVE_BG = "#14170f"
GOLD = "#d4af37"
GOLD_BRIGHT = "#f2d675"
GREEN = "#2e9e5c"
GREEN_BRIGHT = "#42c47a"
DANGER = "#c73e3e"

BG = "#f6efdd"
BG_PANEL = "#eee1c3"
BG_FIELD = "#fffdf6"
TEXT = "#2c2413"
TEXT_MUTED = "#7d735c"
BORDER = "#cdbf9c"
GOLD_TEXT = "#8a5d0d"
GREEN_TEXT = "#1f7a44"

# Helvetica (base14) няма кирилски глифи -> "лв." излизаше като квадратчета.
# Регистрираме Arial (има кирилица) за текста в PDF-а.
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
_WIN_FONTS = r"C:\Windows\Fonts"
try:
    registerFont(TTFont("AppFont", os.path.join(_WIN_FONTS, "arial.ttf")))
    registerFont(TTFont("AppFont-Bold", os.path.join(_WIN_FONTS, "arialbd.ttf")))
    FONT_REGULAR = "AppFont"
    FONT_BOLD = "AppFont-Bold"
except Exception:
    pass

# Podredba na etiketite na A4 list (8 kolони, kompaktni etiketi, bez ramka -
# tochno kakto v obraztsoviq PDF, koyto aptekata veche izpolzva)
PAGE_W, PAGE_H = A4
MARGIN = 10 * mm
COLS = 8
ROW_H = 16 * mm
ROWS = int((PAGE_H - 2 * MARGIN) // ROW_H)
CELL_W = (PAGE_W - 2 * MARGIN) / COLS

# Tochni razmeri na fizicheskiya list samozalepvashti stikeri, koito aptekata
# realno izpolzva za "Svobodna prodazhba" / "Rachno vavezhdane" (izmereni s
# linийka ot potrebitelya - 8 kolони x 20 reda = 160 stikera na list, 22x12mm).
LBL_COLS = 8
LBL_ROWS = 20

# Nachalni (podrazbirashti se) razmeri v mm - polzvat se samo ako potrebitelyat
# oshte ne e zapazil sobstveni v Nastroyki na etiketite (kalibratsia).
# Svobodna prodazhba i NZOK imat SEPARATE nastroyki (razlichen broy redove
# tekst i razlichni kodove -> mozhe da im trqbva razlichen razmer na shrifta).
DEFAULT_LABEL_LAYOUT = {
    "w": 20.0,
    "h": 12.3,
    "margin_left": 2.0,
    "margin_top": 10.0,
    "gap_h": 6.0,
    "gap_v": 2.0,
    "rotation_deg": 0.0,
    "scale_pct": 100.0,
    "code_font_size": 8.0,
}

LABEL_LAYOUT_KEYS = ("otc", "nzok")


def _label_layout_cfg_key(kind):
    return "label_layout_otc" if kind == "otc" else "label_layout_nzok"


def load_label_layout(kind="otc"):
    cfg = load_config()
    layout = dict(DEFAULT_LABEL_LAYOUT)
    # Migratsia ot staroto edinstveno pole "label_layout" (predi Svobodna/NZOK
    # da imat otdelni nastroyki) - za da ne se gubi veche napravenata kalibratsia.
    if "label_layout_otc" not in cfg and "label_layout" in cfg:
        cfg["label_layout_otc"] = cfg["label_layout"]
    # NZOK etiketite se pechatat na sashtiya fizicheski list stikeri kato
    # Svobodna prodazhba - ako oshte nyama otdelno zapazena NZOK kalibratsia,
    # tragvame ot veche kalibriranata geometria na Svobodna (ne ot goli
    # podrazbirashti se stойности), za da ne se preправя ot nulata.
    if kind == "nzok" and "label_layout_nzok" not in cfg and "label_layout_otc" in cfg:
        layout.update(cfg["label_layout_otc"])
        # NZOK etiketite imat 3 reda tekst (kod/izchislenie/tsena) - kodat
        # trqbva da e po-goljam ot kalibrirania za Svobodna (koyato ima samo
        # 2 reda), zatova pri parvoto avtomatichno prehvarlyane go uvelichavame.
        layout["code_font_size"] = layout.get("code_font_size", DEFAULT_LABEL_LAYOUT["code_font_size"]) + 2.0
    layout.update(cfg.get(_label_layout_cfg_key(kind), {}))
    return layout


def save_label_layout(layout, kind="otc"):
    cfg = load_config()
    cfg[_label_layout_cfg_key(kind)] = layout
    save_config(cfg)


def load_printer_name():
    return load_config().get("printer_name") or None


def save_printer_name(name):
    cfg = load_config()
    cfg["printer_name"] = name
    save_config(cfg)


def list_printers():
    """Vrushta spisak s imenata na instaliranite printeri (cherez PowerShell
    Get-Printer - po-nadezhdno na noви Windows versii ot staroto wmic)."""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "(Get-Printer).Name"],
            capture_output=True, text=True, timeout=10, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        names = [line.strip() for line in result.stdout.splitlines() if line.strip()]
        return names
    except Exception:
        return []

# Ofitsialen fiksiran kurs lv./EUR - samo za proverka (sanity check),
# realnite tsени se cheтat direktno ot faila, ne se presmyatat.
BGN_PER_EUR = 1.95583

NUMBER_KEYWORDS = ["код", "kod", "номер", "nomer", "артикул", "artikul"]
EUR_KEYWORDS = ["eur", "€", "евро", "evro"]
BGN_KEYWORDS = ["лв", "lv", "лева", "bgn"]
QTY_KEYWORDS = ["кол", "kol", "брой", "broy", "quantity", "qty"]
NO_QTY_LABEL = "— без (по 1 бр. на артикул) —"


# ---------- Chetene na failove ----------

def _load_csv_rows(path):
    encodings = ["utf-8-sig", "utf-8", "cp1251"]
    last_error = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                sample = f.read(4096)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
                except csv.Error:
                    class _Fallback(csv.excel):
                        delimiter = ";" if sample.count(";") > sample.count(",") else ","
                    dialect = _Fallback
                reader = csv.reader(f, dialect)
                return [tuple(row) for row in reader]
        except UnicodeDecodeError as e:
            last_error = e
            continue
    raise last_error or ValueError("Failat e prazen ili neprochitaem.")


# Skladovi razpiski (napr. ot Kafarma/Sofarma) chesto nyamat realni linii mezhdu
# redovete s produkti - samo poredeni koloni. Tablichnoto razpoznavane na
# PyMuPDF/pdfplumber vizhda samo zaglavniq red v takiv sluchay. Zatova parsvame
# vseki red s regex, koyto poznava tochniya red na kolonite v tozi format:
# Kod  Naimenovanie+Merka  Kol.  Ed.TSE  TO%  Ed.pok.tsena  Stoynost  TN%  PH%  Prod.tsena(EUR)  Prod.tsena lv.(BGN)  Partida  Srok  Raft
# (Merka ne se izdelya otdelno - mozhe da e "Опак."/"Бр." no sushto taka proizvolen
# tekst kato "EYE DROPS, SOLUTION" - zatova ostava chast ot svobodniya blok "rest",
# a granitsata se namira ot regex-a chrez zadulzhitelnata struktura otdyasno.)
_INVOICE_LINE_RE = re.compile(
    r"^(?P<code>\S+)\s+"
    r"(?P<rest>.+?)\s+"
    r"(?P<qty>\d+)\s+"
    r"(?P<f1>-?[\d,]+)\s+"
    r"(?P<f2>-?[\d,]+)\s+"
    r"(?P<f3>-?[\d,]+)\s+"
    r"(?P<f4>-?[\d,]+)\s+"
    r"(?P<f5>-?[\d,]+)\s+"
    r"(?P<f6>-?[\d,]+)\s+"
    r"(?P<eur>-?[\d,]+)\s+"
    r"(?P<bgn>-?[\d,]+)\s*"
    r"(?P<partida>\S+)\s+"
    r"(?P<date>\d{1,2}\.\d{1,2}\.\d{4})\s*г\.?\s*"
    r"(?P<raft>\S*)\s*$"
)


def _parse_invoice_text_rows(text):
    """Vrushta grid s poznati zaglaviq (Kod/Naimenovanie/Kol./Tsena EUR/Tsena lv.)
    izvlecheni ot obikn текст na PDF (bez tablichni linii), red po red."""
    header = ("Код", "Наименование", "Кол.", "Цена EUR", "Цена лв.")
    rows = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        m = _INVOICE_LINE_RE.match(line)
        if not m:
            continue
        rows.append((
            m.group("code"),
            m.group("rest").strip(),
            m.group("qty"),
            m.group("eur"),
            m.group("bgn"),
        ))
    return [header] + rows


def _reconstruct_visual_lines(page, y_tol=3.0):
    """PDF-tekstoviyat potok chesto ne e v chetivniya (vizualen) red - vseki
    klix/kletka moje da e otdelen 'run' bez garantiran red. Zatova vzimame
    dumite s tehnite koordinati i gi gruparame po red (blizka Y-koordinata),
    sled tova gi nareждаме otляvo nadyasno (po X) - taka vuzstanovyavame
    istinskiya vizualen red na vseki red ot tablitsata."""
    words = page.get_text("words")  # (x0, y0, x1, y1, word, block_no, line_no, word_no)
    if not words:
        return []
    words = sorted(words, key=lambda w: (round(w[1], 1), w[0]))

    lines = []
    current = [words[0]]
    current_y = words[0][1]
    for w in words[1:]:
        if abs(w[1] - current_y) <= y_tol:
            current.append(w)
            current_y = (current_y + w[1]) / 2
        else:
            lines.append(current)
            current = [w]
            current_y = w[1]
    lines.append(current)

    result = []
    for group in lines:
        group_sorted = sorted(group, key=lambda w: w[0])
        result.append(" ".join(w[4] for w in group_sorted))
    return result


def _load_pdf_rows(path):
    import pymupdf

    doc = pymupdf.open(path)
    try:
        all_lines = []
        for page in doc:
            all_lines.extend(_reconstruct_visual_lines(page))
    finally:
        doc.close()

    grid = _parse_invoice_text_rows("\n".join(all_lines))
    if len(grid) <= 1:
        raise ValueError(
            "Не успях да разпозная редовете в този PDF по познатия формат на складова "
            "разписка.\nОпитай с Excel/CSV версия на документа, или прати ми файла да го огледам."
        )
    return grid


def _excel_cell_text(value):
    """Kletka ot Excel kato tekst vuv formata na PDF-a (5.4 -> '5,4', 2.0 -> '2'),
    za da mozhe redat da se parsne sus sashtiya _INVOICE_LINE_RE."""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(round(value, 4)).replace(".", ",")
    return str(value).strip()


def _normalize_invoice_excel(grid):
    """Skladovite razpiski na Kafarma kato .xls/.xlsx imat zaglavieto razdeleno
    na 3 reda i izmesteno sprямo dannite (slyati kletki) - guess_header_row
    vzimashe purviya red s lekarstvo za zaglavie i kolonata za tsena v €
    ne se poznavashe (izbirashe se proizvolna kolona -> greshna kalkulatsiya v
    NZOK). Zatova sglobyavame vseki red kato tekst i go parsvame kato PDF-a.
    Vrushta None, ako failat ne e v tozi format."""
    lines = []
    for row in grid:
        cells = [_excel_cell_text(v) for v in row if v not in (None, "")]
        if cells:
            lines.append(" ".join(c for c in cells if c))
    normalized = _parse_invoice_text_rows("\n".join(lines))
    return normalized if len(normalized) > 1 else None


def load_table(path):
    """Vrushta pulnata reshetka ot redove (tuples), bez da predpolaga koy red e zaglaven."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        grid = list(ws.iter_rows(values_only=True))
        return _normalize_invoice_excel(grid) or grid
    if ext == ".xls":
        import xlrd
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        grid = [tuple(sheet.row_values(r)) for r in range(sheet.nrows)]
        return _normalize_invoice_excel(grid) or grid
    if ext == ".csv":
        return _load_csv_rows(path)
    if ext == ".pdf":
        return _load_pdf_rows(path)
    raise ValueError(f"Неподдържан тип файл: {ext}")


def guess_header_row(grid, max_scan=25):
    """Fakturi/skladovi raztiski chesto imat neskolko reda s obshta informatsiq
    (firma, nomer, data) predi realniq red sus zaglaviq na kolonite. Praviloto:
    zaglaviq red obiknoveno ima nay-mnogo popalneni кletki."""
    best_idx, best_count = 0, -1
    for i, row in enumerate(grid[:max_scan]):
        count = sum(1 for v in row if v not in (None, ""))
        if count > best_count:
            best_count = count
            best_idx = i
    return best_idx


def guess_column(headers, keywords, exclude=()):
    for h in headers:
        if h in exclude:
            continue
        low = h.lower()
        if any(k in low for k in keywords):
            return h
    return None


# ---------- Parsvane na stойности ----------

def clean_code(value):
    """Izvlicha koda na produkta - zapazva bukvi i tsifri (NZOK kodove kato
    'NG050' sa bukvено-tsifreni, ne samo tsifri), stripva interpunktsiya/intervali."""
    if value is None:
        return None
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    code = re.sub(r"[^0-9A-Za-zА-Яа-я]", "", text)
    return code.upper() if code else None


def clean_price(value):
    """Parsva parichna stoynost ot razlichni formati (5.99 / 5,99 / '5.99 lv' / 5.99€)."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    text = re.sub(r"[^0-9,\.]", "", text)
    if not text:
        return None
    if "," in text and "." in text:
        text = text.replace(",", "")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def clean_quantity(value):
    """Parsva broyka na etiketi za edin red (zaokrugleno do tsyalo chislo, minimum 1)."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        q = value
    else:
        text = re.sub(r"[^0-9,\.]", "", str(value).strip())
        if not text:
            return None
        try:
            q = float(text.replace(",", "."))
        except ValueError:
            return None
    q = int(round(q))
    return q if q > 0 else None


def format_label_price(eur, bgn=None):
    eur_text = f"{eur:.2f}€"
    bgn_text = f"{bgn:.2f}лв." if bgn is not None else None
    return eur_text, bgn_text


# ---------- Generirane na PDF ----------

def _fit_price_fonts(eur_text, bgn_text, max_width, base_eur=10, base_bgn=7):
    scale = 1.0
    while scale > 0.4:
        fe = max(6.0, base_eur * scale)
        fb = max(5.0, base_bgn * scale) if bgn_text else 0
        ew = stringWidth(eur_text, FONT_REGULAR, fe)
        bw = stringWidth(bgn_text, FONT_REGULAR, fb) if bgn_text else 0
        if ew + bw <= max_width:
            return fe, fb, ew, bw
        scale -= 0.05
    return fe, fb, ew, bw


def generate_calibration_grid_pdf(output_path, layout=None):
    """Diagnostichen list - chertae ramka (bez zapulvane) i nomer red/kolona na
    vsяka ot 160-te pozitsii spored podadenata (ili zapazenata) kalibratsia.
    Sluzhi za ednokratno kalibrirane sreshtu realen list stikeri - otpechatva
    se direktno varhu tyah i se sravnyava ramkata s realnite rezani kvadratcheta."""
    layout = layout or load_label_layout()
    lbl_w = layout["w"] * mm
    lbl_h = layout["h"] * mm
    margin_left = layout["margin_left"] * mm
    margin_top = layout["margin_top"] * mm
    pitch_x = lbl_w + layout["gap_h"] * mm
    pitch_y = lbl_h + layout["gap_v"] * mm
    rotation_deg = layout.get("rotation_deg", 0.0)
    scale_factor = layout.get("scale_pct", 100.0) / 100.0
    pivot_x, pivot_y = margin_left, PAGE_H - margin_top

    c = canvas.Canvas(output_path, pagesize=A4)
    c.setLineWidth(0.3)
    c.saveState()
    c.translate(pivot_x, pivot_y)
    c.rotate(rotation_deg)
    c.scale(scale_factor, scale_factor)
    c.translate(-pivot_x, -pivot_y)
    for row in range(LBL_ROWS):
        for col in range(LBL_COLS):
            x = margin_left + col * pitch_x
            y = PAGE_H - margin_top - row * pitch_y - lbl_h
            c.rect(x, y, lbl_w, lbl_h, stroke=1, fill=0)
            c.setFont(FONT_REGULAR, 5)
            c.drawString(x + 0.5 * mm, y + lbl_h - 2 * mm, f"{row + 1}.{col + 1}")
    c.restoreState()
    c.showPage()
    c.save()


def generate_labels_pdf(labels, output_path, layout=None, start_row=1, start_col=1):
    """labels: list na (nomer:str, eur:float, bgn:float ili None) - edin zapis
    na etiket (ako artikul ima kolichestvo > 1, se povtarq neskolko puti v
    spisaka predi da se povika tazi funktsia). Ako bgn e None, na etiketa
    izlicha samo tsenata v €. start_row/start_col (1-baznyi) pozvolyavat da se
    zapochne pechatat ot sredata na veche zapochnat list stikeri (napr. ako
    parvite neskolko reda veche sa izpolzvani ot predishen pechat)."""
    layout = layout or load_label_layout()
    lbl_w = layout["w"] * mm
    lbl_h = layout["h"] * mm
    margin_left = layout["margin_left"] * mm
    margin_top = layout["margin_top"] * mm
    pitch_x = lbl_w + layout["gap_h"] * mm
    pitch_y = lbl_h + layout["gap_v"] * mm
    rotation_deg = layout.get("rotation_deg", 0.0)
    scale_factor = layout.get("scale_pct", 100.0) / 100.0
    pivot_x, pivot_y = margin_left, PAGE_H - margin_top

    c = canvas.Canvas(output_path, pagesize=A4)
    per_page = LBL_COLS * LBL_ROWS
    first_offset = max(0, min(per_page - 1, (max(1, int(start_row)) - 1) * LBL_COLS + (max(1, int(start_col)) - 1)))

    idx = 0
    offset = first_offset
    while idx < len(labels):
        capacity = per_page - offset
        page_labels = labels[idx:idx + capacity]
        c.saveState()
        c.translate(pivot_x, pivot_y)
        c.rotate(rotation_deg)
        c.scale(scale_factor, scale_factor)
        c.translate(-pivot_x, -pivot_y)
        for j, (number, eur, bgn) in enumerate(page_labels):
            i = offset + j
            row = i // LBL_COLS
            col = i % LBL_COLS
            x = margin_left + col * pitch_x
            y = PAGE_H - margin_top - row * pitch_y - lbl_h
            cx = x + lbl_w / 2

            max_width = lbl_w - 1.5 * mm

            fn = layout.get("code_font_size", 8.0)
            while stringWidth(number, FONT_REGULAR, fn) > max_width and fn > 5.0:
                fn -= 0.3
            c.setFont(FONT_REGULAR, fn)
            c.drawCentredString(cx, y + lbl_h * 0.62, number)

            eur_text, bgn_text = format_label_price(eur, bgn)

            if bgn_text:
                fe, fb, ew, bw = _fit_price_fonts(eur_text, bgn_text, max_width)
                start_x = cx - (ew + bw) / 2
                price_y = y + lbl_h * 0.18
                c.setFont(FONT_REGULAR, fe)
                c.drawString(start_x, price_y, eur_text)
                c.setFont(FONT_REGULAR, fb)
                c.drawString(start_x + ew, price_y, bgn_text)
            else:
                fe = 10.0
                while stringWidth(eur_text, FONT_REGULAR, fe) > max_width and fe > 6.0:
                    fe -= 0.5
                c.setFont(FONT_REGULAR, fe)
                c.drawCentredString(cx, y + lbl_h * 0.18, eur_text)

        c.restoreState()
        c.showPage()
        idx += capacity
        offset = 0

    c.save()


def generate_nzok_labels_pdf(labels, output_path, layout=None, start_row=1, start_col=1):
    """labels: list na (kod:str, pulna_tsena_eur:float, reimbursatsiya_eur:float).
    Etiket na 3 reda: kod / 'pulna - reimb' / krayna tsena v € i lv. (izchislena
    ot fiksirania kurs, tui kato nyama otdelna kolona za nea v faktura).
    Izpolzva sashtata kalibrirana mrezha (razmeri na stikerite + zavartane) kato
    generate_labels_pdf, za da suvpadat i dvata vida etiketi na edin i sushti
    fizicheski list."""
    layout = layout or load_label_layout("nzok")
    lbl_w = layout["w"] * mm
    lbl_h = layout["h"] * mm
    margin_left = layout["margin_left"] * mm
    margin_top = layout["margin_top"] * mm
    pitch_x = lbl_w + layout["gap_h"] * mm
    pitch_y = lbl_h + layout["gap_v"] * mm
    rotation_deg = layout.get("rotation_deg", 0.0)
    scale_factor = layout.get("scale_pct", 100.0) / 100.0
    pivot_x, pivot_y = margin_left, PAGE_H - margin_top

    c = canvas.Canvas(output_path, pagesize=A4)
    per_page = LBL_COLS * LBL_ROWS
    first_offset = max(0, min(per_page - 1, (max(1, int(start_row)) - 1) * LBL_COLS + (max(1, int(start_col)) - 1)))

    idx = 0
    offset = first_offset
    while idx < len(labels):
        capacity = per_page - offset
        page_labels = labels[idx:idx + capacity]
        c.saveState()
        c.translate(pivot_x, pivot_y)
        c.rotate(rotation_deg)
        c.scale(scale_factor, scale_factor)
        c.translate(-pivot_x, -pivot_y)
        for j, (code, full_eur, reimb_eur) in enumerate(page_labels):
            i = offset + j
            row = i // LBL_COLS
            col = i % LBL_COLS
            x = margin_left + col * pitch_x
            y = PAGE_H - margin_top - row * pitch_y - lbl_h
            cx = x + lbl_w / 2
            max_width = lbl_w - 1.2 * mm

            final_eur = max(0.0, full_eur - reimb_eur)

            # Trite reda sa nareждани blizko edin do drug (kodat e nay-golyam i
            # nay-vazhen, ostanalite dva sa po-malki no vidimi) - fraktsiite na
            # lbl_h sa naroчno sblizeni v sredata na etiketa.
            fn = layout.get("code_font_size", 8.5)
            while stringWidth(code, FONT_REGULAR, fn) > max_width and fn > 4.0:
                fn -= 0.3
            c.setFont(FONT_REGULAR, fn)
            c.drawCentredString(cx, y + lbl_h * 0.66, code)

            calc_text = f"{full_eur:.2f} - {reimb_eur:.2f}"
            fc = 5.0
            while stringWidth(calc_text, FONT_REGULAR, fc) > max_width and fc > 3.5:
                fc -= 0.3
            c.setFont(FONT_REGULAR, fc)
            c.drawCentredString(cx, y + lbl_h * 0.46, calc_text)

            eur_text, _ = format_label_price(final_eur)
            fe = 6.5
            while stringWidth(eur_text, FONT_REGULAR, fe) > max_width and fe > 4.0:
                fe -= 0.3
            price_y = y + lbl_h * 0.27
            c.setFont(FONT_REGULAR, fe)
            c.drawCentredString(cx, price_y, eur_text)

        c.restoreState()
        c.showPage()
        idx += capacity
        offset = 0

    c.save()


def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------- Pechat na PDF (direktno, bez brauzar/viewer) ----------

SUMATRA_PATH = os.path.join(_app_dir(), "SumatraPDF.exe")


def print_pdf_file(path):
    """Praшta PDF-a direktno kum printera po podrazbirane na tochno 100%
    mащab (bez fit-to-page/avtomatichno smalyavane), zaobikalyaiki
    brauzъra ili PDF programata po podrazbirane na kompyutъra - imenno
    tehnite razlichni nastroyki za pechat sa prichinata etiketite da
    izlizat razlichno vseki path pri raznite kompyutri. Ako SumatraPDF.exe
    ne e naliчen do programata, se vrashta kum obiknoveniya nachin."""
    if os.path.exists(SUMATRA_PATH):
        printer_name = load_printer_name()
        printer_args = ["-print-to", printer_name] if printer_name else ["-print-to-default"]
        subprocess.run(
            [
                SUMATRA_PATH,
                *printer_args,
                "-print-settings", "noscale",
                "-silent",
                "-exit-when-done",
                path,
            ],
            check=True,
            timeout=30,
        )
    else:
        os.startfile(path, "print")


# ---------- NZOK reimbursatsii (avtomatichno svalyane ot nhif.bg) ----------

NZOK_UPDATES_PAGE = "https://www.nhif.bg/bg/software/drugstores/updates/{year}"
REIMB_CACHE_PATH = os.path.join(_app_dir(), "nzok_reimbursements.json")


def load_reimbursement_cache():
    """Vrushta (tablitsa {kod: reimbursatsiya_na_opakovka_v_EUR}, mesets_kato_'MM.YYYY')."""
    if os.path.exists(REIMB_CACHE_PATH):
        try:
            with open(REIMB_CACHE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("table", {}), data.get("month")
        except Exception:
            return {}, None
    return {}, None


def save_reimbursement_cache(table, month_label):
    try:
        prev_period = None
        if os.path.exists(REIMB_CACHE_PATH):
            try:
                with open(REIMB_CACHE_PATH, "r", encoding="utf-8") as f:
                    prev_period = json.load(f).get("last_check_period")
            except Exception:
                prev_period = None
        with open(REIMB_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "table": table,
                    "month": month_label,
                    "last_check_period": prev_period,
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
    except Exception:
        pass


def _current_reimb_period():
    """Perioda za avtomatichna proverka - smenya se na 1-vo i 16-o chislo na meseca."""
    today = datetime.date.today()
    half = "A" if today.day < 16 else "B"
    return f"{today.year}-{today.month:02d}-{half}"


def get_last_checked_period():
    if os.path.exists(REIMB_CACHE_PATH):
        try:
            with open(REIMB_CACHE_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get("last_check_period")
        except Exception:
            return None
    return None


def mark_reimb_period_checked():
    """Zapisva, che avtomatichnata proverka za tekushtiya period (1-vo/16-o) e napravena,
    bez da promenya tablitsata s reimbursatsii (izpolzva se i pri neuspeshen fetch,
    za da ne se opitva na vseki start bez internet)."""
    try:
        data = {}
        if os.path.exists(REIMB_CACHE_PATH):
            try:
                with open(REIMB_CACHE_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
        data["last_check_period"] = _current_reimb_period()
        with open(REIMB_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _http_session():
    import requests
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
        "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.8",
    })
    return session


def fetch_nzok_reimbursement_table():
    """Svalya nay-noviya arhiv 'Update Pharma' (nomenklaturata za aptechni
    programi, publikuvana na 1-vo i 16-o chislo na vseki mesets - t.e. tochno
    navreme, za razlika ot staria istochnik 'Spravka 1_PLS1', koyto izlizashe
    s meseци zakasnenie, tui kato e otcheten spisak na realno dispensiranite
    opakovki). Vzima kolona 15 (indeks 14) ot .NDU faila vutre v zip-a - tova e
    referentnata reimbursatsionna stoynost na opakovka v EUR po NZOK kod
    (proverено sreshtu staria istochnik - suvpada tochno za ~80% ot kodovete i
    e blizko za oshte ~12%, s razlika samo pri lekarstva s razlichen % pokritie
    po diagnoza).

    Vrushta (tablitsa, mesets_label, url). Hvurlya izklyuchenie pri greshka -
    povikvashtiyat trqbva da go pokazhe na potrebitelya (nyama internet, sайtat
    se e promenil i t.n.)."""
    import re as _re
    import zipfile
    import io
    from urllib.parse import urljoin

    session = _http_session()
    today = datetime.date.today()

    matches = []
    page_url = None
    for year in (today.year, today.year - 1):
        page_url = NZOK_UPDATES_PAGE.format(year=year)
        r = session.get(page_url, timeout=20)
        r.raise_for_status()
        matches = _re.findall(r'href="(/upload/\d+/Update-apteki-[^"]*?\.zip)"', r.text)
        if matches:
            break

    if not matches:
        raise ValueError(
            "Не намерих линк към списъка \"Update Pharma\" на nhif.bg - страницата може да се е променила."
        )
    file_url = urljoin("https://www.nhif.bg", matches[0])

    r2 = session.get(file_url, timeout=60)
    r2.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(r2.content)) as z:
        ndu_names = [n for n in z.namelist() if n.upper().endswith(".NDU")]
        if not ndu_names:
            raise ValueError("Архивът от НЗОК не съдържа файл с номенклатурата (.NDU) - форматът може да се е променил.")
        content = z.read(ndu_names[0]).decode("cp1251", errors="ignore")

    table = {}
    for line in content.splitlines():
        parts = line.split("|")
        if len(parts) < 19:
            continue
        code = parts[2].strip().upper()
        if not code:
            continue
        try:
            reimb = float(parts[14])
        except ValueError:
            continue
        table[code] = round(reimb, 2)

    if not table:
        raise ValueError("Не разпознах структурата на файла от НЗОК (форматът се е променил).")

    date_match = _re.search(r"(\d{2})-(\d{2})-(\d{4})", matches[0])
    month_label = f"{date_match.group(2)}.{date_match.group(3)}" if date_match else ndu_names[0]
    return table, month_label, file_url


# ---------- Avtomatichno obnovyavane (ot GitHub Releases) ----------

APP_VERSION = "1.1.0"
GITHUB_REPO = "Velislav1/etiketi-apteka"
UPDATE_ASSET_NAME = "EtiketiApteka.zip"
# Failove na potrebitelya, koito NE se prezapisvat pri obnovyavane.
UPDATE_KEEP_FILES = ("config.json", "nzok_reimbursements.json")


def _version_tuple(text):
    return tuple(int(p) for p in re.findall(r"\d+", text or "")[:3])


def check_for_update():
    """Vrushta (versiya, url_na_zip) ako v GitHub ima po-nova versiya, inache None."""
    session = _http_session()
    r = session.get(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest", timeout=10)
    r.raise_for_status()
    data = r.json()
    tag = data.get("tag_name", "")
    if _version_tuple(tag) <= _version_tuple(APP_VERSION):
        return None
    for asset in data.get("assets", []):
        if asset.get("name") == UPDATE_ASSET_NAME:
            return tag.lstrip("v"), asset["browser_download_url"]
    return None


def download_and_stage_update(zip_url):
    """Svalya i razarhivira novata versiya vuv vremenna papka i podgotvya
    .bat, koyto izchakva programata da se zatvori, kopira novite failove
    varhu tekushtata papka (bez config.json / kesha na NZOK) i ya pusка pak.
    Vrushta patya do .bat-a."""
    import tempfile
    import zipfile
    import io

    session = _http_session()
    r = session.get(zip_url, timeout=300)
    r.raise_for_status()

    stage_dir = tempfile.mkdtemp(prefix="etiketi_update_")
    with zipfile.ZipFile(io.BytesIO(r.content)) as z:
        z.extractall(stage_dir)
    src = stage_dir
    entries = os.listdir(stage_dir)
    if len(entries) == 1 and os.path.isdir(os.path.join(stage_dir, entries[0])):
        src = os.path.join(stage_dir, entries[0])
    if not os.path.exists(os.path.join(src, "EtiketiApteka.exe")):
        raise ValueError("Архивът с новата версия не съдържа EtiketiApteka.exe.")

    for name in UPDATE_KEEP_FILES:
        p = os.path.join(src, name)
        if os.path.exists(p) and os.path.exists(os.path.join(_app_dir(), name)):
            os.remove(p)

    app_dir = _app_dir()
    exe_path = os.path.join(app_dir, "EtiketiApteka.exe")
    bat_path = os.path.join(stage_dir, "obnovi.bat")
    # UTF-8 + chcp 65001 - inache kirilitsa v patya (napr. ime na potrebitel) se chupi.
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(
            "@echo off\r\n"
            "chcp 65001 >nul\r\n"
            "echo Obnovyavam programata, molya izchakay...\r\n"
            ":wait\r\n"
            f'tasklist /FI "PID eq {os.getpid()}" | find "{os.getpid()}" >nul\r\n'
            "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)\r\n"
            f'robocopy "{src}" "{app_dir}" /E /NFL /NDL /NJH /NJS /NP >nul\r\n'
            f'start "" "{exe_path}"\r\n'
        )
    return bat_path


# ---------- Interfeys ----------

BaseTk = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title(f"Етикети за аптека  v{APP_VERSION}")
        self.geometry("1020x720")
        self.minsize(860, 600)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

        self.cfg = load_config()

        self.manual_items = []
        self.manual_last_output_path = None

        self._setup_style()
        self._build_ui()

        # Obnovyavaneto raboti samo za .exe versiyata (pri main.py nyama kakvo da se zameni).
        if getattr(sys, "frozen", False):
            self.after(1500, self._start_update_check)

    def _start_update_check(self):
        def worker():
            try:
                found = check_for_update()
            except Exception:
                found = None
            if found:
                self.after(0, lambda: self._offer_update(*found))

        threading.Thread(target=worker, daemon=True).start()

    def _offer_update(self, version, zip_url):
        if not messagebox.askyesno(
            "Нова версия",
            f"Има нова версия на програмата ({version}).\n"
            f"В момента е инсталирана {APP_VERSION}.\n\n"
            "Да я изтегля и инсталирам ли сега? Програмата ще се рестартира сама.",
        ):
            return
        self.config(cursor="watch")

        def worker():
            try:
                bat = download_and_stage_update(zip_url)
                error = None
            except Exception as e:
                bat, error = None, str(e)
            self.after(0, lambda: self._finish_update(bat, error))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_update(self, bat_path, error):
        self.config(cursor="")
        if error:
            messagebox.showerror("Грешка при обновяване", f"Не можах да изтегля новата версия:\n{error}")
            return
        subprocess.Popen(
            ["cmd", "/c", bat_path],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        self.destroy()

    def _setup_style(self):
        self.configure(background=BG)
        self.option_add("*TCombobox*Listbox.background", BG_FIELD)
        self.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", GREEN)
        self.option_add("*TCombobox*Listbox.selectForeground", "#04140a")
        self.option_add("*TCombobox*Listbox.font", ("Segoe UI", 10))

        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", background=BG, foreground=TEXT, font=("Segoe UI", 11))
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("TCheckbutton", background=BG, foreground=TEXT)
        style.map(
            "TCheckbutton",
            background=[("active", BG)],
            foreground=[("active", GOLD_TEXT)],
            indicatorcolor=[("selected", GREEN), ("!selected", BG_FIELD)],
        )

        # Gornata chast (logo/zaglavie) - ostava cherna kato loгото.
        style.configure("Header.TFrame", background=HEADER_BG)

        for lf_style in ("TLabelframe", "Step.TLabelframe"):
            style.configure(
                lf_style, background=BG, bordercolor=BORDER, darkcolor=BG, lightcolor=BG,
                relief="solid", borderwidth=1,
            )
        for lf_label_style in ("TLabelframe.Label", "Step.TLabelframe.Label"):
            style.configure(lf_label_style, background=BG, foreground=GOLD_TEXT, font=("Segoe UI", 12, "bold"))

        style.configure("TNotebook", background=HEADER_BG, bordercolor=HEADER_BG, tabmargins=(6, 6, 6, 0))
        style.configure(
            "TNotebook.Tab", background=TAB_INACTIVE_BG, foreground="#9a9c8f",
            padding=(20, 10), font=("Segoe UI", 11, "bold"), borderwidth=0,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", BG)],
            foreground=[("selected", GOLD_TEXT)],
        )

        style.configure(
            "TButton", background=BG_PANEL, foreground=GOLD_TEXT, bordercolor=GOLD,
            focuscolor=BG, padding=8, font=("Segoe UI", 10, "bold"), relief="flat",
        )
        style.map(
            "TButton",
            background=[("active", GREEN), ("disabled", BG_PANEL)],
            foreground=[("active", "#04140a"), ("disabled", TEXT_MUTED)],
            bordercolor=[("disabled", BORDER)],
        )

        style.configure(
            "Big.TButton", font=("Segoe UI", 12, "bold"), padding=10,
            background=GOLD, foreground=HEADER_BG, bordercolor=GOLD, relief="flat",
        )
        style.map(
            "Big.TButton",
            background=[("active", GREEN_BRIGHT), ("disabled", BG_PANEL)],
            foreground=[("active", "#04140a"), ("disabled", TEXT_MUTED)],
        )

        style.configure(
            "TCombobox", fieldbackground=BG_FIELD, background=BG_FIELD, foreground=TEXT,
            arrowcolor=GOLD_TEXT, bordercolor=BORDER, selectbackground=BG_FIELD, selectforeground=TEXT,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", BG_FIELD), ("disabled", BG_PANEL)],
            foreground=[("readonly", TEXT), ("disabled", TEXT_MUTED)],
            bordercolor=[("focus", GOLD)],
        )

        style.configure("TEntry", fieldbackground=BG_FIELD, foreground=TEXT, insertcolor=TEXT, bordercolor=BORDER)
        style.map("TEntry", bordercolor=[("focus", GOLD)])

        style.configure(
            "TSpinbox", fieldbackground=BG_FIELD, foreground=TEXT, background=BG_FIELD,
            arrowcolor=GOLD_TEXT, bordercolor=BORDER,
        )
        style.map("TSpinbox", bordercolor=[("focus", GOLD)])

        style.configure(
            "Treeview", background=BG_FIELD, fieldbackground=BG_FIELD, foreground=TEXT,
            bordercolor=BORDER, borderwidth=1, rowheight=24,
        )
        style.configure(
            "Treeview.Heading", background=BG_PANEL, foreground=GOLD_TEXT,
            font=("Segoe UI", 10, "bold"), relief="flat",
        )
        style.map("Treeview.Heading", background=[("active", BG_PANEL)])
        style.map("Treeview", background=[("selected", GREEN)], foreground=[("selected", "#04140a")])

        style.configure("Vertical.TScrollbar", background=BG_PANEL, troughcolor=BG, arrowcolor=GOLD_TEXT, bordercolor=BG)
        style.map("Vertical.TScrollbar", background=[("active", GREEN)])

        style.configure("Header.TLabel", background=HEADER_BG, foreground=GOLD_BRIGHT, font=("Georgia", 28, "bold"))
        style.configure("Brand.TLabel", background=HEADER_BG, foreground=GREEN_BRIGHT, font=("Segoe UI", 12, "bold"))
        style.configure("Status.TLabel", background=BG, foreground=TEXT_MUTED, font=("Segoe UI", 10))
        style.configure("Preview.TLabel", background=BG, foreground=GREEN_TEXT, font=("Consolas", 12, "bold"))
        style.configure("Warn.TLabel", background=BG, foreground=DANGER, font=("Segoe UI", 10, "bold"))

    def _load_logo(self):
        self._icon_photo = None
        self._logo_photo = None
        try:
            from PIL import Image, ImageTk
            img = Image.open(LOGO_PATH)

            icon_img = img.resize((256, 256), Image.LANCZOS)
            self._icon_photo = ImageTk.PhotoImage(icon_img)
            self.iconphoto(True, self._icon_photo)

            header_img = img.resize((150, 150), Image.LANCZOS)
            self._logo_photo = ImageTk.PhotoImage(header_img)
        except Exception:
            self._icon_photo = None
            self._logo_photo = None

    def _make_scrollable(self, parent):
        """Obvivа sadarzhanieto na edin tab v vertikalno skrolируема oblast -
        na malki/nizki ekrani (napr. laptopi 1366x768) syderzhanieto na taba
        mozhe da e po-visoko ot vidimия prozorec; bez tova dolnata lenta s
        butonite (Генерирай/Отвори PDF/Принтирай) ostava nedostizhima."""
        canvas = tk.Canvas(parent, background=BG, highlightthickness=0, bd=0)
        vscroll = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(inner_id, width=event.width)

        inner.bind("<Configure>", _on_inner_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        return inner

    def _build_ui(self):
        self._load_logo()

        header_frame = ttk.Frame(self, style="Header.TFrame")
        header_frame.pack(fill="x")

        if self._logo_photo:
            logo_label = tk.Label(header_frame, image=self._logo_photo, bg=HEADER_BG, bd=0, highlightthickness=0)
            logo_label.pack(side="left", padx=(16, 14), pady=14)

        title_box = ttk.Frame(header_frame, style="Header.TFrame")
        title_box.pack(side="left", fill="y", pady=14)
        ttk.Label(title_box, text="Етикети за аптека", style="Header.TLabel").pack(anchor="w", pady=(46, 0))
        ttk.Label(title_box, text="VELISLAV CODE", style="Brand.TLabel").pack(anchor="w", pady=(4, 0))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        otc_tab_outer = ttk.Frame(self.notebook)
        nzok_tab_outer = ttk.Frame(self.notebook)
        manual_tab_outer = ttk.Frame(self.notebook)
        calib_tab_outer = ttk.Frame(self.notebook)
        self.notebook.add(otc_tab_outer, text="Свободна продажба")
        self.notebook.add(nzok_tab_outer, text="НЗОК")
        self.notebook.add(manual_tab_outer, text="Ръчно въвеждане")
        self.notebook.add(calib_tab_outer, text="Калибровка на етикетите")

        otc_tab = self._make_scrollable(otc_tab_outer)
        nzok_tab = self._make_scrollable(nzok_tab_outer)
        manual_tab = self._make_scrollable(manual_tab_outer)
        calib_tab = self._make_scrollable(calib_tab_outer)

        self.otc_panel = FilePanel(
            self, otc_tab, "otc",
            "Качи файл с доставка (свободна продажба) → провери колоните → генерирай етикети за печат.",
        )
        self.nzok_panel = NzokFilePanel(
            self, nzok_tab, "nzok",
            "Качи файл с доставка (НЗОК) → провери колоните → генерирай етикети за печат.",
        )
        self._build_manual_tab(manual_tab)
        self._build_calibration_tab(calib_tab)

    def _build_manual_tab(self, manual_tab):
        intro = ttk.Label(
            manual_tab,
            text="Въведи код и цена на ръка (напр. за допечатване на липсващ етикет) и колко броя ти трябват.",
            style="Status.TLabel",
        )
        intro.pack(anchor="w", padx=16, pady=(10, 10))

        kind_row = ttk.Frame(manual_tab)
        kind_row.pack(anchor="w", padx=16, pady=(0, 6), fill="x")
        ttk.Label(kind_row, text="Вид етикети:").pack(side="left", padx=(0, 8))
        self.manual_kind_var = tk.StringVar(value="Свободна продажба")
        self._manual_kind_map = {"Свободна продажба": "otc", "НЗОК": "nzok"}
        self._manual_current_kind_cache = "otc"
        manual_kind_combo = ttk.Combobox(
            kind_row, textvariable=self.manual_kind_var, state="readonly",
            values=list(self._manual_kind_map.keys()), width=20,
        )
        manual_kind_combo.pack(side="left")
        manual_kind_combo.bind("<<ComboboxSelected>>", lambda e: self._manual_kind_changed())

        form = ttk.LabelFrame(manual_tab, text="Добави артикул", style="Step.TLabelframe", padding=10)
        form.pack(fill="x", padx=16, pady=6)

        ttk.Label(form, text="Код:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=5)
        self.manual_code_var = tk.StringVar()
        code_entry = ttk.Entry(form, textvariable=self.manual_code_var, width=16)
        code_entry.grid(row=0, column=1, padx=(0, 20), pady=5)

        self.manual_price_label = ttk.Label(form, text="Цена (€):")
        self.manual_price_label.grid(row=0, column=2, sticky="w", padx=(0, 6), pady=5)
        self.manual_price_var = tk.StringVar()
        price_entry = ttk.Entry(form, textvariable=self.manual_price_var, width=12)
        price_entry.grid(row=0, column=3, padx=(0, 20), pady=5)

        self.manual_reimb_label = ttk.Label(form, text="Реимбурсация (€):")
        self.manual_reimb_var = tk.StringVar()
        self.manual_reimb_entry = ttk.Entry(form, textvariable=self.manual_reimb_var, width=12)

        ttk.Label(form, text="Брой етикети:").grid(row=0, column=6, sticky="w", padx=(0, 6), pady=5)
        self.manual_qty_var = tk.StringVar(value="1")
        qty_spin = ttk.Spinbox(form, from_=1, to=999, textvariable=self.manual_qty_var, width=6)
        qty_spin.grid(row=0, column=7, padx=(0, 20), pady=5)

        add_btn = ttk.Button(form, text="Добави", style="Big.TButton", command=self._manual_add)
        add_btn.grid(row=0, column=8, pady=5)

        # Enter minava kum sledvashtoto pole (a ne направo "Добави") - inache pri
        # NZOK, Enter sled tsenata skachashe napravo kum zapisvane predi
        # potrebitelyat da e vpisal reimbursatsiyata, i vinagi izlizashe
        # greshka "въведи валидна реимбурсация".
        code_entry.bind("<Return>", lambda e: price_entry.focus_set())
        price_entry.bind(
            "<Return>",
            lambda e: self.manual_reimb_entry.focus_set() if self._current_manual_kind() == "nzok" else qty_spin.focus_set(),
        )
        self.manual_reimb_entry.bind("<Return>", lambda e: qty_spin.focus_set())
        qty_spin.bind("<Return>", lambda e: self._manual_add())

        list_frame = ttk.LabelFrame(manual_tab, text="Списък за печат", style="Step.TLabelframe", padding=10)
        list_frame.pack(fill="x", padx=16, pady=6)

        tree_container = ttk.Frame(list_frame)
        tree_container.pack(fill="x")

        self.manual_tree = ttk.Treeview(tree_container, columns=("code", "price", "qty"), show="headings", height=5)
        self.manual_tree.pack(fill="x", expand=True, side="left")
        manual_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.manual_tree.yview)
        manual_scroll.pack(side="right", fill="y")
        self.manual_tree.configure(yscrollcommand=manual_scroll.set)

        self._manual_reconfigure_fields()

        list_buttons = ttk.Frame(list_frame)
        list_buttons.pack(fill="x", pady=(6, 0))
        ttk.Button(list_buttons, text="Изтрий избраното", command=self._manual_remove_selected).pack(side="left")
        ttk.Button(list_buttons, text="Изчисти всички", command=self._manual_clear).pack(side="left", padx=(8, 0))

        start_frame = ttk.LabelFrame(
            manual_tab, text="Начало на печат (ако листът вече е започнат)", style="Step.TLabelframe", padding=10
        )
        start_frame.pack(fill="x", padx=16, pady=6)
        ttk.Label(start_frame, text="Ред:").pack(side="left", padx=(0, 4))
        self.manual_start_row_var = tk.StringVar(value="1")
        ttk.Spinbox(start_frame, from_=1, to=LBL_ROWS, width=5, textvariable=self.manual_start_row_var).pack(side="left", padx=(0, 16))
        ttk.Label(start_frame, text="Колона:").pack(side="left", padx=(0, 4))
        self.manual_start_col_var = tk.StringVar(value="1")
        ttk.Spinbox(start_frame, from_=1, to=LBL_COLS, width=5, textvariable=self.manual_start_col_var).pack(side="left", padx=(0, 16))
        ttk.Button(
            start_frame, text="Нулирай (1 / 1)",
            command=lambda: (self.manual_start_row_var.set("1"), self.manual_start_col_var.set("1")),
        ).pack(side="left")

        bottom = ttk.Frame(manual_tab, padding=(16, 10))
        bottom.pack(fill="x")

        self.manual_status_label = ttk.Label(bottom, text="", style="Status.TLabel")
        self.manual_status_label.pack(side="left")

        manual_button_bar = ttk.Frame(bottom)
        manual_button_bar.pack(side="right")

        self.manual_print_btn = ttk.Button(
            manual_button_bar, text="Принтирай", command=self._manual_print_last, state="disabled"
        )
        self.manual_print_btn.pack(side="right", padx=(8, 0))

        self.manual_open_btn = ttk.Button(
            manual_button_bar, text="Отвори PDF", command=self._manual_open_last, state="disabled"
        )
        self.manual_open_btn.pack(side="right", padx=(8, 0))

        ttk.Button(
            manual_button_bar, text="Генерирай етикети (PDF)", style="Big.TButton", command=self._manual_generate
        ).pack(side="right")

    # ---------- Ръчно въвеждане ----------

    def _current_manual_kind(self):
        return self._manual_kind_map.get(self.manual_kind_var.get(), "otc")

    def _manual_reconfigure_fields(self):
        kind = self._current_manual_kind()
        if kind == "nzok":
            self.manual_price_label.config(text="Пълна цена (€):")
            self.manual_reimb_label.grid(row=0, column=4, sticky="w", padx=(0, 6), pady=5)
            self.manual_reimb_entry.grid(row=0, column=5, padx=(0, 20), pady=5)
            self.manual_tree.configure(columns=("code", "price", "reimb", "final", "qty"))
            self.manual_tree.heading("code", text="Код")
            self.manual_tree.heading("price", text="Пълна цена (€)")
            self.manual_tree.heading("reimb", text="Реимбурсация (€)")
            self.manual_tree.heading("final", text="Крайна цена (€)")
            self.manual_tree.heading("qty", text="Брой етикети")
            for col, w in (("code", 100), ("price", 110), ("reimb", 130), ("final", 120), ("qty", 110)):
                self.manual_tree.column(col, width=w)
        else:
            self.manual_price_label.config(text="Цена (€):")
            self.manual_reimb_label.grid_remove()
            self.manual_reimb_entry.grid_remove()
            self.manual_tree.configure(columns=("code", "price", "qty"))
            self.manual_tree.heading("code", text="Код")
            self.manual_tree.heading("price", text="Цена (€)")
            self.manual_tree.heading("qty", text="Брой етикети")
            for col, w in (("code", 150), ("price", 150), ("qty", 150)):
                self.manual_tree.column(col, width=w)

    def _manual_kind_changed(self):
        new_kind = self._current_manual_kind()
        if new_kind == self._manual_current_kind_cache:
            return
        if self.manual_items:
            if not messagebox.askyesno(
                "Смяна на вида етикети",
                "Смяната на вида (Свободна продажба / НЗОК) ще изчисти текущия списък за печат.\n\nПродължи ли?",
            ):
                inv = {v: k for k, v in self._manual_kind_map.items()}
                self.manual_kind_var.set(inv[self._manual_current_kind_cache])
                return
        self.manual_items.clear()
        self.manual_tree.delete(*self.manual_tree.get_children())
        self.manual_status_label.config(text="")
        self.manual_last_output_path = None
        self.manual_print_btn.config(state="disabled")
        self.manual_open_btn.config(state="disabled")
        self._manual_current_kind_cache = new_kind
        self._manual_reconfigure_fields()

    def _manual_add(self):
        code = clean_code(self.manual_code_var.get())
        price = clean_price(self.manual_price_var.get())
        try:
            qty = max(1, int(round(float(str(self.manual_qty_var.get()).replace(",", ".")))))
        except ValueError:
            qty = 1

        if code is None:
            messagebox.showwarning("Внимание", "Въведи валиден код (число).")
            return
        if price is None:
            messagebox.showwarning("Внимание", "Въведи валидна цена в €.")
            return

        kind = self._current_manual_kind()
        if kind == "nzok":
            reimb = clean_price(self.manual_reimb_var.get())
            if reimb is None:
                messagebox.showwarning("Внимание", "Въведи валидна реимбурсация в €.")
                return
            final = max(0.0, price - reimb)
            self.manual_items.append((code, price, reimb, qty))
            self.manual_tree.insert("", "end", values=(code, f"{price:.2f}", f"{reimb:.2f}", f"{final:.2f}", qty))
            self.manual_reimb_var.set("")
        else:
            self.manual_items.append((code, price, qty))
            self.manual_tree.insert("", "end", values=(code, f"{price:.2f}", qty))

        self.manual_code_var.set("")
        self.manual_price_var.set("")
        self.manual_qty_var.set("1")
        self.manual_status_label.config(text=f"{len(self.manual_items)} артикула в списъка.")

    def _manual_remove_selected(self):
        selected = self.manual_tree.selection()
        if not selected:
            return
        indices = sorted((self.manual_tree.index(i) for i in selected), reverse=True)
        for iid in selected:
            self.manual_tree.delete(iid)
        for idx in indices:
            del self.manual_items[idx]
        self.manual_status_label.config(text=f"{len(self.manual_items)} артикула в списъка.")

    def _manual_clear(self):
        self.manual_tree.delete(*self.manual_tree.get_children())
        self.manual_items.clear()
        self.manual_status_label.config(text="")

    def _manual_generate(self):
        if not self.manual_items:
            messagebox.showwarning("Внимание", "Добави поне един артикул в списъка.")
            return

        kind = self._current_manual_kind()
        labels = []
        if kind == "nzok":
            for code, full_price, reimb, qty in self.manual_items:
                labels.extend([(code, full_price, reimb)] * qty)
        else:
            for code, price, qty in self.manual_items:
                labels.extend([(code, price, None)] * qty)

        try:
            start_row = max(1, min(LBL_ROWS, int(float(str(self.manual_start_row_var.get()).replace(",", ".")))))
        except ValueError:
            start_row = 1
        try:
            start_col = max(1, min(LBL_COLS, int(float(str(self.manual_start_col_var.get()).replace(",", ".")))))
        except ValueError:
            start_col = 1

        output_path = filedialog.asksaveasfilename(
            title="Запази PDF с етикетите",
            initialdir=self.cfg.get("last_dir") or None,
            initialfile="rachni_etiketi_nzok.pdf" if kind == "nzok" else "rachni_etiketi.pdf",
            defaultextension=".pdf",
            filetypes=[("PDF файлове", "*.pdf")],
        )
        if not output_path:
            return

        try:
            if kind == "nzok":
                generate_nzok_labels_pdf(labels, output_path, layout=load_label_layout("nzok"), start_row=start_row, start_col=start_col)
            else:
                generate_labels_pdf(labels, output_path, start_row=start_row, start_col=start_col)
        except Exception as e:
            messagebox.showerror("Грешка", f"Не можах да генерирам PDF:\n{e}")
            return

        self.manual_last_output_path = output_path
        self.manual_print_btn.config(state="normal")
        self.manual_open_btn.config(state="normal")
        self.manual_status_label.config(text=f"Готово! Генерирани са {len(labels)} етикета.")

    def _manual_open_last(self):
        if self.manual_last_output_path and os.path.exists(self.manual_last_output_path):
            os.startfile(self.manual_last_output_path)

    def _manual_print_last(self):
        if not (self.manual_last_output_path and os.path.exists(self.manual_last_output_path)):
            return
        try:
            print_pdf_file(self.manual_last_output_path)
            self.manual_status_label.config(text="Файлът е изпратен за принтиране.")
        except Exception as e:
            messagebox.showerror(
                "Грешка при печат",
                f"Не можах да принтирам директно:\n{e}\n\nОтвори PDF файла и принтирай ръчно (Ctrl+P).",
            )

    # ---------- Калибровка на етикетите ----------

    def _build_calibration_tab(self, tab):
        intro = ttk.Label(
            tab,
            text=(
                "Тук се задават точните размери на листа със стикери (в мм) и размерът на шрифта. "
                "Свободна продажба и НЗОК имат ОТДЕЛНИ настройки (различен брой редове текст), "
                "избери отдолу коя от двете калибрираш. Промени числата, разпечатай тестов лист, "
                "сравни с реалните квадратчета и пак коригирай, докато съвпадне."
            ),
            style="Status.TLabel",
            wraplength=760,
            justify="left",
        )
        intro.pack(anchor="w", padx=16, pady=(10, 10), fill="x")

        kind_row = ttk.Frame(tab)
        kind_row.pack(anchor="w", padx=16, pady=(0, 10), fill="x")
        ttk.Label(kind_row, text="Калибрирам за:").pack(side="left", padx=(0, 8))
        self.calib_kind_display_var = tk.StringVar(value="Свободна продажба")
        self._calib_kind_map = {"Свободна продажба": "otc", "НЗОК": "nzok"}
        kind_combo = ttk.Combobox(
            kind_row, textvariable=self.calib_kind_display_var, state="readonly",
            values=list(self._calib_kind_map.keys()), width=20,
        )
        kind_combo.pack(side="left")
        kind_combo.bind("<<ComboboxSelected>>", lambda e: self._calib_reload_fields())

        form = ttk.LabelFrame(tab, text="Размери на етикета и шрифт (мм / pt)", style="Step.TLabelframe", padding=14)
        form.pack(fill="x", padx=16, pady=6)

        self.calib_fields = [
            ("w", "Ширина на етикета (мм)"),
            ("h", "Височина на етикета (мм)"),
            ("margin_left", "Начало отляво (мм)"),
            ("margin_top", "Начало отгоре (мм)"),
            ("gap_h", "Разстояние между колоните (мм)"),
            ("gap_v", "Разстояние между редовете (мм)"),
            ("rotation_deg", "Завъртане (градуси, +/-)"),
            ("scale_pct", "Мащаб на печат (%)"),
            ("code_font_size", "Размер на шрифта за код (pt)"),
        ]
        self.calib_vars = {}
        for i, (key, label) in enumerate(self.calib_fields):
            r, c = divmod(i, 2)
            ttk.Label(form, text=f"{label}:").grid(row=r, column=c * 2, sticky="w", padx=(0, 6), pady=6)
            var = tk.StringVar()
            self.calib_vars[key] = var
            entry = ttk.Entry(form, textvariable=var, width=10)
            entry.grid(row=r, column=c * 2 + 1, sticky="w", padx=(0, 24), pady=6)

        ttk.Label(
            tab,
            text=(
                "8 колони × 20 реда = 160 етикета на лист (фиксирано). Ако колоните/редовете излизат "
                "накриво - пробвай малки стойности за завъртане (0.3, -0.3, 0.5...). Ако целият лист "
                "излиза по-голям/малък от реалния - коригирай мащаба на печат (%)."
            ),
            style="Status.TLabel",
            wraplength=760,
            justify="left",
        ).pack(anchor="w", padx=16, pady=(0, 6), fill="x")

        printer_frame = ttk.LabelFrame(tab, text="Принтер", style="Step.TLabelframe", padding=14)
        printer_frame.pack(fill="x", padx=16, pady=6)
        ttk.Label(printer_frame, text="Използвай принтер:").pack(side="left", padx=(0, 8))
        self.printer_var = tk.StringVar(value=load_printer_name() or "(по подразбиране на Windows)")
        self.printer_combo = ttk.Combobox(printer_frame, textvariable=self.printer_var, state="readonly", width=40)
        self.printer_combo.pack(side="left", padx=(0, 10))
        self.printer_combo.bind("<<ComboboxSelected>>", lambda e: self._save_printer_choice())
        ttk.Button(printer_frame, text="Обнови списъка с принтери", command=self._refresh_printer_list).pack(side="left")
        self._refresh_printer_list()

        self.calib_status_label = ttk.Label(tab, text="", style="Status.TLabel")
        self.calib_status_label.pack(anchor="w", padx=16, pady=(6, 10))

        self._calib_reload_fields()

        btn_bar = ttk.Frame(tab, padding=(16, 0))
        btn_bar.pack(fill="x")
        ttk.Button(
            btn_bar, text="Отвори тестов лист (преглед)", style="Big.TButton",
            command=self._calib_print_test,
        ).pack(side="left")
        ttk.Button(
            btn_bar, text="Принтирай тестовия лист директно", command=self._calib_print_direct,
        ).pack(side="left", padx=(10, 0))
        ttk.Button(
            btn_bar, text="Принтирай примерен етикет с текст", command=self._calib_print_sample,
        ).pack(side="left", padx=(10, 0))
        ttk.Button(
            btn_bar, text="Запази настройките", command=self._calib_save,
        ).pack(side="left", padx=(10, 0))

    def _current_calib_kind(self):
        return self._calib_kind_map.get(self.calib_kind_display_var.get(), "otc")

    def _calib_reload_fields(self):
        layout = load_label_layout(self._current_calib_kind())
        for key, var in self.calib_vars.items():
            var.set(str(layout.get(key, DEFAULT_LABEL_LAYOUT.get(key, 0.0))))
        self.calib_status_label.config(text="")

    def _refresh_printer_list(self):
        printers = list_printers()
        values = ["(по подразбиране на Windows)"] + printers
        self.printer_combo.configure(values=values)
        current = load_printer_name()
        if current and current in printers:
            self.printer_var.set(current)
        elif not current:
            self.printer_var.set("(по подразбиране на Windows)")

    def _save_printer_choice(self):
        choice = self.printer_var.get()
        save_printer_name(None if choice == "(по подразбиране на Windows)" else choice)

    def _read_calib_layout(self):
        layout = {}
        for key, var in self.calib_vars.items():
            try:
                layout[key] = float(str(var.get()).replace(",", "."))
            except ValueError:
                messagebox.showwarning("Внимание", "Всички стойности трябва да са числа.")
                return None
        return layout

    def _calib_save(self):
        layout = self._read_calib_layout()
        if layout is None:
            return
        save_label_layout(layout, self._current_calib_kind())
        self._save_printer_choice()
        kind_label = self.calib_kind_display_var.get()
        self.calib_status_label.config(
            text=f"Настройките за \"{kind_label}\" са запазени и ще се ползват за следващите етикети."
        )

    def _calib_print_test(self):
        layout = self._read_calib_layout()
        if layout is None:
            return
        # Avtomatichno zapazvane pri vseki test - inache ako potrebitelyat
        # koригира chislata i samo natisne test/pechat (bez otdelno "Zapazi
        # nastroykite"), popravkite se gubyat pri sledvashto otvaryane.
        save_label_layout(layout, self._current_calib_kind())
        out_path = os.path.join(_app_dir(), "kalibrovka.pdf")
        try:
            generate_calibration_grid_pdf(out_path, layout=layout)
        except Exception as e:
            messagebox.showerror("Грешка", f"Не можах да генерирам тестовия лист:\n{e}")
            return
        self.calib_status_label.config(text=f"Тестовият лист е готов и настройките са запазени: {out_path}")
        try:
            os.startfile(out_path)
        except OSError:
            pass

    def _calib_print_sample(self):
        """Pechata realni primerni etiketi (s tekst, ne samo ramki) s tekushtite
        stойности na shrifta - za da se vizhda direktno dali razmerat na
        shrifta e dobar, ne samo geometriyata."""
        layout = self._read_calib_layout()
        if layout is None:
            return
        self._save_printer_choice()
        save_label_layout(layout, self._current_calib_kind())
        out_path = os.path.join(_app_dir(), "primeren_etiket.pdf")
        try:
            if self._current_calib_kind() == "nzok":
                sample = [("NG050", 5.40, 1.39), ("AF469", 6.20, 1.71), ("CA029", 4.10, 2.75)] * 8
                generate_nzok_labels_pdf(sample, out_path, layout=layout)
            else:
                sample = [("12345", 3.99, None), ("67890", 12.50, None), ("54321", 0.99, None)] * 8
                generate_labels_pdf(sample, out_path, layout=layout)
            print_pdf_file(out_path)
            self.calib_status_label.config(text="Примерният етикет е изпратен за печат и настройките са запазени.")
        except Exception as e:
            messagebox.showerror("Грешка при печат", f"Не можах да принтирам примерния етикет:\n{e}")

    def _calib_print_direct(self):
        layout = self._read_calib_layout()
        if layout is None:
            return
        self._save_printer_choice()
        save_label_layout(layout, self._current_calib_kind())
        out_path = os.path.join(_app_dir(), "kalibrovka.pdf")
        try:
            generate_calibration_grid_pdf(out_path, layout=layout)
            print_pdf_file(out_path)
            self.calib_status_label.config(text="Тестовият лист е изпратен за печат и настройките са запазени.")
        except Exception as e:
            messagebox.showerror("Грешка при печат", f"Не можах да принтирам тестовия лист:\n{e}")


class FilePanel:
    """Nezavisim 'ot fail' panel - izbor na fail, mapirane na koloni, predpregled
    i generirane na etiketi. Izpolzva se za razlichni vidove dokumentatsiya
    (svobodna prodazhba, NZOK i t.n.), vseki sus svoe sustoyanie i svoi
    zapazeni nastroyki v config.json (razlichavani po cfg_prefix)."""

    def __init__(self, app, parent, cfg_prefix, intro_text):
        self.app = app
        self.cfg = app.cfg
        self.cfg_prefix = cfg_prefix

        self.raw_grid = []
        self.header_row_idx = 0
        self.headers = []
        self.rows = []
        self.row_origin = []
        self.checked_rows = set()
        self.qty_overrides = {}
        self._qty_edit_entry = None
        self._qty_edit_row = None
        self.loaded_files = []
        self.file_path = None
        self.last_output_path = None
        self._updating_header_row = False

        self._build_ui(parent, intro_text)

    def _ck(self, name):
        return f"{self.cfg_prefix}_{name}"

    def _build_ui(self, file_tab, intro_text):
        subtitle = ttk.Label(file_tab, text=intro_text, style="Status.TLabel")
        subtitle.pack(anchor="w", padx=16, pady=(10, 10))

        # --- Стъпка 1: файл ---
        step1 = ttk.LabelFrame(file_tab, text="Стъпка 1 — избери файл(ове)", style="Step.TLabelframe", padding=10)
        step1.pack(fill="x", padx=16, pady=6)

        drop_text = "Провлачи един или няколко файла тук  или  кликни за избор\n(поддържани формати: .xlsx  .xls  .csv  .pdf)"
        if not DND_AVAILABLE:
            drop_text = "Кликни тук, за да избереш един или няколко файла\n(поддържани формати: .xlsx  .xls  .csv  .pdf)"

        self.drop_zone = tk.Label(
            step1,
            text=drop_text,
            bg=BG_FIELD,
            fg=GOLD_TEXT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=GOLD,
            highlightcolor=GOLD,
            bd=0,
            font=("Segoe UI", 11),
            cursor="hand2",
            padx=10,
            pady=16,
            justify="center",
        )
        self.drop_zone.pack(fill="x")
        self.drop_zone.bind("<Button-1>", lambda e: self.choose_file())

        if DND_AVAILABLE:
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self.on_drop)

        file_row = ttk.Frame(step1)
        file_row.pack(fill="x", pady=(6, 0))
        self.file_label = ttk.Label(file_row, text="Няма избран файл", style="Status.TLabel", wraplength=700, justify="left")
        self.file_label.pack(side="left", fill="x")

        ttk.Label(file_row, text="   Ред със заглавия на колоните:").pack(side="left", padx=(20, 4))
        self.header_row_var = tk.IntVar(value=1)
        self.header_row_spin = ttk.Spinbox(
            file_row, from_=1, to=1, width=5, textvariable=self.header_row_var,
            command=self._on_header_row_changed, state="disabled",
        )
        self.header_row_spin.pack(side="left")
        self.header_row_spin.bind("<Return>", lambda e: self._on_header_row_changed())
        self.header_row_spin.bind("<FocusOut>", lambda e: self._on_header_row_changed())

        ttk.Button(file_row, text="Изчисти всички файлове", command=self._clear_all_files).pack(side="left", padx=(20, 0))

        # --- Стъпка 2: колони ---
        step2 = ttk.LabelFrame(file_tab, text="Стъпка 2 — коя колона какво е", style="Step.TLabelframe", padding=10)
        step2.pack(fill="x", padx=16, pady=6)
        self.step2 = step2

        ttk.Label(step2, text="Код на продукта:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=5)
        self.number_col_var = tk.StringVar()
        self.number_col_combo = ttk.Combobox(step2, textvariable=self.number_col_var, state="readonly", width=26)
        self.number_col_combo.grid(row=0, column=1, padx=(0, 20), pady=5)

        ttk.Label(step2, text="Цена в €:").grid(row=0, column=2, sticky="w", padx=(0, 6), pady=5)
        self.eur_col_var = tk.StringVar()
        self.eur_col_combo = ttk.Combobox(step2, textvariable=self.eur_col_var, state="readonly", width=26)
        self.eur_col_combo.grid(row=0, column=3, padx=(0, 20), pady=5)

        self.show_bgn_var = tk.BooleanVar(value=self.cfg.get(self._ck("show_bgn"), False))
        self.show_bgn_check = ttk.Checkbutton(
            step2, text="покажи и лв.", variable=self.show_bgn_var, command=self._on_show_bgn_toggled
        )
        self.show_bgn_check.grid(row=0, column=4, sticky="w", padx=(0, 4), pady=5)
        self.bgn_col_var = tk.StringVar()
        self.bgn_col_combo = ttk.Combobox(step2, textvariable=self.bgn_col_var, state="disabled", width=22)
        self.bgn_col_combo.grid(row=0, column=5, padx=(0, 6), pady=5)

        ttk.Label(step2, text="Количество (бр. етикети):").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=5)
        self.qty_col_var = tk.StringVar()
        self.qty_col_combo = ttk.Combobox(step2, textvariable=self.qty_col_var, state="readonly", width=26)
        self.qty_col_combo.grid(row=1, column=1, padx=(0, 20), pady=5)

        for combo in (self.number_col_combo, self.eur_col_combo, self.bgn_col_combo):
            combo.bind("<<ComboboxSelected>>", lambda e: self._update_preview_line())
        # Smqnata na kolonata za kolichestvo trqbva da preizchisli i kolonata
        # "Broy za pechat" v predpregleda (ne samo primernия red).
        self.qty_col_combo.bind("<<ComboboxSelected>>", lambda e: (self._fill_preview(), self._update_preview_line()))

        preview_row = ttk.Frame(step2)
        preview_row.grid(row=2, column=0, columnspan=6, sticky="w", pady=(8, 0))
        ttk.Label(preview_row, text="Пример за етикет:").pack(side="left")
        self.preview_line = ttk.Label(preview_row, text="—", style="Preview.TLabel")
        self.preview_line.pack(side="left", padx=(8, 0))

        self.sanity_label = ttk.Label(step2, text="", style="Warn.TLabel")
        self.sanity_label.grid(row=3, column=0, columnspan=6, sticky="w", pady=(4, 0))

        # --- Предпреглед на данните --- (fiksirana visochina, za da ne izbutva
        # dolnata lenta s butonite izvun ekrana na po-malki monitori)
        step3 = ttk.LabelFrame(file_tab, text="Предпреглед на данните", style="Step.TLabelframe", padding=10)
        step3.pack(fill="x", padx=16, pady=6)

        tree_container = ttk.Frame(step3)
        tree_container.pack(fill="x")

        self.tree = ttk.Treeview(tree_container, show="headings", height=5, selectmode="none")
        self.tree.pack(fill="x", expand=True, side="left")
        self.tree.bind("<Button-1>", self._on_tree_click)
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        preview_buttons = ttk.Frame(step3)
        preview_buttons.pack(fill="x", pady=(6, 0))
        ttk.Button(
            preview_buttons, text="Изтрий маркираните редове", command=self._delete_selected_rows
        ).pack(side="left")
        ttk.Label(
            preview_buttons,
            text=(
                "  (щракни квадратчето пред реда за изтриване; щракни в колоната "
                "\"Брой за печат\", за да смениш броя етикети на този продукт)"
            ),
            style="Status.TLabel",
        ).pack(side="left")

        # --- Начало на печат (за започнат лист стикери) ---
        start_frame = ttk.LabelFrame(
            file_tab, text="Начало на печат (ако листът вече е започнат)", style="Step.TLabelframe", padding=10
        )
        start_frame.pack(fill="x", padx=16, pady=6)
        ttk.Label(start_frame, text="Ред:").pack(side="left", padx=(0, 4))
        self.start_row_var = tk.StringVar(value="1")
        ttk.Spinbox(start_frame, from_=1, to=LBL_ROWS, width=5, textvariable=self.start_row_var).pack(side="left", padx=(0, 16))
        ttk.Label(start_frame, text="Колона:").pack(side="left", padx=(0, 4))
        self.start_col_var = tk.StringVar(value="1")
        ttk.Spinbox(start_frame, from_=1, to=LBL_COLS, width=5, textvariable=self.start_col_var).pack(side="left", padx=(0, 16))
        ttk.Button(start_frame, text="Нулирай (1 / 1)", command=self._reset_start_position).pack(side="left")

        # --- Долна лента ---
        bottom = ttk.Frame(file_tab, padding=(16, 10))
        bottom.pack(fill="x")

        self.status_label = ttk.Label(bottom, text="", style="Status.TLabel")
        self.status_label.pack(side="left")

        button_bar = ttk.Frame(bottom)
        button_bar.pack(side="right")
        self.button_bar = button_bar

        self.print_btn = ttk.Button(button_bar, text="Принтирай", command=self.print_last, state="disabled")
        self.print_btn.pack(side="right", padx=(8, 0))

        self.open_btn = ttk.Button(button_bar, text="Отвори PDF", command=self.open_last, state="disabled")
        self.open_btn.pack(side="right", padx=(8, 0))

        ttk.Button(
            button_bar, text="Генерирай етикети (PDF)", style="Big.TButton", command=self.generate
        ).pack(side="right")

    # ---------- Начало на печат ----------

    def _reset_start_position(self):
        self.start_row_var.set("1")
        self.start_col_var.set("1")

    def _get_start_position(self):
        try:
            row = max(1, min(LBL_ROWS, int(float(str(self.start_row_var.get()).replace(",", ".")))))
        except ValueError:
            row = 1
        try:
            col = max(1, min(LBL_COLS, int(float(str(self.start_col_var.get()).replace(",", ".")))))
        except ValueError:
            col = 1
        return row, col

    # ---------- Файлови действия ----------

    def on_drop(self, event):
        paths = self.app.tk.splitlist(event.data)
        if not paths:
            return
        self._open_paths(list(paths))

    def choose_file(self):
        paths = filedialog.askopenfilenames(
            title="Избери файл(ове) с документацията",
            initialdir=self.cfg.get("last_dir") or None,
            filetypes=[
                ("Поддържани файлове", "*.xlsx *.xlsm *.xls *.csv *.pdf"),
                ("Всички файлове", "*.*"),
            ],
        )
        if not paths:
            return
        # Na nyakoi Tk versii askopenfilenames vrushta niz (ne tuple) - splitlist
        # go razdelya pravilno na otdelni patishta (za razlika ot list(str),
        # koeto bi go razbilo na otdelni buквi).
        if isinstance(paths, str):
            paths = self.app.tk.splitlist(paths)
        self._open_paths(list(paths))

    def _open_paths(self, paths):
        valid_paths = []
        invalid_names = []
        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            if ext in SUPPORTED_EXTENSIONS:
                valid_paths.append(path)
            else:
                invalid_names.append(os.path.basename(path))
        if invalid_names:
            messagebox.showwarning(
                "Неподдържан формат",
                "Пропуснати файлове (трябва да са .xlsx, .xls, .csv или .pdf):\n" + "\n".join(invalid_names),
            )
        if not valid_paths:
            return

        if len(valid_paths) == 1 and not self.loaded_files:
            self._open_path(valid_paths[0])
            return

        # Dobavyane kum veche zaredenite failove (ne gi zamenya) - taka moje
        # da se dobavyat failove i po edin nayednъж, i po neskolko. Zaglavniyat
        # red se opredelya avtomatichno za vseki nov fail (spinbord za rachna
        # korekciya raboti samo dokato ima samo edin zareden fail obshto).
        new_entries = []
        errors = []
        for path in valid_paths:
            try:
                grid = load_table(path)
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {e}")
                continue
            if not grid:
                errors.append(f"{os.path.basename(path)}: файлът е празен")
                continue
            header_idx = guess_header_row(grid)
            header_row = grid[header_idx]
            headers = [str(h) if h not in (None, "") else f"Колона {i + 1}" for i, h in enumerate(header_row)]
            rows = grid[header_idx + 1:]
            new_entries.append({"path": path, "headers": headers, "rows": rows})

        if errors:
            messagebox.showwarning("Проблем с някои файлове", "\n".join(errors))
        if not new_entries:
            return

        self.loaded_files = self.loaded_files + new_entries
        self.file_path = self.loaded_files[0]["path"]
        self.raw_grid = []

        self.cfg["last_dir"] = os.path.dirname(new_entries[-1]["path"])
        save_config(self.cfg)

        self.header_row_spin.config(state="disabled")
        self.file_label.config(text="✓ " + ", ".join(os.path.basename(e["path"]) for e in self.loaded_files))

        self.headers = self.loaded_files[0]["headers"]
        self.rows = []
        self.row_origin = []
        for entry_idx, entry in enumerate(self.loaded_files):
            for row_idx, row in enumerate(entry["rows"]):
                self.rows.append(row)
                self.row_origin.append((entry_idx, row_idx))

        self.number_col_combo["values"] = self.headers
        self.eur_col_combo["values"] = self.headers
        self.bgn_col_combo["values"] = self.headers
        self.bgn_col_combo.config(state="readonly" if self.show_bgn_var.get() else "disabled")
        self.qty_col_combo["values"] = [NO_QTY_LABEL] + self.headers
        self._auto_select_columns()
        self._fill_preview()
        self._update_preview_line()

        self.print_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.last_output_path = None
        self.status_label.config(
            text=(
                f"Добавени {len(new_entries)} файл(а) - общо заредени {len(self.loaded_files)} файла, "
                f"{len(self.rows)} реда с данни (колоните по-долу се избират по ИМЕ и трябва да съвпадат "
                "във всички файлове). Провери колоните по-долу."
            )
        )

    def _clear_all_files(self):
        # Zatvarya (bez zapazvane) otvoreno pole za redaktirane na broя predi
        # da izchistim rows/row_origin - inache _fill_preview po-dolu bi
        # komitnal stойността pod veche nevalid indeks v praznия spisak.
        self._close_qty_edit(commit=False)
        self.loaded_files = []
        self.raw_grid = []
        self.file_path = None
        self.headers = []
        self.rows = []
        self.row_origin = []
        self.qty_overrides = {}
        self.file_label.config(text="Няма избран файл")
        self.header_row_spin.config(state="disabled", to=1)
        self.header_row_var.set(1)
        self.number_col_combo["values"] = []
        self.eur_col_combo["values"] = []
        self.bgn_col_combo["values"] = []
        self.qty_col_combo["values"] = [NO_QTY_LABEL]
        self.number_col_var.set("")
        self.eur_col_var.set("")
        self.bgn_col_var.set("")
        self.qty_col_var.set(NO_QTY_LABEL)
        self._fill_preview()
        self._update_preview_line()
        self.print_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.last_output_path = None
        self.status_label.config(text="")

    def _open_path(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            messagebox.showerror(
                "Неподдържан формат",
                f"Файлът трябва да е .xlsx, .xls, .csv или .pdf (получено: {ext or 'без разширение'}).",
            )
            return

        try:
            grid = load_table(path)
        except Exception as e:
            messagebox.showerror("Грешка при отваряне", f"Не можах да прочета файла:\n{e}")
            return

        if not grid:
            messagebox.showerror("Грешка", "Файлът е празен.")
            return

        self.raw_grid = grid
        self.file_path = path
        self.file_label.config(text=f"✓ {os.path.basename(path)}")

        self.cfg["last_dir"] = os.path.dirname(path)
        save_config(self.cfg)

        max_row = max(1, len(grid))
        self.header_row_spin.config(state="normal", to=max_row)
        self._updating_header_row = True
        self.header_row_var.set(guess_header_row(grid) + 1)
        self._updating_header_row = False

        self._apply_header_row()

    def _on_header_row_changed(self):
        if self._updating_header_row:
            return
        self._apply_header_row()

    def _apply_header_row(self):
        if not self.raw_grid:
            return
        # Smqna na reda sus zaglaviя preprestroяva izcяlo rows/row_origin -
        # zatvarяme (bez zapazvane) otvoreno pole za redaktirane na broя,
        # inache _fill_preview po-dolu bi komitnal stойността pod veche
        # nevalid/razlichen indeks.
        self._close_qty_edit(commit=False)
        idx = max(0, min(self.header_row_var.get() - 1, len(self.raw_grid) - 1))
        self.header_row_idx = idx
        header_row = self.raw_grid[idx]
        self.headers = [str(h) if h not in (None, "") else f"Колона {i + 1}" for i, h in enumerate(header_row)]
        self.rows = list(self.raw_grid[idx + 1:])
        self.loaded_files = [{"path": self.file_path, "headers": self.headers, "rows": self.rows}]
        self.row_origin = [(0, j) for j in range(len(self.rows))]
        self.qty_overrides = {}

        self.number_col_combo["values"] = self.headers
        self.eur_col_combo["values"] = self.headers
        self.bgn_col_combo["values"] = self.headers
        self.bgn_col_combo.config(state="readonly" if self.show_bgn_var.get() else "disabled")
        self.qty_col_combo["values"] = [NO_QTY_LABEL] + self.headers
        self._auto_select_columns()
        self._fill_preview()
        self._update_preview_line()

        self.print_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.last_output_path = None
        self.status_label.config(text=f"Заредени {len(self.rows)} реда с данни. Провери колоните по-долу.")

    def _auto_select_columns(self):
        last_number = self.cfg.get(self._ck("number_column"))
        last_eur = self.cfg.get(self._ck("eur_column"))
        last_bgn = self.cfg.get(self._ck("bgn_column"))
        last_qty = self.cfg.get(self._ck("qty_column"))

        number_col = last_number if last_number in self.headers else guess_column(self.headers, NUMBER_KEYWORDS)
        if not number_col and self.headers:
            number_col = self.headers[0]

        eur_col = last_eur if last_eur in self.headers else guess_column(self.headers, EUR_KEYWORDS, exclude={number_col})
        bgn_col = last_bgn if last_bgn in self.headers else guess_column(self.headers, BGN_KEYWORDS, exclude={number_col, eur_col})

        remaining = [h for h in self.headers if h not in (number_col, eur_col, bgn_col)]
        if not eur_col:
            eur_col = remaining.pop(0) if remaining else (self.headers[0] if self.headers else "")

        # Opredelyame kolichestvoto PREDI "leftover" fallback-a za lv. kolona -
        # inache, ako nyama realna BGN kolona vuv faila, leftover fallback-at
        # by "otkradnal" tochno kolonata s kolichestvo (napr. "Кол.") za bgn_col
        # i qty_col nikoga ne bi se avtoselektiral, dori da e ochevidna.
        if last_qty in self.headers or last_qty == NO_QTY_LABEL:
            qty_col = last_qty
        else:
            qty_col = guess_column(self.headers, QTY_KEYWORDS, exclude={number_col, eur_col, bgn_col}) or None

        if not bgn_col:
            remaining = [h for h in self.headers if h not in (number_col, eur_col, qty_col)]
            bgn_col = remaining[0] if remaining else eur_col

        if qty_col is None:
            qty_col = NO_QTY_LABEL

        self.number_col_var.set(number_col or "")
        self.eur_col_var.set(eur_col or "")
        self.bgn_col_var.set(bgn_col or "")
        self.qty_col_var.set(qty_col)

    def _effective_qty_for_row(self, i, row):
        """Broyat etiketi za red i - ili rachno zadaden (self.qty_overrides),
        ili ot izbraната kolona za kolichestvo (po ime, v sobstveните zaglaviя
        na faila, ot koyto idva tozi red), ili 1 po podrazbirane."""
        if i in self.qty_overrides:
            return self.qty_overrides[i]
        qty_col = self.qty_col_var.get()
        if qty_col and qty_col != NO_QTY_LABEL and i < len(self.row_origin):
            entry_idx, _ = self.row_origin[i]
            if entry_idx < len(self.loaded_files):
                entry_headers = self.loaded_files[entry_idx]["headers"]
                if qty_col in entry_headers:
                    idx = entry_headers.index(qty_col)
                    if idx < len(row):
                        return clean_quantity(row[idx]) or 1
        return 1

    def _fill_preview(self):
        self._close_qty_edit(commit=True)
        self.tree.delete(*self.tree.get_children())
        self.checked_rows = set()
        self.tree["columns"] = ["_check"] + list(self.headers) + ["_qty"]
        self.tree.heading("_check", text="✓")
        self.tree.column("_check", width=34, anchor="center", stretch=False)
        for h in self.headers:
            self.tree.heading(h, text=h)
            self.tree.column(h, width=130)
        self.tree.heading("_qty", text="Брой за печат")
        self.tree.column("_qty", width=110, anchor="center", stretch=False)
        # iid = indeks v self.rows/self.row_origin - taka mojem da svarjem
        # izbran red v tablicata obratno kum tochno kое (fail, red) da iztrием
        # ili da promenим broя za nego.
        for i, row in enumerate(self.rows):
            qty_val = self._effective_qty_for_row(i, row)
            self.tree.insert("", "end", iid=str(i), values=("☐",) + tuple(row) + (str(qty_val),))

    def _qty_column_id(self):
        return f"#{len(self.headers) + 2}"

    def _close_qty_edit(self, commit=True):
        """Zatvarya otvorenoto pole za redaktirane na broя (ako ima takova).
        Po podrazbirane ZAPAZVA vpisanata stойност (commit=True) - inache
        klik na drug red/kvadratche, dokato se редактира broят, bi zagubil
        tihomulkom veche vpisанoto chislo. Samo Escape trqbva da go otmenи
        bez da go zapazi (commit=False)."""
        entry = self._qty_edit_entry
        row_iid = self._qty_edit_row
        if entry is None:
            return
        self._qty_edit_entry = None
        self._qty_edit_row = None
        if commit and self.tree.exists(row_iid):
            text = entry.get().strip().replace(",", ".")
            try:
                new_qty = max(0, int(round(float(text))))
            except ValueError:
                new_qty = None
            if new_qty is not None:
                i = int(row_iid)
                self.qty_overrides[i] = new_qty
                self.tree.set(row_iid, "_qty", str(new_qty))
                self.print_btn.config(state="disabled")
                self.open_btn.config(state="disabled")
                self.last_output_path = None
        entry.destroy()

    def _on_tree_click(self, event):
        """Shtraknuvane vurhu kvadratcheto (parvata kolona) markira/de-markira
        reda za iztrivane - ne trqbva Ctrl/Shift za poveche redove naведнъж.
        Shtraknuvane v poslednata kolona ("Broy za pechat") otvarya pole za
        rachna promяna na broya etiketi na tozi produkt."""
        self._close_qty_edit(commit=True)
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        col = self.tree.identify_column(event.x)
        row_iid = self.tree.identify_row(event.y)
        if not row_iid:
            return
        if col == "#1":
            if row_iid in self.checked_rows:
                self.checked_rows.discard(row_iid)
                self.tree.set(row_iid, "_check", "☐")
            else:
                self.checked_rows.add(row_iid)
                self.tree.set(row_iid, "_check", "☑")
            return
        if col == self._qty_column_id():
            self._start_qty_edit(row_iid)

    def _start_qty_edit(self, row_iid):
        bbox = self.tree.bbox(row_iid, "_qty")
        if not bbox:
            return
        x, y, w, h = bbox
        current_val = self.tree.set(row_iid, "_qty")

        entry = tk.Entry(self.tree, justify="center")
        entry.insert(0, current_val)
        entry.select_range(0, "end")
        entry.focus_set()
        entry.place(x=x, y=y, width=w, height=h)
        self._qty_edit_entry = entry
        self._qty_edit_row = row_iid

        entry.bind("<Return>", lambda e: self._close_qty_edit(commit=True))
        entry.bind("<FocusOut>", lambda e: self._close_qty_edit(commit=True))
        entry.bind("<Escape>", lambda e: self._close_qty_edit(commit=False))

    def _delete_selected_rows(self):
        # Zatvarya (bez zapazvane) vsяko otvoreno pole za redaktirane na broя -
        # indeksite se promeнят sled iztrivane, taka che commit tuk bi mogal
        # da zapishe broя pod veche pogreshen (izместен) indeks.
        self._close_qty_edit(commit=False)
        if not self.checked_rows:
            messagebox.showinfo(
                "Няма маркирани редове",
                "Щракни квадратчето пред реда(овете), които искаш да изтриеш, после натисни бутона пак.",
            )
            return
        indices = sorted((int(iid) for iid in self.checked_rows), reverse=True)

        entry_removals = {}
        for i in indices:
            entry_idx, row_idx = self.row_origin[i]
            entry_removals.setdefault(entry_idx, set()).add(row_idx)

        for entry_idx, row_idx_set in entry_removals.items():
            entry_rows = self.loaded_files[entry_idx]["rows"]
            self.loaded_files[entry_idx]["rows"] = [
                r for j, r in enumerate(entry_rows) if j not in row_idx_set
            ]

        self.rows = []
        self.row_origin = []
        for entry_idx, entry in enumerate(self.loaded_files):
            for row_idx, row in enumerate(entry["rows"]):
                self.rows.append(row)
                self.row_origin.append((entry_idx, row_idx))

        # Indeksite na redovete se preместват sled iztrivane - rachno zadadените
        # broevе (po stария indeks) veche bi sochили kum grešen red, zatova gi
        # izchistvame, za da ne se slepят s pogreshen produkt.
        self.qty_overrides = {}

        self._fill_preview()
        self._update_preview_line()
        self.print_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.last_output_path = None
        self.status_label.config(
            text=f"Изтрити {len(indices)} ред(а). Остават {len(self.rows)} реда с данни."
        )

    def _on_show_bgn_toggled(self):
        show_bgn = self.show_bgn_var.get()
        self.bgn_col_combo.config(state="readonly" if show_bgn else "disabled")
        self.cfg[self._ck("show_bgn")] = show_bgn
        save_config(self.cfg)
        self._update_preview_line()

    def _update_preview_line(self):
        number_col = self.number_col_var.get()
        eur_col = self.eur_col_var.get()
        show_bgn = self.show_bgn_var.get()
        bgn_col = self.bgn_col_var.get() if show_bgn else ""

        if not (number_col and eur_col) or (show_bgn and not bgn_col) or not self.rows:
            self.preview_line.config(text="—")
            self.sanity_label.config(text="")
            return

        try:
            ni = self.headers.index(number_col)
            ei = self.headers.index(eur_col)
            bi = self.headers.index(bgn_col) if bgn_col else None
        except ValueError:
            return

        qty_col = self.qty_col_var.get()
        qty_idx = self.headers.index(qty_col) if qty_col and qty_col != NO_QTY_LABEL else None

        for row in self.rows[:30]:
            if max(ni, ei, bi if bi is not None else 0) >= len(row):
                continue
            number = clean_code(row[ni])
            eur = clean_price(row[ei])
            bgn = clean_price(row[bi]) if bi is not None else None
            if number is None or eur is None or (show_bgn and bgn is None):
                continue
            eur_text, bgn_text = format_label_price(eur, bgn if show_bgn else None)
            qty_suffix = ""
            if qty_idx is not None and qty_idx < len(row):
                qty = clean_quantity(row[qty_idx]) or 1
                qty_suffix = f"   ×{qty} бр."
            self.preview_line.config(text=f"{number}   {eur_text}{bgn_text or ''}{qty_suffix}")
            if show_bgn and bgn is not None:
                self._check_rate_sanity(eur, bgn)
            else:
                self.sanity_label.config(text="")
            return

        self.preview_line.config(text="(няма валиден пример в първите редове)")
        self.sanity_label.config(text="")

    def _check_rate_sanity(self, eur, bgn):
        if eur <= 0:
            return
        ratio = bgn / eur
        if abs(ratio - BGN_PER_EUR) / BGN_PER_EUR > 0.03:
            self.sanity_label.config(
                text=(
                    f"⚠ Съотношението €/лв. в примера е {ratio:.3f} (очаквано {BGN_PER_EUR}). "
                    "Провери дали колоните за € и лв. са избрани правилно."
                )
            )
        else:
            self.sanity_label.config(text="")

    # ---------- Генериране ----------

    def generate(self):
        if not self.rows:
            messagebox.showwarning("Внимание", "Първо избери файл със стоковите данни.")
            return

        number_col = self.number_col_var.get()
        eur_col = self.eur_col_var.get()
        show_bgn = self.show_bgn_var.get()
        bgn_col = self.bgn_col_var.get() if show_bgn else ""

        if not number_col or not eur_col or (show_bgn and not bgn_col):
            messagebox.showwarning("Внимание", "Избери колона за код и за цена в €.")
            return
        chosen = {number_col, eur_col} | ({bgn_col} if bgn_col else set())
        if len(chosen) < (3 if show_bgn else 2):
            messagebox.showwarning("Внимание", "Избраните колони трябва да са различни.")
            return

        qty_col = self.qty_col_var.get()

        labels = []
        skipped = 0
        ratios = []
        skipped_files = []
        global_i = 0
        for entry in self.loaded_files:
            entry_headers = entry["headers"]
            try:
                e_number_idx = entry_headers.index(number_col)
                e_eur_idx = entry_headers.index(eur_col)
                e_bgn_idx = entry_headers.index(bgn_col) if bgn_col else None
                e_qty_idx = entry_headers.index(qty_col) if qty_col and qty_col != NO_QTY_LABEL else None
            except ValueError:
                skipped_files.append(os.path.basename(entry["path"]))
                global_i += len(entry["rows"])
                continue
            for row in entry["rows"]:
                needed = [e_number_idx, e_eur_idx] + ([e_bgn_idx] if e_bgn_idx is not None else [])
                if max(needed) >= len(row):
                    skipped += 1
                    global_i += 1
                    continue
                number = clean_code(row[e_number_idx])
                eur = clean_price(row[e_eur_idx])
                bgn = clean_price(row[e_bgn_idx]) if e_bgn_idx is not None else None
                if number is None or eur is None or (show_bgn and bgn is None):
                    skipped += 1
                    global_i += 1
                    continue
                # Rachno zadadeniyat broy (ot predpregleda) ima prioritet pred
                # stойността v samия fail - taka moje da se smaли/uveliчи broя
                # etiketi za konkreten produkt bez da se pipa failat.
                qty = self.qty_overrides.get(global_i)
                if qty is None:
                    qty = 1
                    if e_qty_idx is not None and e_qty_idx < len(row):
                        qty = clean_quantity(row[e_qty_idx]) or 1
                if qty > 0:
                    labels.extend([(number, eur, bgn if show_bgn else None)] * qty)
                if show_bgn and bgn is not None and eur > 0:
                    ratios.append(bgn / eur)
                global_i += 1

        if skipped_files:
            messagebox.showwarning(
                "Пропуснати файлове",
                "Следните файлове нямат избраните колони (по име) и бяха пропуснати:\n" + "\n".join(skipped_files),
            )

        if not labels:
            messagebox.showerror(
                "Грешка",
                "Не се намериха валидни редове с код и цени в избраните колони.\nПровери дали си избрал правилните колони.",
            )
            return

        if ratios:
            median_ratio = sorted(ratios)[len(ratios) // 2]
            if abs(median_ratio - BGN_PER_EUR) / BGN_PER_EUR > 0.03:
                proceed = messagebox.askyesno(
                    "Проверка на цените",
                    f"Съотношението лв./€ в данните е ~{median_ratio:.3f}, а официалният курс е {BGN_PER_EUR}.\n"
                    "Изглежда избраните колони за € и лв. може да не са верни.\n\nДа продължа ли все пак?",
                )
                if not proceed:
                    return

        if len(self.loaded_files) > 1:
            default_name = f"etiketi_{len(self.loaded_files)}_faila.pdf"
        else:
            default_name = os.path.splitext(os.path.basename(self.file_path))[0] + "_etiketi.pdf"
        output_path = filedialog.asksaveasfilename(
            title="Запази PDF с етикетите",
            initialdir=self.cfg.get("last_dir") or None,
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF файлове", "*.pdf")],
        )
        if not output_path:
            return

        start_row, start_col = self._get_start_position()
        try:
            generate_labels_pdf(labels, output_path, start_row=start_row, start_col=start_col)
        except Exception as e:
            messagebox.showerror("Грешка", f"Не можах да генерирам PDF:\n{e}")
            return

        self.cfg[self._ck("number_column")] = number_col
        self.cfg[self._ck("eur_column")] = eur_col
        if bgn_col:
            self.cfg[self._ck("bgn_column")] = bgn_col
        self.cfg[self._ck("qty_column")] = qty_col
        self.cfg[self._ck("show_bgn")] = show_bgn
        save_config(self.cfg)

        self.last_output_path = output_path
        self.print_btn.config(state="normal")
        self.open_btn.config(state="normal")

        msg = f"Готово! Генерирани са {len(labels)} етикета."
        if skipped:
            msg += f"  (Пропуснати {skipped} реда без валидни данни.)"
        self.status_label.config(text=msg)

    def open_last(self):
        if self.last_output_path and os.path.exists(self.last_output_path):
            os.startfile(self.last_output_path)

    def print_last(self):
        if not (self.last_output_path and os.path.exists(self.last_output_path)):
            return
        try:
            print_pdf_file(self.last_output_path)
            self.status_label.config(text="Файлът е изпратен за принтиране.")
        except Exception as e:
            messagebox.showerror(
                "Грешка при печат",
                f"Не можах да принтирам директно:\n{e}\n\nОтвори PDF файла и принтирай ръчно (Ctrl+P).",
            )


class NzokFilePanel(FilePanel):
    """Fail panel spetsialno za NZOK dostavki - dobavya avtomatichno svalyane na
    reimbursatsii ot nhif.bg (spravochnik po NZOK kod, kesh mezhdu startirania)
    i generira etiketi na 3 reda: kod / (tsena - reimbursatsiya) / krayna tsena."""

    def __init__(self, app, parent, cfg_prefix, intro_text):
        self.reimbursements, self.reimb_month = load_reimbursement_cache()
        self._refreshing = False
        super().__init__(app, parent, cfg_prefix, intro_text)

    def _build_ui(self, file_tab, intro_text):
        super()._build_ui(file_tab, intro_text)

        # Skrivame BGN-tumblera na bazovия panel - NZOK etiketite vinagi pokazvat
        # i lv. (izchisleno ot fiksiraniya kurs sled prispadane na reimbursatsiyata).
        self.show_bgn_check.grid_remove()
        self.bgn_col_combo.grid_remove()

        reimb_row = ttk.Frame(self.step2)
        reimb_row.grid(row=4, column=0, columnspan=6, sticky="w", pady=(10, 0))

        self.reimb_status_label = ttk.Label(reimb_row, text="", style="Status.TLabel")
        self.reimb_status_label.pack(side="left")

        self.reimb_refresh_btn = ttk.Button(
            reimb_row, text="Обнови реимбурсации от НЗОК", command=self._refresh_reimbursements
        )
        self.reimb_refresh_btn.pack(side="left", padx=(10, 0))

        self._update_reimb_status_label()
        self.app.after(800, self._maybe_auto_refresh)

    def _update_reimb_status_label(self):
        if self.reimb_month:
            self.reimb_status_label.config(
                text=f"Реимбурсации: заредени за {self.reimb_month} ({len(self.reimbursements)} кода)"
            )
        else:
            self.reimb_status_label.config(text="Реимбурсации: не са заредени още")

    def _maybe_auto_refresh(self):
        """Avtomatichna proverka pri start na programata - NZOK publikuva nov spisak
        okolo 1-vo i 16-o chislo na meseca, tova proverqva dali veche sme svalyali
        za tozi period i ako ne - svalya sam, bez potrebitelyat da natiska nishto."""
        if get_last_checked_period() == _current_reimb_period():
            return
        self._refresh_reimbursements(silent=True)

    def _refresh_reimbursements(self, silent=False):
        if self._refreshing:
            return
        self._refreshing = True
        self.reimb_refresh_btn.config(state="disabled")
        if not silent:
            self.reimb_status_label.config(text="Свалям последния списък от nhif.bg...")

        def worker():
            try:
                table, month_label, _url = fetch_nzok_reimbursement_table()
                error = None
            except Exception as e:
                table, month_label, error = None, None, str(e)
            self.app.after(0, lambda: self._on_refresh_done(table, month_label, error, silent))

        threading.Thread(target=worker, daemon=True).start()

    def _on_refresh_done(self, table, month_label, error, silent=False):
        self._refreshing = False
        self.reimb_refresh_btn.config(state="normal")
        if error:
            self._update_reimb_status_label()
            if not silent:
                messagebox.showerror(
                    "Грешка при обновяване",
                    f"Не можах да сваля списъка от НЗОК:\n{error}\n\n"
                    "Провери интернет връзката, или опитай пак по-късно.",
                )
            else:
                # Ne markirame perioda kato proveren, za da opitame otnovo pri
                # sledvashto pusкане na programata (napr. ako sega nyama internet).
                pass
            return

        self.reimbursements = table
        self.reimb_month = month_label
        save_reimbursement_cache(table, month_label)
        mark_reimb_period_checked()
        self._update_reimb_status_label()
        self._update_preview_line()
        if not silent:
            messagebox.showinfo(
                "Готово", f"Обновени реимбурсации за {month_label} - {len(table)} кода."
            )

    def _update_preview_line(self):
        number_col = self.number_col_var.get()
        eur_col = self.eur_col_var.get()

        if not (number_col and eur_col) or not self.rows:
            self.preview_line.config(text="—")
            self.sanity_label.config(text="")
            return

        try:
            ni = self.headers.index(number_col)
            ei = self.headers.index(eur_col)
        except ValueError:
            return

        qty_col = self.qty_col_var.get()
        qty_idx = self.headers.index(qty_col) if qty_col and qty_col != NO_QTY_LABEL else None

        for row in self.rows[:30]:
            if max(ni, ei) >= len(row):
                continue
            code = clean_code(row[ni])
            eur = clean_price(row[ei])
            if code is None or eur is None:
                continue
            reimb = self.reimbursements.get(code, 0.0)
            final_eur = max(0.0, eur - reimb)
            qty_suffix = ""
            if qty_idx is not None and qty_idx < len(row):
                qty = clean_quantity(row[qty_idx]) or 1
                qty_suffix = f"   ×{qty} бр."
            missing_note = "  (без реимб. - непознат код)" if code not in self.reimbursements else ""
            self.preview_line.config(
                text=(
                    f"{code}   {eur:.2f} - {reimb:.2f}   "
                    f"{final_eur:.2f}€{qty_suffix}{missing_note}"
                )
            )
            self.sanity_label.config(text="")
            return

        self.preview_line.config(text="(няма валиден пример в първите редове)")
        self.sanity_label.config(text="")

    def generate(self):
        if not self.rows:
            messagebox.showwarning("Внимание", "Първо избери файл със стоковите данни.")
            return

        number_col = self.number_col_var.get()
        eur_col = self.eur_col_var.get()
        if not number_col or not eur_col:
            messagebox.showwarning("Внимание", "Избери колона за код и за цена в €.")
            return
        if number_col == eur_col:
            messagebox.showwarning("Внимание", "Избраните колони трябва да са различни.")
            return

        if not self.reimbursements:
            proceed = messagebox.askyesno(
                "Няма заредени реимбурсации",
                "Още не си сваляла(л) списъка с реимбурсации от НЗОК - всички етикети "
                "ще излязат с пълна цена (без приспадане).\n\nДа продължа ли така?",
            )
            if not proceed:
                return

        qty_col = self.qty_col_var.get()

        labels = []
        skipped = 0
        missing_reimb = set()
        skipped_files = []
        global_i = 0
        for entry in self.loaded_files:
            entry_headers = entry["headers"]
            try:
                e_number_idx = entry_headers.index(number_col)
                e_eur_idx = entry_headers.index(eur_col)
                e_qty_idx = entry_headers.index(qty_col) if qty_col and qty_col != NO_QTY_LABEL else None
            except ValueError:
                skipped_files.append(os.path.basename(entry["path"]))
                global_i += len(entry["rows"])
                continue
            for row in entry["rows"]:
                needed = [e_number_idx, e_eur_idx]
                if max(needed) >= len(row):
                    skipped += 1
                    global_i += 1
                    continue
                code = clean_code(row[e_number_idx])
                eur = clean_price(row[e_eur_idx])
                if code is None or eur is None:
                    skipped += 1
                    global_i += 1
                    continue
                qty = self.qty_overrides.get(global_i)
                if qty is None:
                    qty = 1
                    if e_qty_idx is not None and e_qty_idx < len(row):
                        qty = clean_quantity(row[e_qty_idx]) or 1

                reimb = self.reimbursements.get(code)
                if reimb is None:
                    missing_reimb.add(code)
                    reimb = 0.0
                if qty > 0:
                    labels.extend([(code, eur, reimb)] * qty)
                global_i += 1

        if skipped_files:
            messagebox.showwarning(
                "Пропуснати файлове",
                "Следните файлове нямат избраните колони (по име) и бяха пропуснати:\n" + "\n".join(skipped_files),
            )

        if not labels:
            messagebox.showerror(
                "Грешка",
                "Не се намериха валидни редове с код и цена в избраните колони.\n"
                "Провери дали си избрал правилните колони.",
            )
            return

        if missing_reimb and self.reimbursements:
            shown = ", ".join(sorted(missing_reimb)[:15])
            more = "..." if len(missing_reimb) > 15 else ""
            proceed = messagebox.askyesno(
                "Липсваща реимбурсация",
                f"{len(missing_reimb)} кода не бяха намерени в списъка от НЗОК (ще излязат "
                f"с пълна цена, без приспадане):\n{shown}{more}\n\nДа продължа ли така?",
            )
            if not proceed:
                return

        if len(self.loaded_files) > 1:
            default_name = f"etiketi_nzok_{len(self.loaded_files)}_faila.pdf"
        else:
            default_name = os.path.splitext(os.path.basename(self.file_path))[0] + "_etiketi.pdf"
        output_path = filedialog.asksaveasfilename(
            title="Запази PDF с етикетите",
            initialdir=self.cfg.get("last_dir") or None,
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF файлове", "*.pdf")],
        )
        if not output_path:
            return

        start_row, start_col = self._get_start_position()
        try:
            generate_nzok_labels_pdf(labels, output_path, start_row=start_row, start_col=start_col)
        except Exception as e:
            messagebox.showerror("Грешка", f"Не можах да генерирам PDF:\n{e}")
            return

        self.cfg[self._ck("number_column")] = number_col
        self.cfg[self._ck("eur_column")] = eur_col
        self.cfg[self._ck("qty_column")] = qty_col
        save_config(self.cfg)

        self.last_output_path = output_path
        self.print_btn.config(state="normal")
        self.open_btn.config(state="normal")

        msg = f"Готово! Генерирани са {len(labels)} етикета."
        if skipped:
            msg += f"  (Пропуснати {skipped} реда без валидни данни.)"
        self.status_label.config(text=msg)


if __name__ == "__main__":
    app = App()
    app.mainloop()
