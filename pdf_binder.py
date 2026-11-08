"""
PDF Binder - collect PDF, Word, Excel, PowerPoint, images and text files
into one PDF.

  * every document: preview of its pages, rotation of the whole document or
    of single pages (+90 / -90, as often as needed) - the rotation is written
    into the saved PDF
  * every document: check box "In contents", title and level
    (1 = section, 2 = subsection, 3 = sub-subsection)
  * on Save: table of contents (clickable, also as PDF bookmarks),
    page numbers, header and footer - asked every time

Conversion of Office files: MS Office (Windows, pywin32) if installed,
else LibreOffice (soffice). PDF, images and text need nothing else.

Run:   python pdf_binder.py
Needs: pip install pymupdf pyside6        (Windows + MS Office: pywin32)
"""

import datetime
import os
import shutil
import subprocess
import sys
import tempfile

import pymupdf
from PySide6.QtCore import QSize, Qt, QThread, Signal
from PySide6.QtGui import QIcon, QImage, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog,
    QDialogButtonBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox, QProgressDialog, QPushButton, QSlider, QSpinBox, QSplitter,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

APP_NAME = "PDF Binder"
APP_VERSION = "0.0.1"
AUTHOR = "Ievgen German"
AUTHOR_URL = "https://www.linkedin.com/in/ievgen-german-945b16b6/"      # LinkedIn profile

PDF_EXT = {".pdf"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif", ".webp"}
TEXT_EXT = {".txt"}
WORD_EXT = {".doc", ".docx", ".docm", ".rtf", ".odt"}
EXCEL_EXT = {".xls", ".xlsx", ".xlsm", ".ods", ".csv"}
PPT_EXT = {".ppt", ".pptx", ".odp"}
OTHER_LO_EXT = {".html", ".htm", ".xml", ".vsd", ".vsdx"}
ALL_EXT = PDF_EXT | IMAGE_EXT | TEXT_EXT | WORD_EXT | EXCEL_EXT | PPT_EXT | OTHER_LO_EXT


# =====================================================================
# fonts (Cyrillic needs a real TTF font - the PDF base fonts have none)
# =====================================================================
_FONT_CANDIDATES = [
    r"C:\Windows\Fonts\arial.ttf",
    r"C:\Windows\Fonts\calibri.ttf",
    r"C:\Windows\Fonts\segoeui.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
]


def find_font():
    for p in _FONT_CANDIDATES:
        if os.path.isfile(p):
            return p
    return None


FONT_FILE = find_font()
FONT_NAME = "binderfont"


def add_font(page):
    """Font for text on a page; returns the font name to use."""
    if FONT_FILE:
        page.insert_font(fontname=FONT_NAME, fontfile=FONT_FILE)
        return FONT_NAME
    return "helv"


def text_width(text, size):
    if FONT_FILE:
        return pymupdf.Font(fontfile=FONT_FILE).text_length(text, fontsize=size)
    return pymupdf.get_text_length(text, fontname="helv", fontsize=size)


# =====================================================================
# conversion to PDF
# =====================================================================
class ConvertError(Exception):
    pass


def find_soffice():
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.isfile(c):
            return c
    return None


def _office_convert(src, out_pdf):
    """MS Office through COM (Windows). True = done, False = Office not usable."""
    if not sys.platform.startswith("win"):
        return False
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    ext = os.path.splitext(src)[1].lower()
    src = os.path.abspath(src)
    out_pdf = os.path.abspath(out_pdf)
    pythoncom.CoInitialize()
    try:
        if ext in WORD_EXT:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            try:
                doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(out_pdf, 17)            # wdExportFormatPDF
                doc.Close(False)
            finally:
                app.Quit()
        elif ext in EXCEL_EXT:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                wb = app.Workbooks.Open(src, ReadOnly=True, UpdateLinks=0)
                wb.ExportAsFixedFormat(0, out_pdf)              # xlTypePDF, all sheets
                wb.Close(False)
            finally:
                app.Quit()
        elif ext in PPT_EXT:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                pres = app.Presentations.Open(src, ReadOnly=True, WithWindow=False)
                pres.SaveAs(out_pdf, 32)                        # ppSaveAsPDF
                pres.Close()
            finally:
                app.Quit()
        else:
            return False
        return os.path.isfile(out_pdf)
    except Exception:
        return False
    finally:
        pythoncom.CoUninitialize()


def _libreoffice_convert(src, out_pdf):
    soffice = find_soffice()
    if not soffice:
        return False
    outdir = tempfile.mkdtemp(prefix="pdfbinder_lo_")
    try:
        subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "pdf",
                        "--outdir", outdir, src],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       timeout=300, check=False)
        produced = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
        if os.path.isfile(produced):
            shutil.move(produced, out_pdf)
            return True
        return False
    except Exception:
        return False
    finally:
        shutil.rmtree(outdir, ignore_errors=True)


def _image_to_pdf(src, out_pdf):
    img = pymupdf.open(src)
    pdf_bytes = img.convert_to_pdf()
    img.close()
    with open(out_pdf, "wb") as f:
        f.write(pdf_bytes)


def _text_to_pdf(src, out_pdf):
    with open(src, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "cp1251", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    doc = pymupdf.open()
    w, h = pymupdf.paper_size("a4")
    margin, size = 50, 10
    lines = text.replace("\t", "    ").splitlines() or [""]
    line_h = size * 1.35
    per_page = int((h - 2 * margin) // line_h)
    for start in range(0, len(lines), per_page):
        page = doc.new_page(width=w, height=h)
        font = add_font(page)
        y = margin + size
        for ln in lines[start:start + per_page]:
            page.insert_text((margin, y), ln, fontname=font, fontsize=size)
            y += line_h
    doc.save(out_pdf)
    doc.close()


def convert_to_pdf(src, workdir):
    """Returns the path of a PDF with the content of src (src itself for a PDF)."""
    ext = os.path.splitext(src)[1].lower()
    if ext in PDF_EXT:
        return src
    out_pdf = os.path.join(workdir, f"{len(os.listdir(workdir)):04d}_" +
                           os.path.splitext(os.path.basename(src))[0] + ".pdf")
    if ext in IMAGE_EXT:
        _image_to_pdf(src, out_pdf)
        return out_pdf
    if ext in TEXT_EXT:
        _text_to_pdf(src, out_pdf)
        return out_pdf
    if ext in WORD_EXT | EXCEL_EXT | PPT_EXT:
        if _office_convert(src, out_pdf):
            return out_pdf
    if _libreoffice_convert(src, out_pdf):
        return out_pdf
    raise ConvertError("Cannot convert this file: no MS Office (with pywin32) and no LibreOffice found.")


# =====================================================================
# model
# =====================================================================
class DocItem:
    def __init__(self, src, pdf_path):
        self.src = src
        self.pdf_path = pdf_path
        d = pymupdf.open(pdf_path)
        self.page_count = d.page_count
        d.close()
        self.extra_rot = [0] * self.page_count       # added rotation per page (0, 90, 180, 270)
        self.in_toc = False
        self.title = os.path.splitext(os.path.basename(src))[0]
        self.level = 1

    def rotate_pages(self, pages, delta):
        for p in pages:
            self.extra_rot[p] = (self.extra_rot[p] + delta) % 360


# =====================================================================
# building the PDF
# =====================================================================
class BuildOptions:
    def __init__(self):
        self.toc = False
        self.toc_title = "Contents"
        self.page_numbers = False
        self.number_pos = "bottom-center"          # bottom-center / bottom-right / top-right
        self.number_format = "{page}"              # {page}, {pages}
        self.count_toc_pages = False               # numbering starts after the contents
        self.header = ""                           # placeholders {page} {pages} {title} {date} {file}
        self.footer = ""
        self.font_size = 9
        self.marks_on_toc = False                  # header / footer / number on the contents pages


def _fix_levels(levels):
    """Bookmark levels must start at 1 and grow by one step at most."""
    out, prev = [], 0
    for lv in levels:
        lv = max(1, min(lv, prev + 1))
        out.append(lv)
        prev = lv
    return out


def _place(page, x_vis, y_vis, text, font, size):
    """Writes text so it is upright in the viewer, also on rotated pages.
    (x_vis, y_vis) = baseline start in the visible (rotated) page."""
    p = pymupdf.Point(x_vis, y_vis) * page.derotation_matrix
    page.insert_text(p, text, fontname=font, fontsize=size, rotate=page.rotation)


def _toc_lines_per_page(height, top, bottom, line_h):
    return max(1, int((height - top - bottom) // line_h))


def build_pdf(items, out_path, opt, progress=None):
    """items: list of DocItem in order. Writes out_path."""
    out = pymupdf.open()
    entries = []                                    # (level, title, first page index in the body)
    for it in items:
        start = out.page_count
        src = pymupdf.open(it.pdf_path)
        out.insert_pdf(src)
        src.close()
        for i in range(it.page_count):
            pg = out[start + i]
            pg.set_rotation((pg.rotation + it.extra_rot[i]) % 360)
        if it.in_toc and it.title.strip():
            entries.append([it.level, it.title.strip(), start])
        if progress:
            progress(f"Added {os.path.basename(it.src)}")

    body_titles = {}                                # body page -> title of its document
    starts = [(e[2], e[1]) for e in entries]
    for k, (s, t) in enumerate(starts):
        end = starts[k + 1][0] if k + 1 < len(starts) else out.page_count
        for p in range(s, end):
            body_titles[p] = t

    # ---- contents pages ----
    n_toc = 0
    toc_w, toc_h = pymupdf.paper_size("a4")
    top, bottom, left, right = 80, 60, 60, 60
    size, line_h = 11, 20
    if opt.toc and entries:
        per_page = _toc_lines_per_page(toc_h, top + 30, bottom, line_h)
        n_toc = (len(entries) + per_page - 1) // per_page
        for k in range(n_toc):
            out.new_page(pno=k, width=toc_w, height=toc_h)
        for e in entries:
            e[2] += n_toc
        shown = (lambda phys: phys + 1) if opt.count_toc_pages else (lambda phys: phys + 1 - n_toc)
        for k in range(n_toc):
            page = out[k]
            font = add_font(page)
            y = top
            if k == 0:
                page.insert_text((left, y), opt.toc_title, fontname=font, fontsize=16)
            y += 30
            for lv, title, target in entries[k * per_page:(k + 1) * per_page]:
                indent = left + (lv - 1) * 18
                num = str(shown(target))
                num_w = text_width(num, size)
                x_num = toc_w - right - num_w
                avail = x_num - 10 - indent
                t = title
                while text_width(t, size) > avail and len(t) > 3:
                    t = t[:-2]
                if t != title:
                    t = t[:-1] + "..."
                tw = text_width(t, size)
                page.insert_text((indent, y), t, fontname=font, fontsize=size)
                dots_x0 = indent + tw + 6
                dot_w = text_width(".", size) or 3
                n_dots = int((x_num - 6 - dots_x0) // dot_w)
                if n_dots > 0:
                    page.insert_text((dots_x0, y), "." * n_dots, fontname=font, fontsize=size,
                                     color=(0.5, 0.5, 0.5))
                page.insert_text((x_num, y), num, fontname=font, fontsize=size)
                page.insert_link({"kind": pymupdf.LINK_GOTO, "page": target,
                                  "from": pymupdf.Rect(indent, y - size, toc_w - right, y + 4),
                                  "to": pymupdf.Point(0, 0)})
                y += line_h

    # ---- page numbers, header, footer ----
    total = out.page_count
    shown_total = total if opt.count_toc_pages else total - n_toc
    today = datetime.date.today().strftime("%d.%m.%Y")
    fname = os.path.splitext(os.path.basename(out_path))[0]
    fs = opt.font_size
    if opt.page_numbers or opt.header or opt.footer:
        for i in range(total):
            if i < n_toc and not opt.marks_on_toc:
                continue
            page = out[i]
            font = add_font(page)
            w, h = page.rect.width, page.rect.height          # visible size (after rotation)
            num = (i + 1) if opt.count_toc_pages else (i + 1 - n_toc)
            vals = {"page": num, "pages": shown_total, "date": today, "file": fname,
                    "title": body_titles.get(i - n_toc, "") if i >= n_toc else opt.toc_title}

            def fill(s):
                try:
                    return s.format(**vals)
                except (KeyError, IndexError, ValueError):
                    return s
            if opt.header:
                _place(page, 30, 22, fill(opt.header), font, fs)
            if opt.footer:
                _place(page, 30, h - 14, fill(opt.footer), font, fs)
            if opt.page_numbers and num >= 1:
                s = fill(opt.number_format)
                sw = text_width(s, fs)
                if opt.number_pos == "bottom-right":
                    _place(page, w - 30 - sw, h - 14, s, font, fs)
                elif opt.number_pos == "top-right":
                    _place(page, w - 30 - sw, 22, s, font, fs)
                else:
                    _place(page, (w - sw) / 2, h - 14, s, font, fs)

    # ---- bookmarks ----
    if entries:
        levels = _fix_levels([e[0] for e in entries])
        out.set_toc([[lv, e[1], e[2] + 1] for lv, e in zip(levels, entries)])
    out.set_metadata({"title": fname, "creator": APP_NAME})
    out.save(out_path, garbage=3, deflate=True)
    out.close()


# =====================================================================
# GUI
# =====================================================================
def page_pixmap(pdf_path, pno, extra_rot, height):
    d = pymupdf.open(pdf_path)
    try:
        page = d[pno]
        h = page.rect.height if extra_rot in (0, 180) else page.rect.width
        zoom = max(0.05, height / max(h, 1))
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom).prerotate(extra_rot), alpha=False)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
        return QPixmap.fromImage(img)
    finally:
        d.close()


class ThumbWorker(QThread):
    done = Signal(int, int, QImage)            # generation, page, image

    def __init__(self, gen, pdf_path, rots, height, pages):
        super().__init__()
        self.gen, self.pdf_path, self.rots, self.height, self.pages = gen, pdf_path, rots, height, pages
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            d = pymupdf.open(self.pdf_path)
        except Exception:
            return
        for p in self.pages:
            if self._stop:
                break
            page = d[p]
            rot = self.rots[p]
            h = page.rect.height if rot in (0, 180) else page.rect.width
            zoom = max(0.05, self.height / max(h, 1))
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom).prerotate(rot), alpha=False)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
            self.done.emit(self.gen, p, img)
        d.close()


class SaveDialog(QDialog):
    """Asked on every Save: contents, page numbers, header, footer."""

    def __init__(self, parent, opt):
        super().__init__(parent)
        self.setWindowTitle("Save PDF - options")
        lay = QVBoxLayout(self)

        g1 = QGroupBox("Table of contents")
        f1 = QFormLayout(g1)
        self.chk_toc = QCheckBox("Add a contents page (only documents with 'In contents' ticked)")
        self.chk_toc.setChecked(opt.toc)
        self.ed_toc = QLineEdit(opt.toc_title)
        f1.addRow(self.chk_toc)
        f1.addRow("Title:", self.ed_toc)
        lay.addWidget(g1)

        g2 = QGroupBox("Page numbers")
        f2 = QFormLayout(g2)
        self.chk_num = QCheckBox("Add page numbers")
        self.chk_num.setChecked(opt.page_numbers)
        self.cmb_pos = QComboBox()
        self.cmb_pos.addItems(["bottom-center", "bottom-right", "top-right"])
        self.cmb_pos.setCurrentText(opt.number_pos)
        self.cmb_fmt = QComboBox()
        self.cmb_fmt.setEditable(True)
        self.cmb_fmt.addItems(["{page}", "Page {page}", "Page {page} of {pages}", "{page} / {pages}"])
        self.cmb_fmt.setCurrentText(opt.number_format)
        self.chk_count = QCheckBox("Count the contents pages (else page 1 = first document page)")
        self.chk_count.setChecked(opt.count_toc_pages)
        f2.addRow(self.chk_num)
        f2.addRow("Position:", self.cmb_pos)
        f2.addRow("Format:", self.cmb_fmt)
        f2.addRow(self.chk_count)
        lay.addWidget(g2)

        g3 = QGroupBox("Header and footer   (placeholders: {page} {pages} {title} {date} {file})")
        f3 = QFormLayout(g3)
        self.ed_head = QLineEdit(opt.header)
        self.ed_head.setPlaceholderText("empty = no header, e.g.  {title}")
        self.ed_foot = QLineEdit(opt.footer)
        self.ed_foot.setPlaceholderText("empty = no footer, e.g.  {file}   {date}")
        self.spn_size = QSpinBox()
        self.spn_size.setRange(6, 16)
        self.spn_size.setValue(opt.font_size)
        self.chk_marks_toc = QCheckBox("Also on the contents pages")
        self.chk_marks_toc.setChecked(opt.marks_on_toc)
        f3.addRow("Header:", self.ed_head)
        f3.addRow("Footer:", self.ed_foot)
        f3.addRow("Font size:", self.spn_size)
        f3.addRow(self.chk_marks_toc)
        lay.addWidget(g3)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def result_options(self):
        o = BuildOptions()
        o.toc = self.chk_toc.isChecked()
        o.toc_title = self.ed_toc.text().strip() or "Contents"
        o.page_numbers = self.chk_num.isChecked()
        o.number_pos = self.cmb_pos.currentText()
        o.number_format = self.cmb_fmt.currentText() or "{page}"
        o.count_toc_pages = self.chk_count.isChecked()
        o.header = self.ed_head.text()
        o.footer = self.ed_foot.text()
        o.font_size = self.spn_size.value()
        o.marks_on_toc = self.chk_marks_toc.isChecked()
        return o


COL_TOC, COL_TITLE, COL_LEVEL, COL_FILE, COL_PAGES = range(5)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION}")
        self.resize(1250, 760)
        self.setAcceptDrops(True)
        self.items = []
        self.workdir = tempfile.mkdtemp(prefix="pdfbinder_")
        self.options = BuildOptions()
        self.thumb_gen = 0
        self.worker = None
        self._filling = False

        # ---- left: documents ----
        left = QWidget()
        ll = QVBoxLayout(left)
        bar = QHBoxLayout()
        b_add = QPushButton("Add files...")
        b_add.clicked.connect(self.add_files_dialog)
        self.b_up = QPushButton("Up")
        self.b_up.clicked.connect(lambda: self.move_doc(-1))
        self.b_down = QPushButton("Down")
        self.b_down.clicked.connect(lambda: self.move_doc(1))
        self.b_del = QPushButton("Remove")
        self.b_del.clicked.connect(self.remove_doc)
        for b in (b_add, self.b_up, self.b_down, self.b_del):
            bar.addWidget(b)
        bar.addStretch()
        ll.addLayout(bar)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["In contents", "Title", "Level", "File", "Pages"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(COL_TOC, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(COL_TITLE, QHeaderView.Stretch)
        hh.setSectionResizeMode(COL_LEVEL, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(COL_FILE, QHeaderView.Interactive)
        hh.setSectionResizeMode(COL_PAGES, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(True)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.itemSelectionChanged.connect(self.show_thumbs)
        ll.addWidget(self.table)
        hint = QLabel("Drag files into the window. Tick 'In contents' and edit the title for the "
                      "table of contents; Level 1 = section, 2 = subsection, 3 = sub-subsection.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray")
        ll.addWidget(hint)

        # ---- right: preview ----
        right = QWidget()
        rl = QVBoxLayout(right)
        rbar = QHBoxLayout()
        rbar.addWidget(QLabel("Document:"))
        for txt, d in (("Rotate -90", -90), ("Rotate +90", 90)):
            b = QPushButton(txt)
            b.clicked.connect(lambda _=False, d=d: self.rotate(d, whole=True))
            rbar.addWidget(b)
        rbar.addSpacing(20)
        rbar.addWidget(QLabel("Selected pages:"))
        for txt, d in (("Rotate -90", -90), ("Rotate +90", 90)):
            b = QPushButton(txt)
            b.clicked.connect(lambda _=False, d=d: self.rotate(d, whole=False))
            rbar.addWidget(b)
        rbar.addStretch()
        rbar.addWidget(QLabel("Size"))
        self.zoom = QSlider(Qt.Horizontal)
        self.zoom.setRange(100, 600)
        self.zoom.setValue(220)
        self.zoom.setFixedWidth(120)
        self.zoom.sliderReleased.connect(self.show_thumbs)
        rbar.addWidget(self.zoom)
        rl.addLayout(rbar)

        self.thumbs = QListWidget()
        self.thumbs.setViewMode(QListWidget.IconMode)
        self.thumbs.setResizeMode(QListWidget.Adjust)
        self.thumbs.setMovement(QListWidget.Static)
        self.thumbs.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.thumbs.setSpacing(8)
        self.thumbs.itemDoubleClicked.connect(self.open_big)
        rl.addWidget(self.thumbs)
        rhint = QLabel("Click / Ctrl+click / Shift+click pages to select them. Double click = large view.")
        rhint.setStyleSheet("color: gray")
        rl.addWidget(rhint)

        split = QSplitter()
        split.addWidget(left)
        split.addWidget(right)
        split.setSizes([520, 730])

        # ---- bottom ----
        central = QWidget()
        cl = QVBoxLayout(central)
        cl.addWidget(split)
        bot = QHBoxLayout()
        self.lbl_info = QLabel("")
        bot.addWidget(self.lbl_info)
        bot.addStretch()
        b_about = QPushButton("i")
        b_about.setToolTip("About")
        b_about.setFixedSize(30, 30)
        b_about.setStyleSheet("font: bold italic 14px 'Georgia'; border-radius: 15px;"
                              "border: 1px solid #888")
        b_about.clicked.connect(self.show_about)
        bot.addWidget(b_about)
        b_save = QPushButton("Save PDF...")
        b_save.setStyleSheet("font-weight: bold; padding: 6px 18px")
        b_save.clicked.connect(self.save_pdf)
        bot.addWidget(b_save)
        cl.addLayout(bot)
        self.setCentralWidget(central)
        self.update_info()

    # ---------------- files ----------------
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        files = []
        for p in paths:
            if os.path.isdir(p):
                for root, _, names in os.walk(p):
                    files += [os.path.join(root, n) for n in sorted(names)]
            else:
                files.append(p)
        self.add_files(files)

    def add_files_dialog(self):
        flt = "Supported files (" + " ".join("*" + e for e in sorted(ALL_EXT)) + ");;All files (*)"
        files, _ = QFileDialog.getOpenFileNames(self, "Add files", "", flt)
        self.add_files(files)

    def add_files(self, files):
        files = [f for f in files if os.path.splitext(f)[1].lower() in ALL_EXT]
        if not files:
            return
        dlg = QProgressDialog("Converting...", "Cancel", 0, len(files), self)
        dlg.setWindowModality(Qt.WindowModal)
        dlg.setMinimumDuration(300)
        errors = []
        for k, f in enumerate(files):
            if dlg.wasCanceled():
                break
            dlg.setLabelText(f"Converting {os.path.basename(f)} ...")
            dlg.setValue(k)
            QApplication.processEvents()
            try:
                pdf = convert_to_pdf(f, self.workdir)
                item = DocItem(f, pdf)
                if item.page_count == 0:
                    raise ConvertError("no pages")
                self.items.append(item)
                self.append_row(item)
            except Exception as ex:
                errors.append(f"{os.path.basename(f)}: {ex}")
        dlg.setValue(len(files))
        self.update_info()
        if self.table.currentRow() < 0 and self.items:
            self.table.selectRow(0)
        if errors:
            QMessageBox.warning(self, APP_NAME, "Not added:\n\n" + "\n".join(errors))

    # ---------------- table ----------------
    def append_row(self, item):
        self._filling = True
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.fill_row(r, item)
        self._filling = False

    def fill_row(self, r, item):
        c = QTableWidgetItem()
        c.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        c.setCheckState(Qt.Checked if item.in_toc else Qt.Unchecked)
        self.table.setItem(r, COL_TOC, c)
        self.table.setItem(r, COL_TITLE, QTableWidgetItem(item.title))
        spn = QSpinBox()
        spn.setRange(1, 3)
        spn.setValue(item.level)
        spn.valueChanged.connect(lambda v, it=item: setattr(it, "level", v))
        self.table.setCellWidget(r, COL_LEVEL, spn)
        f = QTableWidgetItem(os.path.basename(item.src))
        f.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        f.setToolTip(item.src)
        self.table.setItem(r, COL_FILE, f)
        n = QTableWidgetItem(str(item.page_count))
        n.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.table.setItem(r, COL_PAGES, n)

    def on_item_changed(self, cell):
        if self._filling:
            return
        r = cell.row()
        if not 0 <= r < len(self.items):
            return
        it = self.items[r]
        if cell.column() == COL_TOC:
            it.in_toc = cell.checkState() == Qt.Checked
        elif cell.column() == COL_TITLE:
            it.title = cell.text()

    def refill_table(self, select=None):
        self._filling = True
        self.table.setRowCount(0)
        for it in self.items:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.fill_row(r, it)
        self._filling = False
        if select is not None and 0 <= select < len(self.items):
            self.table.selectRow(select)
        self.update_info()

    def move_doc(self, d):
        r = self.table.currentRow()
        n = r + d
        if r < 0 or not 0 <= n < len(self.items):
            return
        self.items[r], self.items[n] = self.items[n], self.items[r]
        self.refill_table(n)

    def remove_doc(self):
        r = self.table.currentRow()
        if r < 0:
            return
        del self.items[r]
        self.refill_table(min(r, len(self.items) - 1))
        if not self.items:
            self.thumbs.clear()

    def update_info(self):
        pages = sum(i.page_count for i in self.items)
        self.lbl_info.setText(f"{len(self.items)} document(s), {pages} page(s)")

    # ---------------- preview ----------------
    def current_item(self):
        r = self.table.currentRow()
        return self.items[r] if 0 <= r < len(self.items) else None

    def show_thumbs(self, keep_selection=None):
        it = self.current_item()
        if self.worker:
            self.worker.stop()
            self.worker.wait()
        self.thumb_gen += 1
        self.thumbs.clear()
        if not it:
            return
        h = self.zoom.value()
        self.thumbs.setIconSize(QSize(int(h * 0.8), h))
        for p in range(it.page_count):
            li = QListWidgetItem(f"{p + 1}" + (f"  ({it.extra_rot[p]} deg)" if it.extra_rot[p] else ""))
            li.setData(Qt.UserRole, p)
            li.setSizeHint(QSize(int(h * 0.8) + 16, h + 28))
            self.thumbs.addItem(li)
            if keep_selection and p in keep_selection:
                li.setSelected(True)
        self.worker = ThumbWorker(self.thumb_gen, it.pdf_path, list(it.extra_rot), h,
                                  list(range(it.page_count)))
        self.worker.done.connect(self.on_thumb)
        self.worker.start()

    def on_thumb(self, gen, p, img):
        if gen != self.thumb_gen or p >= self.thumbs.count():
            return
        self.thumbs.item(p).setIcon(QIcon(QPixmap.fromImage(img)))

    def rotate(self, delta, whole):
        it = self.current_item()
        if not it:
            return
        sel = sorted(li.data(Qt.UserRole) for li in self.thumbs.selectedItems())
        if whole:
            pages = range(it.page_count)
        else:
            if not sel:
                QMessageBox.information(self, APP_NAME, "Select one or more pages first.")
                return
            pages = sel
        it.rotate_pages(pages, delta)
        self.show_thumbs(keep_selection=set(sel))

    def open_big(self, li):
        it = self.current_item()
        if not it:
            return
        p = li.data(Qt.UserRole)
        dlg = QDialog(self)
        dlg.setWindowTitle(f"{os.path.basename(it.src)} - page {p + 1}")
        lay = QVBoxLayout(dlg)
        lbl = QLabel()
        scr = QApplication.primaryScreen().availableGeometry()
        lbl.setPixmap(page_pixmap(it.pdf_path, p, it.extra_rot[p], int(scr.height() * 0.8)))
        lay.addWidget(lbl)
        dlg.exec()

    # ---------------- about ----------------
    def show_about(self):
        dlg = QDialog(self)
        dlg.setWindowTitle(f"About {APP_NAME}")
        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(28, 24, 28, 18)
        lbl = QLabel(
            f"<h2 style='margin-bottom:2px'>{APP_NAME}</h2>"
            f"<div style='color:gray'>Version {APP_VERSION}</div>"
            "<p style='margin-top:16px; line-height:140%'>"
            "This program was made with love for my dear daughter <b>Nastia</b>,<br>"
            "so that she could hand in her homework a little faster<br>"
            "and have a little more time for what she truly loves &mdash;<br>"
            "instead of the dull work of converting pages.</p>"
            f"<p style='margin-top:18px'>&mdash; <a href='{AUTHOR_URL}'>{AUTHOR}</a></p>")
        lbl.setTextFormat(Qt.RichText)
        lbl.setOpenExternalLinks(True)
        lay.addWidget(lbl)
        bb = QDialogButtonBox(QDialogButtonBox.Ok)
        bb.accepted.connect(dlg.accept)
        lay.addWidget(bb)
        dlg.exec()

    # ---------------- save ----------------
    def save_pdf(self):
        if not self.items:
            QMessageBox.information(self, APP_NAME, "Add documents first.")
            return
        # titles typed but not yet committed
        self.table.setCurrentCell(-1, -1)
        dlg = SaveDialog(self, self.options)
        if dlg.exec() != QDialog.Accepted:
            return
        self.options = dlg.result_options()
        path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "binder.pdf", "PDF (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            build_pdf(self.items, path, self.options)
        except Exception as ex:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, APP_NAME, f"Not saved:\n{ex}")
            return
        QApplication.restoreOverrideCursor()
        QMessageBox.information(self, APP_NAME, f"Saved:\n{path}")

    def closeEvent(self, e):
        if self.worker:
            self.worker.stop()
            self.worker.wait(2000)
        shutil.rmtree(self.workdir, ignore_errors=True)
        super().closeEvent(e)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    w = MainWindow()
    if len(sys.argv) > 1:
        w.add_files(sys.argv[1:])
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
