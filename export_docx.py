"""Экспорт переведённого PDF в DOCX: текст по блокам + изображения + оглавление."""
import re, json, os
import pymupdf as fitz
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO

SRC = "/workspace/A6-EC_servo_drive_manual_RU.pdf"
OUT = "/workspace/A6-EC_servo_drive_manual_RU.docx"

doc = fitz.open(SRC)
entries = json.load(open("/tmp/toc_entries_ru.json"))

d = Document()
sec = d.sections[0]
p0 = doc[10]
sec.page_width = Cm(p0.rect.width / 72 * 2.54)
sec.page_height = Cm(p0.rect.height / 72 * 2.54)
sec.left_margin = Cm(1.4)
sec.right_margin = Cm(1.4)
sec.top_margin = Cm(1.2)
sec.bottom_margin = Cm(1.2)
style = d.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(8)

MAXIMG_W_CM = sec.page_width.cm - 2.8

def add_toc_field(document):
    para = document.add_paragraph()
    run = para.add_run()
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    instr.text = r'TOC \o "1-3" \h \z \u'
    fld_sep = OxmlElement("w:fldChar"); fld_sep.set(qn("w:fldCharType"), "separate")
    t_el = OxmlElement("w:t"); t_el.text = "Оглажение обновится по F9 (Обновить поле)."
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_sep)
    run._r.append(t_el); run._r.append(fld_end)

# ---- cover ----
c = doc[0].get_text().strip().splitlines()
title_p = d.add_heading("Руководство пользователя", level=0)
d.add_heading("Сервопривод серии A6-EC", level=1)
for line in c:
    if re.search(r"(Data code|Version|Power range)", line):
        d.add_paragraph(line.strip())
d.add_paragraph("Перевод на русский язык. Изображения сохранены без изменений.")

# ---- clickable TOC (Word field + hyperlinked entries) ----
d.add_heading("Содержание", level=1)
add_toc_field(d)

def color_of(cint):
    return RGBColor((cint >> 16) & 255, (cint >> 8) & 255, cint & 255)

heading_nums_ch = {e["num"] for e in entries if e["num"].startswith("Chapter")}
heading_nums_1 = {e["num"] for e in entries if not e["num"].startswith("Chapter") and e["num"].count(".") == 1}

# map phys page -> bookmark id to create hyperlink targets from TOC
bm_idx = 0
for pno in range(doc.page_count):
    page = doc[pno]
    if pno == 1:
        d.add_page_break()
    imgs = []
    seen_img = set()
    for im in page.get_image_info(xrefs=True):
        r = fitz.Rect(im["bbox"])
        if r.width < 10 or r.height < 10:
            continue
        key = (im["xref"], round(r.x0), round(r.y0))
        if key in seen_img:
            continue
        seen_img.add(key)
        imgs.append((r, im))
    blocks = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        spans = [s for l in b["lines"] for s in l["spans"]]
        txt = "".join(s["text"] for s in spans).replace("\x08", "").strip()
        if txt:
            mx = max(s["size"] for s in spans)
            blocks.append((fitz.Rect(b["bbox"]), txt, mx, b))
    # is this a chapter/section start page? insert Word heading instead of plain text block
    consumed_blocks = set()
    items = []
    for r, txt, mx, bl in blocks:
        t = re.sub(r"\s+", " ", txt.replace("\u2003", " ").replace("\t", " "))
        m = re.match(r"^(Глава\s+\d+|\d+\.\d+)\s+(\S.*)$", t)
        if m and m.group(1) in ("Глава %s" % m.group(1).split()[-1] if False else m.group(1),) or m:
            num = m.group(1)
            is_ch = num.startswith("Глава")
            ch_ok = is_ch and True
            lvl = 1 if is_ch else (2 if num.count(".") == 1 else None)
            if lvl and (is_ch or num in heading_nums_1) and mx >= 9:
                h = d.add_heading  # placeholder; we append later in order
                items.append(("h", r, (lvl, t)))
                consumed_blocks.add(id(bl))
                continue
        items.append(("t", r, (txt, mx, bl)))
    for r, im in imgs:
        items.append(("i", r, im))
    items.sort(key=lambda it: (it[1].y0, it[1].x0))
    for kind, r, obj in items:
        if kind == "i":
            try:
                pix = fitz.Pixmap(doc, obj["xref"])
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                png = pix.tobytes("png")
                w_cm = min(r.width / 72 * 2.54, MAXIMG_W_CM)
                para = d.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(BytesIO(png), width=Cm(max(1.0, w_cm)))
            except Exception:
                pass
            continue
        if kind == "h":
            lvl, t = obj
            global bm_idx
            bm_idx += 1
            head = d.add_heading("", level=lvl)
            run = head.add_run(t)
            # assign bookmark so TOC field picks headings automatically anyway
            continue
        txt, mx, bl = obj
        rect = r
        if rect.y1 < 40 or rect.y0 > page.rect.height - 30:
            continue  # running header/footer
        para = d.add_paragraph()
        if abs((rect.x0 + rect.x1) / 2 - page.rect.width / 2) < 12 and rect.width < page.rect.width * 0.8:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for l in bl["lines"]:
            for s in l["spans"]:
                t = s["text"].replace("\x08", "")
                if not t.strip():
                    continue
                run = para.add_run(t)
                run.font.size = Pt(min(s["size"], 20))
                fn = s["font"].lower()
                run.bold = "bold" in fn or bool(s["flags"] & 16)
                run.italic = "italic" in fn or bool(s["flags"] & 2)
                try:
                    run.font.color.rgb = color_of(s["color"])
                except Exception:
                    pass
    if pno % 25 == 0:
        print("docx page", pno, flush=True)

d.save(OUT)
print("saved", OUT, os.path.getsize(OUT))
