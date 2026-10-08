"""
Перевод руководства A6-EC на русский язык с сохранением вёрстки.
- Текст перекрывается (redaction) и набирается заново кириллическим шрифтом.
- Изображения остаются как есть (не переводятся).
- Оглавление кликабельно (внутренние ссылки) + закладки PDF.
"""
import re, sys, json, os
import pymupdf as fitz
from translate_lib import translate_many, norm

SRC = "/workspace/A6-EC_series_servo_drive_manual.pdf"
OUT = "/workspace/A6-EC_servo_drive_manual_RU.pdf"

LIB = "/usr/share/fonts/truetype/liberation"
FR = LIB + "/LiberationSans-Regular.ttf"
FB = LIB + "/LiberationSans-Bold.ttf"
FI = LIB + "/LiberationSans-Italic.ttf"
FBI = LIB + "/LiberationSans-BoldItalic.ttf"
for p in (FR, FB, FI, FBI):
    assert os.path.exists(p), p

TOC_PAGES = (2, 3, 4, 5, 6)

class Line:
    __slots__ = ("bbox", "size", "color", "bold", "italic", "text", "ru", "align")
    def __init__(self, bbox, size, color, bold, italic, text):
        self.bbox, self.size, self.color, self.bold, self.italic = bbox, size, color, bold, italic
        self.text, self.ru, self.align = text, None, 0

SKIP_PATTERNS = [
    re.compile(r"^[\d\s.,;:%°±\-–—/\\|·•\u2605\u2606\u25cf\u25ca\u25a0\u25cb\u25d8\u2192\u2190\u2191\u2193]+$"),
    re.compile(r"^[A-Za-z]?[\d.]+[hHkKmMgGµnpf]\d*$"),
    re.compile(r"^(0x|0b)[0-9a-fA-F]+$"),
    re.compile(r"^\s*[\u2460-\u24ff]\s*$"),
]

def has_letters(t):
    return bool(re.search(r"[A-Za-zА-Яа-я]", t))

def needs_translation(t):
    if not has_letters(t):
        return False
    letters = re.sub(r"[^A-Za-z]", "", t)
    if len(letters) < 2:
        return False
    for pat in SKIP_PATTERNS:
        if pat.match(t.strip()):
            return False
    return True

def font_is_bold(fn):
    fn = fn.lower()
    return "bold" in fn or "+mi" in fn

def font_is_italic(fn):
    fn = fn.lower()
    return "italic" in fn or "oblique" in fn

def extract_page_lines(page):
    d = page.get_text("dict")
    spans = []
    for b in d["blocks"]:
        if b["type"] != 0:
            continue
        for l in b["lines"]:
            for s in l["spans"]:
                t = s["text"].replace("\x08", "")
                if not t.strip():
                    continue
                spans.append({
                    "bbox": tuple(s["bbox"]),
                    "origin": s["origin"],
                    "size": round(s["size"], 2),
                    "color": s["color"],
                    "font": s["font"],
                    "bold": font_is_bold(s["font"]) or bool(s["flags"] & 16),
                    "italic": font_is_italic(s["font"]) or bool(s["flags"] & 2),
                    "text": t,
                })
    spans.sort(key=lambda s: (round(s["origin"][1], 1), s["bbox"][0]))
    lines = []
    for s in spans:
        placed = False
        for ln in lines:
            if abs(ln["y"] - s["origin"][1]) <= max(1.5, 0.3 * s["size"]) and \
               abs(s["bbox"][0] - ln["x1"]) < 25:
                if abs(s["size"] - ln["size"]) > 1.5 or s["bold"] != ln["bold"]:
                    break
                ln["segs"].append(s)
                ln["x1"] = max(ln["x1"], s["bbox"][2])
                ln["bbox"] = (min(ln["bbox"][0], s["bbox"][0]), min(ln["bbox"][1], s["bbox"][1]),
                              max(ln["bbox"][2], s["bbox"][2]), max(ln["bbox"][3], s["bbox"][3]))
                placed = True
                break
        if not placed:
            lines.append({"y": s["origin"][1], "size": s["size"], "bold": s["bold"],
                          "italic": s["italic"], "color": s["color"],
                          "bbox": s["bbox"], "x1": s["bbox"][2], "segs": [s]})
    out = []
    for ln in lines:
        segs = sorted(ln["segs"], key=lambda s: s["bbox"][0])
        txt = ""
        prev = None
        for s in segs:
            if prev is not None and s["bbox"][0] - prev > 1.2 and txt and not txt.endswith(" "):
                gap = s["bbox"][0] - prev
                txt += " " if gap < 15 else "    "
            txt += s["text"]
            prev = s["bbox"][2]
        out.append(Line(ln["bbox"], ln["size"], ln["color"], ln["bold"], ln["italic"], txt.strip()))
    return out

JUST_LEFT, JUST_CENTER, JUST_RIGHT = 0, 1, 2

def assign_align(lines, pagew, margins=(57, 362)):
    L, R = margins
    for ln in lines:
        w = ln.bbox[2] - ln.bbox[0]
        x0, x1 = ln.bbox[0], ln.bbox[2]
        center_off = abs((x0 + x1) / 2 - (L + R) / 2)
        if x0 >= L - 3 and x1 >= R - 6 and w > 0.85 * (R - L):
            ln.align = JUST_LEFT
        elif center_off < 10 and x0 > L - 3:
            ln.align = JUST_CENTER
        elif x1 >= R - 6 and x0 > L + 20:
            ln.align = JUST_RIGHT
        else:
            ln.align = JUST_LEFT

_fonts_cache = {}
def get_font(name):
    if name not in _fonts_cache:
        _fonts_cache[name] = fitz.Font(fontfile={"ru": FR, "rub": FB, "rui": FI, "rubi": FBI}[name])
    return _fonts_cache[name]

def int_to_color(c):
    return (((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255)

def layout_line(ln, avail):
    fname = "rub" if (ln.bold and not ln.italic) else ("rubi" if (ln.bold and ln.italic) else ("rui" if ln.italic else "ru"))
    fnt = get_font(fname)
    size = ln.size
    def width_of(s, sz):
        return fnt.text_length(s, fontsize=sz)
    w = width_of(ln.ru, size)
    if w > avail:
        size = max(4.0, size * avail / w)
    words = ln.ru.split(" ")
    chunks = [""]
    for wd in words:
        trial = (chunks[-1] + " " + wd).strip()
        if width_of(trial, size) <= avail or not chunks[-1]:
            chunks[-1] = trial
        else:
            chunks.append(wd)
    return [(ch, size, fnt) for ch in chunks if ch]

def rewrite_page(doc, page, lines):
    changed = [ln for ln in lines if ln.ru]
    if not changed:
        return
    for ln in changed:
        r = fitz.Rect(ln.bbox) + (-0.6, -0.6, 0.6, 0.6)
        page.add_redact_annot(r)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    writers = {}
    for ln in changed:
        r = fitz.Rect(ln.bbox)
        avail = r.width + 8
        rows = layout_line(ln, avail)
        col = int_to_color(ln.color)
        if col not in writers:
            writers[col] = fitz.TextWriter(page.rect)
        tw = writers[col]
        n = len(rows)
        lh = rows[0][1] * 1.13
        need_h = n * lh
        y_top = r.y1 - need_h if need_h > r.height else r.y0
        for k, (ch, size, fnt) in enumerate(rows):
            wch = fnt.text_length(ch, fontsize=size)
            if ln.align == JUST_CENTER:
                x = r.x0 + (r.width - wch) / 2
            elif ln.align == JUST_RIGHT:
                x = r.x1 - wch
            else:
                x = r.x0
            y = y_top + k * lh + size * 0.82
            tw.append((x, y), ch, font=fnt, fontsize=size)
    for col, tw in writers.items():
        tw.write_text(page, color=col)

def find_heading_spans(doc, entry):
    """Return (rect, title_x, size) of the heading on its body page, or None."""
    num = entry["num"]
    page = doc[entry["phys"]]
    want = num.replace("Chapter", "").strip() if num.startswith("Chapter") else num
    pat_num = r"Chapter\s+\d+" if num.startswith("Chapter") else re.escape(num)
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        spans_all = [s for l in b["lines"] for s in l["spans"]]
        txt = "".join(s["text"] for s in spans_all).replace("\u2003", " ").replace("\t", " ")
        t = re.sub(r"\s+", " ", txt).strip()
        m = re.match(r"^(%s)\s+(\S.*)$" % pat_num, t)
        if not m:
            continue
        sizes = [s["size"] for s in spans_all]
        mx = max(sizes)
        if num.startswith("Chapter"):
            if mx < 14:
                continue
        else:
            dots = num.count(".")
            hs = 10.0 if dots == 1 else None
            if hs is None or abs(mx - hs) > 1.5:
                continue
            # must be near top of a text block, not inside a table row far down
        # locate first span whose text starts with the number
        x_title = None
        acc = ""
        for s in sorted(spans_all, key=lambda z: z["bbox"][0]):
            acc += s["text"].replace("\t", " ").replace("\u2003", " ")
            if re.match(r"^\s*(%s)\s+\S" % pat_num, acc):
                x_title = s["bbox"][2]
                break
        if x_title is None:
            x_title = spans_all[0]["bbox"][2]
        return fitz.Rect(b["bbox"]), x_title, mx
    return None

def rewrite_toc_page(doc, page, entries, pno):
    page_entries = [e for e in entries if e["toc_page"] == pno]
    if not page_entries:
        return
    for e in page_entries:
        r = fitz.Rect(e["x"], e["y"] - 1, 362, e["y"] + 11)
        page.add_redact_annot(r)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    tw = fitz.TextWriter(page.rect)
    fnt = get_font("ru")
    fb = get_font("rub")
    for e in page_entries:
        lvl = 0 if e["num"].startswith("Chapter") else min(e["num"].count("."), 3)
        size = 8
        x = e["x"]
        y = e["y"] + 8.5
        numtxt = e["num"].replace("Chapter", "Глава")
        use_font = fb if lvl == 0 else fnt
        tw.append((x, y), numtxt, font=use_font, fontsize=size)
        wnum = use_font.text_length(numtxt, fontsize=size)
        tx = x + wnum + (6 if lvl == 0 else 4)
        tw.append((tx, y), e["ru_title"], font=use_font, fontsize=size)
        link = fitz.Rect(x - 2, e["y"] - 1, 378, e["y"] + 11)
        page.insert_link({"kind": fitz.LINK_GOTO, "from": link,
                          "page": e["phys"], "to": fitz.Point(0, 0), "zoom": 0})
    tw.write_text(page, color=int_to_color(0x231F20))

def main():
    doc = fitz.open(SRC)
    entries = json.load(open("/tmp/toc_entries.json"))

    pages_lines = {}
    for i in range(doc.page_count):
        pages_lines[i] = extract_page_lines(doc[i])
        assign_align(pages_lines[i], doc[i].rect.width)

    texts = []
    for i in range(doc.page_count):
        if i in TOC_PAGES:
            continue
        for ln in pages_lines[i]:
            if needs_translation(ln.text):
                texts.append(ln.text)
    toc_texts = [e["title"] for e in entries if needs_translation(e["title"])]
    print(f"units: body={len(texts)}, toc={len(toc_texts)}", flush=True)

    cache_path = "/tmp/translated_body.json"
    body_map, toc_map = {}, {}
    if os.path.exists(cache_path):
        saved = json.load(open(cache_path))
        body_map, toc_map = saved.get("body", {}), saved.get("toc", {})

    todo = [t for t in dict.fromkeys(texts) if norm(t) not in body_map]
    B = 40
    for s in range(0, len(todo), B):
        batch = todo[s:s + B]
        trs = translate_many(batch)
        for a, b in zip(batch, trs):
            body_map[norm(a)] = b
        json.dump({"body": body_map, "toc": toc_map}, open(cache_path, "w"), ensure_ascii=False)
        print(f"body {min(s+B,len(todo))}/{len(todo)}", flush=True)

    todot = [t for t in dict.fromkeys(toc_texts) if norm(t) not in toc_map]
    if todot:
        trs = translate_many(todot)
        for a, b in zip(todot, trs):
            toc_map[norm(a)] = b
        json.dump({"body": body_map, "toc": toc_map}, open(cache_path, "w"), ensure_ascii=False)

    for i in range(doc.page_count):
        if i in TOC_PAGES:
            continue
        for ln in pages_lines[i]:
            key = norm(ln.text)
            if key in body_map and needs_translation(ln.text):
                ln.ru = body_map[key]
    for e in entries:
        e["ru_title"] = toc_map.get(norm(e["title"]), e["title"])

    total = doc.page_count
    for n in range(total):
        page = doc[n]
        if n in TOC_PAGES:
            rewrite_toc_page(doc, page, entries, n)
        else:
            rewrite_page(doc, page, pages_lines[n])
        if n % 25 == 0:
            print(f"rewrite {n}/{total}", flush=True)

    # chapter/section headings RU on their pages (only x.y and Chapter N)
    heading_done = []
    for e in entries:
        if not e["num"].startswith("Chapter") and e["num"].count(".") != 1:
            continue
        hb = find_heading_spans(doc, e)
        if not hb:
            continue
        rect, x_title, mx = hb
        page = doc[e["phys"]]
        num_ru = e["num"].replace("Chapter", "Глава")
        is_ch = e["num"].startswith("Chapter")
        page.add_redact_annot(rect + (-0.5, -0.5, 0.5, 0.5))
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        fnt = get_font("rub" if is_ch else "ru")
        size = mx
        avail = max(30.0, rect.x1 - x_title + 12)
        w = fnt.text_length(e["ru_title"], fontsize=size)
        if w > avail:
            size = max(6.0, size * avail / w)
        tw = fitz.TextWriter(page.rect)
        tw.append((rect.x0, rect.y1 - mx * 0.22), num_ru, font=fnt, fontsize=mx)
        tw.append((x_title, rect.y1 - size * 0.22), e["ru_title"], font=fnt, fontsize=size)
        tw.write_text(page, color=int_to_color(0x231F20))
        heading_done.append(e["num"])

    toc = [[1, "Титульный лист", 1], [1, "Юридическая информация", 2],
           [1, "Содержание", 3], [1, "Информация по безопасности", 8]]
    for e in entries:
        lvl = 1 if e["num"].startswith("Chapter") else min(1 + e["num"].count("."), 4)
        title = (("Глава " + e["num"].split()[1] + ": ") if e["num"].startswith("Chapter")
                 else e["num"] + " ") + e["ru_title"]
        toc.append([lvl, title, e["phys"] + 1])
    doc.set_toc(toc)
    doc.set_metadata({"title": "Руководство пользователя. Сервопривод серии A6-EC",
                      "author": "", "creator": "A6-EC manual RU"})
    doc.save(OUT, garbage=4, deflate=True, clean=True)
    print("saved", OUT, os.path.getsize(OUT))
    json.dump(entries, open("/tmp/toc_entries_ru.json", "w"), ensure_ascii=False)

if __name__ == "__main__":
    main()
