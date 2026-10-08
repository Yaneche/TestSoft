#!/usr/bin/env python3
"""
vml2png — конвертер векторных схем (VML-групп) из word/document.xml в PNG.

Использует: lxml (разбор XML), Pillow (растеризация), numpy, PIL ImageFont (шрифты).
Автоматически определяет все top-level v:group внутри w:pict и рендерит каждый
как отдельный PNG с наложением подписей (v:textbox -> w:t) и встроенных картинок
(v:imagedata r:id -> word/media/*).

Запуск:  python3 vml2png.py <document.xml> <media_dir> <out_dir> [--min-shapes N] [--width PX]
"""
import re, os, sys, argparse
from lxml import etree
from PIL import Image, ImageDraw, ImageFont
import numpy as np

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
V = '{urn:schemas-microsoft-com:vml}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# ---------------- helpers ----------------
def parse_len(val, default=None):
    """'12pt', '3mm', '.25pt', '17', '4px' -> float в pt; если без единиц — трактуем как pt по умолчанию None."""
    if val is None: return default
    m = re.match(r'^(-?[\d.]+)\s*(pt|mm|cm|in|px)?$', val.strip())
    if not m: return default
    x = float(m.group(1)); u = m.group(2)
    if u == 'mm': return x * 72 / 25.4
    if u == 'cm': return x * 72 / 2.54
    if u == 'in': return x * 72
    if u == 'px': return x * 72 / 96
    return x  # pt или без единиц

def style_get(st, prop):
    m = re.search(r'(?<![-\w])' + prop + r'\s*:\s*(-?[\d.]+)(pt|mm|cm|in|px)?', st or '')
    if not m: return None
    v = float(m.group(1)); u = m.group(2)
    if u == 'mm': return v * 72 / 25.4
    if u == 'cm': return v * 72 / 2.54
    if u == 'in': return v * 72
    if u == 'px': return v * 72 / 96
    return v  # pt либо без единиц (VML path coords -> единицы координат группы, см. вызывающий код)

FONT_PATHS = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
]
def get_font(px):
    px = max(int(round(px)), 6)
    for fp in FONT_PATHS:
        if os.path.exists(fp):
            try: return ImageFont.truetype(fp, px)
            except Exception: pass
    return ImageFont.load_default()

# ---------------- VML path parser ----------------
VML_CMDS = set('mnhvqbcsxre')  # все командные буквы VML path

def tokenize_path(p):
    """Разбивает VML path на токены: однобуквенные команды (m,l,c,r,v,h,x,q,s,t,n,e...) и числа.
    'q'/'y'/'x' внутри путей — маркеры relative-to-origin, но они никогда не встречаются
    без соседних букв-команд; реальные многосимвольные команды: 'nl','ns','ae','qb','th','ar'.
    Разбиваем greedily: буква = команда; если за ней сразу идёт ещё буква и пара образует
    известную 2-буквенную команду — объединяем."""
    TWO = {'nl','ns','ae','qb','th','ar'}
    toks = []
    i, n = 0, len(p)
    while i < n:
        ch = p[i]
        if ch in ' \t\n\r,':
            i += 1; continue
        if ch.isdigit():
            j = i
            while j < n and p[j].isdigit(): j += 1
            toks.append(p[i:j]); i = j; continue
        if ch == '-':
            j = i + 1
            while j < n and p[j].isdigit(): j += 1
            if j > i + 1:
                toks.append(p[i:j]); i = j; continue
        if ch.isalpha():
            cmd = ch.lower()
            if i + 1 < n and p[i+1].isalpha():
                pair = (cmd + p[i+1].lower())
                if pair in TWO:
                    toks.append(pair); i += 2; continue
            toks.append(cmd); i += 1; continue
        i += 1
    return toks

def parse_vml_path(path, coordsize, coordorigin):
    """
    Возвращает список примитивов в ЕДИНИЦАХ КООРДИНАТ ГРУППЫ:
      ('line', [(x,y),(x,y)])
      ('rect', (x0,y0,x1,y1))
      ('poly', [(x,y),...], closed)
      ('bezier', pts_flattened, closed)
    """
    cx, cy = coordsize
    ox, oy = coordorigin
    toks = tokenize_path(path)
    prims = []
    i = 0
    cur = None          # текущая точка
    start_pt = None     # начало текущего подпути (для 'x' close)
    poly = []           # накапливаемый ломаный
    def is_num(t):
        return re.match(r'^-?\d+$', t)
    def read_nums(j, count=None):
        nums = []
        while j < len(toks) and is_num(toks[j]):
            nums.append(int(toks[j])); j += 1
            if count and len(nums) == count: break
        return nums, j
    def flush_poly(closed=False):
        nonlocal poly
        if len(poly) >= 2:
            prims.append(('poly', list(poly), closed))
        poly = []
    while i < len(toks):
        t = toks[i]
        if is_num(t):
            i += 1; continue
        cmd = t
        if cmd == 'm':
            flush_poly()
            nums, j = read_nums(i+1, 2)
            if len(nums) < 2: i += 1; continue
            cur = (nums[0], nums[1]); start_pt = cur; poly = [cur]; i = j
        elif cmd == 'l':
            nums, j = read_nums(i+1)
            k = 0
            while k+1 < len(nums):
                cur = (nums[k], nums[k+1]); poly.append(cur); k += 2
            i = j if nums else i+1
        elif cmd == 'c':
            nums, j = read_nums(i+1, 6)
            if len(nums) >= 6 and cur:
                x0,y0 = cur
                c1x,c1y,c2x,c2y,ex,ey = nums[:6]
                N = 16
                pts = []
                for s_ in range(N+1):
                    tt = s_/N; mt = 1-tt
                    bx = mt**3*x0 + 3*mt**2*tt*c1x + 3*mt*tt**2*c2x + tt**3*ex
                    by = mt**3*y0 + 3*mt**2*tt*c1y + 3*mt*tt**2*c2y + tt**3*ey
                    pts.append((bx,by))
                prims.append(('bezier', pts, False))
                cur = (ex, ey)
                i = j
            else:
                i += 1
        elif cmd == 'qb':   # квадратичная Безье: controlX controlY endX endY
            nums, j = read_nums(i+1, 4)
            if len(nums) >= 4 and cur:
                x0,y0 = cur
                qx,qy,ex,ey = nums[:4]
                N = 12
                pts = []
                for s_ in range(N+1):
                    tt = s_/N; mt = 1-tt
                    bx = mt*mt*x0 + 2*mt*tt*qx + tt*tt*ex
                    by = mt*mt*y0 + 2*mt*tt*qy + tt*tt*ey
                    pts.append((bx,by))
                prims.append(('bezier', pts, False))
                cur = (ex, ey)
                i = j
            else:
                i += 1
        elif cmd == 'th':   # штриховая линия (dash): та же семантика что l
            nums, j = read_nums(i+1)
            k = 0
            while k+1 < len(nums):
                cur = (nums[k], nums[k+1]); poly.append(cur); k += 2
            i = j if nums else i+1
        elif cmd == 'r':
            nums, j = read_nums(i+1, 2)
            if len(nums) < 2: i += 1; continue
            x, y = nums; i = j
            if cur:
                x0,y0 = cur
                prims.append(('rect', (min(x0,x), min(y0,y), max(x0,x), max(y0,y))))
        elif cmd == 'v':   # вертикальная линия abs-x текущего, abs-y
            nums, j = read_nums(i+1, 1)
            if not nums: i += 1; continue
            y = nums[0]; i = j
            if cur: cur = (cur[0], y); poly.append(cur)
        elif cmd == 'h':
            nums, j = read_nums(i+1, 1)
            if not nums: i += 1; continue
            x = nums[0]; i = j
            if cur: cur = (x, cur[1]); poly.append(cur)
        elif cmd == 'ar':  # elliptical arc: eX eY wR hR angS angE [clos] [rev] [ax ay]
            nums, j = read_nums(i+1)
            if len(nums) >= 6 and cur:
                ex_, ey_, wr, hr, as_, ae_ = nums[:6]
                import math
                cx_e = ex_ - wr*math.cos(math.radians(as_))
                cy_e = ey_ - hr*math.sin(math.radians(as_))
                a0, a1 = math.radians(as_), math.radians(ae_)
                if a1 < a0: a1 += 2*math.pi
                steps = max(8, int(abs(a1-a0)/ (math.pi/18)))
                pts = []
                for s_ in range(steps+1):
                    aa = a0 + (a1-a0)*s_/steps
                    pts.append((cx_e + wr*math.cos(aa), cy_e + hr*math.sin(aa)))
                prims.append(('bezier', pts, False))
                cur = pts[-1]
                i = j
            else:
                i += 1
        elif cmd == 'q':   # quadratic control point marker (part of qb handled above)
            i += 1
        elif cmd == 's' or cmd == 't':
            i += 1
        elif cmd == 'x':   # закрыть подпуть
            if len(poly) >= 2:
                if start_pt: poly.append(start_pt)
                prims.append(('poly', list(poly), True))
                poly = []
            i += 1
        elif cmd == 'nl':  # begin new line at current point
            if cur: poly = [cur]
            i += 1
        elif cmd == 'ns':  # begin new sector
            if cur: poly = [start_pt or cur, cur]
            i += 1
        elif cmd == 'ae':  # close + end
            if len(poly) >= 2:
                if start_pt: poly.append(start_pt)
                prims.append(('poly', list(poly), True))
                poly = []
            i += 1
        elif cmd == 'e':   # конец пути
            flush_poly()
            i += 1
        else:
            i += 1
    flush_poly()
    return prims

# ---------------- rasterizer ----------------
class SchemeRenderer:
    def __init__(self, group, target_width_px, media_dir, rel_map):
        self.group = group
        self.media_dir = media_dir
        self.rel_map = rel_map
        cs = group.get('coordsize')
        parts = [x for x in cs.split(',') if x.strip() != '']
        gx, gy = (float(parts[0]), float(parts[1])) if len(parts) >= 2 else (1.0, 1.0)
        if gx <= 0: gx = 1.0
        if gy <= 0: gy = gx
        self.gw, self.gh = gx, gy
        origin = group.get('coordorigin')
        self.ox, self.oy = 0.0, 0.0
        if origin:
            try:
                op = [float(x) for x in origin.replace(',', ' ').split() if x.strip() != '']
                if len(op) >= 2: self.ox, self.oy = op[0], op[1]
                elif len(op) == 1: self.ox = op[0]
            except ValueError:
                pass
        self.S = target_width_px / gx     # px per coord unit
        self.W = int(gx * self.S)
        self.H = int(gy * self.S) + 1
        # ширина группы в pt для перевода strokeweight
        stw = style_get(group.get('style'), 'width')
        self.pt_per_unit = (stw / gx) if (stw and stw > 0) else (1.0 / 15.0)  # fallback ~ стандарт Word

    def box_pt(self, el):
        """Возвращает (left, top, width, height) элемента в PT из его style.
        Если юниты не указаны — координаты считаются в единицах coordsize группы -> переводим в pt."""
        st = el.get('style') or ''
        g = self.group
        def conv(val, unit, prop):
            v = float(val)
            if unit == 'pt': return v
            if unit == 'mm': return v * 72 / 25.4
            if unit == 'cm': return v * 72 / 2.54
            if unit == 'in': return v * 72
            if unit == 'px': return v * 72 / 96
            if unit == '%':
                base = self.gw if prop in ('left', 'width') else self.gh
                return v / 100.0 * base * self.pt_per_unit
            # без единиц: если элемент — сама группа, это pt; иначе единицы coordsize
            if el is g:
                return v
            return v * self.pt_per_unit
        out = []
        for prop in ('left', 'top', 'width', 'height'):
            m = re.search(r'(?<![-\w])' + prop + r'\s*:\s*(-?[\d.]+(?:e-?\d+)?)(pt|mm|cm|in|px|%)?', st)
            out.append(conv(m.group(1), m.group(2), prop) if m else None)
        return tuple(out)

    def box_units(self, el):
        """(l,t,w,h) в ЕДИНИЦАХ КООРДИНАТ ГРУППЫ (для to_px)."""
        l, t, w, h = self.box_pt(el)
        inv = lambda v: (v / self.pt_per_unit) if v is not None else None
        return inv(l), inv(t), inv(w), inv(h)

    def to_px(self, x, y):
        return ((x - self.ox) * self.S, (y - self.oy) * self.S)

    def lw_px(self, shape):
        sw = shape.get('strokeweight')
        if sw is None:
            return max(round(0.75 * 72 * self.pt_per_unit * self.S), 1)
        pt = parse_len(sw, 0.75)
        units = pt / self.pt_per_unit
        return max(int(round(units * self.S)), 1)

    def render(self):
        im = Image.new('RGB', (self.W, self.H), 'white')
        dr = ImageDraw.Draw(im)
        textboxes = []
        pics = []
        for sh in self.group.iter(V + 'shape'):
            tb = sh.find(V + 'textbox')
            if tb is not None:
                textboxes.append((sh, tb))
                continue
            idata = sh.find(V + 'imagedata')
            if idata is not None:
                pics.append((sh, idata))
                continue
            self.draw_shape(dr, sh)
        for sh, idata in pics:
            self.paste_pic(im, sh, idata)
        for sh, tb in textboxes:
            self.draw_textbox(im, dr, sh, tb)
        return im

    def content_bbox_px(self):
        """Ограничивающий прямоугольник содержимого в px (для кадрирования)."""
        import math
        minx, miny, maxx, maxy = math.inf, math.inf, -math.inf, -math.inf
        def add(x0, y0, x1, y1):
            nonlocal minx, miny, maxx, maxy
            minx = min(minx, x0); miny = min(miny, y0)
            maxx = max(maxx, x1); maxy = max(maxy, y1)
        for sh in self.group.iter(V + 'shape'):
            st = sh.get('style') or ''
            def raw(prop):
                m = re.search(r'(?<![-\w])'+prop+r'\s*:\s*(-?[\d.]+)(\w*)', st)
                return float(m.group(1)) if m and not m.group(2) else None
            l, t, w, h = raw('left'), raw('top'), raw('width'), raw('height')
            if None in (l, t, w, h):
                continue  # position relative/absolute без явного box — пропускаем
            add(*self.to_px(l, t), *self.to_px(l + w, t + h))
        # subgroups со своими style
        for sg in self.group.findall('.//' + V + 'group'):
            if sg is self.group: continue
            st = sg.get('style') or ''
            def raw(prop):
                m = re.search(r'(?<![-\w])'+prop+r'\s*:\s*(-?[\d.]+)(\w*)', st)
                return float(m.group(1)) if m and not m.group(2) else None
            l, t, w, h = raw('left'), raw('top'), raw('width'), raw('height')
            if None in (l, t, w, h): continue
            add(*self.to_px(l, t), *self.to_px(l + w, t + h))
        if minx == math.inf:
            return None
        pad = 4
        return (max(int(minx) - pad, 0), max(int(miny) - pad, 0),
                min(int(math.ceil(maxx)) + pad, self.W), min(int(math.ceil(maxy)) + pad, self.H))

    def shape_coord_system(self, sh):
        """Возвращает (coordsize_x, coordsize_y, origin_x, origin_y) в которых задан path формы,
        и масштаб этих единиц в pt."""
        cs = sh.get('coordsize')
        org = sh.get('coordorigin')
        if cs and ',' in cs:
            try:
                cx, cy = [float(x) for x in cs.split(',') if x.strip() != ''][:2]
            except ValueError:
                return None
            ox, oy = 0.0, 0.0
            if org and ',' in org:
                try:
                    p = [float(x) for x in org.split(',') if x.strip() != '']
                    if len(p) >= 2: ox, oy = p[0], p[1]
                except ValueError:
                    pass
            if cx > 0 and cy > 0:
                return cx, cy, ox, oy
        return None

    def draw_shape(self, dr, sh):
        path = sh.get('path') or ''
        stroke = sh.get('strokecolor') or '#000000'
        if stroke.lower() in ('none', ''): 
            stroke = None
        fillc = sh.get('fillcolor')
        filled = sh.get('filled') == 't'
        lw = self.lw_px(sh)
        # box формы в единицах группы (style может быть в pt или без единиц = coordsize units)
        l, t, w, h = self.box_units(sh)
        scs = self.shape_coord_system(sh)
        if scs and path.strip():
            cx, cy, ox, oy = scs
            prims = parse_vml_path(path, (cx, cy), (ox, oy))
            if l is None: l = 0.0
            if t is None: t = 0.0
            if w is None: w = self.gw
            if h is None: h = self.gh
            bx0, by0 = self.to_px(l, t)
            bx1, by1 = self.to_px(l + w, t + h)
            bw = max(bx1 - bx0, 1); bh = max(by1 - by0, 1)
            sx = bw / max(cx, 1); sy = bh / max(cy, 1)
            def f2px(x, y):
                return (bx0 + (x - ox) * sx, by0 + (y - oy) * sy)
            for prim in prims:
                kind = prim[0]
                if kind == 'rect':
                    X0,Y0,X1,Y1 = prim[1]
                    a = f2px(X0,Y0); b = f2px(X1,Y1)
                    box = [a[0],a[1],b[0],b[1]]
                    if filled and fillc:
                        try: dr.rectangle(box, fill=fillc)
                        except Exception: pass
                    if stroke:
                        try: dr.rectangle(box, outline=stroke, width=lw)
                        except Exception: dr.rectangle(box, outline='black', width=1)
                elif kind in ('poly','bezier'):
                    pts = prim[1]; closed = prim[2] if len(prim)>2 else False
                    sp = [f2px(x,y) for x,y in pts]
                    if filled and fillc and len(sp)>=3:
                        try: dr.polygon(sp, fill=fillc)
                        except Exception: pass
                    if stroke:
                        if closed and len(sp)>=3:
                            sp2 = sp + [sp[0]]
                        else:
                            sp2 = sp
                        try: dr.line(sp2, fill=stroke, width=lw)
                        except Exception: pass
            return
        # нет path/coordsize — рисуем box из style в единицах группы
        if None in (l, t, w, h):
            return
        a = self.to_px(l, t); b = self.to_px(l + w, t + h)
        box=[a[0],a[1],b[0],b[1]]
        if filled and fillc:
            try: dr.rectangle(box, fill=fillc)
            except Exception: pass
        if stroke:
            try: dr.rectangle(box, outline=stroke, width=lw)
            except Exception: pass

    def paste_pic(self, im, sh, idata):
        rid = idata.get(R + 'id')
        if not rid or rid not in self.rel_map: return
        fn = self.rel_map[rid]
        path = os.path.join(self.media_dir, fn)
        if not os.path.exists(path): return
        try:
            pic = Image.open(path).convert('RGBA')
        except Exception:
            return
        l,t,w,h = self.box_units(sh)
        if None in (l,t,w,h) or w<=0 or h<=0:
            l,t = 0,0; w,h = self.gw, self.gh
        x0,y0 = self.to_px(l,t); x1,y1 = self.to_px(l+w,t+h)
        tw, th = max(int(x1-x0),1), max(int(y1-y0),1)
        pic = pic.resize((tw,th))
        bg = Image.new('RGBA', pic.size, (255,255,255,255))
        bg.paste(pic, mask=pic.split()[3])
        im.paste(bg.convert('RGB'), (int(x0), int(y0)))

    def draw_textbox(self, im, dr, sh, tb):
        # позиция бокса в единицах группы
        l,t,w,h = self.box_units(sh)
        if None in (l,t,w,h): return
        x0,y0 = self.to_px(l,t); x1,y1 = self.to_px(l+w,t+h)
        box_w = x1-x0; box_h = y1-y0
        # собираем абзацы
        paras = []
        for p in sh.iter(W+'p'):
            runs=[]
            align='left'; line=None
            pPr = p.find(W+'pPr')
            if pPr is not None:
                jc = pPr.find(W+'jc')
                if jc is not None: align = jc.get(W+'val') or 'left'
                sp = pPr.find(W+'spacing')
                if sp is not None:
                    lv = sp.get(W+'line')
                    if lv: line = int(lv)/240.0  # в twentieths? line=240 => single
            for r in p.findall(W+'r'):
                tel = r.find(W+'t')
                txt = tel.text if tel is not None else ''
                sz_pt = 8.0; color='#000000'; bold=False
                rPr = r.find(W+'rPr')
                if rPr is not None:
                    sze = rPr.find(W+'sz')
                    if sze is not None: sz_pt = int(sze.get(W+'val'))/2.0
                    cole = rPr.find(W+'color')
                    if cole is not None:
                        cv = cole.get(W+'val')
                        if cv and cv!='auto': color='#'+cv
                    if rPr.find(W+'b') is not None: bold=True
                runs.append((txt or '', sz_pt, color, bold))
            paras.append((runs, align, line))
        # раскладка: строки сверху вниз, выравнивание по центру/правому краю
        inset = tb.get('inset') or '0,0,0,0'
        ins = [parse_len(x,0) or 0 for x in inset.split(',')]  # pt
        pad_l = ins[3]*self.pt_per_unit*self.S
        pad_t = ins[0]*self.pt_per_unit*self.S
        cursor = y0 + pad_t
        for runs, align, line in paras:
            if not runs:
                continue
            text = ''.join(rr[0] for rr in runs)
            if text.strip()=='' and len(runs)==1:
                # пустая строка — пропуск минимальной высоты
                pass
            sz_pt = runs[0][1] if runs else 8.0
            color = runs[0][2] if runs else '#000000'
            font_px = max(int(round(sz_pt / self.pt_per_unit * self.S)), 6)
            # safety cap: шрифт не должен превышать высоту бокса
            font_px = min(font_px, max(int(box_h), 6))
            font = get_font(font_px)
            asc_desc = font.getbbox('Ag')
            lh = (asc_desc[3]-asc_desc[1]) * (line or 1.0) + 2
            # разбиваем на слова для переноса
            words = text.split(' ')
            lines=[]; cur_line=''
            for wd in words:
                trial = (cur_line+' '+wd).strip()
                if dr.textlength(trial, font=font) <= box_w - 2*pad_l or not cur_line:
                    cur_line = trial
                else:
                    lines.append(cur_line); cur_line = wd
            if cur_line: lines.append(cur_line)
            for ln in lines:
                tw = dr.textlength(ln, font=font)
                if align=='center': tx = x0 + (box_w - tw)/2
                elif align=='right': tx = x1 - tw - pad_l
                else: tx = x0 + pad_l
                if cursor + lh > y1 + 2: break
                dr.text((tx, cursor), ln, font=font, fill=color)
                cursor += lh
        # border?
        if sh.get('stroked')=='t':
            stroke = sh.get('strokecolor') or '#000'
            dr.rectangle([x0,y0,x1,y1], outline=stroke, width=self.lw_px(sh))

# ---------------- main ----------------
def load_rel_map(docx_path):
    z = etree.parse(os.path.join(os.path.dirname(docx_path),'_rels','document.xml.rels')) \
        if False else None
    return z

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('document_xml')
    ap.add_argument('media_dir')
    ap.add_argument('out_dir')
    ap.add_argument('--min-shapes', type=int, default=100, help='мин. число v:shape в группе для конвертации')
    ap.add_argument('--width', type=int, default=1600, help='ширина PNG в px')
    args = ap.parse_args()

    tree = etree.parse(args.document_xml)
    root = tree.getroot()
    body = root.find(W+'body')

    # relationship map rId -> file name
    rels_path = os.path.join(os.path.dirname(args.document_xml), '_rels', 'document.xml.rels')
    rel_map = {}
    if os.path.exists(rels_path):
        rt = etree.parse(rels_path).getroot()
        for rel in rt:
            tgt = rel.get('Target','')
            rel_map[rel.get('Id')] = os.path.basename(tgt.replace('../',''))

    allg = body.findall('.//'+V+'group')
    def in_group(el):
        p = el.getparent()
        while p is not None:
            if p.tag == V+'group': return True
            p = p.getparent()
        return False
    tops = [g for g in allg if not in_group(g)]

    os.makedirs(args.out_dir, exist_ok=True)
    done = 0
    for idx, g in enumerate(tops):
        nshapes = len(g.findall('.//'+V+'shape'))
        if nshapes < args.min_shapes:
            continue
        cs = g.get('coordsize')
        if not cs or ',' not in cs:
            continue
        try:
            rend = SchemeRenderer(g, args.width, args.media_dir, rel_map)
            img = rend.render()
            bbox = rend.content_bbox_px()
            if bbox and (bbox[2]-bbox[0]) > 8 and (bbox[3]-bbox[1]) > 8:
                img = img.crop(bbox)
        except Exception as e:
            print(f"[skip {idx}] error: {e}")
            continue
        out = os.path.join(args.out_dir, f"scheme_{idx:03d}_shapes{nshapes}.png")
        img.save(out, optimize=True)
        print(f"[ok {idx}] shapes={nshapes} size={img.size} -> {out}")
        done += 1
    print(f"converted: {done} schemes")

if __name__ == '__main__':
    main()
