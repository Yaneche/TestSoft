#!/usr/bin/env python3
"""
Оптимизация DOCX (A6-EC manual):
1) document.xml: раскрытие "обёрточных" VML-групп (v:group из 1 child) с пересчётом
   геометрии и path-координат; удаление дублей v:path.
2) word/media: сжатие PNG/JPEG, удаление неиспользуемых файлов, чистка rels/Content_Types.
3) Пересборка ZIP с deflate level 9.
Использование: python3 optimize_v2.py <in.docx> <out.docx>
"""
import sys, os, re, io, shutil, zipfile, hashlib
from lxml import etree
from PIL import Image

SRC = sys.argv[1] if len(sys.argv) > 1 else '/workspace/A6-EC_series_servo_drive_manual.docx'
DST = sys.argv[2] if len(sys.argv) > 2 else '/workspace/A6-EC_series_servo_drive_manual_optimized.docx'
WORK = '/workspace/analysis/work2'
EXTRACT = os.path.join(WORK, 'src')

W   = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
VML = '{urn:schemas-microsoft-com:vml}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

def log(*a): print(*a, flush=True)

# ---------------- geometry helpers ----------------
NUM = re.compile(r'^(-?\d+(?:\.\d+)?)(pt|px|cm|mm|in)?$')

def plen(s):
    """Parse length -> (value, unit). Returns None on failure."""
    if s is None: return None
    m = NUM.match(s.strip())
    if not m: return None
    return float(m.group(1)), (m.group(2) or '')

def pnum(s):
    """Parse bare number (coord space)."""
    try: return float(s)
    except (TypeError, ValueError): return None

def fmt_num(v):
    iv = int(round(v))
    return str(iv)

def fmt_len(vu):
    v, u = vu
    s = ('%.4f' % v).rstrip('0').rstrip('.')
    return s + u

def parse_style(style):
    d = {}
    for part in (style or '').split(';'):
        if ':' in part:
            k, v = part.split(':', 1)
            d[k.strip().lower()] = v.strip()
    return d

def build_style(d, order=('position','left','top','width','height','z-index')):
    parts = []
    for k in order:
        if k in d: parts.append(f'{k}:{d[k]}')
    for k, v in d.items():
        if k not in order: parts.append(f'{k}:{v}')
    return ';'.join(parts)

def transform_path(path, sx, sy, ox, oy):
    """Scale+translate a VML path string. Absolute cmds: m l t q c ar ae; relative: r v nb."""
    toks = re.findall(r'[a-zA-Z]+|-?\d+(?:\.\d+)?', path)
    res = []
    i = 0
    n = len(toks)
    def numbers_until_cmd():
        js = i
        out = []
        while js < n and re.fullmatch(r'-?\d+(?:\.\d+)?', toks[js]):
            out.append(float(toks[js])); js += 1
        return out, js
    while i < n:
        t = toks[i]
        if re.fullmatch(r'[a-zA-Z]+', t):
            cmd = t.lower()
            res.append(t)
            i += 1
            if cmd == 'ar':
                nums, i = numbers_until_cmd()
                # x1 y1 xr yr x2 y2
                if len(nums) >= 6:
                    x1,y1,_,_,x2,y2 = nums[:6]
                    nx1, ny1 = ox + x1*sx, oy + y1*sy
                    nx2, ny2 = ox + x2*sx, oy + y2*sy
                    rx, ry = abs(nx2-nx1)/2, abs(ny2-ny1)/2
                    res += [fmt_num(nx1+rx), fmt_num(ny1+ry), fmt_num(rx), fmt_num(ry),
                            fmt_num(nx1-rx), fmt_num(ny1-ry)]
                    res += [fmt_num(x) for x in nums[6:]]
                else:
                    res += [fmt_num(x) for x in nums]
            elif cmd == 'ae':
                nums, i = numbers_until_cmd()
                if len(nums) >= 4:
                    cx, cy, rx, ry = nums[:4]
                    res += [fmt_num(ox+cx*sx), fmt_num(oy+cy*sy), fmt_num(rx*sx), fmt_num(ry*sy)]
                    res += [fmt_num(x) for x in nums[4:]]
                else:
                    res += [fmt_num(x) for x in nums]
            elif cmd in ('m','l','t','q','c','vb'):
                nums, i = numbers_until_cmd()
                for j in range(0, len(nums)//2*2, 2):
                    res += [fmt_num(ox + nums[j]*sx), fmt_num(oy + nums[j+1]*sy)]
                if len(nums) % 2: res.append(fmt_num(nums[-1]))
            elif cmd in ('r','v','nb'):
                nums, i = numbers_until_cmd()
                for j in range(0, len(nums)//2*2, 2):
                    res += [fmt_num(nums[j]*sx), fmt_num(nums[j+1]*sy)]
                if len(nums) % 2: res.append(fmt_num(nums[-1]))
            else:
                # commands with no coords (e,x,nf,..) – nothing to consume
                pass
        else:
            res.append(t); i += 1
    return ','.join(res) if False else ''.join(
        # VML path uses commas only inside pairs sometimes; safest: reproduce spacing-free like input
        [tok if re.fullmatch(r'[a-zA-Z]+', tok) else tok for tok in res])

def transform_path_join(path, sx, sy, ox, oy):
    """Same as transform_path but joins tokens preserving command letters adjacency."""
    out = []
    toks = re.findall(r'[a-zA-Z]+|,-?-?\d+(?:\.\d+)?(?:,-?\d+(?:\.\d+)?)?|-?\d+(?:\.\d+)?', path)
    # simpler: use char-level rebuild from transform_path token list
    return transform_path(path, sx, sy, ox, oy)

class Ctx:
    __slots__ = ('x','y','w','h')
    def __init__(self, x, y, w, h):
        self.x, self.y, self.w, self.h = x, y, w, h

def resolve(sty, ctx, key, scale_axis):
    """Resolve style value for key ('left','top','width','height') against context box.
       scale_axis: 'x' or 'y'. Returns (abs_pos_or_size_in_pt_units, ok)."""
    raw = sty.get(key)
    if raw is None:
        return None, False
    raw = raw.strip()
    if raw.endswith('%'):
        pct = pnum(raw[:-1])
        if pct is None: return None, False
        base = ctx.w if scale_axis == 'x' else ctx.h
        return base * pct / 100.0, True
    pl = plen(raw)
    if pl is None: return None, False
    return pl[0], True   # treat pt/px numbers as-is in the same unit space (Word mixes them; groups here use pt)

def unwrap_pict(pict):
    """Flatten wrapper groups inside one w:pict. Returns count of removed groups."""
    root_groups = [c for c in pict if isinstance(c.tag, str) and c.tag == f'{VML}group']
    if not root_groups:
        return 0
    psty = parse_style(root_groups[0].get('style'))
    # root group defines absolute box in pt
    rw, ok1 = resolve(psty, Ctx(0,0,1,1), 'width', 'x')
    rh, ok2 = resolve(psty, Ctx(0,0,1,1), 'height', 'y')
    rl, _ = resolve(psty, Ctx(0,0,1,1), 'left', 'x')
    rt, _ = resolve(psty, Ctx(0,0,1,1), 'top', 'y')
    if not (ok1 and ok2):
        return 0
    # strip position:absolute & page-relative props kept on root group itself (unchanged)
    removed = 0
    def walk(g, ctx):
        nonlocal removed
        kids = [c for c in g if isinstance(c.tag, str)]
        newkids = []
        for ch in kids:
            if ch.tag == f'{VML}group':
                st = parse_style(ch.get('style'))
                # try unwrap if single shape/group child
                sub = [c for c in ch if isinstance(c.tag, str)]
                can = False
                if len(sub) == 1 and sub[0].tag in (f'{VML}shape', f'{VML}group'):
                    child = sub[0]
                    extra = [k for k in ch.attrib if k.split('}')[-1] not in
                             ('id','style','coordsize','coordorigin')]
                    csty_raw = child.get('style','')
                    if (not extra and 'rotation' not in csty_raw and 'flip' not in csty_raw
                            and child.find(f'{VML}textbox') is None
                            and ch.find(f'{VML}textbox') is None):
                        can = True
                if can:
                    cx0, cy0 = ctx.x, ctx.y
                    cw, okw = resolve(st, ctx, 'width', 'x')
                    chh, okh = resolve(st, ctx, 'height', 'y')
                    cl, okl = resolve(st, ctx, 'left', 'x')
                    ct, okt = resolve(st, ctx, 'top', 'y')
                    if okw and okh and okl and okt:
                        nctx = Ctx(cx0+cl, cy0+ct, cw, chh)
                        removed += walk(ch, nctx)  # flatten inner first
                        # now unwrap this group: lift its (single remaining) child
                        subs = [c for c in ch if isinstance(c.tag, str)]
                        if len(subs) == 1:
                            child = subs[0]
                            sx = cw / ctx.w if ctx.w else 1.0
                            sy = chh / ctx.h if ctx.h else 1.0
                            ox = (cx0 + cl) - ctx.x
                            oy = (cy0 + ct) - ctx.y
                            cst = parse_style(child.get('style'))
                            # rewrite child style to parent space
                            def conv(key, axis):
                                raw = cst.get(key)
                                if raw is None: return
                                v, ok = resolve(cst, Ctx(ctx.x, ctx.y, ctx.w, ctx.h), key, axis) if raw.strip().endswith('%') else (plen(raw)[0] if plen(raw) else (None, False))[0] if True else None
                                # simpler explicit:
                                raw_s = raw.strip()
                                if raw_s.endswith('%'):
                                    base = ctx.w if axis=='x' else ctx.h
                                    nv = base * pnum(raw_s[:-1])/100.0
                                else:
                                    pv = plen(raw_s)
                                    if pv is None: return None
                                    nv = pv[0]
                                if key in ('width','height'):
                                    nv = nv * (sx if axis=='x' else sy)
                                else:
                                    off = ox if axis=='x' else oy
                                    nv = nv + off
                                cst[key] = fmt_len((nv, ''))
                                return True
                            if not (conv('left','x') and conv('top','y') and conv('width','x') and conv('height','y')):
                                newkids.append(ch); continue
                            # coord attrs: keep child's own (they are in child local space) — correct since
                            # child keeps its internal coord system; only outer box changes.
                            # BUT shapes drawn via 'path' in GROUP coord space need scaling:
                            if child.tag == f'{VML}shape':
                                pth = child.get('path')
                                co = child.get('coordorigin'); csz = child.get('coordsize')
                                if pth is not None:
                                    # path coords live in CHILD's coord system defined by ITS coordorigin/coordsize?
                                    # In Word exports shape coord* == group coord*, so path IS in group space -> scale it.
                                    if co == ch.get('coordorigin') and csz == ch.get('coordsize'):
                                        child.set('path', transform_path(pth, sx, sy, ox, oy))
                                        # also update imagedata aspect? images unaffected
                                    else:
                                        # unknown mapping -> abort unwrap for this node
                                        newkids.append(ch); continue
                            # apply new style
                            child.set('style', build_style(cst))
                            parent = ch.getparent()
                            idx = list(parent).index(ch)
                            parent.remove(ch)
                            parent.insert(idx, child)
                            removed += 1
                            newkids.append(child)
                            continue
                # not unwrappable -> recurse keeping group
                cw, okw = resolve(st, ctx, 'width', 'x')
                chh, okh = resolve(st, ctx, 'height', 'y')
                cl, okl = resolve(st, ctx, 'left', 'x')
                ct, okt = resolve(st, ctx, 'top', 'y')
                if okw and okh and okl and okt:
                    walk(ch, Ctx(ctx.x+cl, ctx.y+ct, cw, chh))
                newkids.append(ch)
            else:
                newkids.append(ch)
        return removed
    for rg in root_groups:
        walk(rg, Ctx(rl or 0, rt or 0, rw, rh))
    return removed

def dedup_vpath(tree):
    n = 0
    for shp in tree.iter(f'{VML}shape'):
        p1 = shp.find(f'{VML}path')
        if p1 is not None and len(p1.attrib) <= 1 and 'path' in shp.attrib:
            shp.remove(p1); n += 1
    return n

# ---------------- main ----------------
shutil.rmtree(WORK, ignore_errors=True)
os.makedirs(EXTRACT)
with zipfile.ZipFile(SRC) as z:
    names = z.namelist()
    z.extractall(EXTRACT)
log(f'[0] extracted {len(names)} entries, src={os.path.getsize(SRC)/1e6:.1f} MB')

total_unwrap = 0
for name in names:
    if not (name.endswith('.xml') or name.endswith('.rels')):
        continue
    path = os.path.join(EXTRACT, name)
    raw = open(path, 'rb').read()
    if b'v:group' not in raw:
        continue
    tree = etree.fromstring(raw)
    cnt = 0
    for pict in tree.iter(f'{W}pict'):
        cnt += unwrap_pict(pict)
    d = dedup_vpath(tree)
    if cnt or d:
        open(path, 'wb').write(etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True))
        total_unwrap += cnt
        log(f'[1] {name}: removed {cnt} wrapper groups, {d} dup v:path')
log(f'[1] total wrapper groups removed: {total_unwrap}')

# images
media_dir = os.path.join(EXTRACT, 'word', 'media')
saved = 0
if os.path.isdir(media_dir):
    for f in sorted(os.listdir(media_dir)):
        p = os.path.join(media_dir, f)
        orig = os.path.getsize(p)
        try:
            low = f.lower()
            if low.endswith('.png'):
                im = Image.open(p); im.load()
                has_alpha = im.mode in ('RGBA','LA') or (im.mode=='P' and 'transparency' in im.info)
                if has_alpha:
                    im2 = im.convert('RGBA')
                    buf = io.BytesIO(); im2.save(buf,'PNG',optimize=True)
                else:
                    rgb = im.convert('RGB')
                    q = rgb.quantize(colors=256, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE)
                    buf = io.BytesIO(); q.save(buf,'PNG',optimize=True)
                new = buf.getvalue()
                if len(new) < orig*0.92:
                    open(p,'wb').write(new); saved += orig-len(new)
            elif low.endswith(('.jpg','.jpeg')):
                im = Image.open(p).convert('RGB')
                buf = io.BytesIO(); im.save(buf,'JPEG',quality=82,optimize=True,progressive=True)
                new = buf.getvalue()
                if len(new) < orig*0.92:
                    open(p,'wb').write(new); saved += orig-len(new)
        except Exception as e:
            log(f'[2] skip {f}: {type(e).__name__} {e}')
log(f'[2] media compressed, saved {saved/1e6:.2f} MB')

# prune unreferenced media
used = set()
rels_files = [os.path.join(EXTRACT,'word','_rels',n.split('/')[-1])
              for n in names if n.startswith('word/_rels/')]
for rf in rels_files:
    if not os.path.exists(rf): continue
    for rel in etree.parse(rf).getroot():
        tgt = rel.get('Target','')
        if rel.get('TargetMode')=='External': continue
        used.add(tgt.replace('\\','/').split('/')[-1])
removed_media = 0
if os.path.isdir(media_dir):
    for f in list(sorted(os.listdir(media_dir))):
        if f not in used:
            os.remove(os.path.join(media_dir,f)); removed_media += 1
log(f'[3] removed {removed_media} unused media files')

def prune_rels(rf):
    if not os.path.exists(rf): return
    t = etree.parse(rf); r = t.getroot(); changed=False
    for rel in list(r):
        tgt = rel.get('Target','')
        fn = tgt.replace('\\','/').split('/')[-1]
        if 'media/' in tgt and not os.path.exists(os.path.join(media_dir,fn)):
            r.remove(rel); changed=True
    if changed: t.write(rf, xml_declaration=True, encoding='UTF-8')
for rf in rels_files: prune_rels(rf)

ctp = os.path.join(EXTRACT,'[Content_Types].xml')
if os.path.exists(ctp):
    t = etree.parse(ctp); r=t.getroot()
    for ov in list(r):
        pn = ov.get('PartName','')
        if pn.startswith('/word/media/') and not os.path.exists(os.path.join(media_dir,pn.split('/')[-1])):
            r.remove(ov)
    t.write(ctp, xml_declaration=True, encoding='UTF-8')

# rebuild zip
tmp = DST+'.tmp'
with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as z:
    ordered = ['[Content_Types].xml'] + [n for n in names if n.startswith('_rels/')] \
              + [n for n in names if n not in ('[Content_Types].xml',) and not n.startswith('_rels/')]
    seen=set()
    for n in ordered:
        lp = os.path.join(EXTRACT,n)
        if os.path.exists(lp) and n not in seen:
            z.write(lp,n); seen.add(n)
os.replace(tmp,DST)
old,new = os.path.getsize(SRC)/1e6, os.path.getsize(DST)/1e6
log(f'[4] {old:.1f} MB -> {new:.1f} MB ({(1-new/old)*100:.0f}% smaller)')

try:
    import docx
    d = docx.Document(DST)
    log(f'[V] python-docx OK: paras={len(d.paragraphs)} tables={len(d.tables)}')
except Exception as e:
    log(f'[V] FAILED: {e}')
