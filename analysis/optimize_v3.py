#!/usr/bin/env python3
"""
Оптимизация DOCX (A6-EC manual). Безопасная версия:
1) document.xml: удаление VML-обёрток v:group, которые НИЧЕГО не меняют
   (дочерний элемент имеет идентичные style/coordorigin/coordsize => координаты
   уже в абсолютной системе группы, трансформации не требуются).
   Плюс удаление дублирующего пустого <v:path/> внутри v:shape.
2) media: сжатие PNG (без потери для малых/палитра) и JPEG q=82; удаление неиспользуемых.
3) Пересборка ZIP, deflate 9.
Использование: python3 optimize_v3.py <in.docx> <out.docx>
"""
import sys, os, re, io, shutil, zipfile
from lxml import etree
from PIL import Image

SRC = sys.argv[1] if len(sys.argv) > 1 else '/workspace/A6-EC_series_servo_drive_manual.docx'
DST = sys.argv[2] if len(sys.argv) > 2 else '/workspace/A6-EC_series_servo_drive_manual_optimized.docx'
WORK = '/workspace/analysis/work3'
EXTRACT = os.path.join(WORK, 'src')

W   = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
VML = '{urn:schemas-microsoft-com:vml}'

def log(*a): print(*a, flush=True)

def norm_style(s):
    """Canonical form of style string for comparison."""
    parts = []
    for p in (s or '').split(';'):
        p = p.strip()
        if not p or ':' not in p: continue
        k, v = p.split(':', 1)
        parts.append(f'{k.strip().lower()}:{v.strip().lower()}')
    return ';'.join(sorted(parts))

def unwrap_identical_groups(tree):
    """Remove v:group wrappers that are provably no-ops for rendering. Two accepted cases:
    A) strict identity: child style == group style and coordorigin/coordsize equal;
    B) pure geometric wrapper: single-child group whose child is a v:shape WITHOUT imagedata,
       with identical coordorigin/coordsize, no textbox descendants anywhere in the subtree,
       and only position/left/top/width/height/z-index style props on the group.
       (left/top/width/height are equal by VML spec => mapping is exactly the identity.)
    Iterate until stable (cascading removal)."""
    removed = 0
    while True:
        changed = False
        for g in tree.iter(f'{VML}group'):
            parent = g.getparent()
            if parent is None:
                continue
            kids = [c for c in g if isinstance(c.tag, str)]
            if len(kids) != 1:
                continue
            ch = kids[0]
            if ch.tag not in (f'{VML}shape', f'{VML}group'):
                continue
            # group must carry no attributes besides id/style/coord*
            extra = [k for k in g.attrib if k.split('}')[-1] not in
                     ('id', 'style', 'coordsize', 'coordorigin')]
            if extra:
                continue
            # coord systems must match -> path coordinates need no transform
            if g.get('coordorigin') != ch.get('coordorigin'):
                continue
            if g.get('coordsize') != ch.get('coordsize'):
                continue
            gs = norm_style(g.get('style'))
            cs = norm_style(ch.get('style'))
            case_a = (gs == cs)
            case_b = False
            if not case_a and ch.tag == f'{VML}shape':
                gd = {}
                for p in gs.split(';'):
                    if ':' in p:
                        k, v = p.split(':', 1); gd[k] = v
                cd = {}
                for p in cs.split(';'):
                    if ':' in p:
                        k, v = p.split(':', 1); cd[k] = v
                allowed = {'position', 'left', 'top', 'width', 'height', 'z-index'}
                if set(gd.keys()) <= allowed and 'position:absolute' in gd.get('position', '') + ';':
                    try:
                        def num(v): return float(re.sub(r'(pt|px)$', '', v))
                        same_pos = (abs(num(gd['left']) - num(cd.get('left', gd['left']))) < 0.05 and
                                    abs(num(gd['top']) - num(cd.get('top', gd['top']))) < 0.05 and
                                    abs(num(gd['width']) - num(cd.get('width', gd['width']))) < 0.05 and
                                    abs(num(gd['height']) - num(cd.get('height', gd['height']))) < 0.05)
                    except (ValueError, KeyError):
                        same_pos = False
                    if same_pos and ch.find(f'{VML}textbox') is None \
                            and g.find(f'{VML}imagedata') is None \
                            and ch.find(f'{VML}imagedata') is None:
                        case_b = True
            if not (case_a or case_b):
                continue
            # never lift subtrees containing textboxes or images via the group itself
            if ch.find(f'{VML}textbox') is not None:
                continue
            if g.findall(f'{VML}imagedata') and not ch.findall(f'.//{VML}imagedata'):
                continue
            idx = list(parent).index(g)
            parent.remove(g)
            parent.insert(idx, ch)
            removed += 1
            changed = True
        if not changed:
            break
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
    cnt = unwrap_identical_groups(tree)
    d = dedup_vpath(tree)
    if cnt or d:
        open(path, 'wb').write(etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True))
        total_unwrap += cnt
        log(f'[1] {name}: removed {cnt} wrapper groups, {d} dup v:path; new size {os.path.getsize(path)/1e6:.1f} MB')
log(f'[1] total wrapper groups removed: {total_unwrap}')

# images: conservative compression
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
                has_alpha = im.mode in ('RGBA', 'LA') or (im.mode == 'P' and 'transparency' in im.info)
                buf = io.BytesIO()
                if orig < 20_000:
                    # small icons: just re-save optimized (lossless)
                    im.save(buf, 'PNG', optimize=True)
                elif has_alpha:
                    im.convert('RGBA').save(buf, 'PNG', optimize=True)
                else:
                    rgb = im.convert('RGB')
                    q = rgb.quantize(colors=256, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE)
                    q.save(buf, 'PNG', optimize=True)
                new = buf.getvalue()
                if len(new) < orig * 0.9:
                    open(p, 'wb').write(new); saved += orig - len(new)
            elif low.endswith(('.jpg', '.jpeg')):
                im = Image.open(p).convert('RGB')
                buf = io.BytesIO(); im.save(buf, 'JPEG', quality=82, optimize=True, progressive=True)
                new = buf.getvalue()
                if len(new) < orig * 0.9:
                    open(p, 'wb').write(new); saved += orig - len(new)
        except Exception as e:
            log(f'[2] skip {f}: {type(e).__name__} {e}')
log(f'[2] media compressed, saved {saved/1e6:.2f} MB')

# prune unreferenced media
used = set()
rels_files = [os.path.join(EXTRACT, 'word', '_rels', n.split('/')[-1])
              for n in names if n.startswith('word/_rels/')]
for rf in rels_files:
    if not os.path.exists(rf): continue
    for rel in etree.parse(rf).getroot():
        if rel.get('TargetMode') == 'External': continue
        used.add(rel.get('Target', '').replace('\\', '/').split('/')[-1])
removed_media = 0
if os.path.isdir(media_dir):
    for f in list(sorted(os.listdir(media_dir))):
        if f not in used:
            os.remove(os.path.join(media_dir, f)); removed_media += 1
log(f'[3] removed {removed_media} unused media files')

def prune_rels(rf):
    if not os.path.exists(rf): return
    t = etree.parse(rf); r = t.getroot(); changed = False
    for rel in list(r):
        tgt = rel.get('Target', '')
        fn = tgt.replace('\\', '/').split('/')[-1]
        if 'media/' in tgt and not os.path.exists(os.path.join(media_dir, fn)):
            r.remove(rel); changed = True
    if changed: t.write(rf, xml_declaration=True, encoding='UTF-8')
for rf in rels_files: prune_rels(rf)

ctp = os.path.join(EXTRACT, '[Content_Types].xml')
if os.path.exists(ctp):
    t = etree.parse(ctp); r = t.getroot()
    for ov in list(r):
        pn = ov.get('PartName', '')
        if pn.startswith('/word/media/') and not os.path.exists(os.path.join(media_dir, pn.split('/')[-1])):
            r.remove(ov)
    t.write(ctp, xml_declaration=True, encoding='UTF-8')

# rebuild zip
tmp = DST + '.tmp'
with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    ordered = ['[Content_Types].xml'] \
        + [n for n in names if n.startswith('_rels/')] \
        + [n for n in names if n != '[Content_Types].xml' and not n.startswith('_rels/')]
    seen = set()
    for n in ordered:
        lp = os.path.join(EXTRACT, n)
        if os.path.exists(lp) and n not in seen:
            z.write(lp, n); seen.add(n)
os.replace(tmp, DST)
old, new = os.path.getsize(SRC)/1e6, os.path.getsize(DST)/1e6
log(f'[4] {old:.1f} MB -> {new:.1f} MB ({(1-new/old)*100:.0f}% smaller)')

try:
    import docx
    d = docx.Document(DST)
    log(f'[V] python-docx OK: paras={len(d.paragraphs)} tables={len(d.tables)}')
except Exception as e:
    log(f'[V] FAILED: {e}')
