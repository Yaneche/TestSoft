#!/usr/bin/env python3
"""
build_final.py — финальная сборка оптимизированного DOCX:
1. Находит все top-level v:group (векторные схемы) в word/document.xml базового optimized-DOCX.
2. Группы с >= MIN_SHAPES фигур растеризуются vml2png.SchemeRenderer в PNG.
3. В XML группа заменяется на <w:pict><v:rect><v:imagedata r:blip?/></v:rect></w:pict>-конструкцию
   (VML image, совместимый со старым содержимым w:pict), размеры берутся из style группы (pt).
4. PNG добавляются в word/media/, связи — в document.xml.rels, [Content_Types] дополняется png.
5. Пересобирается ZIP -> итоговый docx.

Запуск: python3 build_final.py <base.docx> <out.docx> [--min-shapes 100] [--width 1600]
"""
import re, os, sys, shutil, zipfile, argparse, traceback
from lxml import etree
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vml2png import SchemeRenderer, W, V, R, style_get

PIL_NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R_NSPREFIX = 'r'

def parse_pt(val, default=None):
    if val is None: return default
    m = re.match(r'^(-?[\d.]+)\s*(pt|mm|cm|in|px)?$', val.strip())
    if not m: return default
    x = float(m.group(1)); u = m.group(2)
    if u == 'mm': return x * 72 / 25.4
    if u == 'cm': return x * 72 / 2.54
    if u == 'in': return x * 72
    if u == 'px': return x * 72 / 96
    return x

def group_display_pt(g):
    """Размеры группы на странице в pt из её style."""
    st = g.get('style') or ''
    w = parse_pt((re.search(r'(?<![-\w])width\s*:\s*(-?[\d.]+(?:\.\d+)?)(pt|mm|cm|in|px)?', st) or [None,None])[1] if re.search(r'(?<![-\w])width\s*:', st) else None)
    mw = re.search(r'(?<![-\w])width\s*:\s*(-?[\d.]+(?:[eE]-?\d+)?)(pt|mm|cm|in|px)?', st)
    mh = re.search(r'(?<![-\w])height\s*:\s*(-?[\d.]+(?:[eE]-?\d+)?)(pt|mm|cm|in|px)?', st)
    def conv(m):
        if not m: return None
        v = float(m.group(1)); u = m.group(2)
        if u == 'mm': return v * 72 / 25.4
        if u == 'cm': return v * 72 / 2.54
        if u == 'in': return v * 72
        if u == 'px': return v * 72 / 96
        return v  # pt или без единиц (VML style по умолчанию pt)
    return conv(mw), conv(mh)

EMU_PER_PT = 12700

def make_pict_image(el_id, rid, w_pt, h_pt):
    """Строит <pic:pic> DrawingML? — нет: используем VML imagedata внутри w:pict (уже namespace v объявлен)."""
    xml = (
        f'<w:pict>'
        f'<v:rect id="Rasterized{el_id}" o:gfxdata="" style="position:absolute;width:{w_pt:.2f}pt;height:{h_pt:.2f}pt;z-index:1" '
        f'strokecolor="white" filled="f" stroked="f">'
        f'<v:fill opacity="0"/>'
        f'<v:imagedata r:id="{rid}" o:title="" cropbottom="0" croptop="0" cropleft="0" cropright="0"/>'
        f'</v:rect>'
        f'</w:pict>'
    )
    return etree.fromstring(
        xml.replace('<w:pict>', '<w:pict xmlns:w="%s" xmlns:v="%s" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="%s">' % (
            W.strip('{}'), V.strip('{}'), REL_NS))
    )

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('base_docx')
    ap.add_argument('out_docx')
    ap.add_argument('--min-shapes', type=int, default=100)
    ap.add_argument('--min-bytes', type=int, default=7000, help='растеризовать группы с XML-размером больше этого')
    ap.add_argument('--width', type=int, default=1600)
    args = ap.parse_args()

    workdir = '/workspace/analysis/final_build'
    if os.path.exists(workdir): shutil.rmtree(workdir)
    os.makedirs(workdir)
    with zipfile.ZipFile(args.base_docx) as z:
        z.extractall(workdir)

    doc_path = os.path.join(workdir, 'word', 'document.xml')
    media_dir = os.path.join(workdir, 'word', 'media')
    os.makedirs(media_dir, exist_ok=True)
    rels_path = os.path.join(workdir, 'word', '_rels', 'document.xml.rels')

    parser = etree.XMLParser(huge_tree=True)
    tree = etree.parse(doc_path, parser)
    root = tree.getroot()
    body = root.find(W + 'body')

    rt = etree.parse(rels_path, parser)
    rroot = rt.getroot()
    existing_ids = set()
    for rel in rroot:
        existing_ids.add(rel.get('Id'))
    def next_rid():
        i = 9000
        while f'rId{i}' in existing_ids: i += 1
        rid = f'rId{i}'
        existing_ids.add(rid)
        return rid

    rel_map = {}
    for rel in rroot:
        tgt = rel.get('Target', '')
        rel_map[rel.get('Id')] = os.path.basename(tgt.replace('../', ''))

    allg = body.findall('.//' + V + 'group')
    def in_group(el):
        p = el.getparent()
        while p is not None:
            if p.tag == V + 'group': return True
            p = p.getparent()
        return False
    tops = [g for g in allg if not in_group(g)]
    print(f'top-level groups: {len(tops)}')

    # --- 1. Удаление скрытых объектов (hidden="t") и их содержимого ---
    hidden_removed = 0
    for h in body.findall('.//' + V + 'shape') + body.findall('.//' + V + 'group'):
        st = h.get('style') or ''
        if 'visibility:hidden' in st.replace(' ', '') or h.get('hideway') == 't':
            par = h.getparent()
            if par is not None:
                par.remove(h)
                hidden_removed += 1
    # скрытые w:pict целиком (атрибут hidden у o:spMax/hid или style visibility:hidden)
    for p_ in body.findall('.//' + W + 'pict'):
        st_all = etree.tostring(p_, encoding='unicode')[:2000]
        if 'visibility:hidden' in st_all.replace(' ', '').replace('&quot;','"') and '<v:group' not in st_all:
            pass  # оставляем: массовое удаление risky, скрываем только явные shape/group выше
    print(f'removed hidden shapes/groups: {hidden_removed}')

    # перечитать список групп после удаления скрытых
    allg = body.findall('.//' + V + 'group')
    tops = [g for g in allg if not in_group(g)]

    # --- 2. Очистка мусора: пустые textbox-параграфы внутри оставшихся групп ---
    tb_cleared = 0
    for tb in body.findall('.//' + V + 'textbox'):
        texts = ''.join((t.text or '') for t in tb.iter(W + 't')).strip()
        if not texts:
            for ch in list(tb):
                tb.remove(ch)
            tb_cleared += 1
    print(f'cleared empty textboxes: {tb_cleared}')

    # --- 2b. Удаление дублирующихся w:pict (одинаковый XML внутри одного родителя) ---
    dedup_removed = 0
    for parent in list(body.iter()):
        seen = {}
        for ch in list(parent):
            if ch.tag == W + 'pict':
                key = etree.tostring(ch)
                if len(key) > 500:
                    if key in seen:
                        parent.remove(ch)
                        dedup_removed += 1
                    else:
                        seen[key] = ch
    print(f'deduplicated identical w:pict: {dedup_removed}')

    # перечитать список групп после дедупликации
    allg = body.findall('.//' + V + 'group')
    tops = [g for g in allg if not in_group(g)]

    # --- 3. Статистика по числу фигур для выбора порога ---
    from collections import Counter
    shape_counts = []
    for g in tops:
        shape_counts.append(len(g.findall('.//' + V + 'shape')))
    sc = sorted(shape_counts, reverse=True)
    total_shapes = sum(sc)
    cum = 0
    for i, v in enumerate(sc):
        cum += v
        if cum >= 0.95 * total_shapes:
            print(f'95% фигур содержатся в {i+1} крупнейших группах; порог = {v} фигур')
            break
    print('top20 sizes:', sc[:20])

    replaced = 0
    png_total = 0
    failed_idx = []
    for idx, g in enumerate(tops):
        nshapes = len(g.findall('.//' + V + 'shape'))
        gbytes = len(etree.tostring(g))
        # конвертируем если группа большая по XML ИЛИ содержит много фигур
        if nshapes < args.min_shapes and gbytes < args.min_bytes:
            continue
        cs = g.get('coordsize')
        if not cs or ',' not in cs:
            continue
        try:
            rend = SchemeRenderer(g, args.width, media_dir, rel_map)
            img = rend.render()
            bbox = rend.content_bbox_px()
            if bbox and (bbox[2] - bbox[0]) > 8 and (bbox[3] - bbox[1]) > 8:
                img = img.crop(bbox)
        except Exception as e:
            print(f'[skip {idx}] render error: {e}')
            traceback.print_exc()
            failed_idx.append(idx)
            continue

        # защита от полностью белого рендера (неудачная растеризация) — оставляем вектор
        try:
            import numpy as _np
            arr = _np.asarray(img.convert('L'))
            nonwhite = int((arr < 245).sum())
            if nonwhite < max(30, int(arr.size * 0.0002)):
                print(f'[skip {idx}] blank render ({nonwhite} non-white px), keeping vector')
                failed_idx.append(idx)
                continue
        except Exception:
            pass

        # масштаб растра: минимум для чёткости, но не избыточно
        w_pt_disp, h_pt_disp = group_display_pt(g)
        dpi_target = 150.0
        if w_pt_disp and w_pt_disp > 0:
            want_w = int(w_pt_disp / 72.0 * dpi_target)
            if want_w < img.size[0]:
                factor = want_w / img.size[0]
                new_h = max(int(img.size[1] * factor), 16)
                img = img.resize((max(want_w, 16), new_h), Image.LANCZOS)
        # очень узкие/низкие изображения не уменьшаем ниже разумного
        if img.size[0] < 40 or img.size[1] < 20:
            pass  # оставляем как есть

        w_pt, h_pt = w_pt_disp, h_pt_disp
        if not w_pt or not h_pt or w_pt <= 0 or h_pt <= 0:
            # fallback: из пикселей при целевом DPI
            w_pt = img.size[0] * 72 / dpi_target
            h_pt = img.size[1] * 72 / dpi_target
        # сохраняем PNG
        fname = f'image_vec{idx:03d}.png'
        outp = os.path.join(media_dir, fname)
        img.save(outp, optimize=True)
        png_total += os.path.getsize(outp)

        rid = next_rid()
        rel_el = etree.SubElement(rroot, 'Relationship')
        rel_el.set('Id', rid)
        rel_el.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        rel_el.set('Target', f'media/{fname}')

        newpict = make_pict_image(f'{idx:03d}', rid, w_pt, h_pt)
        parent = g.getparent()
        # идём вверх до элемента w:pict (группа обычно внутри w:pict)
        anc = g
        pict_el = None
        while anc is not None:
            if anc.tag == W + 'pict':
                pict_el = anc; break
            anc = anc.getparent()
        if pict_el is not None:
            pict_parent = pict_el.getparent()
            # заменяем весь w:pict на новый растровый pict
            pict_parent.replace(pict_el, newpict)
        else:
            parent.replace(g, newpict)
        replaced += 1
        print(f'[ok {idx}] shapes={nshapes} png={img.size} {os.path.getsize(outp)//1024}KB -> {rid}')

    print(f'replaced {replaced} vector groups with raster images; total PNG {png_total/1e6:.1f} MB')

    # content types: добавить png если нет
    ct_path = os.path.join(workdir, '[Content_Types].xml')
    ctree = etree.parse(ct_path, parser)
    croot = ctree.getroot()
    CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
    has_png = any(el.get('Extension', '').lower() == 'png' for el in croot)
    if not has_png and replaced:
        el = etree.SubElement(croot, CT + 'Default')
        el.set('Extension', 'png')
        el.set('ContentType', 'image/png')
    ctree.write(ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    rt.write(rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    # пересборка zip
    if os.path.exists(args.out_docx): os.remove(args.out_docx)
    with zipfile.ZipFile(args.out_docx, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        # mimetype-подобный порядок не важен для docx; просто всё
        for dirpath, dirnames, filenames in os.walk(workdir):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                arc = os.path.relpath(full, workdir)
                z.write(full, arc)
    sz_in = os.path.getsize(args.base_docx)
    sz_out = os.path.getsize(args.out_docx)
    print(f'base: {sz_in/1e6:.2f} MB -> final: {sz_out/1e6:.2f} MB')
    print(f'written: {args.out_docx}')

if __name__ == '__main__':
    main()
