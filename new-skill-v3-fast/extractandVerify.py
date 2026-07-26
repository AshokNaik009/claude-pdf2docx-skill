#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_plumber.py — pdfplumber-style inspection for .docx WITHOUT LibreOffice/Word.

Python 3.9.6 compatible. Pure stdlib (zipfile + xml.etree). No pip installs needed.
Optional: if fontTools is installed AND you point --font at a real .ttf, character
widths become exact instead of table-based.

WHAT IT GIVES YOU
  1. Package validation  : required OOXML parts, relationship integrity,
                           well-formed XML for every part (catches corrupt docs).
  2. Geometry            : page size, margins, orientation, columns per section.
  3. Structure           : paragraphs (style, alignment, indent, spacing),
                           runs (font, size, bold/italic — resolved through
                           docDefaults -> style chain -> direct formatting),
                           tables (grid, widths, merges), inline images (px + EMU size).
  4. ESTIMATED layout    : line-wrap + pagination computed from font metrics,
                           i.e. an approximation of what a layout engine would do.
                           OOXML stores no page breaks — this is physically the
                           best any renderer-less tool can do.

USAGE
  python docx_plumber.py file.docx                 # human summary
  python docx_plumber.py file.docx --json out.json # full machine-readable report
  python docx_plumber.py file.docx --strict        # exit 1 on validation errors
"""
import sys, os, json, zipfile, argparse, posixpath, struct
import xml.etree.ElementTree as ET

W  = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R  = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
A  = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

TWIPS_PER_PT = 20.0
EMU_PER_PX   = 9525

# ---------------------------------------------------------------- font metrics
# EXACT per-character advance widths in 1/1000 em, extracted with fontTools from
# the official Plus Jakarta Sans variable font (Google Fonts, OFL), instantiated
# at wght=400 (regular) and wght=700 (bold). unitsPerEm=1000.
# Regenerate with --dump-metrics <font.ttf> if you switch typeface.
JAKARTA = {' ':170,'!':282,'"':395,'#':790,'$':646,'%':974,'&':734,"'":245,'(':295,
')':295,'*':479,'+':606,',':270,'-':498,'.':286,'/':429,'0':732,'1':371,
'2':600,'3':609,'4':630,'5':616,'6':597,'7':543,'8':628,'9':597,':':286,
';':306,'<':606,'=':606,'>':606,'?':499,'@':929,'A':652,'B':693,'C':781,
'D':742,'E':612,'F':581,'G':821,'H':736,'I':256,'J':383,'K':654,'L':531,
'M':836,'N':736,'O':878,'P':643,'Q':878,'R':641,'S':646,'T':512,'U':700,
'V':652,'W':984,'X':597,'Y':623,'Z':584,'[':315,'\\':429,']':315,
'^':606,'_':598,'`':352,'a':572,'b':667,'c':600,'d':667,'e':615,'f':393,
'g':662,'h':573,'i':229,'j':229,'k':548,'l':229,'m':908,'n':573,'o':655,
'p':667,'q':667,'r':338,'s':509,'t':388,'u':573,'v':517,'w':797,'x':487,
'y':536,'z':449,'{':317,'|':300,'}':317,'~':606}

JAKARTA_BOLD = {' ':178,'!':371,'"':498,'#':915,'$':647,'%':1046,'&':793,"'":311,
'(':379,')':379,'*':540,'+':660,',':357,'-':600,'.':378,'/':506,'0':708,
'1':403,'2':597,'3':611,'4':654,'5':615,'6':602,'7':559,'8':632,'9':602,
':':378,';':398,'<':660,'=':660,'>':660,'?':583,'@':927,'A':712,'B':691,
'C':774,'D':740,'E':595,'F':586,'G':808,'H':733,'I':279,'J':393,'K':679,
'L':543,'M':893,'N':741,'O':878,'P':648,'Q':878,'R':661,'S':647,'T':542,
'U':718,'V':697,'W':1020,'X':661,'Y':660,'Z':571,'[':416,'\\':506,
']':416,'^':660,'_':663,'`':381,'a':580,'b':672,'c':608,'d':672,'e':612,
'f':406,'g':652,'h':593,'i':252,'j':252,'k':577,'l':252,'m':924,'n':593,
'o':652,'p':672,'q':672,'r':374,'s':514,'t':413,'u':593,'v':566,'w':883,
'x':558,'y':586,'z':481,'{':398,'|':395,'}':398,'~':660}

# The typeface the exact width tables above describe.
METRICS_FONT = "Plus Jakarta Sans"
# Used only when a document declares NO font anywhere. Word and LibreOffice both
# fall back to a Times-metric serif (LO substitutes Liberation Serif).
# Override with --fallback-font if your renderer differs.
FALLBACK_FONT = "Times New Roman"
DEFAULT_FONT = METRICS_FONT
AVG      = 540    # measured avg lowercase advance for Plus Jakarta Sans
AVG_BOLD = 574

# Width of other fonts relative to Plus Jakarta Sans (Jakarta is a wide
# geometric sans; Calibri/Times are considerably narrower).
FONT_SCALE = {"plus jakarta sans":1.0,"jakarta sans":1.0,"plusjakartasans":1.0,
              "calibri":0.87,"arial":0.90,"helvetica":0.90,"times new roman":0.82,
              "cambria":0.85,"georgia":0.91,"verdana":1.02,"courier new":1.04,
              "tahoma":0.91,"segoe ui":0.89}

# Single-spaced line height in em, per font. Plus Jakarta Sans is taller than
# Calibri: hhea ascent 1038 / descent -222 / lineGap 0 -> 1.26 em natural.
# Calibri value calibrated against LibreOffice rendering.
LINE_HEIGHT = {"plus jakarta sans":1.26,"jakarta sans":1.26,"plusjakartasans":1.26,
               "calibri":1.15,"arial":1.15,"helvetica":1.15,"times new roman":1.15,
               "cambria":1.17,"georgia":1.14,"verdana":1.21,"courier new":1.13,
               "tahoma":1.21,"segoe ui":1.33}
LINE_HEIGHT_EM = 1.26   # default when font unknown

def line_height_em(font=None):
    return LINE_HEIGHT.get((font or DEFAULT_FONT).lower().strip(), LINE_HEIGHT_EM)

def text_width_pt(text, size_pt, font=None, bold=False):
    """Advance width in points. Exact for Plus Jakarta Sans; scaled for others."""
    name = (font or DEFAULT_FONT).lower().strip()
    scale = FONT_SCALE.get(name, 1.0)
    table, avg = (JAKARTA_BOLD, AVG_BOLD) if bold else (JAKARTA, AVG)
    units = sum(table.get(ch, avg) for ch in text)
    return units / 1000.0 * size_pt * scale

def dump_metrics(ttf_path, wght=None):
    """Regenerate the width tables from any TTF/OTF (needs fontTools)."""
    from fontTools.ttLib import TTFont
    f = TTFont(ttf_path)
    if wght is not None and "fvar" in f:
        from fontTools.varLib import instancer
        f = instancer.instantiateVariableFont(f, {"wght": wght}, inplace=False)
    up = f["head"].unitsPerEm
    cmap, hmtx = f.getBestCmap(), f["hmtx"]
    out = {}
    for c in range(32, 127):
        g = cmap.get(c)
        if g: out[chr(c)] = round(hmtx[g][0] * 1000.0 / up)
    hh = f["hhea"]
    return out, (hh.ascent - hh.descent + hh.lineGap) / up

# ---------------------------------------------------------------- package layer
REQUIRED = ["[Content_Types].xml", "_rels/.rels", "word/document.xml"]

def validate_package(zf):
    issues, names = [], set(zf.namelist())
    for p in REQUIRED:
        if p not in names:
            issues.append(("ERROR", "missing required part: " + p))
    for n in names:
        if n.endswith((".xml", ".rels")):
            try:
                ET.fromstring(zf.read(n))
            except ET.ParseError as e:
                issues.append(("ERROR", "malformed XML in %s: %s" % (n, e)))
    # every relationship target must exist
    for n in [x for x in names if x.endswith(".rels")]:
        base = posixpath.dirname(posixpath.dirname(n))
        try:
            for rel in ET.fromstring(zf.read(n)):
                if rel.get("TargetMode") == "External":
                    continue
                tgt = posixpath.normpath(posixpath.join(base, rel.get("Target", "")))
                if tgt.lstrip("/") not in names and tgt not in names:
                    issues.append(("ERROR", "%s -> dangling target %s" % (n, tgt)))
        except ET.ParseError:
            pass
    if "[Content_Types].xml" in names:
        ct = ET.fromstring(zf.read("[Content_Types].xml"))
        declared_exts = {d.get("Extension","").lower() for d in ct.iter(CT+"Default")}
        overrides = {o.get("PartName","").lstrip("/") for o in ct.iter(CT+"Override")}
        for n in names:
            ext = n.rsplit(".",1)[-1].lower() if "." in n else ""
            if n not in overrides and ext not in declared_exts and ext != "":
                issues.append(("WARN", "no content type for part: " + n))
    return issues

# ---------------------------------------------------------------- style resolution
def _shd_props(container, prefix):
    """Read <w:shd> fill into {prefix+'fill': '#RRGGBB', prefix+'shading_pattern': ...}.
    w:fill='auto' means 'no fill' and is skipped. Also resolves w:themeFill."""
    out = {}
    if container is None: return out
    shd = container.find(W+"shd")
    if shd is None: return out
    fill = shd.get(W+"fill")
    pattern = shd.get(W+"val")           # clear, solid, pct10, horzStripe, ...
    theme_fill = shd.get(W+"themeFill")
    if fill and fill.lower() not in ("auto",):
        out[prefix+"fill"] = "#" + fill.upper()
    elif theme_fill:
        slot = THEME_COLOR_KEYS.get(theme_fill, theme_fill)
        resolved = THEME_COLORS.get(slot)
        out[prefix+"fill"] = ("#" + resolved.upper()) if resolved else ("theme:" + theme_fill)
        out[prefix+"fill_from_theme"] = True
    # A pattern other than clear/nil paints even with fill=auto (uses w:color)
    if pattern and pattern not in ("clear", "nil") and prefix+"fill" not in out:
        pc = shd.get(W+"color")
        if pc and pc.lower() != "auto":
            out[prefix+"fill"] = "#" + pc.upper()
    if pattern and pattern not in ("clear", "nil"):
        out[prefix+"shading_pattern"] = pattern
    return out

def _rpr_props(rpr):
    p = {}
    if rpr is None: return p
    sz = rpr.find(W+"sz")
    if sz is not None and sz.get(W+"val"): p["size_pt"] = float(sz.get(W+"val"))/2.0
    fo = rpr.find(W+"rFonts")
    if fo is not None:
        explicit = fo.get(W+"ascii") or fo.get(W+"hAnsi") or fo.get(W+"cs")
        if explicit:
            p["font"] = explicit
        else:   # theme indirection: asciiTheme="minorHAnsi" etc.
            for att in ("asciiTheme","hAnsiTheme","cstheme"):
                tv = fo.get(W+att)
                if tv:
                    slot = THEME_KEYS.get(tv, "minor")
                    if THEME.get(slot):
                        p["font"] = THEME[slot]; p["font_from_theme"] = True
                    break
    for tag,key in ((W+"b","bold"),(W+"i","italic"),(W+"u","underline")):
        el = rpr.find(tag)
        if el is not None: p[key] = el.get(W+"val","true") not in ("false","0","none")
    hl = rpr.find(W+"highlight")           # named highlight colors (yellow, green, ...)
    if hl is not None and hl.get(W+"val","none") != "none":
        p["highlight"] = hl.get(W+"val")
    col = rpr.find(W+"color")              # glyph color
    if col is not None and col.get(W+"val","auto") != "auto":
        p["color"] = "#" + col.get(W+"val").upper()
    p.update(_shd_props(rpr, "run_"))      # run-level shading fill
    return p

def _ppr_props(ppr):
    p = {}
    if ppr is None: return p
    jc = ppr.find(W+"jc")
    if jc is not None: p["align"] = jc.get(W+"val")
    ind = ppr.find(W+"ind")
    if ind is not None:
        for k in ("left","right","firstLine","hanging"):
            v = ind.get(W+k)
            if v: p["indent_"+k+"_pt"] = float(v)/TWIPS_PER_PT
    sp = ppr.find(W+"spacing")
    if sp is not None:
        for k in ("before","after"):
            v = sp.get(W+k)
            if v: p["space_"+k+"_pt"] = float(v)/TWIPS_PER_PT
        ln = sp.get(W+"line")
        if ln:
            p["line_pt" if sp.get(W+"lineRule") in ("exact","atLeast") else "line_mult"] = \
                float(ln)/TWIPS_PER_PT if sp.get(W+"lineRule") in ("exact","atLeast") else float(ln)/240.0
    p.update(_shd_props(ppr, "para_"))     # paragraph-level shading fill
    return p

THEME = {}
THEME_COLORS = {}   # slot name (accent1, dk1, ...) -> RRGGBB

THEME_KEYS = {"minorHAnsi":"minor","minorAscii":"minor","minorBidi":"minor",
              "majorHAnsi":"major","majorAscii":"major","majorBidi":"major"}

# w:themeFill / w:themeColor tokens -> theme clrScheme slot names
THEME_COLOR_KEYS = {"accent1":"accent1","accent2":"accent2","accent3":"accent3",
    "accent4":"accent4","accent5":"accent5","accent6":"accent6",
    "text1":"dk1","text2":"dk2","background1":"lt1","background2":"lt2",
    "dark1":"dk1","dark2":"dk2","light1":"lt1","light2":"lt2",
    "hyperlink":"hlink","followedHyperlink":"folHlink"}

def load_theme_fonts(zf):
    """Resolve <a:latin> typefaces and clrScheme colors for theme slots."""
    global THEME_COLORS
    out = {}
    for n in zf.namelist():
        if n.startswith("word/theme/") and n.endswith(".xml"):
            try: root = ET.fromstring(zf.read(n))
            except ET.ParseError: continue
            for slot, tag in (("major","majorFont"),("minor","minorFont")):
                el = next(iter(root.iter(A+tag)), None)
                if el is not None:
                    lat = el.find(A+"latin")
                    if lat is not None and lat.get("typeface"):
                        out[slot] = lat.get("typeface")
            scheme = next(iter(root.iter(A+"clrScheme")), None)
            if scheme is not None:
                for child in scheme:
                    slot = child.tag.split("}")[1]     # dk1, lt1, accent1, ...
                    srgb = child.find(A+"srgbClr")
                    sysc = child.find(A+"sysClr")
                    if srgb is not None and srgb.get("val"):
                        THEME_COLORS[slot] = srgb.get("val")
                    elif sysc is not None and sysc.get("lastClr"):
                        THEME_COLORS[slot] = sysc.get("lastClr")
            break
    return out

def load_styles(zf):
    styles, defaults = {}, {"font":None,"size_pt":11.0}
    if "word/styles.xml" not in zf.namelist():
        return styles, defaults
    root = ET.fromstring(zf.read("word/styles.xml"))
    dd = root.find(W+"docDefaults")
    if dd is not None:
        rpd = dd.find(W+"rPrDefault")
        if rpd is not None:
            defaults.update(_rpr_props(rpd.find(W+"rPr")))
        ppd = dd.find(W+"pPrDefault")
        if ppd is not None:
            defaults.update(_ppr_props(ppd.find(W+"pPr")))
    for st in root.iter(W+"style"):
        sid = st.get(W+"styleId")
        based = st.find(W+"basedOn")
        styles[sid] = {
            "basedOn": based.get(W+"val") if based is not None else None,
            "run":  _rpr_props(st.find(W+"rPr")),
            "para": _ppr_props(st.find(W+"pPr")),
        }
    return styles, defaults

def resolve_chain(styles, sid, key):
    out, seen = {}, set()
    chain = []
    while sid and sid in styles and sid not in seen:
        seen.add(sid); chain.append(sid); sid = styles[sid]["basedOn"]
    for s in reversed(chain):
        out.update(styles[s][key])
    return out

# ---------------------------------------------------------------- document walk
def png_size(data):
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w,h = struct.unpack(">II", data[16:24]); return w,h
    return None

def _read_sectpr(sect):
    geom = {"page_w_pt":612.0,"page_h_pt":792.0,
            "margins_pt":{"top":72,"bottom":72,"left":72,"right":72},
            "header_pt":36.0,"footer_pt":36.0,
            "orientation":"portrait","columns":1,"col_space_pt":36.0,
            "header_refs":[],"footer_refs":[]}
    if sect is None: return geom
    pg = sect.find(W+"pgSz")
    if pg is not None:
        geom["page_w_pt"] = float(pg.get(W+"w",12240))/TWIPS_PER_PT
        geom["page_h_pt"] = float(pg.get(W+"h",15840))/TWIPS_PER_PT
        # Trust geometry over the declared attribute — generators often stamp
        # 'portrait' on swapped (landscape) dimensions.
        geom["orientation"] = "landscape" if geom["page_w_pt"] > geom["page_h_pt"] else "portrait"
        geom["orientation_declared"] = pg.get(W+"orient")
    mg = sect.find(W+"pgMar")
    if mg is not None:
        geom["margins_pt"] = {k: float(mg.get(W+k,1440))/TWIPS_PER_PT
                              for k in ("top","bottom","left","right")}
        geom["header_pt"] = float(mg.get(W+"header",720))/TWIPS_PER_PT
        geom["footer_pt"] = float(mg.get(W+"footer",720))/TWIPS_PER_PT
    cols = sect.find(W+"cols")
    if cols is not None:
        geom["columns"] = int(cols.get(W+"num",1))
        geom["col_space_pt"] = float(cols.get(W+"space",720))/TWIPS_PER_PT
    for tag,key in ((W+"headerReference","header_refs"),(W+"footerReference","footer_refs")):
        for ref in sect.findall(tag):
            geom[key].append(ref.get(R+"id"))
    return geom

def load_rels(zf):
    rels = {}
    if "word/_rels/document.xml.rels" in zf.namelist():
        for rel in ET.fromstring(zf.read("word/_rels/document.xml.rels")):
            rels[rel.get("Id")] = "word/" + rel.get("Target","").lstrip("/")
    return rels

def scan_fields(zf, part):
    """Return set of field codes (PAGE, NUMPAGES, ...) used in a header/footer part."""
    fields = set()
    if part not in zf.namelist(): return fields
    root = ET.fromstring(zf.read(part))
    for f in root.iter(W+"fldSimple"):
        fields.add((f.get(W+"instr") or "").strip().split(" ")[0])
    for it in root.iter(W+"instrText"):
        if it.text: fields.add(it.text.strip().split(" ")[0])
    return {f for f in fields if f}

def parse_document(zf):
    global THEME
    THEME = load_theme_fonts(zf)
    styles, defaults = load_styles(zf)
    if not defaults.get("font"):
        defaults["font"] = FALLBACK_FONT
        defaults["font_source"] = "renderer-fallback (document declares none)"
    else:
        defaults["font_source"] = "theme" if THEME else "docDefaults"
    rels = load_rels(zf)
    root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(W+"body")
    sections, cur_blocks = [], []
    for el in body:
        tag = el.tag.split("}")[1]
        if tag == "p":
            para = parse_paragraph(el, styles, defaults, zf)
            ppr = el.find(W+"pPr")
            sect = ppr.find(W+"sectPr") if ppr is not None else None
            cur_blocks.append(para)
            if sect is not None:                     # paragraph that ENDS a section
                sections.append((_read_sectpr(sect), cur_blocks)); cur_blocks = []
        elif tag == "tbl":
            cur_blocks.append(parse_table(el, styles, defaults, zf))
        elif tag == "sectPr":                        # final section
            sections.append((_read_sectpr(el), cur_blocks)); cur_blocks = []
    if cur_blocks:
        sections.append((_read_sectpr(None), cur_blocks))
    # resolve header/footer parts + detect page-number fields per section
    out_sections = []
    for geom, blocks in sections:
        hf = {"header_parts":[],"footer_parts":[],"fields":[]}
        for rid in geom.pop("header_refs"):
            p = rels.get(rid)
            if p: hf["header_parts"].append(p); hf["fields"] += sorted(scan_fields(zf,p))
        for rid in geom.pop("footer_refs"):
            p = rels.get(rid)
            if p: hf["footer_parts"].append(p); hf["fields"] += sorted(scan_fields(zf,p))
        hf["fields"] = sorted(set(hf["fields"]))
        hf["page_numbering"] = any(f in ("PAGE","NUMPAGES","SECTIONPAGES") for f in hf["fields"])
        out_sections.append({"geometry":geom,"headers_footers":hf,"blocks":blocks})
    return out_sections, styles, defaults

def parse_paragraph(p, styles, defaults, zf):
    ppr = p.find(W+"pPr")
    sid = None
    if ppr is not None:
        ps = ppr.find(W+"pStyle")
        if ps is not None: sid = ps.get(W+"val")
    para_props = dict(defaults)
    para_props.update(resolve_chain(styles, sid, "para"))
    para_props.update(_ppr_props(ppr))
    style_run = dict(defaults)
    style_run.update(resolve_chain(styles, sid, "run"))
    runs, images, page_break = [], [], False
    for r in p.iter(W+"r"):
        rp = dict(style_run); rp.update(_rpr_props(r.find(W+"rPr")))
        txt = "".join(t.text or "" for t in r.iter(W+"t"))
        if r.find(W+"br[@"+W+"type='page']") is not None: page_break = True
        for blip in r.iter(A+"blip"):
            rid = blip.get(R+"embed")
            img = {"rel_id": rid}
            ext = next(iter(r.iter(A+"ext")), None)
            if ext is not None:
                img["display_w_pt"] = int(ext.get("cx",0))/12700.0
                img["display_h_pt"] = int(ext.get("cy",0))/12700.0
            images.append(img)
        if txt or images:
            run = {"text": txt, "font": rp.get("font") or defaults.get("font") or FALLBACK_FONT,
                   "size_pt": rp.get("size_pt",11.0),
                   "bold": rp.get("bold",False), "italic": rp.get("italic",False)}
            for k in ("color","highlight","run_fill","run_shading_pattern"):
                if rp.get(k): run[k] = rp[k]
            runs.append(run)
    return {"type":"paragraph","style":sid,"props":para_props,"runs":runs,
            "images":images,"explicit_page_break":page_break,
            "text":"".join(r["text"] for r in runs)}

def parse_table(tbl, styles, defaults, zf):
    # Only THIS table's own grid — tbl.iter() would recurse into nested tables'
    # gridCol elements and wildly overstate the column count/width.
    tg = tbl.find(W+"tblGrid")
    grid = [float(g.get(W+"w",0))/TWIPS_PER_PT for g in (tg.findall(W+"gridCol") if tg is not None else [])]
    rows = []
    for tr in tbl.findall(W+"tr"):
        cells = []
        for tc in tr.findall(W+"tc"):
            tcpr = tc.find(W+"tcPr")
            span = 1; vmerge = None; w_pt = None
            if tcpr is not None:
                gs = tcpr.find(W+"gridSpan")
                if gs is not None: span = int(gs.get(W+"val",1))
                vm = tcpr.find(W+"vMerge")
                if vm is not None: vmerge = vm.get(W+"val","continue")
                tw = tcpr.find(W+"tcW")
                if tw is not None and tw.get(W+"type")=="dxa":
                    w_pt = float(tw.get(W+"w",0))/TWIPS_PER_PT
            tdir = None
            if tcpr is not None:
                td = tcpr.find(W+"textDirection")
                if td is not None: tdir = td.get(W+"val")
            shd = _shd_props(tcpr, "cell_")   # cell background fill
            paras = [parse_paragraph(p, styles, defaults, zf) for p in tc.findall(W+"p")]
            # NESTED TABLES: a cell may itself contain tables (the deterministic-fast
            # renderer lays multi-column pages out as an outer 1-row table whose cells
            # hold the column content — bars and data tables nested inside). Parse them
            # so their text, images and cell fills are not invisible to the inspector.
            nested = [parse_table(t, styles, defaults, zf) for t in tc.findall(W+"tbl")]
            cell_imgs = [img for p in paras for img in p["images"]]
            for nt in nested:
                for nr in nt["rows"]:
                    for nc in nr:
                        cell_imgs.extend(nc.get("images", []))
            nested_text = " ".join(nc["text"] for nt in nested for nr in nt["rows"] for nc in nr)
            cell = {"span":span,"vmerge":vmerge,"width_pt":w_pt,
                    "text_direction":tdir,"images":cell_imgs,
                    "n_images":len(cell_imgs),"fill":shd.get("cell_fill"),
                    "shading_pattern":shd.get("cell_shading_pattern"),
                    "tables":nested,
                    "text":" ".join([p["text"] for p in paras] + ([nested_text] if nested_text.strip() else [])),
                    "paras":paras}
            cells.append(cell)
        rows.append(cells)
    return {"type":"table","grid_pt":grid,"n_rows":len(rows),
            "n_cols":len(grid) or (len(rows[0]) if rows else 0),"rows":rows}

# ---------------------------------------------------------------- layout estimate
def wrap_lines(runs, avail_w_pt):
    """Greedy word-wrap using per-char metrics.
    Returns (line_count, max_font_size_pt, dominant_font_name)."""
    if not runs: return 1, 11.0, DEFAULT_FONT
    words, max_sz, main_font = [], 0.0, None
    for r in runs:
        if r["size_pt"] >= max_sz: main_font = r.get("font") or DEFAULT_FONT
        max_sz = max(max_sz, r["size_pt"])
        for w_ in r["text"].split(" "):
            words.append((w_, r))
    space_cache = {}
    lines, cur = 1, 0.0
    for w_, r in words:
        ww = text_width_pt(w_, r["size_pt"], r["font"], r["bold"])
        key = (r["size_pt"], r["font"], r["bold"])
        if key not in space_cache:
            space_cache[key] = text_width_pt(" ", r["size_pt"], r["font"], r["bold"])
        sw = space_cache[key] if cur > 0 else 0.0
        if cur + sw + ww > avail_w_pt and cur > 0:
            lines += 1; cur = ww
        else:
            cur += sw + ww
    return lines, (max_sz or 11.0), (main_font or DEFAULT_FONT)

def estimate_layout(sections):
    page, placements = 0, []
    total_table_imgs = 0
    for si, sec in enumerate(sections):
        g = sec["geometry"]
        ncol = max(g.get("columns",1),1)
        col_w = (g["page_w_pt"] - g["margins_pt"]["left"] - g["margins_pt"]["right"]
                 - (ncol-1)*g.get("col_space_pt",36.0)) / ncol
        avail_h = g["page_h_pt"] - g["margins_pt"]["top"] - g["margins_pt"]["bottom"]
        page += 1; col, y = 1, 0.0            # each section starts on a new page
        def advance(h, forced=False):
            nonlocal page, col, y
            if forced or (y + h > avail_h and y > 0):
                if col < ncol: col += 1
                else: page += 1; col = 1
                y = 0.0
            y += h
        for i, b in enumerate(sec["blocks"]):
            if b["type"] == "paragraph":
                props = b["props"]
                indent = props.get("indent_left_pt",0)+props.get("indent_right_pt",0)
                lines, sz, fnt = wrap_lines(b["runs"], max(col_w - indent, 36.0))
                line_h = props.get("line_pt") or sz*line_height_em(fnt)*props.get("line_mult",1.0)
                img_h = sum(img.get("display_h_pt",0) for img in b["images"])
                if b["explicit_page_break"]:
                    advance(0, forced=True)
                start_pg = page
                # flow line-by-line so paragraphs split across columns/pages
                y += props.get("space_before_pt",0)
                for _ in range(lines):
                    if y + line_h > avail_h and y > 0:
                        advance(0, forced=True)
                    y += line_h
                if img_h:
                    advance(img_h)
                y += props.get("space_after_pt",0)
                placements.append({"section":si,"type":"paragraph","page":start_pg,
                                   "page_end":page,"column":col,
                                   "est_height_pt":round(lines*line_h+img_h,1),
                                   "est_lines":lines,"text_preview":b["text"][:60]})
            else:
                start = page
                imgs_in_tbl = 0
                for row in b["rows"]:
                    row_h = 0.0
                    for cell in row:
                        cw = cell["width_pt"] or (sum(b["grid_pt"])/max(b["n_cols"],1))
                        if cell.get("text_direction") in ("btLr","tbRl","tbRlV","btLrV"):
                            cw = max(cw, 200)   # rotated text wraps along cell height
                        cell_h = 0.0
                        for p_ in cell["paras"]:
                            ln, sz, fnt = wrap_lines(p_["runs"], max(cw-10.8,20))
                            cell_h += ln*sz*line_height_em(fnt)
                        for img in cell.get("images",[]):
                            cell_h += img.get("display_h_pt",0); imgs_in_tbl += 1
                        row_h = max(row_h, cell_h + 4)
                    advance(row_h)
                total_table_imgs += imgs_in_tbl
                placements.append({"section":si,"type":"table","page_start":start,"page_end":page,
                                   "n_rows":b["n_rows"],"n_cols":b["n_cols"],
                                   "images_in_cells":imgs_in_tbl,
                                   "rotated_cells":sum(1 for r in b["rows"] for c in r
                                                       if c.get("text_direction"))})
    return {"estimated_pages":page,"placements":placements,
            "images_in_table_cells":total_table_imgs}

# ---------------------------------------------------------------- main
def inspect(path):
    zf = zipfile.ZipFile(path)
    issues = validate_package(zf)
    fatal = any(s == "ERROR" for s, _ in issues)
    if fatal:
        sections, styles, defaults = [], {}, {}
        layout = {"estimated_pages":0,"placements":[],"images_in_table_cells":0}
    else:
        sections, styles, defaults = parse_document(zf)
        layout = estimate_layout(sections)
    imgs = []
    for n in zf.namelist():
        if n.startswith("word/media/"):
            data = zf.read(n)
            e = {"part":n,"bytes":len(data)}
            px = png_size(data)
            if px: e["px"] = list(px)
            imgs.append(e)
    all_blocks = [b for s in sections for b in s["blocks"]]

    def _all_cells(tbl):
        """Yield every cell in a table AND in any table nested inside its cells."""
        for row in tbl["rows"]:
            for cell in row:
                yield cell
                for nt in cell.get("tables", []):
                    yield from _all_cells(nt)

    # Word count includes nested-table cell text (cell["text"] already folds it in,
    # but only recursion reaches deeply nested tables' own cells).
    n_words = 0
    for b in all_blocks:
        if b["type"] == "paragraph":
            n_words += len(b["text"].split())
        else:
            for cell in _all_cells(b):
                n_words += len(" ".join(p["text"] for p in cell.get("paras", [])).split())
    # tally shading so color changes are deterministic, not vision-dependent
    palette = {}
    cell_fills = para_fills = run_marks = 0
    for b in all_blocks:
        if b["type"] == "paragraph":
            f = b["props"].get("para_fill")
            if f: para_fills += 1; palette[f] = palette.get(f,0)+1
            for r in b["runs"]:
                for k in ("run_fill","highlight","color"):
                    if r.get(k):
                        run_marks += 1
                        v = r[k] if k!="highlight" else "highlight:"+r[k]
                        palette[v] = palette.get(v,0)+1
        else:
            for cell in _all_cells(b):
                if cell.get("fill"):
                    cell_fills += 1; palette[cell["fill"]] = palette.get(cell["fill"],0)+1
    shading = {"cell_fills":cell_fills,"paragraph_fills":para_fills,
               "run_marks":run_marks,
               "distinct_colors":sorted(palette, key=lambda k:-palette[k]),
               "palette_counts":palette}
    return {
        "file": os.path.basename(path),
        "validation": {"ok": not fatal,
                       "issues":[{"level":s,"msg":m} for s,m in issues]},
        "sections":[{"geometry":s["geometry"],"headers_footers":s["headers_footers"],
                     "n_blocks":len(s["blocks"])} for s in sections],
        "counts":{"sections":len(sections),"blocks":len(all_blocks),
                  "paragraphs":sum(1 for b in all_blocks if b["type"]=="paragraph"),
                  "tables":sum(1 for b in all_blocks if b["type"]=="table"),
                  "words":n_words,"media_files":len(imgs),
                  "styles_defined":len(styles),
                  "shaded_cells":cell_fills,"shaded_paragraphs":para_fills,
                  "run_color_marks":run_marks},
        "shading": shading,
        "default_run": defaults, "media": imgs,
        "estimated_layout": layout,
        "blocks": all_blocks,
    }

def main():
    ap = argparse.ArgumentParser(description="pdfplumber-style docx inspector (no soffice needed)")
    ap.add_argument("docx"); ap.add_argument("--json"); ap.add_argument("--strict", action="store_true")
    ap.add_argument("--fallback-font", default=FALLBACK_FONT,
                    help="font assumed when the document declares none (default: %s)" % FALLBACK_FONT)
    ap.add_argument("--dump-metrics", metavar="TTF",
                    help="regenerate width tables from a TTF/OTF and exit")
    ap.add_argument("--wght", type=float, default=None, help="variable-font weight for --dump-metrics")
    args = ap.parse_args()
    if args.dump_metrics:
        tbl, lh = dump_metrics(args.dump_metrics, args.wght)
        print(json.dumps({"widths":tbl,"natural_line_height_em":round(lh,3)}, indent=1)); return
    globals()["FALLBACK_FONT"] = args.fallback_font
    rep = inspect(args.docx)
    print("=== %s ===" % rep["file"])
    print("validation : %s (%d issue(s))" % ("OK" if rep["validation"]["ok"] else "FAILED",
                                             len(rep["validation"]["issues"])))
    for it in rep["validation"]["issues"][:10]:
        print("  [%s] %s" % (it["level"], it["msg"]))
    if not rep["validation"]["ok"]:
        sys.exit(1 if args.strict else 0)
    c, L = rep["counts"], rep["estimated_layout"]
    for i, s in enumerate(rep["sections"]):
        g, hf = s["geometry"], s["headers_footers"]
        print("section %d  : %.0fx%.0f pt %s, %d col(s), margins T%.0f/B%.0f/L%.0f/R%.0f"
              % (i+1, g["page_w_pt"], g["page_h_pt"], g["orientation"], g["columns"],
                 g["margins_pt"]["top"], g["margins_pt"]["bottom"],
                 g["margins_pt"]["left"], g["margins_pt"]["right"]))
        if hf["header_parts"] or hf["footer_parts"]:
            print("             headers/footers: %s | fields: %s | page numbering: %s"
                  % (", ".join(hf["header_parts"]+hf["footer_parts"]) or "-",
                     ", ".join(hf["fields"]) or "-", hf["page_numbering"]))
    print("metrics    : exact for %s | fallback %s | doc default %s (%s)"
          % (METRICS_FONT, FALLBACK_FONT, rep["default_run"].get("font"),
             rep["default_run"].get("font_source","-")))
    print("content    : %d paragraphs, %d tables, %d words, %d media file(s)"
          % (c["paragraphs"], c["tables"], c["words"], c["media_files"]))
    sh = rep["shading"]
    if sh["cell_fills"] or sh["paragraph_fills"] or sh["run_marks"]:
        print("shading    : %d shaded cell(s), %d shaded para(s), %d run mark(s) | colors: %s"
              % (sh["cell_fills"], sh["paragraph_fills"], sh["run_marks"],
                 ", ".join(sh["distinct_colors"][:8]) or "-"))
    for pl in L["placements"]:
        if pl["type"] == "table":
            print("table      : %dx%d on page %d-%d, %d image(s) in cells, %d rotated cell(s)"
                  % (pl["n_rows"], pl["n_cols"], pl["page_start"], pl["page_end"],
                     pl["images_in_cells"], pl["rotated_cells"]))
    print("ESTIMATED  : %d page(s)" % L["estimated_pages"])
    if args.json:
        with open(args.json, "w") as f:
            json.dump(rep, f, indent=2, ensure_ascii=False)
        print("full report: %s" % args.json)

if __name__ == "__main__":
    main()