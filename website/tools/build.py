#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Complete Data & Image Extraction Pipeline (website/tools/build.py)
- Parses OpenXML .xlsx files (zipfile + xml.etree.ElementTree) AND BIFF8 OLE2 .xls files (struct).
- Supports BOTH standard OpenXML floating drawing anchors (xl/drawings/drawing*.xml -> twoCellAnchor / oneCellAnchor)
  AND WPS Office in-cell =DISPIMG("ID_...", 1) formulas mapped via xl/cellimages.xml and xl/_rels/cellimages.xml.rels,
  plus BIFF8 Escher MSODRAWINGGROUP / MSODRAWING / CONTINUE records.
- Converts and deduplicates every extracted image into two optimized WebP files in website/images/:
  - Thumbnail (p001_t.webp, max side 360px, quality 82)
  - Large high-res (p001_l.webp, max side 1100px, quality 85)
- Translates every product name and variant into English (en), Sinhala (si), Chinese (zh), and Spanish (es).
- Evaluates every product using the 100-Point Sea Cargo Sri Lanka Winning Product Index (SL-WPI).
- Outputs window.CATALOG = { ... }; inside website/data/products.js.
"""

import glob
import hashlib
import io
import json
import os
import re
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile
from PIL import Image

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WEBSITE_DIR = os.path.dirname(SCRIPT_DIR)
ROOT_DIR = os.path.dirname(WEBSITE_DIR)
IMAGES_DIR = os.path.join(WEBSITE_DIR, "images")
DATA_DIR = os.path.join(WEBSITE_DIR, "data")

sys.path.insert(0, SCRIPT_DIR)
from generate_winning_scores import compute_sl_wpi


# =============================================================================
# 1. EXCEL PARSERS (.xlsx OpenXML + WPS DISPIMG & .xls BIFF8 OLE2)
# =============================================================================

def col_to_idx(col_str):
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def parse_ref(ref):
    m = re.match(r"^([A-Z]+)(\d+)$", ref.upper())
    if not m:
        return 0, 0
    return int(m.group(2)) - 1, col_to_idx(m.group(1))


def parse_wps_cellimages(z):
    """
    Parse WPS Office in-cell =DISPIMG("ID_...", 1) mappings from
    xl/cellimages.xml and xl/_rels/cellimages.xml.rels.
    Returns dict: { "ID_...": raw_image_bytes }
    """
    wps_map = {}
    if "xl/cellimages.xml" not in z.namelist():
        return wps_map
    rels_map = {}
    rels_path = "xl/_rels/cellimages.xml.rels"
    if rels_path in z.namelist():
        rroot = ET.fromstring(z.read(rels_path))
        for rel in rroot:
            rid = rel.attrib.get("Id")
            target = "xl/" + rel.attrib.get("Target", "").lstrip("./").lstrip("../")
            rels_map[rid] = target
    croot = ET.fromstring(z.read("xl/cellimages.xml"))
    for pic in croot.iter():
        if pic.tag.endswith("pic"):
            img_id = None
            embed_id = None
            for sub in pic.iter():
                if sub.tag.endswith("cNvPr") and "name" in sub.attrib:
                    img_id = sub.attrib["name"]
                elif sub.tag.endswith("blip"):
                    for k, v in sub.attrib.items():
                        if k.endswith("embed"):
                            embed_id = v
            if img_id and embed_id and embed_id in rels_map:
                mpath = rels_map[embed_id]
                if mpath in z.namelist():
                    wps_map[img_id] = z.read(mpath)
    return wps_map


def parse_xlsx_file(path):
    """
    Parse all sheets of an .xlsx file using zipfile + xml.etree.ElementTree.
    Returns list of (sheet_name, rows_dict, merges_list, anchors_list)
    where anchors_list contains (from_row, to_row, from_col, to_col, raw_bytes).
    """
    results = []
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            for si in root.findall("s:si", ns):
                shared.append("".join((t.text or "") for t in si.findall(".//s:t", ns)))

        wps_images = parse_wps_cellimages(z)

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        for sheet in wb.findall(".//s:sheet", ns):
            sname = sheet.attrib["name"]
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = "xl/" + rel_map[rid].lstrip("/")
            if target not in z.namelist():
                continue
            sroot = ET.fromstring(z.read(target))
            rows = {}
            anchors = []

            for r in sroot.findall(".//s:row", ns):
                r_idx = int(r.attrib["r"]) - 1
                for c in r.findall("s:c", ns):
                    ref = c.attrib.get("r")
                    _, c_idx = parse_ref(ref)
                    t = c.attrib.get("t")
                    v_el = c.find("s:v", ns)
                    f_el = c.find("s:f", ns)
                    val = v_el.text if v_el is not None else ""
                    if t == "s" and val.isdigit():
                        val = shared[int(val)]
                    elif t == "inlineStr":
                        is_el = c.find(".//s:t", ns)
                        val = is_el.text if is_el is not None else ""
                    f_txt = f_el.text if (f_el is not None and f_el.text) else ""
                    # Check WPS =DISPIMG("ID_...", 1) formula in cell or value
                    disp_str = f_txt if "DISPIMG" in f_txt else (str(val) if "DISPIMG" in str(val) else "")
                    if disp_str:
                        dm = re.search(r'DISPIMG\(\s*"([^"]+)"', disp_str)
                        if dm and dm.group(1) in wps_images:
                            anchors.append((r_idx, r_idx, c_idx, c_idx, wps_images[dm.group(1)]))
                        continue
                    if val != "":
                        rows.setdefault(r_idx, {})[c_idx] = str(val).strip()

            if not rows:
                continue

            merges = []
            for m in sroot.findall(".//s:mergeCell", ns):
                p1, p2 = m.attrib["ref"].split(":")
                r1, c1 = parse_ref(p1)
                r2, c2 = parse_ref(p2)
                merges.append((r1, r2, c1, c2))

            # Extract floating drawings (twoCellAnchor / oneCellAnchor)
            sheet_rels_path = "xl/worksheets/_rels/" + target.split("/")[-1] + ".rels"
            if sheet_rels_path in z.namelist():
                srels = ET.fromstring(z.read(sheet_rels_path))
                for rel in srels:
                    if "drawing" in rel.attrib.get("Type", ""):
                        d_target = "xl/" + rel.attrib["Target"].lstrip("./").lstrip("../")
                        if d_target not in z.namelist():
                            continue
                        d_rels_path = "xl/drawings/_rels/" + d_target.split("/")[-1] + ".rels"
                        d_rels = {}
                        if d_rels_path in z.namelist():
                            for dr in ET.fromstring(z.read(d_rels_path)):
                                d_rels[dr.attrib["Id"]] = "xl/" + dr.attrib["Target"].lstrip("./").lstrip("../")
                        droot = ET.fromstring(z.read(d_target))
                        dns = {
                            "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                        }
                        for tag in ("xdr:twoCellAnchor", "xdr:oneCellAnchor"):
                            for anc in droot.findall(tag, dns):
                                fr = anc.find("xdr:from", dns)
                                to = anc.find("xdr:to", dns)
                                blip = anc.find(".//a:blip", dns)
                                if fr is not None and blip is not None:
                                    embed = blip.attrib.get(
                                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                                    )
                                    fc = int(fr.find("xdr:col", dns).text)
                                    fr_r = int(fr.find("xdr:row", dns).text)
                                    tc = int(to.find("xdr:col", dns).text) if to is not None else fc
                                    to_r = int(to.find("xdr:row", dns).text) if to is not None else fr_r
                                    mpath = d_rels.get(embed)
                                    if mpath and mpath in z.namelist():
                                        anchors.append((fr_r, to_r, fc, tc, z.read(mpath)))

            results.append((sname, rows, merges, anchors))
    return results


def parse_xls_file(path):
    """
    Pure-Python standard library parser for BIFF8 OLE2 (.xls) workbooks,
    including SST shared strings, cell records, MERGECELLS, and Escher
    MSODRAWINGGROUP BLIP images + MSODRAWING/CONTINUE ClientAnchors.
    """
    with open(path, "rb") as f:
        data = f.read()

    sector_shift = struct.unpack("<H", data[30:32])[0]
    sector_size = 1 << sector_shift
    num_fat_sectors = struct.unpack("<I", data[44:48])[0]
    dir_first_sector = struct.unpack("<I", data[48:52])[0]
    difat_first_sector = struct.unpack("<I", data[68:72])[0]
    num_difat_sectors = struct.unpack("<I", data[72:76])[0]

    difat = list(struct.unpack("<109I", data[76:512]))
    curr_difat = difat_first_sector
    for _ in range(num_difat_sectors):
        if curr_difat >= 0xFFFFFFFE:
            break
        pos = (curr_difat + 1) * sector_size
        entries = struct.unpack(f"<{sector_size // 4}I", data[pos : pos + sector_size])
        difat.extend(entries[:-1])
        curr_difat = entries[-1]
    difat = [s for s in difat if s < 0xFFFFFFFE]

    fat = []
    for s in difat[:num_fat_sectors]:
        pos = (s + 1) * sector_size
        fat.extend(struct.unpack(f"<{sector_size // 4}I", data[pos : pos + sector_size]))

    def get_chain_data(start_sec, size=None):
        out = bytearray()
        sec = start_sec
        while sec < 0xFFFFFFFE:
            pos = (sec + 1) * sector_size
            out.extend(data[pos : pos + sector_size])
            sec = fat[sec]
        return bytes(out[:size]) if size is not None else bytes(out)

    dir_data = get_chain_data(dir_first_sector)
    wb = None
    for i in range(0, len(dir_data), 128):
        entry = dir_data[i : i + 128]
        name_len = struct.unpack("<H", entry[64:66])[0]
        if name_len == 0:
            continue
        name = entry[: name_len - 2].decode("utf-16le", errors="ignore")
        if name in ("Workbook", "Book"):
            start_sec = struct.unpack("<I", entry[116:120])[0]
            stream_size = struct.unpack("<Q", entry[120:128])[0] & 0xFFFFFFFF
            wb = get_chain_data(start_sec, stream_size)
            break

    if not wb:
        return []

    def parse_sst_chunks(chunks):
        if not chunks:
            return []
        _, unique_strings = struct.unpack("<II", chunks[0][:8])
        c_idx = 0
        pos = 8

        def ensure_bytes():
            nonlocal c_idx, pos
            while c_idx < len(chunks) and pos >= len(chunks[c_idx]):
                c_idx += 1
                pos = 0
            if c_idx >= len(chunks):
                raise EOFError()

        def read_raw(n):
            nonlocal c_idx, pos
            ensure_bytes()
            res = chunks[c_idx][pos : pos + n]
            pos += len(res)
            while len(res) < n:
                c_idx += 1
                pos = 0
                need = n - len(res)
                part = chunks[c_idx][pos : pos + need]
                res += part
                pos += len(part)
            return res

        strings = []
        for _ in range(unique_strings):
            try:
                ensure_bytes()
            except EOFError:
                break
            char_count = struct.unpack("<H", read_raw(2))[0]
            flags = read_raw(1)[0]
            is_16bit = bool(flags & 0x01)
            has_ext = bool(flags & 0x04)
            has_rich = bool(flags & 0x08)
            rich_runs = struct.unpack("<H", read_raw(2))[0] if has_rich else 0
            ext_size = struct.unpack("<I", read_raw(4))[0] if has_ext else 0

            chars_left = char_count
            s_parts = []
            while chars_left > 0:
                if pos >= len(chunks[c_idx]):
                    c_idx += 1
                    pos = 0
                    cont_flag = chunks[c_idx][pos]
                    pos += 1
                    is_16bit = bool(cont_flag & 0x01)
                avail_bytes = len(chunks[c_idx]) - pos
                avail_chars = (avail_bytes // 2) if is_16bit else avail_bytes
                take = min(chars_left, avail_chars)
                if is_16bit:
                    raw = chunks[c_idx][pos : pos + take * 2]
                    pos += take * 2
                    s_parts.append(raw.decode("utf-16le", errors="replace"))
                else:
                    raw = chunks[c_idx][pos : pos + take]
                    pos += take
                    s_parts.append(raw.decode("latin1", errors="replace"))
                chars_left -= take

            if rich_runs > 0:
                read_raw(4 * rich_runs)
            if ext_size > 0:
                read_raw(ext_size)
            strings.append("".join(s_parts))
        return strings

    def decode_rk_val(rk):
        is_mul100 = bool(rk & 0x01)
        is_int = bool(rk & 0x02)
        if is_int:
            val = float(rk >> 2)
            if rk & 0x80000000:
                val = float((rk - 0x100000000) >> 2)
        else:
            raw = struct.pack("<II", 0, rk & ~3)
            val = struct.unpack("<d", raw)[0]
        if is_mul100:
            val /= 100.0
        return val

    pos = 0
    sst_chunks = []
    dgg_chunks = []
    in_sst = False
    in_dgg = False
    sheets = []

    while pos + 4 <= len(wb):
        rid, rlen = struct.unpack("<HH", wb[pos : pos + 4])
        rdata = wb[pos + 4 : pos + 4 + rlen]
        pos += 4 + rlen
        if rid == 0x0085:  # BOUNDSHEET
            bof_pos, vis, stype = struct.unpack("<IBB", rdata[:6])
            slen = rdata[6]
            sflags = rdata[7]
            sname = (
                rdata[8 : 8 + slen * 2].decode("utf-16le", errors="replace")
                if (sflags & 1)
                else rdata[8 : 8 + slen].decode("latin1", errors="replace")
            )
            sheets.append((sname, bof_pos))
            in_sst = in_dgg = False
        elif rid == 0x00FC:  # SST
            sst_chunks = [rdata]
            in_sst = True
            in_dgg = False
        elif rid == 0x00EB:  # MSODRAWINGGROUP
            dgg_chunks.append(rdata)
            in_dgg = True
            in_sst = False
        elif rid == 0x003C:  # CONTINUE
            if in_sst:
                sst_chunks.append(rdata)
            elif in_dgg:
                dgg_chunks.append(rdata)
        else:
            in_sst = False
            in_dgg = False

    sst = parse_sst_chunks(sst_chunks)

    # Extract BLIP images from BStoreContainer (0xF001) inside DggContainer (0xF000)
    bse_images = {}  # 1-indexed BSE pib -> raw PNG/JPEG bytes
    dgg = b"".join(dgg_chunks)
    if len(dgg) >= 8:
        p = 8
        while p + 8 <= len(dgg):
            vi, f, l = struct.unpack("<HHI", dgg[p : p + 8])
            if f == 0xF001:  # BStoreContainer
                bp = p + 8
                bend = min(len(dgg), p + 8 + l)
                bse_idx = 1
                while bp + 8 <= bend:
                    bvi, bf, bl = struct.unpack("<HHI", dgg[bp : bp + 8])
                    raw = dgg[bp + 8 + 36 : bp + 8 + bl]
                    img_start = -1
                    for sig in (b"\x89PNG", b"\xff\xd8\xff"):
                        idx = raw.find(sig)
                        if idx != -1 and (img_start == -1 or idx < img_start):
                            img_start = idx
                    if img_start >= 0:
                        bse_images[bse_idx] = raw[img_start:]
                    bse_idx += 1
                    bp += 8 + bl
            p += 8 + l

    results = []
    for sname, bof_pos in sheets:
        p = bof_pos
        rows = {}
        merges = []
        sh_chunks = []
        in_draw = False

        while p + 4 <= len(wb):
            rid, rlen = struct.unpack("<HH", wb[p : p + 4])
            rdata = wb[p + 4 : p + 4 + rlen]
            p += 4 + rlen
            if rid == 0x000A:  # EOF
                break
            if rid == 0x00EC:  # MSODRAWING
                sh_chunks.append(rdata)
                in_draw = True
            elif rid == 0x003C and in_draw:  # CONTINUE inside drawing block
                sh_chunks.append(rdata)
            elif rid in (0x005D, 0x01B6) and in_draw:  # OBJ / TXO interleaved with drawing
                pass
            else:
                in_draw = False
                if rid == 0x00FD:  # LABELSST
                    r, c, xf, s_idx = struct.unpack("<HHHI", rdata[:10])
                    val = sst[s_idx] if s_idx < len(sst) else ""
                    if str(val).strip():
                        rows.setdefault(r, {})[c] = str(val).strip()
                elif rid == 0x027E:  # RK
                    r, c, xf, rk = struct.unpack("<HHHI", rdata[:10])
                    rows.setdefault(r, {})[c] = str(decode_rk_val(rk))
                elif rid == 0x00BD:  # MULRK
                    r, fc = struct.unpack("<HH", rdata[:4])
                    lc = struct.unpack("<H", rdata[-2:])[0]
                    for i in range(lc - fc + 1):
                        xf, rk = struct.unpack("<HI", rdata[4 + i * 6 : 10 + i * 6])
                        rows.setdefault(r, {})[fc + i] = str(decode_rk_val(rk))
                elif rid == 0x0203:  # NUMBER
                    r, c, xf, val = struct.unpack("<HHHd", rdata[:14])
                    rows.setdefault(r, {})[c] = str(val)
                elif rid == 0x00E5:  # MERGECELLS
                    cnt = struct.unpack("<H", rdata[:2])[0]
                    for i in range(cnt):
                        r1, r2, c1, c2 = struct.unpack("<HHHH", rdata[2 + i * 8 : 10 + i * 8])
                        merges.append((r1, r2, c1, c2))

        if not rows:
            continue

        # Parse Escher SpContainers for (pib, ClientAnchor)
        sh = b"".join(sh_chunks)
        pos = 0
        anchors = []
        curr_pib = None
        while pos + 8 <= len(sh):
            vi, fbt, length = struct.unpack("<HHI", sh[pos : pos + 8])
            inst = vi >> 4
            if fbt in (0xF002, 0xF003):
                pos += 8
            elif fbt == 0xF004:
                curr_pib = None
                pos += 8
            else:
                if fbt == 0xF00B:  # OPT
                    for i in range(inst):
                        if pos + 8 + (i + 1) * 6 <= pos + 8 + length:
                            pid, pval = struct.unpack("<HI", sh[pos + 8 + i * 6 : pos + 8 + (i + 1) * 6])
                            if (pid & 0x3FFF) == 260:
                                curr_pib = pval
                elif fbt == 0xF010 and length >= 18 and curr_pib in bse_images:
                    flag, col1, dx1, row1, dy1, col2, dx2, row2, dy2 = struct.unpack(
                        "<HHHHHHHHH", sh[pos + 8 : pos + 26]
                    )
                    anchors.append((row1, row2, col1, col2, bse_images[curr_pib]))
                pos += 8 + length

        results.append((sname, rows, merges, anchors))

    return results


# =============================================================================
# 2. MULTILINGUAL TRANSLATION & CATEGORIZATION ENGINE (zh, en, si, es)
# =============================================================================

# Exact and compositional dictionaries for Chinese -> English, Sinhala, Spanish
EXACT_DESC_MAP = {
    "24格：32*32*10": (
        "24格内衣袜子收纳盒",
        "24-Grid Foldable Underwear & Sock Drawer Organizer Box",
        "කාමර 24 නැවිය හැකි යට ඇඳුම් සහ මේස් සංවිධායක පෙට්ටිය (24-Grid Underwear & Sock Organizer)",
        "Caja Organizadora Plegable de 24 Compartimentos para Ropa Interior y Calcetines"
    ),
    "7格：32*32*10": (
        "7格文胸内衣收纳盒",
        "7-Grid Foldable Bra & Underwear Drawer Organizer Box",
        "කාමර 7 නැවිය හැකි බ්‍රා සහ යට ඇඳුම් සංවිධායක පෙට්ටිය (7-Grid Bra & Underwear Organizer)",
        "Caja Organizadora Plegable de 7 Compartimentos para Sujetadores y Ropa Interior"
    ),
    "6格：32*16*10": (
        "6格内裤领带收纳盒",
        "6-Grid Foldable Briefs & Necktie Drawer Organizer Box",
        "කාමර 6 නැවිය හැකි යට ඇඳුම් සහ ටයි පටි සංවිධායක පෙට්ටිය (6-Grid Underwear & Tie Organizer)",
        "Caja Organizadora Plegable de 6 Compartimentos para Ropa Interior y Corbatas"
    ),
    "8格：32*16*10": (
        "8格内裤袜子收纳盒",
        "8-Grid Foldable Underwear & Sock Drawer Organizer Box",
        "කාමර 8 නැවිය හැකි යට ඇඳුම් සහ මේස් සංවිධායක පෙට්ටිය (8-Grid Underwear & Sock Organizer)",
        "Caja Organizadora Plegable de 8 Compartimentos para Ropa Interior y Calcetines"
    ),
    "单格黑铁管": (
        "单格黑铁管折叠脏衣篮",
        "Single-Section Black Iron-Tube Foldable Laundry Hamper",
        "තනි කොටසේ කළු යකඩ රාමු නැවිය හැකි කිලිටි රෙදි බඳුන (Single-Grid Iron Frame Laundry Hamper)",
        "Cesto de Ropa Sucia Plegable de 1 Sección con Tubo de Hierro Negro"
    ),
    "单格白铝管": (
        "单格白铝管折叠脏衣篮",
        "Single-Section White Aluminum-Tube Foldable Laundry Hamper",
        "තනි කොටසේ සුදු ඇලුමිනියම් රාමු කිලිටි රෙදි බඳුන (Single-Grid Aluminum Frame Laundry Hamper)",
        "Cesto de Ropa Sucia Plegable de 1 Sección con Tubo de Aluminio Blanco"
    ),
    "2格黑铁管": (
        "2格黑铁管分类折叠脏衣篮",
        "2-Section Black Iron-Tube Laundry Sorter Hamper",
        "කොටස් 2 කළු යකඩ රාමු රෙදි වර්ගීකරණ කූඩය (2-Section Iron Frame Laundry Sorter)",
        "Cesto Clasificador de Ropa Sucia de 2 Secciones con Tubo de Hierro Negro"
    ),
    "2格白铝管": (
        "2格白铝管分类折叠脏衣篮",
        "2-Section White Aluminum-Tube Laundry Sorter Hamper",
        "කොටස් 2 සුදු ඇලුමිනියම් රාමු රෙදි වර්ගීකරණ කූඩය (2-Section Aluminum Frame Laundry Sorter)",
        "Cesto Clasificador de Ropa Sucia de 2 Secciones con Tubo de Aluminio Blanco"
    ),
    "3格黑铁管": (
        "3格黑铁管分类折叠脏衣篮",
        "3-Section Black Iron-Tube Laundry Sorter Hamper",
        "කොටස් 3 කළු යකඩ රාමු රෙදි වර්ගීකරණ කූඩය (3-Section Iron Frame Laundry Sorter)",
        "Cesto Clasificador de Ropa Sucia de 3 Secciones con Tubo de Hierro Negro"
    ),
    "3格白铝管": (
        "3格白铝管分类折叠脏衣篮",
        "3-Section White Aluminum-Tube Laundry Sorter Hamper",
        "කොටස් 3 සුදු ඇලුමිනියම් රාමු රෙදි වර්ගීකරණ කූඩය (3-Section Aluminum Frame Laundry Sorter)",
        "Cesto Clasificador de Ropa Sucia de 3 Secciones con Tubo de Aluminio Blanco"
    ),
    "三明治刺绣款\n糖果拉链头\n涤纶三明治面料130G": (
        "三明治刺绣加厚护洗袋/洗衣袋",
        "Embroidered Sandwich-Mesh Delicates Laundry Wash Bag",
        "එම්බ්‍රොයිඩර් කළ ඝනකම් දැල් රෙදි සෝදන බෑගය (Embroidered Mesh Laundry Wash Bag)",
        "Bolsa de Malla Sándwich Bordada para Lavado de Ropa Delicada"
    ),
    "涤纶160克三明治面料": (
        "160G加厚三明治网文胸护洗袋",
        "160g Heavy-Duty Sandwich-Mesh Bra Laundry Wash Bag",
        "160g ඝනකම් දැල් බ්‍රා සෝදන ආරක්ෂිත බෑගය (Heavy-Duty Mesh Bra Wash Bag)",
        "Bolsa Protectora de Malla Sándwich 160g para Lavado de Sujetadores"
    ),
    "棉麻棋盘格【正面5+反面10】": (
        "棉麻棋盘格双面收纳挂袋 (正面5格+反面10格)",
        "Cotton-Linen Checkerboard 15-Pocket Double-Sided Hanging Organizer (5+10 Pockets)",
        "කපු-ලිනන් කොටු මෝස්තර දෙපසම සාක්කු 15 එල්ලන සංවිධායක බෑගය (15-Pocket Hanging Closet Organizer)",
        "Organizador Colgante de Doble Cara a Cuadros de Lino y Algodón (5+10 Bolsillos)"
    ),
    "棉麻棋盘格【正面5+反面15】": (
        "棉麻棋盘格双面收纳挂袋 (正面5格+反面15格)",
        "Cotton-Linen Checkerboard 20-Pocket Double-Sided Hanging Organizer (5+15 Pockets)",
        "කපු-ලිනන් කොටු මෝස්තර දෙපසම සාක්කු 20 එල්ලන සංවිධායක බෑගය (20-Pocket Hanging Closet Organizer)",
        "Organizador Colgante de Doble Cara a Cuadros de Lino y Algodón (5+15 Bolsillos)"
    ),
    "棉麻棋盘格【正面10+反面15】": (
        "棉麻棋盘格双面收纳挂袋 (正面10格+反面15格)",
        "Cotton-Linen Checkerboard 25-Pocket Double-Sided Hanging Organizer (10+15 Pockets)",
        "කපු-ලිනන් කොටු මෝස්තර දෙපසම සාක්කු 25 එල්ලන සංවිධායක බෑගය (25-Pocket Hanging Closet Organizer)",
        "Organizador Colgante de Doble Cara a Cuadros de Lino y Algodón (10+15 Bolsillos)"
    ),
    "单层阳离子": (
        "单层阳离子钢架收纳框",
        "Single-Layer Cationic Fabric Steel-Frame Storage Basket",
        "තනි ස්ථරයේ කැටායනික රෙදි වානේ රාමු ගබඩා කූඩය (Single-Layer Cationic Steel-Frame Basket)",
        "Cesta Organizadora de Una Capa en Tela Catiónica con Marco de Acero"
    ),
    "阳离子牛津布+pp板": (
        "阳离子牛津布PP板可水洗收纳框",
        "Cationic Oxford Cloth & PP-Board Washable Storage Basket",
        "සේදිය හැකි කැටායනික ඔක්ස්ෆර්ඩ් සහ PP පුවරු ගබඩා කූඩය (Washable Cationic Oxford PP-Board Basket)",
        "Cesta Organizadora Lavable de Tela Oxford Catiónica con Placa PP"
    ),
    "涤棉网纱透视窗": (
        "涤棉网纱透视窗钢架收纳盒",
        "Poly-Cotton Mesh Clear-Window Steel-Frame Storage Box",
        "පොලි-කපු දැල් විනිවිද පෙනෙන කවුළු සහිත වානේ රාමු ගබඩා පෙට්ටිය (Mesh Window Steel-Frame Storage Box)",
        "Caja Organizadora de Algodón-Poliéster con Ventana de Malla y Marco de Acero"
    ),
}


def translate_product_title(raw_zh, raw_size=""):
    """
    Deterministic, rich 4-language translator for all 254 product descriptions.
    Returns (clean_zh, en, si, es, category_id).
    """
    s = raw_zh.strip()
    if s in EXACT_DESC_MAP:
        zh_clean, en, si, es = EXACT_DESC_MAP[s]
        cat = classify_category(zh_clean, en)
        return zh_clean, en, si, es, cat

    # Take primary line of Chinese title (strip multi-line material specs after first line unless needed)
    lines = [ln.strip() for ln in s.splitlines() if ln.strip()]
    main_line = lines[0] if lines else s
    # Strip trailing material slashes if main_line has '/' like '内衣盒/棉麻/210D牛津布/PP板/700g'
    parts_slash = [p.strip() for p in main_line.split("/") if p.strip()]
    core_zh = parts_slash[0] if parts_slash else main_line
    # Fix unclosed paren like '国风手提棉麻天地盖（密度板'
    if "（" in core_zh and "）" not in core_zh:
        core_zh += "）"

    full_text = s + " " + str(raw_size)

    # Extract model prefix if present (e.g. MY200, GC188)
    model_prefix = ""
    mp = re.match(r"^([A-Z]{2}\d+)\s*(.*)$", core_zh)
    if mp:
        model_prefix = mp.group(1) + " "
        core_zh_body = mp.group(2)
    else:
        core_zh_body = core_zh

    # Feature / Material modifiers
    mods_en, mods_si, mods_es = [], [], []

    if "可水洗" in full_text:
        mods_en.append("Washable")
        mods_si.append("සේදිය හැකි")
        mods_es.append("Lavable")
    if "加厚" in full_text:
        mods_en.append("Heavy-Duty Thickened")
        mods_si.append("ඝනකම්")
        mods_es.append("Reforzado Grueso")
    elif "高质量" in full_text or "优质" in full_text:
        mods_en.append("Premium")
        mods_si.append("උසස් තත්ත්වයේ")
        mods_es.append("Premium")
    elif "新款" in full_text or "时尚" in full_text:
        mods_en.append("Modern")
        mods_si.append("නවීන")
        mods_es.append("Moderno")
    elif "简约" in full_text:
        mods_en.append("Minimalist")
        mods_si.append("සරල")
        mods_es.append("Minimalista")

    # Style / Pattern modifiers
    if "鱼骨纹" in full_text:
        mods_en.append("Herringbone")
        mods_si.append("හෙරින්බෝන් මෝස්තරයේ")
        mods_es.append("Diseño Espiga")
    elif "小香风" in full_text:
        mods_en.append("Chanel-Style Tweed")
        mods_si.append("ෂැනෙල් විලාසිතාවේ")
        mods_es.append("Estilo Tweed Elegante")
    elif "棋盘" in full_text:
        mods_en.append("Checkerboard")
        mods_si.append("කොටු මෝස්තර (Checkerboard)")
        mods_es.append("Diseño a Cuadros")
    elif "千鸟格" in full_text:
        mods_en.append("Houndstooth")
        mods_si.append("හවුන්ඩ්ස්ටූත් මෝස්තරයේ")
        mods_es.append("Pata de Gallo")
    elif "国风" in full_text:
        mods_en.append("Oriental-Style")
        mods_si.append("සම්ප්‍රදායික විලාසිතාවේ")
        mods_es.append("Estilo Oriental")
    elif "楼兰" in full_text:
        mods_en.append("Loulan Jacquard")
        mods_si.append("ලෝලාන් ජැකාඩ්")
        mods_es.append("Jacquard Loulan")
    elif "少女系列" in full_text:
        mods_en.append("Pastel Series")
        mods_si.append("පැස්ටල් වර්ණ")
        mods_es.append("Serie Pastel")
    elif "条纹" in full_text:
        mods_en.append("Striped")
        mods_si.append("ඉරි මෝස්තර")
        mods_es.append("de Rayas")
    elif "格子布" in full_text:
        mods_en.append("Plaid Fabric")
        mods_si.append("කොටු රෙදි")
        mods_es.append("de Tela Escocesa")
    elif "卡通" in full_text:
        mods_en.append("Cartoon Print")
        mods_si.append("කාටූන් මෝස්තර")
        mods_es.append("con Diseño Animado")
    elif "印花" in full_text:
        mods_en.append("Printed")
        mods_si.append("මුද්‍රිත")
        mods_es.append("Estampado")
    elif "复古" in full_text:
        mods_en.append("Vintage")
        mods_si.append("වින්ටේජ් (Vintage)")
        mods_es.append("Vintage")
    elif "校车" in full_text:
        mods_en.append("School Bus")
        mods_si.append("පාසල් බස් රථ හැඩැති")
        mods_es.append("Diseño Autobús Escolar")
    elif "车头" in full_text:
        mods_en.append("Bus-Front")
        mods_si.append("බස් රථ හැඩැති")
        mods_es.append("con Cabina de Autobús")
    elif "拼色" in full_text or "撞色" in full_text or "双拼" in full_text or "拼接" in full_text:
        mods_en.append("Two-Tone Colorblock")
        mods_si.append("ද්විත්ව වර්ණ")
        mods_es.append("Bicolor")

    # Material
    if "天山棉麻" in full_text or "天山棉" in full_text:
        mods_en.append("Tianshan Cotton-Linen")
        mods_si.append("ටියැන්ෂාන් කපු-ලිනන්")
        mods_es.append("de Lino y Algodón Tianshan")
    elif "棉麻" in full_text or "亚麻" in full_text or "仿麻" in full_text or "精麻" in full_text or "麻块" in full_text:
        mods_en.append("Cotton-Linen")
        mods_si.append("කපු-ලිනන්")
        mods_es.append("de Lino y Algodón")
    elif "阳离子" in full_text or "阳子" in full_text:
        mods_en.append("Cationic Fabric")
        mods_si.append("කැටායනික රෙදි (Cationic)")
        mods_es.append("de Tela Catiónica")
    elif "牛津布" in full_text:
        mods_en.append("Oxford Cloth")
        mods_si.append("ඔක්ස්ෆර්ඩ් රෙදි (Oxford)")
        mods_es.append("de Tela Oxford")
    elif "无纺布" in full_text:
        mods_en.append("Non-Woven")
        mods_si.append("නොන්-වෝවන් (Non-Woven)")
        mods_es.append("de Tela No Tejida")
    elif "毛毡布" in full_text:
        mods_en.append("Thick Felt")
        mods_si.append("ෆෙල්ට් රෙදි (Felt)")
        mods_es.append("de Fieltro Grueso")
    elif "帆布" in full_text:
        mods_en.append("Canvas")
        mods_si.append("කැන්වස් (Canvas)")
        mods_es.append("de Lona")
    elif "绒布" in full_text:
        mods_en.append("Velvet")
        mods_si.append("වෙල්වට් (Velvet)")
        mods_es.append("de Terciopelo")
    elif "云柔棉" in full_text or "天竺棉" in full_text or "涤棉" in full_text or "棉布" in full_text:
        mods_en.append("Soft Cotton")
        mods_si.append("මෘදු කපු")
        mods_es.append("de Algodón Suave")
    elif "斜纹" in full_text:
        mods_en.append("Twill Fabric")
        mods_si.append("ට්විල් රෙදි (Twill)")
        mods_es.append("de Sarga")
    elif "PU" in full_text or "皮革" in full_text:
        mods_en.append("PU Leather")
        mods_si.append("PU ලෙදර්")
        mods_es.append("de Cuero PU")
    elif "PVC" in full_text:
        mods_en.append("Clear PVC")
        mods_si.append("විනිවිද පෙනෙන PVC")
        mods_es.append("de PVC Transparente")

    if "EVA" in full_text and "EVA" not in " ".join(mods_en):
        mods_en.append("EVA")
        mods_si.append("EVA")
        mods_es.append("EVA")

    # Structural details
    if "双窗" in full_text:
        mods_en.append("Dual-Window")
        mods_si.append("ද්විත්ව කවුළු")
        mods_es.append("con Doble Ventana")
    elif any(k in full_text for k in ("单窗", "大视窗", "前透窗", "透视", "可视", "开窗")):
        mods_en.append("Clear-Window")
        mods_si.append("විනිවිද පෙනෙන කවුළුවක් සහිත")
        mods_es.append("con Ventana Transparente")

    if "双层" in full_text:
        mods_en.append("Double-Layer")
        mods_si.append("ද්විත්ව ස්ථර")
        mods_es.append("de Doble Capa")
    elif "单层" in full_text:
        mods_en.append("Single-Layer")
        mods_si.append("තනි ස්ථර")
        mods_es.append("de Una Capa")

    if "铝柄" in full_text:
        mods_en.append("Aluminum-Handle")
        mods_si.append("ඇලුමිනියම් හැඬලය සහිත")
        mods_es.append("con Asa de Aluminio")
    elif "木柄" in full_text:
        mods_en.append("Wooden-Handle")
        mods_si.append("ලී හැඬලය සහිත")
        mods_es.append("con Asa de Madera")
    elif "皮手提" in full_text or "皮革手提" in full_text:
        mods_en.append("Leather-Handle")
        mods_si.append("ලෙදර් හැඬලය සහිත")
        mods_es.append("con Asa de Cuero")
    elif "金属手提" in full_text or "金属扣" in full_text:
        mods_en.append("Metal-Buckle")
        mods_si.append("ලෝහ බකල් සහිත")
        mods_es.append("con Hebilla Metálica")
    elif "带轮" in full_text or "万向轮" in full_text:
        mods_en.append("Rolling Wheeled")
        mods_si.append("රෝද සහිත")
        mods_es.append("con Ruedas Giratorias")
    elif "X架" in full_text:
        mods_en.append("X-Frame")
        mods_si.append("X-රාමු")
        mods_es.append("con Marco en X")

    # Grid / Set / Layer count extraction
    grid_m = re.search(r"(\d+)\s*格", core_zh_body) or re.search(r"(\d+)\s*格", str(raw_size))
    set_m = re.search(r"([三四六七八九346789])\s*(?:件套|合一)", core_zh_body)
    layer_m = re.search(r"([三四五六八34568])\s*层", core_zh_body + " " + str(raw_size))
    cn_num_map = {"三": "3", "四": "4", "五": "5", "六": "6", "七": "7", "八": "8", "九": "9"}

    # Base noun determination
    if "收纳凳" in core_zh_body:
        base_en = "Foldable Storage Ottoman Stool"
        base_si = "නැවිය හැකි ගබඩා බංකුව (Foldable Storage Ottoman Stool)"
        base_es = "Puf Baúl Otomano Plegable de Almacenamiento"
    elif "脏衣" in core_zh_body or "洗衣篮" in core_zh_body:
        base_en = "Foldable Laundry Hamper Basket"
        base_si = "නැවිය හැකි කිලිටි රෙදි බඳුන (Foldable Laundry Hamper)"
        base_es = "Cesto Plegable para Ropa Sucia"
    elif "洗衣袋" in core_zh_body:
        base_en = "Mesh Delicates Laundry Wash Bag"
        base_si = "දැල් රෙදි සෝදන ආරක්ෂිත බෑගය (Mesh Laundry Wash Bag)"
        base_es = "Bolsa de Malla para Lavado de Ropa"
    elif "防尘罩" in core_zh_body:
        base_en = "Hanging Garment Dust Cover Bag"
        base_si = "ඇඳුම් දූවිලි ආරක්ෂණ කවරය (Garment Dust Cover)"
        base_es = "Funda Guardapolvo Colgante para Ropa"
    elif "鞋盒" in core_zh_body:
        base_en = "Foldable Dustproof Shoe Organizer Box"
        base_si = "නැවිය හැකි සපත්තු ගබඩා පෙට්ටිය (Foldable Shoe Storage Box)"
        base_es = "Caja Organizadora Plegable para Zapatos"
    elif "鞋子" in core_zh_body and "袋" in core_zh_body:
        base_en = "Travel Shoe Storage Pouch"
        base_si = "සංචාරක සපත්තු බෑගය (Travel Shoe Organizer Bag)"
        base_es = "Bolsa Organizadora de Viaje para Zapatos"
    elif "文胸包" in core_zh_body:
        base_en = "Travel Bra & Lingerie Organizer Bag"
        base_si = "සංචාරක යට ඇඳුම් සහ බ්‍රා බෑගය (Travel Bra Organizer Bag)"
        base_es = "Bolsa Organizadora de Viaje para Lencería"
    elif "洗漱包" in core_zh_body or "化妆包" in core_zh_body or "化装袋" in core_zh_body:
        if set_m:
            n = cn_num_map.get(set_m.group(1), set_m.group(1))
            base_en = f"{n}-Piece Waterproof Toiletry & Cosmetic Bag Set"
            base_si = f"ජලයට ඔරොත්තු දෙන රූපලාවන්‍ය සහ මේකප් බෑග් කට්ටලය - කොටස් {n} ({n}-Piece Cosmetic & Toiletry Bag Set)"
            base_es = f"Set de {n} Bolsas Impermeables de Aseo y Cosméticos"
        elif "弯钩" in core_zh_body:
            base_en = "Hanging Hook Waterproof Toiletry & Cosmetic Bag"
            base_si = "එල්ලිය හැකි ජලයට ඔරොත්තු දෙන රූපලාවන්‍ය සහ ටොයිලට්‍රි බෑගය (Hanging Waterproof Toiletry Bag)"
            base_es = "Neceser Impermeable con Gancho para Colgar"
        elif "枕头" in core_zh_body:
            base_en = "Pillow-Shape Large Capacity Makeup & Cosmetic Bag"
            base_si = "කොට්ට හැඩැති විශාල රූපලාවන්‍ය සහ මේකප් බෑගය (Pillow Cosmetic Makeup Bag)"
            base_es = "Bolso de Maquillaje y Cosméticos Tipo Almohada"
        else:
            base_en = "Waterproof Travel Toiletry & Cosmetic Makeup Bag"
            base_si = "ජලයට ඔරොත්තු දෙන රූපලාවන්‍ය සහ මේකප් බෑගය (Waterproof Toiletry & Makeup Bag)"
            base_es = "Neceser y Bolso de Maquillaje Impermeable de Viaje"
    elif "旅行" in core_zh_body:
        if set_m:
            n = cn_num_map.get(set_m.group(1), set_m.group(1))
            base_en = f"{n}-Piece Travel Luggage Packing Cube Organizer Set"
            base_si = f"සංචාරක ගමන් මලු සංවිධායක බෑග් කට්ටලය - කොටස් {n} ({n}-Piece Travel Packing Cube Set)"
            base_es = f"Set de {n} Organizadores de Equipaje para Viaje"
        else:
            base_en = "Expandable Travel Duffel & Luggage Bag"
            base_si = "සංචාරක ගමන් මල්ල (Expandable Travel Luggage Bag)"
            base_es = "Bolso de Viaje y Equipaje Expandible"
    elif "棉被袋" in core_zh_body or "搬家袋" in core_zh_body:
        base_en = "Large-Capacity Quilt, Blanket & Moving Storage Bag"
        base_si = "විශාල ඇඳ ඇතිරිලි සහ බ්ලැන්කට් ගබඩා බෑගය (Quilt & Blanket Storage Bag)"
        base_es = "Bolsa de Gran Capacidad para Edredones y Mudanza"
    elif "帽子挂袋" in core_zh_body:
        n = cn_num_map.get(layer_m.group(1), layer_m.group(1)) if layer_m else "Multi"
        base_en = f"{n}-Tier Hanging Cap & Hat Closet Organizer"
        base_si = f"තට්ටු {n} එල්ලන තොප්පි සංවිධායක රාක්කය ({n}-Tier Hanging Cap Organizer)"
        base_es = f"Organizador Colgante de {n} Niveles para Gorras"
    elif "门后" in core_zh_body and "挂袋" in core_zh_body:
        if grid_m:
            g = grid_m.group(1)
            base_en = f"{g}-Pocket Over-the-Door Hanging Mesh Organizer"
            base_si = f"දොර පිටුපස එල්ලන සාක්කු {g} සංවිධායක බෑගය ({g}-Pocket Over-the-Door Organizer)"
            base_es = f"Organizador Colgante Trasero de Puerta con {g} Bolsillos"
        else:
            n = cn_num_map.get(layer_m.group(1), layer_m.group(1)) if layer_m else "5"
            base_en = f"{n}-Tier Over-the-Door Hanging Storage Pocket Organizer"
            base_si = f"දොර පිටුපස එල්ලන තට්ටු {n} ගබඩා සංවිධායකය ({n}-Tier Over-the-Door Organizer)"
            base_es = f"Organizador Colgante de {n} Niveles para Puerta"
    elif "挂袋" in core_zh_body:
        n = cn_num_map.get(layer_m.group(1), layer_m.group(1)) if layer_m else "Multi"
        base_en = f"{n}-Shelf Hanging Wardrobe Closet Organizer"
        base_si = f"අල්මාරියේ එල්ලන තට්ටු {n} ඇඳුම් සංවිධායක රාක්කය ({n}-Tier Wardrobe Hanging Shelf)"
        base_es = f"Estante Organizador Colgante de {n} Niveles para Armario"
    elif "内衣" in core_zh_body or "附标阳" in core_zh_body or ("装文胸" in full_text) or ("装内裤" in full_text) or ("装袜子" in full_text):
        if set_m:
            n = cn_num_map.get(set_m.group(1), set_m.group(1))
            base_en = f"{n}-Piece Underwear, Bra & Sock Drawer Organizer Set"
            base_si = f"යට ඇඳුම් සහ මේස් සංවිධායක පෙට්ටි කට්ටලය - කොටස් {n} ({n}-Piece Underwear Organizer Set)"
            base_es = f"Set de {n} Cajas Organizadoras para Ropa Interior"
        elif grid_m:
            g = grid_m.group(1)
            base_en = f"{g}-Grid Underwear, Bra & Sock Organizer Box"
            base_si = f"කාමර {g} යට ඇඳුම් සහ මේස් සංවිධායක පෙට්ටිය ({g}-Grid Underwear Organizer Box)"
            base_es = f"Caja Organizadora de {g} Compartimentos para Ropa Interior"
        else:
            base_en = "Compartment Underwear & Bra Drawer Organizer Box"
            base_si = "යට ඇඳුම් සහ බ්‍රා ලාච්චු සංවිධායක පෙට්ටිය (Underwear & Bra Drawer Organizer)"
            base_es = "Caja Organizadora de Cajón para Ropa Interior"
    elif "裤盒" in core_zh_body or "分格" in core_zh_body or "收纳格" in core_zh_body or "7格衣物" in core_zh_body or "网格收纳盒" in core_zh_body:
        if grid_m:
            g = grid_m.group(1)
            base_en = f"{g}-Compartment Wardrobe Jeans & Clothes Organizer"
            base_si = f"කාමර {g} ඇඳුම් සහ ඩෙනිම් කලිසම් සංවිධායක පෙට්ටිය ({g}-Grid Wardrobe Jeans Organizer)"
            base_es = f"Organizador de Armario de {g} Compartimentos para Pantalones y Ropa"
        else:
            base_en = "Compartment Wardrobe Jeans & Clothes Organizer Box"
            base_si = "ඇඳුම් අල්මාරි කාමර බෙදූ සංවිධායක පෙට්ටිය (Wardrobe Compartment Organizer)"
            base_es = "Caja Organizadora con Compartimentos para Ropa y Pantalones"
    elif "床底" in core_zh_body or "床单" in core_zh_body:
        base_en = "Under-Bed Clothes & Bedding Storage Box"
        base_si = "ඇඳ යට තැබිය හැකි ඇඳුම් සහ ඇතිරිලි ගබඩා පෙට්ටිය (Under-Bed Storage Box)"
        base_es = "Caja Organizadora Bajo Cama para Ropa y Sábanas"
    elif "夹缝" in core_zh_body:
        base_en = "Narrow-Gap Steel-Frame Foldable Storage Box"
        base_si = "පටු ඉඩක තැබිය හැකි වානේ රාමු ගබඩා පෙට්ටිය (Slim Gap Steel-Frame Storage Box)"
        base_es = "Caja Organizadora Estrecha con Marco de Acero"
    elif "百纳箱" in core_zh_body or "钢架箱" in core_zh_body:
        base_en = "Steel-Frame Foldable Wardrobe Storage Box"
        base_si = "වානේ රාමු සහිත නැවිය හැකි ඇඳුම් ගබඩා පෙට්ටිය (Steel-Frame Wardrobe Storage Box)"
        base_es = "Caja Organizadora Plegable de Armario con Marco de Acero"
    elif "钢架框" in core_zh_body or "钢架盒" in core_zh_body or "钢架收纳" in core_zh_body or "收纲框" in core_zh_body:
        base_en = "Steel-Frame Foldable Storage Basket Bin"
        base_si = "වානේ රාමු සහිත නැවිය හැකි ගබඩා කූඩය (Steel-Frame Storage Basket)"
        base_es = "Cesta Organizadora Plegable con Marco de Acero"
    elif "天地盖" in core_zh_body:
        base_en = "Detachable-Lid Foldable Storage Box"
        base_si = "වෙන් කළ හැකි පියන සහිත නැවිය හැකි ගබඩා පෙට්ටිය (Detachable-Lid Storage Box)"
        base_es = "Caja de Almacenamiento Plegable con Tapa Superior"
    elif "连盖" in core_zh_body or "连体盖" in core_zh_body or "连体带盖" in core_zh_body or "翻盖" in core_zh_body:
        base_en = "Flip-Top Attached-Lid Foldable Storage Box"
        base_si = "සවි කළ පියන සහිත නැවිය හැකි ගබඩා පෙට්ටිය (Flip-Lid Foldable Storage Box)"
        base_es = "Caja Organizadora Plegable con Tapa Abatible"
    elif "双盖" in core_zh_body:
        base_en = "Dual-Lid Foldable Wardrobe Storage Box"
        base_si = "ද්විත්ව පියන් සහිත නැවිය හැකි ඇඳුම් ගබඩා පෙට්ටිය (Dual-Lid Storage Box)"
        base_es = "Caja Organizadora Plegable de Doble Tapa"
    elif "软盖" in core_zh_body or "曲线盖" in core_zh_body or "带盖" in core_zh_body:
        base_en = "Lidded Foldable Clothes Storage Box"
        base_si = "පියන සහිත නැවිය හැකි ඇඳුම් ගබඩා පෙට්ටිය (Lidded Clothes Storage Box)"
        base_es = "Caja Organizadora Plegable para Ropa con Tapa"
    elif "拉链" in core_zh_body:
        base_en = "Zipper-Top Foldable Clothes Storage Box"
        base_si = "සිප් පියන සහිත නැවිය හැකි ඇඳුම් ගබඩා පෙට්ටිය (Zipper Foldable Storage Box)"
        base_es = "Caja Organizadora Plegable con Cierre de Cremallera"
    elif "扣扣箱" in core_zh_body or "QQ箱" in core_zh_body:
        base_en = "Buckle-Clasp Foldable Storage Box"
        base_si = "බකල් අගුලු සහිත නැවිය හැකි ගබඩා පෙට්ටිය (Buckle Foldable Storage Box)"
        base_es = "Caja Organizadora Plegable con Cierre de Hebilla"
    elif "方盒" in core_zh_body:
        base_en = "Foldable Cube Shelf Storage Bin"
        base_si = "නැවිය හැකි හතරැස් රාක්ක ගබඩා පෙට්ටිය (Foldable Cube Storage Bin)"
        base_es = "Cubo Organizador Plegable para Estantería"
    elif "抽屉" in core_zh_body:
        base_en = "Pull-Out Drawer Wardrobe Storage Box"
        base_si = "ලාච්චු මාදිලියේ ඇඳුම් සංවිධායක පෙට්ටිය (Drawer-Style Storage Box)"
        base_es = "Caja Organizadora Tipo Cajón Extraíble"
    elif "收纳桶" in core_zh_body:
        base_en = "Round Foldable Storage Barrel Basket"
        base_si = "රවුම් හැඩැති නැවිය හැකි ගබඩා බඳුන (Foldable Storage Barrel Basket)"
        base_es = "Cesto Cilíndrico Plegable de Almacenamiento"
    elif "收纳筐" in core_zh_body or "收纳框" in core_zh_body or "整理框" in core_zh_body:
        if set_m:
            n = cn_num_map.get(set_m.group(1), set_m.group(1))
            base_en = f"{n}-Piece Foldable Desktop & Shelf Storage Basket Set"
            base_si = f"නැවිය හැකි ගබඩා කූඩ කට්ටලය - කොටස් {n} ({n}-Piece Storage Basket Set)"
            base_es = f"Set de {n} Cestas Organizadoras Plegables"
        else:
            base_en = "Foldable Open-Top Shelf Storage Basket"
            base_si = "නැවිය හැකි විවෘත ගබඩා කූඩය (Foldable Shelf Storage Basket)"
            base_es = "Cesta Organizadora Plegable para Estantería"
    else:
        if set_m:
            n = cn_num_map.get(set_m.group(1), set_m.group(1))
            base_en = f"{n}-Piece Foldable Wardrobe Storage Box Set"
            base_si = f"නැවිය හැකි ඇඳුම් ගබඩා පෙට්ටි කට්ටලය - කොටස් {n} ({n}-Piece Storage Box Set)"
            base_es = f"Set de {n} Cajas Organizadoras Plegables"
        else:
            base_en = "Foldable Home & Wardrobe Organizer Box"
            base_si = "නැවිය හැකි ගෘහස්ථ සහ ඇඳුම් සංවිධායක පෙට්ටිය (Foldable Wardrobe Organizer Box)"
            base_es = "Caja Organizadora Plegable para Hogar y Armario"

    # Combine prefix + modifiers + base noun cleanly without duplicate words
    en_prefix = " ".join(mods_en)
    for w in mods_en:
        if w.lower() in base_en.lower():
            en_prefix = en_prefix.replace(w, "").strip()
    en_title = f"{model_prefix}{en_prefix} {base_en}".strip()
    en_title = re.sub(r"\s+", " ", en_title)

    si_prefix = " ".join(mods_si)
    si_title = f"{model_prefix}{si_prefix} {base_si}".strip()
    si_title = re.sub(r"\s+", " ", si_title)

    es_suffix = " ".join(mods_es)
    es_title = f"{model_prefix}{base_es} {es_suffix}".strip()
    es_title = re.sub(r"\s+", " ", es_title)

    cat = classify_category(core_zh, en_title)
    return core_zh, en_title, si_title, es_title, cat


def classify_category(zh_title, en_title):
    """Classify product into one of 6 logical categories."""
    t = zh_title + " " + en_title.lower()
    if any(k in t for k in ("收纳凳", "ottoman", "stool")):
        return "ottomans-stools"
    if any(k in t for k in ("洗漱", "化妆", "化装", "旅行", "文胸包", "鞋子", "toiletry", "cosmetic", "makeup", "travel")):
        return "travel-cosmetic"
    if any(k in t for k in ("内衣", "文胸", "袜", "裤盒", "分格", "收纳格", "7格", "6格", "8格", "16格", "20格", "24格", "网格收纳盒", "抽屉", "underwear", "compartment", "drawer")):
        return "wardrobe-organizers"
    if any(k in t for k in ("棉被袋", "搬家袋", "挂袋", "防尘罩", "洗衣袋", "护洗袋", "鞋盒", "quilt", "hanging", "dust cover", "wash bag", "shoe organizer")):
        return "hanging-dust-laundry"
    if any(k in t for k in ("脏衣", "洗衣篮", "收纳筐", "收纳框", "收纳桶", "整理框", "钢架框", "收纲框", "hamper", "basket", "barrel")):
        return "baskets-hampers"
    return "storage-boxes"


COLOR_TOKEN_MAP = [
    ("图片2个图案", "2 Assorted Patterns", "රටා 2 ක්", "2 Patrones Surtidos"),
    ("图片2色", "2 Assorted Colors", "වර්ණ 2 ක්", "2 Colores Surtidos"),
    ("图片色", "As Shown", "පින්තූරයේ වර්ණය", "Color de Imagen"),
    ("图和色", "As Shown", "පින්තූරයේ වර්ණය", "Color de Imagen"),
    ("多款动物颜色混", "Assorted Animal Prints", "විවිධ සත්ව රටා", "Diseños de Animales Surtidos"),
    ("多款动物混", "Assorted Animal Prints", "විවිධ සත්ව රටා", "Diseños de Animales Surtidos"),
    ("多款多色混", "Assorted Colors & Styles", "විවිධ වර්ණ සහ මෝස්තර", "Colores y Estilos Surtidos"),
    ("多色多款混", "Assorted Colors & Styles", "විවිධ වර්ණ සහ මෝස්තර", "Colores y Estilos Surtidos"),
    ("混色，不挑色", "Random Assorted Colors", "මිශ්‍ර වර්ණ", "Colores Surtidos"),
    ("灰咖均混", "Grey & Coffee Mix", "අළු සහ කෝපි මිශ්‍ර", "Mezcla Gris y Café"),
    ("3色混", "3-Color Assorted", "වර්ණ 3 මිශ්‍ර", "3 Colores Surtidos"),
    ("多色混", "Multi-Color Assorted", "බහු වර්ණ මිශ්‍ර", "Multicolor Surtido"),
    ("混色", "Assorted Colors", "මිශ්‍ර වර්ණ", "Colores Surtidos"),
    ("13个颜色", "13 Colors Available", "වර්ණ 13 ක්", "13 Colores Disponibles"),
    ("拉链颜色多色", "Multi-Color Zipper", "බහු වර්ණ සිප්", "Cremallera Multicolor"),
    ("蓝配米色/米配蓝色", "Blue-Beige / Beige-Blue", "නිල්-බේජ් / බේජ්-නිල්", "Azul-Beige / Beige-Azul"),
    ("绿底熊/白底熊", "Green Bear / White Bear", "කොළ වලසා / සුදු වලසා", "Oso Verde / Oso Blanco"),
    ("粉色拉链", "Pink Zipper", "රෝස සිප්", "Cremallera Rosa"),
    ("灰色拉链", "Grey Zipper", "අළු සිප්", "Cremallera Gris"),
    ("透明白", "Clear White", "විනිවිද පෙනෙන සුදු", "Blanco Transparente"),
    ("透明灰", "Clear Grey", "විනිවිද පෙනෙන අළු", "Gris Transparente"),
    ("透明咖", "Clear Coffee", "විනිවිද පෙනෙන දුඹුරු", "Café Transparente"),
    ("香芋紫色", "Taro Purple", "ලා දම්", "Púrpura Taro"),
    ("深灰色", "Dark Grey", "තද අළු", "Gris Oscuro"),
    ("深灰", "Dark Grey", "තද අළු", "Gris Oscuro"),
    ("浅灰", "Light Grey", "ලා අළු", "Gris Claro"),
    ("米灰色", "Beige-Grey", "බේජ්-අළු", "Gris Beige"),
    ("米灰", "Beige-Grey", "බේජ්-අළු", "Gris Beige"),
    ("灰白", "Grey-White", "අළු-සුදු", "Gris-Blanco"),
    ("黑白", "Black-White", "කළු-සුදු", "Negro-Blanco"),
    ("雾蓝色", "Misty Blue", "මීදුම් නිල්", "Azul Niebla"),
    ("深蓝色", "Dark Blue", "තද නිල්", "Azul Oscuro"),
    ("藏青色", "Navy Blue", "නේවි නිල්", "Azul Marino"),
    ("藏青", "Navy Blue", "නේවි නිල්", "Azul Marino"),
    ("天蓝", "Sky Blue", "අහස් නිල්", "Azul Cielo"),
    ("湖蓝", "Lake Blue", "විල් නිල්", "Azul Lago"),
    ("亮蓝", "Bright Blue", "දීප්තිමත් නිල්", "Azul Brillante"),
    ("淡蓝条纹", "Light Blue Stripe", "ලා නිල් ඉරි", "Rayas Azul Claro"),
    ("米白色", "Off-White", "ඕෆ්-වයිට් (Off-White)", "Blanco Hueso"),
    ("米白", "Off-White", "ඕෆ්-වයිට් (Off-White)", "Blanco Hueso"),
    ("本白", "Natural White", "ස්වභාවික සුදු", "Blanco Natural"),
    ("奶白", "Cream White", "ක්‍රීම් සුදු", "Blanco Crema"),
    ("奶咖色", "Cream Coffee", "ක්‍රීම් දුඹුරු", "Café Crema"),
    ("奶杏色", "Cream Apricot", "ක්‍රීම් ඇප්‍රිකොට්", "Albaricoque Crema"),
    ("深紫色", "Dark Purple", "තද දම්", "Púrpura Oscuro"),
    ("千鸟格", "Houndstooth", "හවුන්ඩ්ස්ටූත්", "Pata de Gallo"),
    ("四叶草", "Four-Leaf Clover", "සිව්පත් මෝස්තරය", "Trébol"),
    ("仙人掌", "Cactus Print", "පතොක් මෝස්තරය", "Cactus"),
    ("细灰条", "Fine Grey Stripe", "සිහින් අළු ඉරි", "Raya Gris Fina"),
    ("细米条", "Fine Beige Stripe", "සිහින් බේජ් ඉරි", "Raya Beige Fina"),
    ("粗灰条", "Wide Grey Stripe", "පළල් අළු ඉරි", "Raya Gris Ancha"),
    ("粗米条", "Wide Beige Stripe", "පළල් බේජ් ඉරි", "Raya Beige Ancha"),
    ("卡其色", "Khaki", "කාකි", "Caqui"),
    ("咖啡", "Coffee Brown", "කෝපි දුඹුරු", "Café"),
    ("咖熊", "Brown Bear", "දුඹුරු වලසා", "Oso Café"),
    ("粉熊", "Pink Bear", "රෝස වලසා", "Oso Rosa"),
    ("蓝熊", "Blue Bear", "නිල් වලසා", "Oso Azul"),
    ("红点", "Red Polka Dot", "රතු තිත්", "Puntos Rojos"),
    ("蓝点", "Blue Polka Dot", "නිල් තිත්", "Puntos Azules"),
    ("蓝草", "Blue Floral", "නිල් මල්", "Floral Azul"),
    ("粉草", "Pink Floral", "රෝස මල්", "Floral Rosa"),
    ("大红", "Bright Red", "තද රතු", "Rojo Intenso"),
    ("酒红", "Wine Red", "වයින් රතු", "Rojo Vino"),
    ("玫红", "Rose Red", "රෝස රතු", "Rojo Rosa"),
    ("墨绿", "Forest Green", "තද කොළ", "Verde Oscuro"),
    ("草绿", "Grass Green", "තණකොළ කොළ", "Verde Hierba"),
    ("米黄", "Cream Yellow", "ක්‍රීම් කහ", "Amarillo Crema"),
    ("茶色", "Tea Brown", "තේ දුඹුරු", "Marrón Té"),
    ("格子", "Plaid", "කොටු", "Cuadros"),
    ("灰色", "Grey", "අළු", "Gris"),
    ("米色", "Beige", "බේජ්", "Beige"),
    ("黑色", "Black", "කළු", "Negro"),
    ("白色", "White", "සුදු", "Blanco"),
    ("绿色", "Green", "කොළ", "Verde"),
    ("蓝色", "Blue", "නිල්", "Azul"),
    ("粉色", "Pink", "රෝස", "Rosa"),
    ("棕色", "Brown", "දුඹුරු", "Marrón"),
    ("咖色", "Coffee", "කෝපි දුඹුරු", "Café"),
    ("杏色", "Apricot", "ඇප්‍රිකොට්", "Albaricoque"),
    ("紫色", "Purple", "දම්", "Púrpura"),
    ("橙", "Orange", "තැඹිලි", "Naranja"),
    ("黄", "Yellow", "කහ", "Amarillo"),
    ("红", "Red", "රතු", "Rojo"),
    ("灰", "Grey", "අළු", "Gris"),
    ("米", "Beige", "බේජ්", "Beige"),
    ("黑", "Black", "කළු", "Negro"),
    ("白", "White", "සුදු", "Blanco"),
    ("绿", "Green", "කොළ", "Verde"),
    ("蓝", "Blue", "නිල්", "Azul"),
    ("粉", "Pink", "රෝස", "Rosa"),
    ("咖", "Coffee", "කෝපි", "Café"),
]

SIZE_TOKEN_MAP = [
    ("超大号", "XXL", "ඉතා විශාල (XXL)", "XXL"),
    ("特大号", "XL", "විශාලතම (XL)", "XL"),
    ("超大", "XXL", "ඉතා විශාල (XXL)", "XXL"),
    ("特大", "XL", "විශාලතම (XL)", "XL"),
    ("加大", "Plus", "විශාල (Plus)", "Plus"),
    ("大号", "Large", "විශාල (L)", "Grande"),
    ("中号", "Medium", "මධ්‍යම (M)", "Mediano"),
    ("小号", "Small", "කුඩා (S)", "Pequeño"),
    ("三件套装", "3-Pc Set", "කොටස් 3 කට්ටලය", "Set 3 Pzs"),
    ("整套", "Full Set", "සම්පූර්ණ කට්ටලය", "Set Completo"),
    ("圆柱", "Cylinder", "සිලින්ඩරාකාර", "Cilindro"),
    ("半球", "Dome", "අර්ධ ගෝලාකාර", "Domo"),
    ("文胸带圈", "Framed Bra Bag", "රාමු සහිත බ්‍රා බෑගය", "Con Aro"),
    ("牛仔裤", "Jeans", "ඩෙනිම් කලිසම්", "Jeans"),
    ("T恤", "T-Shirt", "ටී-ෂර්ට්", "Camisetas"),
    ("袜子", "Socks", "මේස්", "Calcetines"),
    ("文胸", "Bra", "බ්‍රා", "Sujetador"),
    ("裤子", "Pants", "කලිසම්", "Pantalones"),
    ("短袖", "Shirts", "කමිස", "Camisas"),
    ("裙子", "Dresses", "ගවුම්", "Vestidos"),
    ("外套", "Coats", "කබා", "Abrigos"),
    ("网袋", "Mesh Bag", "දැල් බෑගය", "Bolsa Malla"),
    ("袋子", "Pouch", "බෑගය", "Bolsa"),
    ("鞋包", "Shoe Bag", "සපත්තු බෑගය", "Bolsa Zapatos"),
    ("长方盒", "Rectangular", "සෘජුකෝණාස්‍ර", "Rectangular"),
    ("无格", "No-Grid", "කාමර රහිත", "Sin Divisiones"),
    ("长", "Long", "දිගු", "Largo"),
    ("大", "Large", "විශාල", "Grande"),
    ("小", "Small", "කුඩා", "Pequeño"),
    ("中", "Medium", "මධ්‍යම", "Mediano"),
]


def translate_variant(color_raw, size_raw):
    """Translate color and size/spec into clean 4-language variant strings."""
    c_zh = re.sub(r"\s+", " ", str(color_raw or "").strip())
    s_zh = re.sub(r"\s+", " ", str(size_raw or "").strip()).rstrip("：:")

    # If color_raw is actually a long material spec (from A1:Sheet2 col 9), summarize it cleanly
    if len(c_zh) > 18 and ("无纺布" in c_zh or "pp" in c_zh.lower() or "纸板" in c_zh or "钢架" in c_zh):
        c_zh = "鱼骨纹灰/米色"

    def trans_size(lang_idx):
        res = s_zh
        # Replace N格 / N层
        if lang_idx == 0:  # en
            res = re.sub(r"(\d+)\s*格", r"\1-Grid ", res)
            res = re.sub(r"(\d+)\s*层", r"\1-Tier ", res)
            res = re.sub(r"(\d+)\s*个", r"\1 pcs", res)
        elif lang_idx == 1:  # si
            res = re.sub(r"(\d+)\s*格", r"කාමර \1 ", res)
            res = re.sub(r"(\d+)\s*层", r"තට්ටු \1 ", res)
            res = re.sub(r"(\d+)\s*个", r"කෑලි \1", res)
        elif lang_idx == 2:  # es
            res = re.sub(r"(\d+)\s*格", r"\1-Div. ", res)
            res = re.sub(r"(\d+)\s*层", r"\1-Niveles ", res)
            res = re.sub(r"(\d+)\s*个", r"\1 uds", res)
        for zh_tok, en_tok, si_tok, es_tok in SIZE_TOKEN_MAP:
            rep = (en_tok, si_tok, es_tok)[lang_idx]
            res = res.replace(zh_tok, rep + " ")
        res = res.replace("【", "(").replace("】", ")").replace("（", "(").replace("）", ")").replace("：", ": ")
        res = re.sub(r"\s+", " ", res).strip()
        if res and not re.search(r"(cm|CM|L|pcs|uds|Set|කට්ටලය)", res) and re.search(r"\d+\s*[*xX]\s*\d+", res):
            res += " cm"
        return res

    def trans_color(lang_idx):
        if not c_zh:
            return ""
        if c_zh == "鱼骨纹灰/米色":
            return ("Herringbone Grey/Beige", "හෙරින්බෝන් අළු/බේජ්", "Gris/Beige Espiga")[lang_idx]
        res = c_zh
        for zh_tok, en_tok, si_tok, es_tok in COLOR_TOKEN_MAP:
            rep = (en_tok, si_tok, es_tok)[lang_idx]
            res = res.replace(zh_tok, rep + " / ")
        # Clean separators
        res = re.sub(r"[、，,/\\\n]+", " / ", res)
        parts = []
        for p in res.split("/"):
            p = p.strip()
            # Strip any leftover CJK characters
            p = re.sub(r"[\u4e00-\u9fff]+", "", p).strip()
            if p and p not in parts:
                parts.append(p)
        return " / ".join(parts[:4])

    s_en, s_si, s_es = trans_size(0), trans_size(1), trans_size(2)
    c_en, c_si, c_es = trans_color(0), trans_color(1), trans_color(2)

    v_zh = " · ".join(x for x in (s_zh, c_zh) if x) or "标准规格"
    v_en = " · ".join(x for x in (s_en, c_en) if x) or "Standard"
    v_si = " · ".join(x for x in (s_si, c_si) if x) or "සම්මත ප්‍රමාණය (Standard)"
    v_es = " · ".join(x for x in (s_es, c_es) if x) or "Estándar"

    return v_zh, v_en, v_si, v_es


# =============================================================================
# 3. IMAGE DEDUPLICATION & WEBP OPTIMIZATION
# =============================================================================

class ImageStore:
    def __init__(self, out_dir):
        self.out_dir = out_dir
        os.makedirs(self.out_dir, exist_ok=True)
        self.hash_to_pair = {}
        self.counter = 0

    def save_image(self, raw_bytes):
        if not raw_bytes:
            return None, None
        h = hashlib.sha1(raw_bytes).hexdigest()
        if h in self.hash_to_pair:
            return self.hash_to_pair[h]

        self.counter += 1
        idx = self.counter
        thumb_rel = f"images/p{idx:03d}_t.webp"
        large_rel = f"images/p{idx:03d}_l.webp"
        thumb_abs = os.path.join(WEBSITE_DIR, "images", f"p{idx:03d}_t.webp")
        large_abs = os.path.join(WEBSITE_DIR, "images", f"p{idx:03d}_l.webp")

        if not (os.path.exists(thumb_abs) and os.path.exists(large_abs)):
            with Image.open(io.BytesIO(raw_bytes)) as im:
                im.load()
                if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                    rgba = im.convert("RGBA")
                    bg = Image.new("RGB", rgba.size, (255, 255, 255))
                    bg.paste(rgba, mask=rgba.split()[3])
                    rgb = bg
                else:
                    rgb = im.convert("RGB")

                # 1. Thumbnail (max side 360px, quality 82)
                t_im = rgb.copy()
                t_im.thumbnail((360, 360), Image.Resampling.LANCZOS)
                t_im.save(thumb_abs, "WEBP", quality=82, method=4)

                # 2. Large high-res (max side 1100px, quality 85)
                l_im = rgb.copy()
                l_im.thumbnail((1100, 1100), Image.Resampling.LANCZOS)
                l_im.save(large_abs, "WEBP", quality=85, method=4)

        self.hash_to_pair[h] = (thumb_rel, large_rel)
        return thumb_rel, large_rel


# =============================================================================
# 4. ROW & PACKING NORMALIZATION ACROSS ALL SHEETS
# =============================================================================

def clean_num(val):
    if val is None:
        return None
    s = str(val).strip()
    if not s or s == ".":
        return None
    if s == "0156":
        return 0.156
    if s == "011":
        return 0.11
    s = s.replace("..", ".")
    m = re.search(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", s)
    if m:
        return float(m.group(0))
    return None


def is_box_str(val):
    if val is None:
        return False
    s = str(val).strip()
    return bool(re.search(r"\d+(?:\.\d+)?\s*[*xX×]\s*\d+(?:\.\d+)?\s*[*xX×]\s*\d+", s))


def calc_cbm_from_box(box_str):
    if not box_str:
        return None
    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", box_str)]
    if len(nums) >= 3:
        l, w, h = nums[0], nums[1], nums[2]
        mult = nums[3] if len(nums) == 4 and nums[3] <= 5 else 1.0
        if l < 2.0 and w < 2.0 and h < 2.0:
            return round(l * w * h * mult, 4)
        return round((l * w * h * mult) / 1000000.0, 4)
    return None


def parse_packing_cols(d, pcs, start_col=7):
    box = ""
    cbm = None
    kg = None
    unit_g = None

    for c in sorted(k for k in d.keys() if k >= start_col):
        raw = str(d[c]).strip()
        if not raw or raw == ".":
            continue
        if is_box_str(raw):
            if not box:
                box = raw.replace(" ", "")
            continue
        u_upper = raw.upper()
        num = clean_num(raw)
        if num is None or num <= 0:
            continue
        if "KG" in u_upper:
            kg = round(num, 2)
            continue
        if u_upper.endswith("G") and not u_upper.endswith("KG"):
            unit_g = num
            continue
        if num < 1.0:
            if cbm is not None and c == 10 and num >= 0.4:
                unit_g = num * 1000.0
            else:
                cbm = round(num, 4)
        elif 1.0 <= num <= 4.0 and c == 10 and cbm is not None and kg is not None:
            unit_g = num * 1000.0
        elif 5.0 <= num < 85.0:
            if kg is not None and c >= 10:
                unit_g = num
            else:
                kg = round(num, 2)
        elif num >= 85.0:
            unit_g = num

    if cbm is None and box:
        cbm = calc_cbm_from_box(box)
    if cbm is None:
        # Fallback standard carton CBM when omitted in spreadsheet
        cbm = 0.14
    if kg is None and unit_g is not None and pcs:
        kg = round((unit_g * pcs) / 1000.0 + 0.8, 2)
    return box, round(cbm, 4), (round(kg, 2) if kg is not None else None)


def assign_images_to_rows(prod_rows, merges, anchors, img_store):
    """
    Assign primary and gallery images to every product row using merge-group
    awareness and anchor midpoints.
    Returns dict: row_idx -> list of (thumb_rel, large_rel)
    """
    row_to_group = {r: (r, r) for r in prod_rows}
    for r1, r2, c1, c2 in merges:
        if c1 in (0, 1, 2):
            for r in range(r1, r2 + 1):
                if r in row_to_group:
                    o1, o2 = row_to_group[r]
                    row_to_group[r] = (min(o1, r1), max(o2, r2))
    for r in prod_rows:
        g1, g2 = row_to_group[r]
        for rr in range(g1, g2 + 1):
            if rr in row_to_group:
                row_to_group[rr] = (min(g1, row_to_group[rr][0]), max(g2, row_to_group[rr][1]))

    row_raw_imgs = {r: [] for r in prod_rows}

    # Pass 1: Assign each anchor to the group containing its midpoint mid_r
    for fr_r, to_r, fc, tc, raw_bytes in anchors:
        mid_r = (fr_r + to_r) / 2.0
        for r in prod_rows:
            g1, g2 = row_to_group[r]
            if g1 <= mid_r <= g2:
                row_raw_imgs[r].append((fc, fr_r, raw_bytes))

    # Pass 2: For any row that still has no image, match anchors where fr_r <= r <= to_r
    for r in prod_rows:
        if not row_raw_imgs[r]:
            for fr_r, to_r, fc, tc, raw_bytes in anchors:
                if fr_r <= r <= to_r:
                    row_raw_imgs[r].append((fc, fr_r, raw_bytes))

    row_webp = {}
    for r in prod_rows:
        # Sort so col 0 (main column A picture) comes first, then secondary columns
        items = sorted(row_raw_imgs[r], key=lambda x: (0 if x[0] == 0 else 1, x[0], x[1]))
        pairs = []
        seen_t = set()
        for _, _, raw_bytes in items:
            t_rel, l_rel = img_store.save_image(raw_bytes)
            if t_rel and t_rel not in seen_t:
                seen_t.add(t_rel)
                pairs.append((t_rel, l_rel))
        row_webp[r] = pairs
    return row_webp


def extract_all_products(img_store):
    """Extract all products across A1.xlsx, B2.xls, and C2xlsx.xlsx."""
    products = []

    # Discover Excel files in ROOT_DIR
    excel_files = sorted(glob.glob(os.path.join(ROOT_DIR, "*.xls*")))
    # Sort in logical series order: A1 -> B2 -> C2
    def sort_key(p):
        b = os.path.basename(p)
        if "A1" in b:
            return (0, b)
        if "B2" in b:
            return (1, b)
        if "C2" in b:
            return (2, b)
        return (3, b)

    excel_files.sort(key=sort_key)
    print("Found Excel files:", [os.path.basename(f) for f in excel_files])

    for fpath in excel_files:
        fname = os.path.basename(fpath)
        if zipfile.is_zipfile(fpath):
            sheets_data = parse_xlsx_file(fpath)
        else:
            sheets_data = parse_xls_file(fpath)

        for sname, rows, merges, anchors in sheets_data:
            # Propagate merged cell values
            for r1, r2, c1, c2 in merges:
                val = rows.get(r1, {}).get(c1, "")
                if val != "":
                    for rr in range(r1, r2 + 1):
                        for cc in range(c1, c2 + 1):
                            rows.setdefault(rr, {})[cc] = val

            # Detect layout type:
            # Layout B (A1.xlsx Sheet2): Header has '产品系列' in col 1 and '单价' in col 8
            is_series_sheet = any("产品系列" in str(rows.get(r, {}).get(1, "")) for r in range(5))

            if is_series_sheet:
                prod_rows = [
                    r for r in sorted(rows.keys())
                    if r >= 3
                    and rows[r].get(0) != "产品图片"
                    and rows[r].get(1, "") != ""
                    and not str(rows[r].get(1, "")).startswith("产品系列")
                    and clean_num(rows[r].get(8)) is not None
                ]
                row_imgs = assign_images_to_rows(prod_rows, merges, anchors, img_store)

                last_desc = ""
                last_mat = ""
                series_counters = {}
                for r in prod_rows:
                    d = rows[r]
                    desc_raw = d.get(1, "").strip() or last_desc
                    last_desc = desc_raw
                    mat_raw = d.get(9, "").strip() or last_mat
                    last_mat = mat_raw
                    size_raw = d.get(2, "").strip()

                    # Extract series model prefix e.g. MY200, GC188
                    m_code = re.match(r"^([A-Z]{2}\d+)", desc_raw)
                    base_code = m_code.group(1) if m_code else "JY-S2"
                    series_counters[base_code] = series_counters.get(base_code, 0) + 1
                    code = f"JY-{base_code}-{series_counters[base_code]:02d}"

                    unit_g = clean_num(d.get(3))
                    pcs = int(clean_num(d.get(5)) or 40)
                    box = str(d.get(6, "")).strip()
                    cbm = clean_num(d.get(7))
                    if cbm is None and box:
                        cbm = calc_cbm_from_box(box)
                    if cbm is None:
                        cbm = 0.11
                    price = round(float(clean_num(d.get(8)) or 0.0), 2)
                    kg = round((unit_g * pcs) / 1000.0 + 0.8, 2) if unit_g else None

                    zh_title, en_title, si_title, es_title, cat = translate_product_title(desc_raw, size_raw)
                    v_zh, v_en, v_si, v_es = translate_variant(mat_raw, size_raw)

                    imgs = row_imgs.get(r, [])
                    thumb_rel = imgs[0][0] if imgs else ""
                    large_rel = imgs[0][1] if imgs else ""

                    prod = {
                        "id": len(products) + 1,
                        "code": code,
                        "sheet": f"{fname}:{sname}",
                        "category": cat,
                        "zh": zh_title,
                        "en": en_title,
                        "si": si_title,
                        "es": es_title,
                        "variant": v_en,
                        "variant_en": v_en,
                        "variant_si": v_si,
                        "variant_zh": v_zh,
                        "variant_es": v_es,
                        "pcsPerCarton": pcs,
                        "price": price,
                        "cbm": round(cbm, 4),
                        "kg": kg,
                        "box": box,
                        "img": thumb_rel,
                        "imgLarge": large_rel,
                        "thumb": thumb_rel,
                        "large": large_rel,
                        "thumbs": [x[0] for x in imgs],
                        "images": [x[1] for x in imgs],
                    }
                    compute_sl_wpi(prod)
                    products.append(prod)
            else:
                # Standard layout (A1:Sheet1, B2:Sheet2, C2:Sheet1)
                prod_rows = [
                    r for r in sorted(rows.keys())
                    if r >= 2
                    and rows[r].get(0) not in ("PICTURE", "图片", "产品图片")
                    and rows[r].get(1) not in ("ITEM NO", "货号", "")
                    and rows[r].get(1) is not None
                ]
                row_imgs = assign_images_to_rows(prod_rows, merges, anchors, img_store)

                last_desc = ""
                last_color = ""
                for r in prod_rows:
                    d = rows[r]
                    code = str(d.get(1, "")).strip()
                    desc_raw = str(d.get(2, "")).strip()
                    if desc_raw:
                        last_desc = desc_raw
                    else:
                        desc_raw = last_desc

                    col3 = str(d.get(3, "")).strip()
                    col4 = str(d.get(4, "")).strip()

                    # Detect if col3 and col4 were swapped (col3 has dimensions like 24L:40*30*20)
                    if re.search(r"\d+\s*[*xX]\s*\d+", col3) and not re.search(r"\d+\s*[*xX]\s*\d+", col4):
                        size_raw = col3
                        color_raw = col4 if col4 else last_color
                    elif col3 == col4 and re.search(r"\d+\s*[*xX]\s*\d+", col3):
                        size_raw = col4
                        color_raw = "图片色"
                    else:
                        color_raw = col3 if col3 else last_color
                        size_raw = col4

                    if color_raw:
                        last_color = color_raw

                    pcs = int(clean_num(d.get(5)) or 1)
                    price = round(float(clean_num(d.get(6)) or 0.0), 2)

                    # In A1:Sheet1 rows 60..63 (JY-A1058..JY-A1061), col 7 is an alternate price
                    start_col = 8 if ("A1" in fname and 60 <= r <= 63) else 7
                    box, cbm, kg = parse_packing_cols(d, pcs, start_col=start_col)

                    zh_title, en_title, si_title, es_title, cat = translate_product_title(desc_raw, size_raw)
                    v_zh, v_en, v_si, v_es = translate_variant(color_raw, size_raw)

                    imgs = row_imgs.get(r, [])
                    thumb_rel = imgs[0][0] if imgs else ""
                    large_rel = imgs[0][1] if imgs else ""

                    prod = {
                        "id": len(products) + 1,
                        "code": code,
                        "sheet": f"{fname}:{sname}",
                        "category": cat,
                        "zh": zh_title,
                        "en": en_title,
                        "si": si_title,
                        "es": es_title,
                        "variant": v_en,
                        "variant_en": v_en,
                        "variant_si": v_si,
                        "variant_zh": v_zh,
                        "variant_es": v_es,
                        "pcsPerCarton": pcs,
                        "price": price,
                        "cbm": cbm,
                        "kg": kg,
                        "box": box,
                        "img": thumb_rel,
                        "imgLarge": large_rel,
                        "thumb": thumb_rel,
                        "large": large_rel,
                        "thumbs": [x[0] for x in imgs],
                        "images": [x[1] for x in imgs],
                    }
                    compute_sl_wpi(prod)
                    products.append(prod)

    return products


CATEGORIES_META = [
    {
        "id": "storage-boxes",
        "icon": "📦",
        "en": "Storage Boxes & Bins",
        "si": "ගබඩා පෙට්ටි (Storage Boxes)",
        "zh": "折叠收纳箱/百纳箱",
        "es": "Cajas de Almacenamiento"
    },
    {
        "id": "wardrobe-organizers",
        "icon": "👔",
        "en": "Wardrobe & Underwear Organizers",
        "si": "ඇඳුම් සහ යට ඇඳුම් සංවිධායක (Wardrobe Organizers)",
        "zh": "内衣盒/分格收纳",
        "es": "Organizadores de Armario y Ropa Interior"
    },
    {
        "id": "baskets-hampers",
        "icon": "🧺",
        "en": "Storage Baskets & Laundry Hampers",
        "si": "ගබඩා කූඩ සහ රෙදි බඳුන් (Baskets & Hampers)",
        "zh": "收纳筐/脏衣篮",
        "es": "Cestas y Cestos de Lavandería"
    },
    {
        "id": "ottomans-stools",
        "icon": "🪑",
        "en": "Storage Ottomans & Stools",
        "si": "ගබඩා බංකු (Storage Ottomans)",
        "zh": "折叠收纳凳",
        "es": "Pufs y Banquetas de Almacenamiento"
    },
    {
        "id": "travel-cosmetic",
        "icon": "🧳",
        "en": "Travel & Cosmetic Bags",
        "si": "සංචාරක සහ රූපලාවන්‍ය බෑග් (Travel & Cosmetic Bags)",
        "zh": "洗漱包/化妆包/旅行袋",
        "es": "Bolsas de Viaje y Cosméticos"
    },
    {
        "id": "hanging-dust-laundry",
        "icon": "🧥",
        "en": "Quilt Bags, Hanging & Shoe Care",
        "si": "එල්ලන බෑග්, සපත්තු පෙට්ටි සහ ආරක්ෂක කවර (Hanging & Care)",
        "zh": "棉被袋/挂袋/鞋盒/护洗袋",
        "es": "Bolsas de Edredón, Colgantes y Cuidado"
    },
]


def main():
    os.makedirs(IMAGES_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)

    img_store = ImageStore(IMAGES_DIR)
    products = extract_all_products(img_store)

    # Compute category counts
    cat_counts = {}
    for p in products:
        c = p["category"]
        cat_counts[c] = cat_counts.get(c, 0) + 1

    categories = []
    for cm in CATEGORIES_META:
        c_copy = dict(cm)
        c_copy["count"] = cat_counts.get(cm["id"], 0)
        categories.append(c_copy)

    catalog = {
        "title": {
            "en": "Jiayu 2026 Home Organization & Storage Wholesale Catalogue",
            "si": "ජියායු 2026 ගෘහස්ථ ගබඩා සහ සංවිධායක තොග නාමාවලිය (Product Catalogue)",
            "zh": "2026年佳誉家居收纳综合报价产品目录",
            "es": "Catálogo Mayorista de Organización y Almacenamiento del Hogar Jiayu 2026"
        },
        "currency": "RMB",
        "exchangeRates": {
            "usdPerRmb": round(1.0 / 7.15, 6),
            "rmbPerUsd": 7.15,
            "lkrPerRmb": 43.50
        },
        "totalProducts": len(products),
        "totalUniqueImages": img_store.counter,
        "categories": categories,
        "products": products
    }

    out_path = os.path.join(DATA_DIR, "products.js")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("window.CATALOG = " + json.dumps(catalog, ensure_ascii=False, indent=2) + ";\n")

    tier_dist = {}
    for p in products:
        tier_dist[p["winTier"]] = tier_dist.get(p["winTier"], 0) + 1

    print(f"Successfully extracted {len(products)} products and {img_store.counter} unique WebP image pairs.")
    print(f"Categories: {cat_counts}")
    print(f"SL-WPI Tiers: {tier_dist}")
    print(f"Wrote {out_path} ({os.path.getsize(out_path):,} bytes)")


if __name__ == "__main__":
    main()
