import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from copy import copy
from io import BytesIO
import xlrd
import re

st.set_page_config(page_title="BOM to Trial Template Filler", layout="wide")
st.title("BOM to Trial Template Auto-Filler")

def clean_header(text):
    return re.sub(r"\s+", " ", str(text)).strip().lower() if text is not None else ""

def rgb_to_hex(rgb):
    return "{:02X}{:02X}{:02X}".format(*rgb) if rgb and len(rgb) == 3 else None

def get_xls_styles(wb, sh, r, c):
    try:
        xf = wb.xf_list[sh.cell_xf_index(r, c)]
        bg_idx = xf.background.pattern_colour_index
        bg_hex = rgb_to_hex(wb.colour_map.get(bg_idx))
        if bg_hex in (None, "FFFFFF", "000000") and bg_idx in (64, 65, 9):
            bg_hex = None
        font = wb.font_list[xf.font_index]
        font_hex = rgb_to_hex(wb.colour_map.get(font.colour_index))
        if font_hex in (None, "000000") and font.colour_index in (32767, 8, 64):
            font_hex = None
        return bg_hex, font_hex
    except Exception:
        return None, None

def get_xlsx_styles(ws, r, c):
    try:
        cell = ws.cell(row=r + 1, column=c + 1)
        bg_hex, font_hex = None, None
        if cell.fill and cell.fill.fill_type == "solid" and cell.fill.start_color:
            rgb = str(cell.fill.start_color.rgb)
            if len(rgb) == 8 and rgb[2:] not in ("000000", "FFFFFF"):
                bg_hex = rgb[2:]
        if cell.font and cell.font.color and cell.font.color.rgb:
            frgb = str(cell.font.color.rgb)
            if len(frgb) == 8 and frgb[2:] != "000000":
                font_hex = frgb[2:]
        return bg_hex, font_hex
    except Exception:
        return None, None

def extract_bom_data(bom_file):
    fb = bom_file.getvalue()
    is_xls = bom_file.name.lower().endswith(".xls")
    xls_pd = pd.ExcelFile(BytesIO(fb))
    sname = next((s for s in xls_pd.sheet_names if "vse_bom" in s.lower() or "bom" in s.lower()), xls_pd.sheet_names[0])
    df = pd.read_excel(xls_pd, sheet_name=sname, header=None)

    if is_xls:
        wb_fmt = xlrd.open_workbook(file_contents=fb, formatting_info=True)
        sh_fmt = wb_fmt.sheet_by_name(sname)
    else:
        wb_fmt = openpyxl.load_workbook(BytesIO(fb), data_only=True)
        sh_fmt = wb_fmt[sname]

    meta = dict(customer_id="", ship_site="")
    hdr_idx = None
    for r in range(min(20, len(df))):
        rvals = [clean_header(v) for v in df.iloc[r]]
        for c, val in enumerate(rvals):
            for key, label in [("customer_id", "customer id"), ("ship_site", "ship site")]:
                if label in val:
                    for off in range(1, 4):
                        if c + off < len(df.columns) and pd.notna(df.iloc[r, c + off]):
                            meta[key] = str(df.iloc[r, c + off]).strip()
                            break
        if any("vse item" in v for v in rvals) and any("manufacturer" in v for v in rvals):
            hdr_idx = r
            break

    if hdr_idx is None:
        raise ValueError("Could not find 'VSE Item' and 'Manufacturer' headers in the BOM file.")

    bh = [clean_header(v) for v in df.iloc[hdr_idx]]
    cmap = dict()
    rules = [
        ("lvl", lambda h: h == "lvl"), ("vse_item", lambda h: h == "vse item"),
        ("cust_item", lambda h: h == "cust item"), ("description", lambda h: h == "description"),
        ("manufacturer", lambda h: h == "manufacturer"), ("mpn", lambda h: "manufacturer p/n" in h or "mpn" in h),
        ("mfg_id", lambda h: "manufacturer id" in h), ("planner", lambda h: "planner" in h),
        ("vendor_id", lambda h: "vendor id" in h), ("unit_cost", lambda h: "unit cost" in h),
    ]
    for idx, h in enumerate(bh):
        for k, fn in rules:
            if k not in cmap and fn(h):
                cmap[k] = idx

    last_row = hdr_idx
    for r in range(hdr_idx + 1, len(df)):
        if any(pd.notna(df.iloc[r, cmap[k]]) and str(df.iloc[r, cmap[k]]).strip() != "" for k in ["lvl", "vse_item", "manufacturer", "mpn"] if k in cmap):
            last_row = r

    records, parent = [], dict()
    for r in range(hdr_idx + 1, last_row + 1):
        row = df.iloc[r]
        val = lambda k: ("" if (str(row[cmap[k]]).strip() == "0" and k == "vendor_id") else str(row[cmap[k]]).strip()) if (k in cmap and pd.notna(row[cmap[k]])) else ""
        lvl, vse, mfg, mpn = val("lvl"), val("vse_item"), val("manufacturer"), val("mpn")
        ccol = cmap.get("mpn", cmap.get("vse_item", 0))
        bg, fg = get_xls_styles(wb_fmt, sh_fmt, r, ccol) if is_xls else get_xlsx_styles(sh_fmt, r, ccol)

        if not lvl and not vse and not mfg and not mpn:
            parent = dict()
            records.append(dict(bg_hex=bg, font_hex=fg))
            continue

        if lvl or vse:
            rc = row[cmap["unit_cost"]] if "unit_cost" in cmap and pd.notna(row[cmap["unit_cost"]]) else ""
            parent = {
                "vse item": vse, "cust item": val("cust_item"), "description": val("description"),
                "planner codes": val("planner"), "vendor id": val("vendor_id"), "unit cost": rc if rc != 0 else ""
            }

        has_part = bool(vse or mpn or mfg)
        records.append({
            "manufacturer id": val("mfg_id"), "manufacturer p/n": mpn, "manufatures": mfg, "manufacturer": mfg,
            "vse item": parent.get("vse item", ""), "description": parent.get("description", ""),
            "unit cost": parent.get("unit cost", ""), "planner codes": parent.get("planner codes", ""),
            "customer id": meta["customer_id"] if has_part else "", "cust item": parent.get("cust item", ""),
            "vendor id pricing from": parent.get("vendor id", ""), "vendor id": parent.get("vendor id", ""),
            "stock location": meta["ship_site"] if has_part else "", "bg_hex": bg, "font_hex": fg
        })
    return records

def fill_trial_template(trial_file, records):
    wb = openpyxl.load_workbook(trial_file)
    ws = wb.active
    th = [clean_header(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
    sec = set(["main", "planning", "controls", "vse", "item coset", "item cost", "customer contract", "item mfgr"])
    bf = copy(ws.cell(2, 1).font) if ws.max_row >= 2 else Font(name="Calibri", size=11)

    for r_idx, rec in enumerate(records, start=2):
        bg, fg = rec.get("bg_hex"), rec.get("font_hex")
        rfill = PatternFill(start_color=bg, end_color=bg, fill_type="solid") if bg else None
        for c_idx, hc in enumerate(th, start=1):
            v = "" if (hc in sec or not hc) else (rec.get("vendor id", "") if "vendor id" in hc else rec.get(hc, next((val for k, val in rec.items() if k not in ("bg_hex", "font_hex") and k and k in hc), "")))
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cf = copy(bf)
            if fg:
                cf.color = fg
            cell.font = cf
            if rfill:
                cell.fill = copy(rfill)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

c1, c2 = st.columns(2)
with c1:
    t_up = st.file_uploader("1. Upload Trial Template File (.xlsx)", type=["xlsx"])
with c2:
    b_up = st.file_uploader("2. Upload BOM File (.xls / .xlsx)", type=["xls", "xlsx"])

if t_up and b_up and st.button("Fill Template & Download", type="primary"):
    try:
        with st.spinner("Extracting BOM items and preserving formatting..."):
            recs = extract_bom_data(b_up)
            f_xl = fill_trial_template(t_up, recs)
            st.success(f"Successfully populated {len(recs)} rows with original colors and empty rows!")
            st.dataframe(pd.DataFrame(recs).drop(columns=["bg_hex", "font_hex"], errors="ignore").head(20), use_container_width=True)
            st.download_button("Download Filled Excel", f_xl, b_up.name.rsplit(".", 1)[0] + "_Trial_Filled.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Error: {e}")