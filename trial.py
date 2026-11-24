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

def apply_tint(hex_color, tint):
    if not hex_color or len(hex_color) != 6 or not tint:
        return hex_color
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    if tint < 0:
        r, g, b = int(r * (1 + tint)), int(g * (1 + tint)), int(b * (1 + tint))
    else:
        r = int(r + (255 - r) * tint)
        g = int(g + (255 - g) * tint)
        b = int(b + (255 - b) * tint)
    return "{:02X}{:02X}{:02X}".format(max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))

def get_xlsx_theme_colors(wb):
    theme_colors = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]
    try:
        if wb.loaded_theme:
            import xml.etree.ElementTree as ET
            root = ET.fromstring(wb.loaded_theme)
            ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            clr_scheme = root.find('.//a:clrScheme', ns)
            if clr_scheme is not None:
                order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
                extracted = []
                for tag in order:
                    node = clr_scheme.find(f'a:{tag}', ns)
                    if node is not None:
                        srgb = node.find('a:srgbClr', ns)
                        sysclr = node.find('a:sysClr', ns)
                        if srgb is not None and 'val' in srgb.attrib:
                            extracted.append(srgb.attrib['val'].upper())
                        elif sysclr is not None and 'lastClr' in sysclr.attrib:
                            extracted.append(sysclr.attrib['lastClr'].upper())
                if len(extracted) == 10:
                    theme_colors = extracted
    except Exception:
        pass
    return theme_colors

def resolve_openpyxl_color(color_obj, theme_colors, is_bg=True):
    if not color_obj:
        return None
    ctype = getattr(color_obj, "type", None)
    if ctype == "rgb":
        rgb = str(color_obj.rgb)
        if len(rgb) == 8:
            alpha, hex_val = rgb[:2],
