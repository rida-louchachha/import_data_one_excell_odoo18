# -*- coding: utf-8 -*-
import unicodedata
from odoo import _

from . import constants as C


def slug(s):
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    for ch in (" ", "_", ".", "-", "’", "'"):
        s = s.replace(ch, "")
    return s


def alias_map(target):
    amap = {}
    for canon, _label in C.HEADER_LABELS.get(target, []):
        amap[slug(canon)] = canon
    for canon, aliases in C.HEADER_ALIASES.get(target, {}).items():
        for a in aliases:
            amap[slug(a)] = canon
    for canon, label in C.HEADER_LABELS.get(target, []):
        amap[slug(label)] = canon
    return amap


def required_cols(target):
    if target == "product.category":
        return ["name"]
    if target in ("product.raw", "product.semi_finished", "product.finished"):
        return ["name", "uom_name"]
    if target in ("mrp.bom.semi_finished", "mrp.bom.finished"):
        return ["bom_product_name", "component_name", "component_qty"]
    if target == "stock.location":
        return ["name", "usage"]
    return []


def build_header_index(header_cells, target):
    amap = alias_map(target)
    hmap, unknown = {}, []
    for idx, raw in enumerate(header_cells):
        key = amap.get(slug(raw or ""))
        if key and key not in hmap:
            hmap[key] = idx
        elif not key:
            unknown.append(raw or "")
    missing = [k for k in required_cols(target) if k not in hmap]
    return hmap, missing, unknown


def col_by_key(hmap, row, key):
    idx = hmap.get(key)
    if idx is None or idx >= len(row):
        return ""
    return row[idx]


def map_usage_canonical(value):
    v = (value or "").strip().lower()
    if v in C.USAGE_CHOICES:
        return v
    return C.USAGE_FR_TO_CANON.get(v, v)


def norm(s):
    return (s or "").strip().casefold()
