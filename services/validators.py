# -*- coding: utf-8 -*-
from odoo import _, exceptions

from .header_utils import col_by_key, map_usage_canonical


def _log(buf, line=None, msg=""):
    text = f"Line {line}: {msg}" if line is not None else msg
    buf.append(text)


def validate_category(hmap, row, idx, buf):
    col = lambda k: col_by_key(hmap, row, k)
    name = col("name")
    parent_name = col("parent_name") or ""
    if not name:
        _log(buf, idx, _("Missing category name"))
    return {"name": name, "parent_name": parent_name}


def validate_product_generic(env, resolvers, hmap, row, idx, buf):
    col = lambda k: col_by_key(hmap, row, k)
    name = col("name")
    uom_name = col("uom_name")
    price_raw = col("list_price") or ""
    cat_name = col("category_name") or ""

    if not name:
        _log(buf, idx, _("Missing product Name"))
    if not uom_name:
        _log(buf, idx, _("Missing product Unit (Unité)"))

    list_price = 0.0
    if price_raw != "":
        try:
            list_price = float(price_raw)
        except Exception:
            _log(buf, idx, _("Invalid Price"))

    # UoM compatibility check vs existing
    if name and uom_name:
        try:
            sheet_uom = resolvers.uom_by_name(env, uom_name, strict=True)
        except Exception:
            _log(buf, idx, _("Unknown Unit of Measure: %s") % uom_name)
            sheet_uom = False

        try:
            tmpl = resolvers.product_tmpl_by_name(env, name, strict=False)
            if tmpl and sheet_uom:
                db_cat = tmpl.uom_id.category_id
                in_cat = sheet_uom.category_id
                if db_cat and in_cat and db_cat.id != in_cat.id:
                    _log(
                        buf, idx,
                        _("UoM '%(uom)s' is incompatible with existing product '%(prod)s' UoM '%(dbuom)s' "
                          "(different categories: %(c1)s vs %(c2)s)") % {
                            "uom": sheet_uom.name,
                            "prod": tmpl.display_name,
                            "dbuom": tmpl.uom_id.name,
                            "c1": in_cat.display_name or in_cat.name,
                            "c2": db_cat.display_name or db_cat.name,
                        }
                    )
        except Exception:
            pass

    return {
        "name": name,
        "uom_name": uom_name,
        "list_price": list_price,
        "category_name": cat_name,
    }


def validate_location(hmap, row, idx, buf):
    col = lambda k: col_by_key(hmap, row, k)
    name = col("name")
    usage = map_usage_canonical(col("usage"))
    parent_name = col("parent_name") or ""
    if not name:
        _log(buf, idx, _("Missing Location Name"))
    from . import constants as C
    if usage not in C.USAGE_CHOICES:
        _log(buf, idx, _("Invalid Type (internal/view/supplier/customer/inventory/production/transit)"))
    return {"name": name, "usage": usage or "internal", "parent_name": parent_name}


def validate_bom_row(env, resolvers, hmap, row, idx, buf):
    if isinstance(row, dict):
        col = lambda k: (row.get(k) or "").strip()
    else:
        col = lambda k: (col_by_key(hmap, row, k) or "").strip()

    p_name = col("bom_product_name")
    b_code = col("bom_code") or ""
    b_qty = col("bom_qty") or "1"
    b_uom = col("bom_uom") or ""
    b_type = "normal"

    c_name = col("component_name")
    c_qty = col("component_qty") or ""
    c_uom = col("component_uom") or ""

    if not p_name:
        _log(buf, idx, _("Missing finished/fabricated product name"))

    if not c_name:
        try:
            b_qty_f = float(b_qty)
            b_qty_f = b_qty_f if b_qty_f > 0 else 1.0
        except Exception:
            b_qty_f = 1.0
        return {
            "bom_product_name": p_name,
            "bom_code": b_code,
            "bom_qty": b_qty_f,
            "bom_uom": b_uom,
            "bom_type": b_type,
            "component_name": "",
            "component_qty": 0.0,
            "component_uom": c_uom,
        }

    try:
        c_qty_f = float(c_qty)
        if c_qty_f <= 0:
            raise Exception
    except Exception:
        _log(buf, idx, _("Invalid component_qty (must be > 0)"))
        c_qty_f = 0.0

    try:
        b_qty_f = float(b_qty)
        if b_qty_f <= 0:
            raise Exception
    except Exception:
        _log(buf, idx, _("Invalid BOM quantity (must be > 0)"))
        b_qty_f = 1.0

    # UoM compatibility checks
    if p_name and b_uom:
        try:
            tmpl = resolvers.product_tmpl_by_name(env, p_name, strict=False)
            bom_uom = resolvers.uom_by_name(env, b_uom, strict=True)
            if tmpl and bom_uom:
                db_cat = tmpl.uom_id.category_id
                in_cat = bom_uom.category_id
                if db_cat and in_cat and db_cat.id != in_cat.id:
                    _log(
                        buf, idx,
                        _("BOM UoM '%(uom)s' is incompatible with semi-finished '%(prod)s' UoM '%(dbuom)s' "
                          "(different categories: %(c1)s vs %(c2)s). Line %(line)d (columns A–D).") % {
                            "uom": bom_uom.name,
                            "prod": tmpl.display_name if tmpl else p_name,
                            "dbuom": tmpl.uom_id.name if tmpl else "-",
                            "c1": in_cat.display_name or in_cat.name,
                            "c2": db_cat.display_name or db_cat.name,
                            "line": idx,
                        }
                    )
        except Exception:
            pass

    if c_name and c_uom:
        try:
            comp = resolvers.product_by_name(env, c_name, strict=False)
            comp_uom_in = resolvers.uom_by_name(env, c_uom, strict=True)
            if comp and comp_uom_in:
                db_cat = comp.uom_id.category_id
                in_cat = comp_uom_in.category_id
                if db_cat and in_cat and db_cat.id != in_cat.id:
                    _log(
                        buf, idx,
                        _("Component UoM '%(uom)s' is incompatible with component '%(comp)s' UoM '%(dbuom)s' "
                          "(different categories: %(c1)s vs %(c2)s). Line %(line)d (columns F–H).") % {
                            "uom": comp_uom_in.name,
                            "comp": comp.display_name if comp else c_name,
                            "dbuom": comp.uom_id.name if comp else "-",
                            "c1": in_cat.display_name or in_cat.name,
                            "c2": db_cat.display_name or db_cat.name,
                            "line": idx,
                        }
                    )
        except Exception:
            pass

    return {
        "bom_product_name": p_name,
        "bom_code": b_code,
        "bom_qty": b_qty_f,
        "bom_uom": b_uom,
        "bom_type": b_type,
        "component_name": c_name,
        "component_qty": c_qty_f,
        "component_uom": c_uom,
    }
