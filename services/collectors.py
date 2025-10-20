# -*- coding: utf-8 -*-
from .header_utils import col_by_key
from . import validators as V


def _is_empty_row(hmap, row):
    for idx in hmap.values():
        if idx is not None and idx < len(row):
            if (row[idx] or "").strip():
                return False
    return True


def _collect_bom_rows(env, hmap, body, idx_start, buf):
    records = []
    ctx = {"bom_product_name": "", "bom_code": "", "bom_qty": "1", "bom_uom": ""}

    for i, row in enumerate(body, start=idx_start):
        def raw(k):
            return (col_by_key(hmap, row, k) or "").strip()

        bp, bcd, bqt, bu = raw("bom_product_name"), raw("bom_code"), raw("bom_qty"), raw("bom_uom")
        cn, cq, cu = raw("component_name"), raw("component_qty"), raw("component_uom")

        if not (bp or bcd or bqt or bu or cn or cq or cu):
            continue

        if bp or bcd or bqt or bu:
            ctx = {
                "bom_product_name": bp or ctx["bom_product_name"],
                "bom_code": bcd or ctx["bom_code"],
                "bom_qty": bqt or ctx["bom_qty"],
                "bom_uom": bu or ctx["bom_uom"],
            }

        if not (cn or cq or cu):
            continue

        if not cn:
            continue

        vals = {
            "bom_product_name": ctx["bom_product_name"],
            "bom_code": ctx["bom_code"],
            "bom_qty": ctx["bom_qty"],
            "bom_uom": ctx["bom_uom"],
            "component_name": cn,
            "component_qty": cq,
            "component_uom": cu,
        }
        rec = V.validate_bom_row(env, resolvers=None, hmap={}, row=vals, idx=i, buf=buf)  # resolvers not needed here
        records.append(rec)

    return records


def collect(env, hmap, body, target):
    """
    Returns: records, log, count, errors
    """
    log, records = [], []
    if target in ("mrp.bom.semi_finished", "mrp.bom.finished"):
        records = _collect_bom_rows(env, hmap, body, idx_start=2, buf=log)
        bad_lines = set()
        for l in log:
            if l.startswith("Line "):
                bad_lines.add(l.split(":")[0].split()[-1])
        return records, log, len(records), len(bad_lines)

    count_nonempty = 0
    from . import resolvers as R  # local import to avoid cycles

    for i, row in enumerate(body, start=2):
        if _is_empty_row(hmap, row):
            continue
        count_nonempty += 1
        if target == "product.category":
            vals = V.validate_category(hmap, row, i, log)
        elif target in ("product.raw", "product.semi_finished", "product.finished"):
            vals = V.validate_product_generic(env, R, hmap, row, i, log)
        elif target == "stock.location":
            vals = V.validate_location(hmap, row, i, log)
        else:
            vals = V.validate_bom_row(env, R, hmap, row, i, log)
        records.append(vals)

    bad_lines = set()
    for l in log:
        if l.startswith("Line "):
            bad_lines.add(l.split(":")[0].split()[-1])
    return records, log, count_nonempty, len(bad_lines)
