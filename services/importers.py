# -*- coding: utf-8 -*-
from odoo import _, exceptions
from odoo.exceptions import ValidationError
from odoo.osv import expression

from .header_utils import norm
from . import resolvers as R


def import_categories(env, rows):
    created = updated = 0
    details = []
    for r in rows:
        v = R.resolve_category_vals(env, r)
        dom = [("name", "=", v["name"])]
        rec, status = R.upsert(env, "product.category", dom, v)
        if status == "created":
            created += 1
        else:
            updated += 1
        details.append(f"product.category({rec.id}) {status}")
    return created, updated, details


def import_products(env, rows, target_key):
    PT = env["product.template"]
    created = updated = 0
    details = []
    for r in rows:
        v = R.resolve_product_vals(env, r, target_key)
        rec = PT.search([("name", "=", v["name"])], limit=1)
        if rec:
            rec.write(v)
            updated += 1
            details.append(f"product.template({rec.id}) updated")
        else:
            rec = PT.create(v)
            created += 1
            details.append(f"product.template({rec.id}) created")
    return created, updated, details


def import_locations(env, rows):
    created = updated = 0
    details = []
    for r in rows:
        v = R.resolve_location_vals(env, r)
        dom = [("name", "=", v["name"])]
        if v.get("location_id"):
            dom.append(("location_id", "=", v["location_id"]))
        rec, status = R.upsert(env, "stock.location", dom, v)
        if status == "created":
            created += 1
        else:
            updated += 1
        details.append(f"stock.location({rec.id}) {status}")
    return created, updated, details


def import_boms_semi_finished(env, rows, allowed_names=None, raw_sheet_names=None):
    MrpBom = env["mrp.bom"]
    PP = env["product.product"]
    created = updated = 0
    details = []

    grouped = {}
    for r in rows:
        pname = norm(r.get("bom_product_name"))
        code = norm(r.get("bom_code")) or "normal"
        key = (pname, code)
        grouped.setdefault(key, {"head": r, "lines": []})
        grouped[key]["lines"].append(r)

    # allowlists
    norm_allowed_names = {norm(x) for x in (allowed_names or set())} if allowed_names is not None else None
    db_purch = {norm(n) for n in PP.search([("purchase_ok", "=", True)]).mapped("name")}
    norm_allowed_comp = db_purch | ({norm(x) for x in (raw_sheet_names or set())})

    for (_pname_norm, _code_norm), pack in grouped.items():
        head = pack["head"]

        bom_prod_name = (head.get("bom_product_name") or "").strip()
        code_raw = (head.get("bom_code") or "").strip()

        if norm_allowed_names is not None and norm(bom_prod_name) not in norm_allowed_names:
            raise ValidationError(
                _("Le produit de nomenclature '%s' n’est pas listé dans l’onglet 'Produits – Semi-finis' du fichier.") %
                (bom_prod_name or "")
            )

        prod = R.product_by_name(env, bom_prod_name, strict=False)
        tmpl = prod.product_tmpl_id if prod else R.product_tmpl_by_name(env, bom_prod_name, strict=True)

        if not (tmpl.type == "consu" and not tmpl.sale_ok):
            raise ValidationError(
                _("Le produit '%s' n'est pas identifié comme 'fabriqué / semi-fini' (type=consu et non vendu).") %
                tmpl.display_name
            )

        try:
            qty = float(head.get("bom_qty") or 1.0)
            if qty <= 0:
                qty = 1.0
        except Exception:
            qty = 1.0
        uom = R.uom_by_name(env, head.get("bom_uom")) or tmpl.uom_id
        btype = "normal"

        domain_base = [("type", "=", btype), ("company_id", "in", [False, env.company.id])]
        any_variant = PP.search([("product_tmpl_id", "=", tmpl.id)], limit=1)
        domain_prod = ['|', ("product_tmpl_id", "=", tmpl.id), ("product_id", "=", any_variant.id or 0)]
        if code_raw:
            domain = expression.AND([domain_base, domain_prod, [("code", "=", code_raw)]])
        else:
            domain = expression.AND([domain_base, domain_prod])

        bom = MrpBom.search(domain, limit=1)

        vals = {
            "product_tmpl_id": tmpl.id,
            "product_qty": qty,
            "product_uom_id": uom.id,
            "code": code_raw or False,
            "type": btype,
        }
        if bom:
            bom.write(vals)
            status = "updated"; updated += 1
        else:
            bom = MrpBom.create(vals)
            status = "created"; created += 1

        line_vals = []
        for r in pack["lines"]:
            cname = (r.get("component_name") or "").strip()
            if cname and norm(cname) not in norm_allowed_comp:
                raise ValidationError(
                    _("Composant '%s' non autorisé. Utilisez une MP existante achetable ou une MP saisie dans 'Produits – MP'.") %
                    cname
                )
            if not cname:
                continue
            comp = R.product_by_name(env, cname, strict=True)
            comp_uom = R.uom_by_name(env, r.get("component_uom")) or comp.uom_id
            try:
                cqty = float(r.get("component_qty") or 0.0)
            except Exception:
                cqty = 0.0
            line_vals.append((0, 0, {
                "product_id": comp.id,
                "product_qty": cqty,
                "product_uom_id": comp_uom.id,
            }))

        if line_vals:
            bom.write({"bom_line_ids": [(5, 0, 0)] + line_vals})

        details.append(f"mrp.bom({bom.id}) {status} for template {tmpl.display_name} (code={code_raw or '-'})")

    return created, updated, details


def import_boms_finished(env, rows, raw_sheet_names=None):
    MrpBom = env["mrp.bom"]
    PP = env["product.product"]
    created = updated = 0
    details = []

    grouped = {}
    for r in rows:
        key = (r.get("bom_product_name") or "", norm(r.get("bom_code") or "phantom"))
        grouped.setdefault(key, {"head": r, "lines": []})["lines"].append(r)

    db_purch = {norm(n) for n in PP.search([("purchase_ok", "=", True)]).mapped("name")}
    norm_allowed_comp = db_purch | ({norm(x) for x in (raw_sheet_names or set())})

    for (_pname, _code_norm), pack in grouped.items():
        head = pack["head"]

        prod = R.product_by_name(env, head.get("bom_product_name"), strict=True)
        tmpl = prod.product_tmpl_id
        if not tmpl.available_in_pos:
            raise ValidationError(_("Le produit fini '%s' n'est pas disponible en PoS (available_in_pos=False).") %
                                  (tmpl.display_name))

        try:
            qty = float(head.get("bom_qty") or 1.0)
            if qty <= 0:
                qty = 1.0
        except Exception:
            qty = 1.0
        uom = R.uom_by_name(env, head.get("bom_uom")) or tmpl.uom_id
        btype = "phantom"
        code = (head.get("bom_code") or "").strip()

        domain_base = [("type", "=", btype), ("company_id", "in", [False, env.company.id])]
        any_variant = PP.search([("product_tmpl_id", "=", tmpl.id)], limit=1)
        domain_prod = ['|', ("product_tmpl_id", "=", tmpl.id), ("product_id", "=", any_variant.id or 0)]
        domain = expression.AND([domain_base, domain_prod, [("code", "=", code)]]) if code else \
                 expression.AND([domain_base, domain_prod])
        bom = MrpBom.search(domain, limit=1)

        vals = {
            "product_tmpl_id": tmpl.id,
            "product_qty": qty,
            "product_uom_id": uom.id,
            "code": code or False,
            "type": btype,
        }
        if bom:
            bom.write(vals); status = "updated"; updated += 1
        else:
            bom = MrpBom.create(vals); status = "created"; created += 1

        line_vals = []
        for r in pack["lines"]:
            cname = (r.get("component_name") or "").strip()
            if cname and norm(cname) not in norm_allowed_comp:
                raise ValidationError(
                    _("Composant '%s' non autorisé. Utilisez une MP existante achetable ou une MP saisie dans 'Produits – MP'.") %
                    cname
                )
            if not cname:
                continue
            comp = R.product_by_name(env, cname, strict=True)
            comp_uom = R.uom_by_name(env, r.get("component_uom")) or comp.uom_id
            try:
                cqty = float(r.get("component_qty") or 0.0)
            except Exception:
                cqty = 0.0
            line_vals.append((0, 0, {"product_id": comp.id, "product_qty": cqty, "product_uom_id": comp_uom.id}))

        if line_vals:
            bom.write({"bom_line_ids": [(5, 0, 0)] + line_vals})

        details.append(f"mrp.bom({bom.id}) {status} for template {tmpl.display_name} (code={code or '-'})")

    return created, updated, details
