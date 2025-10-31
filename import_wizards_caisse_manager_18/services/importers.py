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

def _next_from_sequence(env, code):
    seq = env["ir.sequence"].sudo().search([("code", "=", code)], limit=1)
    if not seq:
        # fallback: create a simple sequence if missing
        seq = env["ir.sequence"].sudo().create({
            "name": code, "code": code, "implementation": "standard", "prefix": "", "padding": 0,
        })
    return seq.next_by_code(code)

def import_products(env, records, sheet_key):
    """
    Import products for product.raw / product.semi_finished / product.finished.
    Also resolves UoM & Category reliably from many possible column aliases.
    """
    PT = env["product.template"]
    PC = env["product.category"]

    created = updated = 0
    details = []

    def _s(v):
        return (v or "").strip()

    def _f(v, default=0.0):
        try:
            if v is None or v == "":
                return default
            return float(str(v).replace(",", "."))
        except Exception:
            return default

    def _pick(vals, *keys):
        # first present & non-empty match
        for k in keys:
            v = vals.get(k)
            if v not in (None, ""):
                return v
        return ""

    def _resolve_category(cat_name_raw):
        name = _s(cat_name_raw)
        if not name:
            return None
        # exact first, then ilike
        return PC.search([("name", "=", name)], limit=1) or PC.search([("name", "ilike", name)], limit=1)

    def _resolve_uom(uom_name_raw):
        name = _s(uom_name_raw)
        return R.uom_by_name(env, name) if name else None

    for vals in records:
        name = _s(vals.get("name"))
        if not name:
            continue

        # price: accept list_price or standard_price
        price = _f(vals.get("list_price"), None)
        if price is None:
            price = _f(vals.get("standard_price"), 0.0)

        # UoM aliases commonly seen from collectors/templates
        uom_name = _s(_pick(
            vals,
            "uom", "uom_name", "uom_id", "uom_label", "uom (name)",
            "Unité", "unite", "unité"
        ))
        uom = _resolve_uom(uom_name)

        # Category aliases commonly seen
        cat_name = _s(_pick(
            vals,
            "category", "category_name", "category_label",
            "categ", "categ_name", "categ_label",
            "Catégorie", "categorie", "catégorie"
        ))
        categ = _resolve_category(cat_name)

        # per-sheet defaults
        defaults = {}
        if sheet_key == "product.raw":
            defaults.update({
                "type": "consu",
                "purchase_ok": True,
                "sale_ok": False,
                "available_in_pos": False,
            })
        elif sheet_key == "product.semi_finished":
            defaults.update({
                "type": "consu",
                "sale_ok": False,
                "purchase_ok": False,
                "available_in_pos": False,
            })
        elif sheet_key == "product.finished":
            defaults.update({
                "type": "consu",
                "sale_ok": True,
                "purchase_ok": False,
                "available_in_pos": True,
            })

        data = {
            "name": name,
            "standard_price": price,
            "list_price": price,
            "is_storable": True,
        }
        if uom:
            data["uom_id"] = uom.id
            data["uom_po_id"] = uom.id
        else:
            if uom_name:
                details.append(f"{name}: UoM introuvable -> '{uom_name}' (inchangé)")
        if categ:
            data["categ_id"] = categ.id
        else:
            if cat_name:
                details.append(f"{name}: Catégorie introuvable -> '{cat_name}' (inchangée)")

        data.update(defaults)

        # find/create and write
        tmpl = PT.search([("name", "=", name)], limit=1) or PT.search([("name", "ilike", name)], limit=1)
        if tmpl:
            tmpl.write(data); updated += 1
        else:
            tmpl = PT.create(data); created += 1

        # barcode logic only for MP (raw)
        if sheet_key == "product.raw":
            mode = _s(_pick(vals, "barcode_mode", "barcode mode", "mode")).lower()
            manual = _s(_pick(vals, "barcode_manual", "barcode manual"))
            new_barcode = None  # None = don't touch; False = clear

            if mode == "auto":
                new_barcode = _next_from_sequence(env, "cm.barcode.raw")
            elif mode in ("manuel", "manual") and manual:
                new_barcode = manual
            elif mode in ("aucun", "none", "no"):
                new_barcode = False

            if new_barcode is not None:
                tmpl.product_variant_ids.write({"barcode": new_barcode})
                if new_barcode is False:
                    details.append(f"{name}: barcode cleared")
                else:
                    details.append(f"{name}: barcode set ({mode or 'n/a'})")

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
    """
    Finished BoM import + barcode handling from header row.

    Rules:
      • Group rows by (bom_product_name, bom_code->phantom).
      • The "header" for each group is the first row whose component_name is empty.
        From it we read:
          - barcode_mode  (Auto / Manuel / Aucun; case-insensitive)
          - barcode_manual (used only if mode == Manuel)
      • Apply barcode to ALL variants of the finished product template.
      • Create/Update a phantom BoM and replace its lines with component rows.
      • Component names must be purchasable products OR present in 'raw' sheet.
    """
    MrpBom = env["mrp.bom"]
    PT = env["product.template"]
    PP = env["product.product"]

    BARCODE_SEQ_CODE = "cm.barcode.finished"  # change if your sequence code differs

    created = updated = 0
    details = []

    # --- helpers -------------------------------------------------------------
    def _norm(s):
        return (s or "").strip().lower()

    def _cell_str(row, key):
        return (row.get(key) or "").strip()

    def _to_float(s, default=0.0):
        try:
            if s is None or s == "":
                return default
            return float(str(s).replace(",", "."))
        except Exception:
            return default

    def _next_barcode():
        # Prefer a specific sequence if you created one for finished goods barcodes
        seq = env["ir.sequence"].next_by_code(BARCODE_SEQ_CODE)
        if not seq:
            # Fallback: use the generic product barcode sequence if you have one,
            # else raise to make the issue explicit.
            seq = env["ir.sequence"].next_by_code("product.barcode")
        if not seq:
            raise ValidationError(_("Aucune séquence de code-barres n'est configurée "
                                     "(%s / product.barcode).") % BARCODE_SEQ_CODE)
        return seq

    # --- group rows by finished product + bom_code ---------------------------
    grouped = {}
    for r in rows:
        pname = _cell_str(r, "bom_product_name")
        if not pname:
            # we'll report later when trying to use; keep grouping but empty pname yields a separate key
            pass
        bcode = _cell_str(r, "bom_code") or "phantom"
        key = (pname, _norm(bcode))
        grouped.setdefault(key, []).append(r)

    # allow-list for components: purchasable OR listed in raw sheet
    db_purch = { _norm(n) for n in PP.search([("purchase_ok", "=", True)]).mapped("name") }
    norm_allowed_comp = db_purch | { _norm(x) for x in (raw_sheet_names or set()) }

    # --- process each group --------------------------------------------------
    for (prod_name, _code_norm), rows_in_group in grouped.items():
        sheet_prod_name = prod_name or ""
        # header = first row having NO component_name (typed empty)
        header = next((r for r in rows_in_group if not _cell_str(r, "component_name")), rows_in_group[0])

        # read barcode mode from header
        mode_raw = _cell_str(header, "barcode_mode") or _cell_str(header, "barcode mode") or _cell_str(header, "mode")
        mode = _norm(mode_raw)  # auto / manuel / aucun
        manual = _cell_str(header, "barcode_manual") or _cell_str(header, "barcode manual")

        # Resolve finished product template (strict=False -> clearer error below)
        prod = R.product_by_name(env, sheet_prod_name, strict=False)
        tmpl = prod.product_tmpl_id if prod else (
            R.product_tmpl_by_name(env, sheet_prod_name, strict=False)
        )
        if not tmpl:
            details.append(f"[{sheet_prod_name}] produit introuvable — BoM/CB ignorés")
            continue

        # Must be a PoS finished product
        if not getattr(tmpl, "available_in_pos", False):
            raise ValidationError(_("Le produit fini '%s' n'est pas disponible en PoS (available_in_pos=False).") %
                                  (tmpl.display_name))

        # --- apply barcode to variants based on mode -------------------------
        if mode in ("auto", "automatique"):
            bc = _next_barcode()
            tmpl.product_variant_ids.write({"barcode": bc})
            details.append(f"[{tmpl.display_name}] code-barres affecté automatiquement ({bc})")
        elif mode in ("manuel", "manual"):
            if manual:
                tmpl.product_variant_ids.write({"barcode": manual})
                details.append(f"[{tmpl.display_name}] code-barres saisi manuellement ({manual})")
            else:
                details.append(f"[{tmpl.display_name}] mode=Manuel mais aucun code saisi — ignoré")
        elif mode in ("aucun", "none", "no"):
            tmpl.product_variant_ids.write({"barcode": False})
            details.append(f"[{tmpl.display_name}] code-barres effacé (mode=Aucun)")
        else:
            # no mode provided -> leave as is
            if mode_raw:
                details.append(f"[{tmpl.display_name}] mode code-barres inconnu '{mode_raw}' — ignoré")

        # --- BoM header values ----------------------------------------------
        qty = _to_float(header.get("bom_qty"), 1.0) or 1.0
        if qty <= 0:
            qty = 1.0
        uom = R.uom_by_name(env, header.get("bom_uom")) or tmpl.uom_id
        btype = "phantom"
        code = _cell_str(header, "bom_code")

        # Find or create BoM (phantom) for this template (or any variant)
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

        # --- BoM lines from component rows (rows that DO have component_name) -
        line_vals = []
        for r in rows_in_group:
            cname = _cell_str(r, "component_name")
            if not cname:
                continue
            if _norm(cname) not in norm_allowed_comp:
                raise ValidationError(
                    _("Composant '%s' non autorisé. Utilisez une MP existante achetable ou une MP saisie dans 'Produits – MP'.") %
                    cname
                )

            comp = R.product_by_name(env, cname, strict=True)
            comp_uom = R.uom_by_name(env, r.get("component_uom")) or comp.uom_id
            cqty = _to_float(r.get("component_qty"), 0.0)
            line_vals.append((0, 0, {
                "product_id": comp.id,
                "product_qty": cqty,
                "product_uom_id": comp_uom.id,
            }))

        if line_vals:
            # replace lines
            bom.write({"bom_line_ids": [(5, 0, 0)] + line_vals})

        details.append(f"mrp.bom({bom.id}) {status} for template {tmpl.display_name} (code={code or '-'})")

    return created, updated, details
