# -*- coding: utf-8 -*-
import base64
from io import BytesIO
from odoo import api, fields, models, _
from odoo.exceptions import UserError, ValidationError
import unicodedata

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

# ----- constants / labels / aliases -----

USAGE_CHOICES = ["internal", "view", "supplier", "customer", "inventory", "production", "transit"]
USAGE_FR_TO_CANON = {
    "interne": "internal",
    "vue": "view",
    "fournisseur": "supplier",
    "client": "customer",
    "inventaire": "inventory",
    "production": "production",
    "transit": "transit",
}
BOM_TYPE_CHOICES = ["normal", "phantom", "subcontracting"]
BOM_TYPES = set(BOM_TYPE_CHOICES)

HEADER_LABELS = {
    # 1) Categories
    "product.category": [
        ("name", "Nom de la catégorie"),
        ("parent_name", "Catégorie parente"),
    ],

    # 2) Raw materials
    "product.raw": [
        ("name", "Nom du produit"),
        ("uom_name", "Unité"),
        ("list_price", "Prix"),
        ("category_name", "Catégorie"),
    ],

    # 2.5) Semi-finished / fabricated products
    "product.semi_finished": [
        ("name", "Nom du produit"),
        ("uom_name", "Unité"),
        ("list_price", "Prix"),
        ("category_name", "Catégorie"),
    ],

    # 3) Finished products
    "product.finished": [
        ("name", "Nom du produit"),
        ("uom_name", "Unité"),
        ("list_price", "Prix"),
        ("category_name", "Catégorie"),
    ],

    # 4) BOM for fabricated (semi-finished) products
    # NOTE: We keep bom_type in the schema but hide it in the Excel; importer forces 'phantom' anyway.
    "mrp.bom.semi_finished": [
        ("bom_product_name", "Produit fabriqué (nom)"),
        ("bom_code", "Code nomenclature"),
        ("bom_qty", "Quantité nomenclature (sortie)"),
        ("bom_uom", "Unité nomenclature"),
        ("bom_type", "Type de nomenclature (normal/phantom/subcontracting)"),
        ("component_name", "Nom composant"),
        ("component_qty", "Quantité composant"),
        ("component_uom", "Unité composant"),
    ],

    # (Optional) BOM for finished products (kept as before)
    "mrp.bom.finished": [
        ("bom_product_name", "Produit fini (nom)"),
        ("bom_code", "Code nomenclature"),
        ("bom_qty", "Quantité nomenclature (sortie)"),
        ("bom_uom", "Unité nomenclature"),
        ("bom_type", "Type de nomenclature (normal/phantom/subcontracting)"),
        ("component_name", "Nom composant"),
        ("component_qty", "Quantité composant"),
        ("component_uom", "Unité composant"),
    ],

    # Locations
    "stock.location": [
        ("name", "Nom de l’emplacement"),
        ("parent_name", "Emplacement parent"),
        ("usage", "Type (internal/view/fournisseur/client/inventory/production/transit)"),
    ],
}

HEADER_ALIASES = {
    "product.category": {
        "name": ["Nom de la catégorie", "Catégorie"],
        "parent_name": ["Catégorie parente", "Parent"],
    },
    "product.raw": {
        "name": ["Nom du produit"],
        "uom_name": ["Unité", "UoM", "Unité de mesure"],
        "list_price": ["Prix", "Prix de vente", "Tarif public"],
        "category_name": ["Catégorie", "Catégorie produit"],
    },
    "product.semi_finished": {
        "name": ["Nom du produit", "Nom produit fabriqué", "Nom semi-fini"],
        "uom_name": ["Unité", "UoM", "Unité de mesure"],
        "list_price": ["Prix", "Prix de vente"],
        "category_name": ["Catégorie", "Catégorie produit"],
    },
    "product.finished": {
        "name": ["Nom du produit", "Nom produit fini"],
        "uom_name": ["Unité", "UoM", "Unité de mesure"],
        "list_price": ["Prix", "Prix de vente"],
        "category_name": ["Catégorie", "Catégorie produit"],
    },
    "mrp.bom.semi_finished": {
        "bom_product_name": ["Produit fabriqué (nom)", "Nom produit fabriqué", "Produit semi-fini (nom)"],
        "bom_code": ["Code nomenclature", "Référence nomenclature"],
        "bom_qty": ["Quantité nomenclature (sortie)", "Qté nomenclature", "Qté BOM"],
        "bom_uom": ["Unité nomenclature", "UoM nomenclature"],
        "bom_type": ["Type de nomenclature", "Type BOM"],
        "component_name": ["Nom composant", "Composant"],
        "component_qty": ["Quantité composant", "Qté composant"],
        "component_uom": ["Unité composant", "UoM composant"],
    },
    "mrp.bom.finished": {
        "bom_product_name": ["Produit fini (nom)", "Nom produit fini"],
        "bom_code": ["Code nomenclature", "Référence nomenclature"],
        "bom_qty": ["Quantité nomenclature (sortie)", "Qté nomenclature", "Qté BOM"],
        "bom_uom": ["Unité nomenclature", "UoM nomenclature"],
        "bom_type": ["Type de nomenclature", "Type BOM"],
        "component_name": ["Nom composant", "Composant"],
        "component_qty": ["Quantité composant", "Qté composant"],
        "component_uom": ["Unité composant", "UoM composant"],
    },
    "stock.location": {
        "name": ["Nom de l’emplacement", "Nom emplacement"],
        "parent_name": ["Emplacement parent", "Parent"],
        "usage": ["Usage", "Type", "Type (internal/view/fournisseur/client/inventory/production/transit)"],
    },
}

_TARGETS = [
    ("product.category", "Categories"),
    ("product.raw", "Raw Materials (Products)"),
    ("product.semi_finished", "Semi-finished (Fabricated) Products"),
    ("product.finished", "Finished Products"),
    ("mrp.bom.semi_finished", "BOM for Fabricated (Semi-finished) Products"),
    ("mrp.bom.finished", "BOM for Finished Products"),
    ("stock.location", "Locations"),
]


class CmImportWizard(models.TransientModel):
    _name = "cm.import.wizard"
    _description = "CM Import Wizard (XLSX)"

    target_model = fields.Selection(selection=_TARGETS, required=False)
    file = fields.Binary("Excel File (.xlsx)", required=True)
    filename = fields.Char()
    log_text = fields.Text(readonly=True)
    line_count = fields.Integer(readonly=True)
    error_count = fields.Integer(readonly=True)
    tested_ok = fields.Boolean(readonly=True)

    # ---------- base helpers ----------
    def _require_openpyxl(self):
        if load_workbook is None:
            raise UserError(_("Missing Python dependency: openpyxl"))

    def _read_xlsx(self):
        self._require_openpyxl()
        if not self.file:
            raise UserError(_("Please upload an .xlsx file."))
        try:
            stream = BytesIO(base64.b64decode(self.file))
            wb = load_workbook(stream, read_only=True, data_only=True)
            ws = wb.active
        except Exception as e:
            raise UserError(_("Could not read file: %s") % e)
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append([("" if v is None else str(v).strip()) for v in row])
        if not data:
            raise UserError(_("The file is empty."))
        return data

    def _log(self, buf, line=None, msg=""):
        text = f"Line {line}: {msg}" if line is not None else msg
        buf.append(text)

    def _is_empty_row(self, hmap, row):
        for idx in hmap.values():
            if idx is not None and idx < len(row):
                if (row[idx] or "").strip():
                    return False
        return True

    def _get_or_create_category_by_name(self, name, parent_name=None):
        Cat = self.env["product.category"]
        cat = Cat.search([("name", "=", name)], limit=1)
        if cat:
            return cat
        vals = {"name": name}
        if parent_name:
            parent = Cat.search([("name", "=", parent_name)], limit=1)
            if parent:
                vals["parent_id"] = parent.id
        return Cat.create(vals)

    def _default_semi_finished_category(self):
        cat = self.env.ref(
            "import_wizards_caisse_manager_18.category_semi_finished_cm",
            raise_if_not_found=False,
        )
        if cat:
            return cat
        return self._get_or_create_category_by_name("Produits semi-finis")

    # ---------- header mapping ----------
    def _slug(self, s):
        if not s:
            return ""
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = s.lower()
        for ch in (" ", "_", ".", "-", "’", "'"):
            s = s.replace(ch, "")
        return s

    def _alias_map(self, target):
        amap = {}
        for canon, _label in HEADER_LABELS.get(target, []):
            amap[self._slug(canon)] = canon
        for canon, aliases in HEADER_ALIASES.get(target, {}).items():
            for a in aliases:
                amap[self._slug(a)] = canon
        for canon, label in HEADER_LABELS.get(target, []):
            amap[self._slug(label)] = canon
        return amap

    def _required_cols(self, target):
        if target == "product.category":
            return ["name"]
        if target in ("product.raw", "product.semi_finished", "product.finished"):
            return ["name", "uom_name"]
        if target in ("mrp.bom.semi_finished", "mrp.bom.finished"):
            return ["bom_product_name", "component_name", "component_qty"]
        if target == "stock.location":
            return ["name", "usage"]
        return []

    def _build_header_index(self, header_cells, target):
        amap = self._alias_map(target)
        hmap, unknown = {}, []
        for idx, raw in enumerate(header_cells):
            key = amap.get(self._slug(raw or ""))
            if key and key not in hmap:
                hmap[key] = idx
            elif not key:
                unknown.append(raw or "")
        required = self._required_cols(target)
        missing = [k for k in required if k not in hmap]
        return hmap, missing, unknown

    def _col_by_key(self, hmap, row, key):
        idx = hmap.get(key)
        if idx is None or idx >= len(row):
            return ""
        return row[idx]

    def _map_usage_canonical(self, value):
        v = (value or "").strip().lower()
        if v in USAGE_CHOICES:
            return v
        return USAGE_FR_TO_CANON.get(v, v)

    def _parse(self):
        rows = self._read_xlsx()
        raw_header = [h if h is not None else "" for h in rows[0]]
        hmap, missing, _unknown = self._build_header_index(raw_header, self.target_model)
        if missing:
            label_map = dict(HEADER_LABELS.get(self.target_model, []))
            pretty = [label_map.get(k, k) for k in missing]
            raise UserError(_("Missing required columns for %s: %s") %
                            (dict(_TARGETS).get(self.target_model), ", ".join(pretty)))
        body = rows[1:]
        return hmap, body

    # ---------- lookups ----------
    def _uom_by_name(self, name, strict=False):
        if not name:
            return False
        Uom = self.env["uom.uom"]
        u = Uom.search([("name", "=", name)], limit=1)
        if not u:
            u = Uom.search([("name", "ilike", name)], limit=1)
        if strict and not u:
            raise ValidationError(_("Unit of Measure not found: %s") % (name or ""))
        return u

    def _category_by_name(self, name, strict=False):
        if not name:
            return False
        Cat = self.env["product.category"]
        c = Cat.search([("name", "=", name)], limit=1)
        if not c:
            c = Cat.search([("name", "ilike", name)], limit=1)
        if strict and not c:
            raise ValidationError(_("Category not found: %s") % (name or ""))
        return c

    def _product_by_name(self, name, strict=False):
        Product = self.env["product.product"]
        p = Product.search([("name", "=", name)], limit=1)
        if not p:
            p = Product.search([("name", "ilike", name)], limit=1)
        if strict and not p:
            raise ValidationError(_("Product not found: %s") % (name or ""))
        return p

    def _product_tmpl_by_name(self, name, strict=False):
        T = self.env["product.template"]
        t = T.search([("name", "=", name)], limit=1)
        if not t:
            t = T.search([("name", "ilike", name)], limit=1)
        if strict and not t:
            raise ValidationError(_("Product template not found: %s") % (name or ""))
        return t

    # ---------- validators ----------
    def _validate_category(self, hmap, row, idx, buf):
        col = lambda k: self._col_by_key(hmap, row, k)
        name = col("name")
        parent_name = col("parent_name") or ""
        if not name:
            self._log(buf, idx, _("Missing category name"))
        return {"name": name, "parent_name": parent_name}

    def _validate_product_generic(self, hmap, row, idx, buf):
        col = lambda k: self._col_by_key(hmap, row, k)
        name = col("name")
        uom_name = col("uom_name")
        price_raw = col("list_price") or ""
        cat_name = col("category_name") or ""
        if not name:
            self._log(buf, idx, _("Missing product Name"))
        if not uom_name:
            self._log(buf, idx, _("Missing product Unit (Unité)"))
        list_price = 0.0
        if price_raw != "":
            try:
                list_price = float(price_raw)
            except Exception:
                self._log(buf, idx, _("Invalid Price"))
        return {
            "name": name,
            "uom_name": uom_name,
            "list_price": list_price,
            "category_name": cat_name,
        }

    def _validate_location(self, hmap, row, idx, buf):
        col = lambda k: self._col_by_key(hmap, row, k)
        name = col("name")
        usage = self._map_usage_canonical(col("usage"))
        parent_name = col("parent_name") or ""
        if not name:
            self._log(buf, idx, _("Missing Location Name"))
        if usage not in USAGE_CHOICES:
            self._log(buf, idx, _("Invalid Type (internal/view/supplier/customer/inventory/production/transit)"))
        return {"name": name, "usage": usage or "internal", "parent_name": parent_name}

    def _validate_bom_row(self, hmap, row, idx, buf):
        col = lambda k: self._col_by_key(hmap, row, k)
        p_name = col("bom_product_name")
        b_code = col("bom_code") or ""
        b_qty = col("bom_qty") or "1"
        b_uom = col("bom_uom") or ""
        # keep reading but we'll ignore user value for semi-finished imports
        b_type = (col("bom_type") or "normal").lower()
        c_name = col("component_name")
        c_qty = col("component_qty") or ""
        c_uom = col("component_uom") or ""

        if not p_name:
            self._log(buf, idx, _("Missing finished/fabricated product name"))
        if not c_name:
            self._log(buf, idx, _("Missing component name"))
        try:
            c_qty_f = float(c_qty)
            if c_qty_f <= 0:
                raise Exception
        except Exception:
            self._log(buf, idx, _("Invalid component_qty (must be > 0)"))
            c_qty_f = 0.0
        try:
            b_qty_f = float(b_qty)
            if b_qty_f <= 0:
                raise Exception
        except Exception:
            self._log(buf, idx, _("Invalid BOM quantity (must be > 0)"))
            b_qty_f = 1.0
        if b_type and b_type not in BOM_TYPES:
            self._log(buf, idx, _("Invalid bom_type (normal/phantom/subcontracting)"))

        return {
            "bom_product_name": p_name,
            "bom_code": b_code,
            "bom_qty": b_qty_f,
            "bom_uom": b_uom,
            "bom_type": b_type or "normal",
            "component_name": c_name,
            "component_qty": c_qty_f,
            "component_uom": c_uom,
        }

    # ---------- collect ----------
    def _collect(self, hmap, body, target):
        log, records = [], []
        count_nonempty = 0
        for i, row in enumerate(body, start=2):
            if self._is_empty_row(hmap, row):
                continue
            count_nonempty += 1
            if target == "product.category":
                vals = self._validate_category(hmap, row, i, log)
            elif target in ("product.raw", "product.semi_finished", "product.finished"):
                vals = self._validate_product_generic(hmap, row, i, log)
            elif target == "stock.location":
                vals = self._validate_location(hmap, row, i, log)
            else:  # mrp.bom.*
                vals = self._validate_bom_row(hmap, row, i, log)
            records.append(vals)
        bad_lines = set()
        for l in log:
            if l.startswith("Line "):
                bad_lines.add(l.split(":")[0].split()[-1])
        return records, log, count_nonempty, len(bad_lines)

    # ---------- resolvers / upserts ----------
    def _upsert(self, model, domain, create_vals, update_vals=None):
        rec = self.env[model].search(domain, limit=1)
        if rec:
            rec.write(update_vals or create_vals)
            return rec, "updated"
        return self.env[model].create(create_vals), "created"

    def _resolve_category_vals(self, vals):
        parent_id = False
        if vals.get("parent_name"):
            parent = self._category_by_name(vals["parent_name"], strict=True)
            parent_id = parent.id
        return {"name": vals["name"], "parent_id": parent_id or False}

    def _resolve_product_vals(self, vals, target):
        uom = self._uom_by_name(vals.get("uom_name"), strict=True)
        provided_cat = self._category_by_name(vals.get("category_name"))
        if provided_cat:
            categ = provided_cat
        elif target == "product.semi_finished":
            categ = self._default_semi_finished_category()
        else:
            categ = self.env.ref("product.product_category_all")

        base = {
            "name": vals["name"],
            "uom_id": uom.id,
            "uom_po_id": uom.id,
            "list_price": vals.get("list_price") or 0.0,
            "categ_id": categ.id,
            "is_storable": True,
            "product_type_cm": "raw",
        }

        if target == "product.raw":
            base.update({
                "type": "consu",
                "purchase_ok": True,
                "sale_ok": False,
                "product_type_cm": "raw",
            })
        elif target == "product.semi_finished":
            base.update({
                "type": "consu",
                "purchase_ok": False,
                "sale_ok": False,
                "product_type_cm": "semi",
            })
        else:  # product.finished
            base.update({
                "type": "consu",
                "purchase_ok": False,
                "sale_ok": True,
                "available_in_pos": True,
                "product_type_cm": "finished",
            })

        return base

    def _resolve_location_vals(self, vals):
        parent_id = False
        if vals.get("parent_name"):
            parent = self.env["stock.location"].search([("name", "=", vals["parent_name"])], limit=1)
            if parent:
                parent_id = parent.id
        return {"name": vals["name"], "usage": vals["usage"], "location_id": parent_id, "active": True}

    # ---------- importers ----------
    def _import_categories(self, rows):
        created = updated = 0
        details = []
        for r in rows:
            v = self._resolve_category_vals(r)
            dom = [("name", "=", v["name"])]
            rec, status = self._upsert("product.category", dom, v)
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

            # NEW: image payload (from wizard)
            img_bytes = r.get("_image_bytes")
            img_fmt = (r.get("_image_fmt") or "").lower()
            if img_bytes and img_fmt in ("png", "jpeg", "jpg"):
                v["image_1920"] = base64.b64encode(img_bytes)

            rec = PT.search([("name", "=", v["name"])], limit=1)
            if rec:
                rec.write(v);
                updated += 1
                details.append(f"product.template({rec.id}) updated")
            else:
                rec = PT.create(v);
                created += 1
                details.append(f"product.template({rec.id}) created")
        return created, updated, details

    def _import_locations(self, rows):
        created = updated = 0
        details = []
        for r in rows:
            v = self._resolve_location_vals(r)
            dom = [("name", "=", v["name"])]
            if v.get("location_id"):
                dom.append(("location_id", "=", v["location_id"]))
            rec, status = self._upsert("stock.location", dom, v)
            if status == "created":
                created += 1
            else:
                updated += 1
            details.append(f"stock.location({rec.id}) {status}")
        return created, updated, details

    def _import_boms_semi_finished(self, rows):
        """BOM for fabricated (semi-finished) products. Always type 'phantom'."""
        MrpBom = self.env["mrp.bom"]
        created = updated = 0
        details = []

        grouped = {}
        for r in rows:
            # group by product + code (type forced to 'phantom', so not part of grouping)
            key = (r["bom_product_name"], r.get("bom_code") or "phantom")
            grouped.setdefault(key, {"head": r, "lines": []})
            grouped[key]["lines"].append(r)

        for key, pack in grouped.items():
            head = pack["head"]

            # accept a variant name in Column A, fallback to template name
            prod = self._product_by_name(head["bom_product_name"], strict=False)
            if prod:
                tmpl = prod.product_tmpl_id
            else:
                tmpl = self._product_tmpl_by_name(head["bom_product_name"], strict=True)

            # Enforce fabricated definition: storable and not sold
            if not (tmpl.type == "product" and not tmpl.sale_ok):
                raise ValidationError(_(
                    "Le produit '%s' n'est pas identifié comme 'fabriqué / semi-fini' (type=product et non vendu)."
                ) % tmpl.display_name)

            try:
                qty = float(head.get("bom_qty") or 1.0)
                if qty <= 0:
                    qty = 1.0
            except Exception:
                qty = 1.0
            uom = self._uom_by_name(head.get("bom_uom")) or tmpl.uom_id
            btype = "phantom"  # <--- forced
            code = head.get("bom_code") or ""

            if code:
                bom = MrpBom.search([("code", "=", code)], limit=1)
            else:
                bom = MrpBom.search([("product_tmpl_id", "=", tmpl.id), ("type", "=", btype)], limit=1)

            bom_vals = {
                "product_tmpl_id": tmpl.id,
                "product_qty": qty,
                "product_uom_id": uom.id,
                "code": code or False,
                "type": btype,
            }

            if bom:
                bom.write(bom_vals)
                bom.bom_line_ids.unlink()
                status = "updated"
                updated += 1
            else:
                bom = MrpBom.create(bom_vals)
                status = "created"
                created += 1

            line_vals = []
            for r in pack["lines"]:
                comp = self._product_by_name(r["component_name"], strict=True)  # variant
                comp_uom = self._uom_by_name(r.get("component_uom")) or comp.uom_id
                line_vals.append((0, 0, {
                    "product_id": comp.id,
                    "product_qty": float(r.get("component_qty") or 0.0),
                    "product_uom_id": comp_uom.id,
                }))
            if line_vals:
                bom.write({"bom_line_ids": line_vals})
            details.append(f"mrp.bom({bom.id}) {status} for template {tmpl.display_name}")

        return created, updated, details

    def _import_boms_finished(self, rows):
        """BOMs on finished goods (unchanged)."""
        MrpBom = self.env["mrp.bom"]
        created = updated = 0
        details = []

        grouped = {}
        for r in rows:
            key = (r["bom_product_name"], r.get("bom_code") or r.get("bom_type") or "normal")
            grouped.setdefault(key, {"head": r, "lines": []})
            grouped[key]["lines"].append(r)

        for key, pack in grouped.items():
            head = pack["head"]
            prod = self._product_by_name(head["bom_product_name"], strict=True)
            tmpl = prod.product_tmpl_id

            try:
                qty = float(head.get("bom_qty") or 1.0)
                if qty <= 0:
                    qty = 1.0
            except Exception:
                qty = 1.0
            uom = self._uom_by_name(head.get("bom_uom")) or tmpl.uom_id
            btype = head.get("bom_type") or "normal"
            code = head.get("bom_code") or ""

            if code:
                bom = MrpBom.search([("code", "=", code)], limit=1)
            else:
                bom = MrpBom.search([("product_tmpl_id", "=", tmpl.id), ("type", "=", btype)], limit=1)

            bom_vals = {
                "product_tmpl_id": tmpl.id,
                "product_qty": qty,
                "product_uom_id": uom.id,
                "code": code or False,
                "type": btype,
            }

            if bom:
                bom.write(bom_vals)
                bom.bom_line_ids.unlink()
                updated += 1
                status = "updated"
            else:
                bom = MrpBom.create(bom_vals)
                created += 1
                status = "created"

            line_vals = []
            for r in pack["lines"]:
                comp = self._product_by_name(r["component_name"], strict=True)
                comp_uom = self._uom_by_name(r.get("component_uom")) or comp.uom_id
                line_vals.append((0, 0, {
                    "product_id": comp.id,
                    "product_qty": float(r.get("component_qty") or 0.0),
                    "product_uom_id": comp_uom.id,
                }))
            if line_vals:
                bom.write({"bom_line_ids": line_vals})
            details.append(f"mrp.bom({bom.id}) {status} for template {tmpl.display_name}")

        return created, updated, details

    # ---------- actions ----------
    def action_test(self):
        self.ensure_one()
        hmap, body = self._parse()
        records, log, count, errors = self._collect(hmap, body, self.target_model)
        summary = [
            _("=== TEST RESULT ==="),
            _("Target: %s") % dict(_TARGETS).get(self.target_model),
            _("Rows found: %s") % count,
            _("Rows with issues: %s") % errors,
        ]
        if log:
            summary.append("\n".join(log))
        self.write({
            "log_text": "\n".join(summary),
            "line_count": count,
            "error_count": errors,
            "tested_ok": errors == 0,
        })
        return {
            "type": "ir.actions.act_window",
            "res_model": self._name,
            "res_id": self.id,
            "view_mode": "form",
            "target": "new",
        }

    def action_import(self):
        self.ensure_one()
        hmap, body = self._parse()
        rows, log, count, errors = self._collect(hmap, body, self.target_model)
        if errors:
            raise ValidationError(_("Please fix errors first (run Test). Issues found: %s") % errors)

        if self.target_model == "product.category":
            created, updated, details = self._import_categories(rows)
        elif self.target_model in ("product.raw", "product.semi_finished", "product.finished"):
            created, updated, details = self._import_products(rows)
        elif self.target_model == "mrp.bom.semi_finished":
            created, updated, details = self._import_boms_semi_finished(rows)
        elif self.target_model == "mrp.bom.finished":
            created, updated, details = self._import_boms_finished(rows)
        elif self.target_model == "stock.location":
            created, updated, details = self._import_locations(rows)
        else:
            raise UserError(_("Unsupported target"))

        summary = [
            _("=== IMPORT RESULT ==="),
            _("Target: %s") % dict(_TARGETS).get(self.target_model),
            _("Rows scanned: %s") % count,
            _("Created: %s") % created,
            _("Updated: %s") % updated,
        ]
        self.write({
            "log_text": "\n".join(summary + ["\n"] + details),
            "tested_ok": True,
        })
        return {
            "type": "ir.actions.act_window",
            "res_model": self._name,
            "res_id": self.id,
            "view_mode": "form",
            "target": "new",
        }

    def action_download_template(self):
        """
        Generate an .xlsx template with FRIENDLY headers for each target.

        For mrp.bom.semi_finished:
          - Column A (Produit fabriqué) = dropdown of existing SEMI-FINISHED product *variants*
          - Column F (Nom composant)    = dropdown of existing RAW MATERIAL *variants*
          - Column E (Type) is hidden and prefilled with 'phantom'
        """
        self.ensure_one()
        if not self.target_model:
            raise UserError(_("Select a target first."))

        import io
        output = io.BytesIO()
        try:
            import xlsxwriter
            use_xlsxwriter = True
        except Exception:
            use_xlsxwriter = False

        headers = HEADER_LABELS.get(self.target_model, [])
        if not headers:
            raise UserError(_("No template defined for %s") % self.target_model)

        # ----- Preload lists -----
        uom_names = []
        category_names = []
        parent_cat_names = []
        parent_loc_names = []

        # lists for BOM(semi_finished)
        semi_finished_names = []  # product.product (variants) whose template product_type_cm = 'semi'
        raw_material_names = []   # product.product (variants) whose template product_type_cm = 'raw'

        # UoMs
        if self.target_model in ("product.raw", "product.semi_finished", "product.finished",
                                 "mrp.bom.semi_finished", "mrp.bom.finished"):
            uom_names = self.env["uom.uom"].search([]).mapped("name")
            seen = set()
            uom_names = [n for n in uom_names if not (n in seen or seen.add(n))]

        if self.target_model in ("product.raw", "product.semi_finished", "product.finished"):
            category_names = self.env["product.category"].search([]).mapped("name")
            seen = set()
            category_names = [n for n in category_names if not (n in seen or seen.add(n))]

        if self.target_model == "product.category":
            parent_cat_names = self.env["product.category"].search([]).mapped("name")
            seen = set()
            parent_cat_names = [n for n in parent_cat_names if not (n in seen or seen.add(n))]

        if self.target_model == "stock.location":
            parent_loc_names = self.env["stock.location"].search([]).mapped("name")
            if "WH" not in parent_loc_names:
                parent_loc_names = ["WH"] + parent_loc_names
            else:
                parent_loc_names = ["WH"] + [n for n in parent_loc_names if n != "WH"]
            seen = set()
            parent_loc_names = [n for n in parent_loc_names if not (n in seen or seen.add(n))]

        # Preload product names for mrp.bom.semi_finished (variants)
        if self.target_model == "mrp.bom.semi_finished":
            PP = self.env["product.product"]
            semi_finished_names = PP.search([("product_tmpl_id.product_type_cm", "=", "semi")]).mapped("name")
            raw_material_names = PP.search([("product_tmpl_id.product_type_cm", "=", "raw")]).mapped("name")
            seen = set()
            semi_finished_names = [n for n in semi_finished_names if not (n in seen or seen.add(n))]
            seen = set()
            raw_material_names = [n for n in raw_material_names if not (n in seen or seen.add(n))]

        MAX_ROWS = 10000

        # ========================= xlsxwriter path =========================
        if use_xlsxwriter:
            import xlsxwriter
            wb = xlsxwriter.Workbook(output, {'in_memory': True})
            ws = wb.add_worksheet("Template")
            fmt_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

            # headers + widths
            for col, (_key, label) in enumerate(headers):
                ws.write(0, col, label, fmt_header)
                ws.set_column(col, col, 28)

            # example row
            if self.target_model == "product.category":
                ex = ["Catégorie A", ""]
            elif self.target_model in ("product.raw", "product.semi_finished", "product.finished"):
                u0 = uom_names[0] if uom_names else ""
                c0 = category_names[0] if category_names else ""
                ex = ["Produit A", u0, "199.99", c0]
            elif self.target_model in ("mrp.bom.semi_finished", "mrp.bom.finished"):
                u0 = uom_names[0] if uom_names else ""
                p0 = semi_finished_names[0] if (
                    self.target_model == "mrp.bom.semi_finished" and semi_finished_names) else "Produit Semi-fini A"
                comp0 = raw_material_names[0] if (
                    self.target_model == "mrp.bom.semi_finished" and raw_material_names) else "Matière Première A"
                # A  B        C   D   E         F           G   H
                ex = [p0, "BOM-001", "1", u0, "phantom", comp0, "2", u0]
            else:
                ex = ["Stock", "WH", "internal"]
            for c, v in enumerate(ex):
                ws.write(1, c, v)

            # data validations / lists
            if self.target_model == "product.category":
                lws = wb.add_worksheet("Lists")
                for i, name in enumerate(parent_cat_names, start=1):
                    lws.write(i - 1, 0, name)
                existing_n = len(parent_cat_names)
                for r in range(1, existing_n + 1):
                    lws.write_formula(r - 1, 2, f"=A{r}")
                for j in range(1, MAX_ROWS + 1):
                    lws.write_formula(existing_n + j - 1, 2, f"=Template!$A${j + 1}")
                lws.hide()
                last_row = existing_n + MAX_ROWS
                ws.data_validation(1, 1, MAX_ROWS, 1, {
                    'validate': 'list',
                    'source': f"=Lists!$C$1:$C${last_row}",
                    'ignore_blank': True,
                    'error_type': 'stop',
                    'error_title': 'Parent invalide',
                    'error_message': 'Choisissez une catégorie parente existante ou déjà saisie en colonne A.',
                })

            if self.target_model in ("product.raw", "product.semi_finished", "product.finished"):
                lws = wb.add_worksheet("Lists")
                for i, name in enumerate(uom_names, start=1):
                    lws.write(i - 1, 0, name)
                for i, name in enumerate(category_names, start=1):
                    lws.write(i - 1, 1, name)
                lws.hide()
                u_last = len(uom_names)
                c_last = len(category_names)
                ws.data_validation(1, 1, MAX_ROWS, 1, {
                    'validate': 'list',
                    'source': f"=Lists!$A$1:$A${u_last}",
                    'ignore_blank': False,
                    'error_type': 'stop',
                    'error_title': 'Unité invalide',
                    'error_message': "Choisissez l'unité depuis la liste.",
                })
                ws.data_validation(1, 3, MAX_ROWS, 3, {
                    'validate': 'list',
                    'source': f"=Lists!$B$1:$B${c_last}",
                    'ignore_blank': True,
                    'error_type': 'stop',
                    'error_title': 'Catégorie invalide',
                    'error_message': "Choisissez une catégorie depuis la liste.",
                })

            if self.target_model in ("mrp.bom.semi_finished", "mrp.bom.finished"):
                lws = wb.add_worksheet("Lists")
                for i, name in enumerate(uom_names, start=1):
                    lws.write(i - 1, 0, name)  # Col A -> UoM
                u_last = len(uom_names)

                if self.target_model == "mrp.bom.semi_finished":
                    for i, name in enumerate(semi_finished_names, start=1):
                        lws.write(i - 1, 1, name)  # Col B -> semi-finished variants
                    for i, name in enumerate(raw_material_names, start=1):
                        lws.write(i - 1, 2, name)  # Col C -> raw variants
                lws.hide()

                # UoM columns D(3) and H(7)
                for col_idx in (3, 7):
                    ws.data_validation(1, col_idx, MAX_ROWS, col_idx, {
                        'validate': 'list',
                        'source': f"=Lists!$A$1:$A${u_last}",
                        'ignore_blank': False,
                        'error_type': 'stop',
                        'error_title': 'Unité invalide',
                        'error_message': "Choisissez l'unité depuis la liste.",
                    })

                # Type column handling:
                if self.target_model == "mrp.bom.semi_finished":
                    # Hide column E and prefill with 'phantom'
                    ws.set_column(4, 4, None, None, {'hidden': True})
                    # no validation for type (forced)
                else:
                    # Finished BOMs: keep visible type validation
                    ws.data_validation(1, 4, MAX_ROWS, 4, {
                        'validate': 'list',
                        'source': BOM_TYPE_CHOICES,
                        'ignore_blank': False,
                        'error_type': 'stop',
                        'error_title': 'Type invalide',
                        'error_message': "Choisissez un type (normal/phantom/subcontracting).",
                    })

                # Dropdowns for product (A) and component (F) in semi-finished sheet
                if self.target_model == "mrp.bom.semi_finished":
                    sf_last = len(semi_finished_names)
                    rm_last = len(raw_material_names)
                    if sf_last:
                        ws.data_validation(1, 0, MAX_ROWS, 0, {
                            'validate': 'list',
                            'source': f"=Lists!$B$1:$B${sf_last}",
                            'ignore_blank': False,
                            'error_type': 'stop',
                            'error_title': 'Produit invalide',
                            'error_message': "Choisissez un produit semi-fini depuis la liste.",
                        })
                    if rm_last:
                        ws.data_validation(1, 5, MAX_ROWS, 5, {
                            'validate': 'list',
                            'source': f"=Lists!$C$1:$C${rm_last}",
                            'ignore_blank': False,
                            'error_type': 'stop',
                            'error_title': 'Composant invalide',
                            'error_message': "Choisissez une matière première depuis la liste.",
                        })

            wb.close()

        # ========================= openpyxl fallback =========================
        else:
            from openpyxl import Workbook
            from openpyxl.utils import get_column_letter
            from openpyxl.worksheet.datavalidation import DataValidation

            wb = Workbook()
            ws = wb.active
            ws.title = "Template"

            for c, (_key, label) in enumerate(headers, start=1):
                ws.cell(row=1, column=c, value=label)
                ws.column_dimensions[get_column_letter(c)].width = 28

            if self.target_model == "product.category":
                ex = ["Catégorie A", ""]
            elif self.target_model in ("product.raw", "product.semi_finished", "product.finished"):
                u0 = uom_names[0] if uom_names else ""
                c0 = category_names[0] if category_names else ""
                ex = ["Produit A", u0, "199.99", c0]
            elif self.target_model in ("mrp.bom.semi_finished", "mrp.bom.finished"):
                u0 = uom_names[0] if uom_names else ""
                p0 = semi_finished_names[0] if (
                    self.target_model == "mrp.bom.semi_finished" and semi_finished_names) else "Produit Semi-fini A"
                comp0 = raw_material_names[0] if (
                    self.target_model == "mrp.bom.semi_finished" and raw_material_names) else "Matière Première A"
                ex = [p0, "BOM-001", "1", u0, "phantom", comp0, "2", u0]
            else:
                ex = ["Stock", "WH", "internal"]
            for c, v in enumerate(ex, start=1):
                ws.cell(row=2, column=c, value=v)

            if self.target_model == "product.category":
                lws = wb.create_sheet("Lists")
                for i, name in enumerate(parent_cat_names, start=1):
                    lws.cell(row=i, column=1, value=name)
                existing_n = len(parent_cat_names)
                for r in range(1, existing_n + 1):
                    lws.cell(row=r, column=3).value = f"=A{r}"
                for j in range(1, MAX_ROWS + 1):
                    lws.cell(row=existing_n + j, column=3).value = f"=Template!$A${j + 1}"
                lws.sheet_state = "hidden"
                last_row = existing_n + MAX_ROWS
                dv = DataValidation(type="list", formula1=f"=Lists!$C$1:$C${last_row}", allow_blank=True,
                                    showErrorMessage=True)
                dv.errorTitle = "Parent invalide"
                dv.error = "Choisissez une catégorie parente existante ou déjà saisie en colonne A."
                ws.add_data_validation(dv)
                dv.add(f"B2:B{MAX_ROWS + 1}")

            if self.target_model in ("product.raw", "product.semi_finished", "product.finished"):
                lws = wb.create_sheet("Lists")
                for i, name in enumerate(uom_names, start=1):  # A
                    lws.cell(row=i, column=1, value=name)
                for i, name in enumerate(category_names, start=1):  # B
                    lws.cell(row=i, column=2, value=name)
                lws.sheet_state = "hidden"
                u_last = len(uom_names)
                c_last = len(category_names)
                dv_u = DataValidation(type="list", formula1=f"=Lists!$A$1:$A${u_last}", allow_blank=False,
                                      showErrorMessage=True)
                dv_u.errorTitle = "Unité invalide"
                dv_u.error = "Choisissez l'unité depuis la liste."
                ws.add_data_validation(dv_u)
                dv_u.add(f"B2:B{MAX_ROWS + 1}")

                dv_c = DataValidation(type="list", formula1=f"=Lists!$B$1:$B${c_last}", allow_blank=True,
                                      showErrorMessage=True)
                dv_c.errorTitle = "Catégorie invalide"
                dv_c.error = "Choisissez une catégorie depuis la liste."
                ws.add_data_validation(dv_c)
                dv_c.add(f"D2:D{MAX_ROWS + 1}")

            if self.target_model in ("mrp.bom.semi_finished", "mrp.bom.finished"):
                lws = wb.create_sheet("Lists")
                for i, name in enumerate(uom_names, start=1):
                    lws.cell(row=i, column=1, value=name)  # A
                u_last = len(uom_names)

                if self.target_model == "mrp.bom.semi_finished":
                    for i, name in enumerate(semi_finished_names, start=1):
                        lws.cell(row=i, column=2, value=name)  # B
                    for i, name in enumerate(raw_material_names, start=1):
                        lws.cell(row=i, column=3, value=name)  # C
                lws.sheet_state = "hidden"

                # UoM validations D and H
                for rng in (f"D2:D{MAX_ROWS + 1}", f"H2:H{MAX_ROWS + 1}"):
                    dv_u = DataValidation(type="list", formula1=f"=Lists!$A$1:$A${u_last}", allow_blank=False,
                                          showErrorMessage=True)
                    dv_u.errorTitle = "Unité invalide"
                    dv_u.error = "Choisissez l'unité depuis la liste."
                    ws.add_data_validation(dv_u)
                    dv_u.add(rng)

                if self.target_model == "mrp.bom.semi_finished":
                    # Hide E and set 'phantom'
                    ws.column_dimensions['E'].hidden = True
                    # no validation for E
                else:
                    dv_t = DataValidation(type="list",
                                          formula1=f'"{",".join(BOM_TYPE_CHOICES)}"',
                                          allow_blank=False, showErrorMessage=True)
                    dv_t.errorTitle = "Type invalide"
                    dv_t.error = "Choisissez un type (normal/phantom/subcontracting)."
                    ws.add_data_validation(dv_t)
                    dv_t.add(f"E2:E{MAX_ROWS + 1}")

                if self.target_model == "mrp.bom.semi_finished":
                    sf_last = len(semi_finished_names)
                    rm_last = len(raw_material_names)
                    if sf_last:
                        dv_p = DataValidation(type="list", formula1=f"=Lists!$B$1:$B${sf_last}",
                                              allow_blank=False, showErrorMessage=True)
                        dv_p.errorTitle = "Produit invalide"
                        dv_p.error = "Choisissez un produit semi-fini depuis la liste."
                        ws.add_data_validation(dv_p)
                        dv_p.add(f"A2:A{MAX_ROWS + 1}")
                    if rm_last:
                        dv_c = DataValidation(type="list", formula1=f"=Lists!$C$1:$C${rm_last}",
                                              allow_blank=False, showErrorMessage=True)
                        dv_c.errorTitle = "Composant invalide"
                        dv_c.error = "Choisissez une matière première depuis la liste."
                        ws.add_data_validation(dv_c)
                        dv_c.add(f"F2:F{MAX_ROWS + 1}")

            wb.save(output)

        content = output.getvalue()
        output.close()

        fname = {
            "product.category": "Modele_Categories.xlsx",
            "product.raw": "Modele_Produits_MP.xlsx",
            "product.semi_finished": "Modele_Produits_SemiFini.xlsx",
            "product.finished": "Modele_Produits_Finis.xlsx",
            "mrp.bom.semi_finished": "Modele_BOM_Produits_Fabriques.xlsx",
            "mrp.bom.finished": "Modele_BOM_Produits_Finis.xlsx",
            "stock.location": "Modele_Emplacements.xlsx",
        }.get(self.target_model, "Modele_Import.xlsx")

        att = self.env["ir.attachment"].create({
            "name": fname,
            "type": "binary",
            "datas": base64.b64encode(content),
            "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "res_model": self._name,
            "res_id": self.id,
        })
        url = f"/web/content/{att.id}?download=true"
        return {"type": "ir.actions.act_url", "url": url, "target": "self"}
