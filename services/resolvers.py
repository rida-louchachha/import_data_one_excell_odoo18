# -*- coding: utf-8 -*-
from odoo import _, exceptions
from odoo.exceptions import ValidationError

from .header_utils import norm


# --- lookups ---
def uom_by_name(env, name, strict=False):
    if not name:
        return False
    Uom = env["uom.uom"]
    u = Uom.search([("name", "=", name)], limit=1) or Uom.search([("name", "ilike", name)], limit=1)
    if strict and not u:
        raise ValidationError(_("Unit of Measure not found: %s") % (name or ""))
    return u


def category_by_name(env, name, strict=False):
    if not name:
        return False
    Cat = env["product.category"]
    c = Cat.search([("name", "=", name)], limit=1) or Cat.search([("name", "ilike", name)], limit=1)
    if strict and not c:
        raise ValidationError(_("Category not found: %s") % (name or ""))
    return c


def product_by_name(env, name, strict=False):
    Product = env["product.product"]
    p = Product.search([("name", "=", name)], limit=1) or Product.search([("name", "ilike", name)], limit=1)
    if strict and not p:
        raise ValidationError(_("Product not found: %s") % (name or ""))
    return p


def product_tmpl_by_name(env, name, strict=False):
    T = env["product.template"]
    t = T.search([("name", "=", name)], limit=1) or T.search([("name", "ilike", name)], limit=1)
    if strict and not t:
        raise ValidationError(_("Product template not found: %s") % (name or ""))
    return t


# --- upserts/helpers ---
def upsert(env, model, domain, create_vals, update_vals=None):
    rec = env[model].search(domain, limit=1)
    if rec:
        rec.write(update_vals or create_vals)
        return rec, "updated"
    return env[model].create(create_vals), "created"


def get_or_create_category_by_name(env, name, parent_name=None):
    Cat = env["product.category"]
    cat = Cat.search([("name", "=", name)], limit=1)
    if cat:
        return cat
    vals = {"name": name}
    if parent_name:
        parent = Cat.search([("name", "=", parent_name)], limit=1)
        if parent:
            vals["parent_id"] = parent.id
    return Cat.create(vals)


def default_semi_finished_category(env):
    cat = env.ref("import_wizards_caisse_manager_18.category_semi_finished_cm", raise_if_not_found=False)
    if cat:
        return cat
    return get_or_create_category_by_name(env, "Produits semi-finis")


def resolve_category_vals(env, vals):
    parent_id = False
    if vals.get("parent_name"):
        parent = category_by_name(env, vals["parent_name"], strict=True)
        parent_id = parent.id
    return {"name": vals["name"], "parent_id": parent_id or False}


def resolve_product_vals(env, vals, target):
    uom = uom_by_name(env, vals.get("uom_name"), strict=True)
    provided_cat = category_by_name(env, vals.get("category_name"))
    if provided_cat:
        categ = provided_cat
    elif target == "product.semi_finished":
        categ = default_semi_finished_category(env)
    else:
        categ = env.ref("product.product_category_all")
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
    else:  # finished
        base.update({
            "type": "consu",
            "purchase_ok": False,
            "sale_ok": True,
            "available_in_pos": True,
            "is_storable": True,
            "product_type_cm": "finished",
        })
    return base


def resolve_location_vals(env, vals):
    parent_id = False
    if vals.get("parent_name"):
        parent = env["stock.location"].search([("name", "=", vals["parent_name"])], limit=1)
        if parent:
            parent_id = parent.id
    return {"name": vals["name"], "usage": vals["usage"], "location_id": parent_id, "active": True}
