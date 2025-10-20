# -*- coding: utf-8 -*-
from odoo import _

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
    "product.category": [
        ("name", _("Nom de la catégorie")),
        ("parent_name", _("Catégorie parente")),
    ],
    "product.raw": [
        ("name", _("Nom du produit")),
        ("uom_name", _("Unité")),
        ("list_price", _("Prix")),
        ("category_name", _("Catégorie")),
        ("image_embed", _("Image (insérer PNG/JPG)")),
    ],
    "product.semi_finished": [
        ("name", _("Nom du produit")),
        ("uom_name", _("Unité")),
        ("list_price", _("Prix")),
        ("category_name", _("Catégorie")),
        ("image_embed", _("Image (insérer PNG/JPG)")),
    ],
    "mrp.bom.semi_finished": [
        ("bom_product_name", _("Produit fabriqué (nom)")),
        ("bom_code", _("Code nomenclature")),
        ("bom_qty", _("Quantité nomenclature (sortie)")),
        ("bom_uom", _("Unité nomenclature")),
        ("bom_type", _("Type de nomenclature (normal/phantom/subcontracting)")),
        ("component_name", _("Nom composant")),
        ("component_qty", _("Quantité composant")),
        ("component_uom", _("Unité composant")),
    ],
    "mrp.bom.finished": [
        ("bom_product_name", _("Produit fini (nom)")),
        ("bom_code", _("Code nomenclature")),
        ("bom_qty", _("Quantité nomenclature (sortie)")),
        ("bom_uom", _("Unité nomenclature")),
        ("bom_type", _("Type de nomenclature (normal/phantom/subcontracting)")),
        ("component_name", _("Nom composant")),
        ("component_qty", _("Quantité composant")),
        ("component_uom", _("Unité composant")),
    ],
    "stock.location": [
        ("name", _("Nom de l’emplacement")),
        ("parent_name", _("Emplacement parent")),
        ("usage", _("Type (internal/view/fournisseur/client/inventory/production/transit)")),
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

SHEET_TITLES = {
    "product.category": "Catégories",
    "product.raw": "Produits – MP",
    "product.semi_finished": "Produits – Semi-finis",
    "mrp.bom.semi_finished": "Nomenclatures – Semi-finis",
    "mrp.bom.finished": "Nomenclatures – Finis",
    "stock.location": "Emplacements",
}

SHEET_IMPORT_ORDER = [
    "product.category",
    "product.raw",
    "product.semi_finished",
    "mrp.bom.semi_finished",
    "mrp.bom.finished",
    "stock.location",
]
