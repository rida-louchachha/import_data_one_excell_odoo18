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

# 3-state values for Excel
BARCODE_MODES = ["Auto", "Manuel", "Aucun"]

HEADER_LABELS = {
    "product.category": [
        ("name", _("Nom de la catégorie")),
        ("parent_name", _("Catégorie parente")),
    ],

    # MP (raw)
    "product.raw": [
        ("name",        _("Nom du produit")),
        ("uom_name",    _("Unité")),
        ("list_price",  _("Prix")),
        ("category_name", _("Catégorie")),
        ("image_embed", _("Image (insérer PNG/JPG)")),
        ("barcode_mode", _("Code-barres (Auto/Manuel/Aucun)")),
        ("barcode_manual", _("Code-barres (si Manuel)")),
    ],

    # Semi-finished (no barcode here)
    "product.semi_finished": [
        ("name",        _("Nom du produit")),
        ("uom_name",    _("Unité")),
        ("list_price",  _("Prix")),
        ("category_name", _("Catégorie")),
        ("image_embed", _("Image (insérer PNG/JPG)")),
    ],

    "mrp.bom.semi_finished": [
        ("bom_product_name", _("Produit fabriqué (nom)")),
        ("bom_code",         _("Code nomenclature")),
        ("bom_qty",          _("Quantité nomenclature (sortie)")),
        ("bom_uom",          _("Unité nomenclature")),
        ("bom_type",         _("Type de nomenclature (normal/phantom/subcontracting)")),
        ("component_name",   _("Nom composant")),
        ("component_qty",    _("Quantité composant")),
        ("component_uom",    _("Unité composant")),
    ],
    "product.finished": [
        ("name",        _("Nom du produit")),
        ("uom_name",    _("Unité")),
        ("list_price",  _("Prix")),
        ("category_name", _("Catégorie")),
        ("image_embed", _("Image (insérer PNG/JPG)")),
    ],

    # Finished BoM (barcode mode on the header row of the merged block)
    "mrp.bom.finished": [
        ("bom_product_name", _("Produit fini (nom)")),
        ("barcode_mode",     _("Code-barres (Auto/Manuel/Aucun)")),
        ("barcode_manual",   _("Code-barres (si Manuel)")),
        ("bom_code",         _("Code nomenclature")),
        ("bom_qty",          _("Quantité nomenclature (sortie)")),
        ("bom_uom",          _("Unité nomenclature")),
        ("bom_type",         _("Type de nomenclature (normal/phantom/subcontracting)")),
        ("component_name",   _("Nom composant")),
        ("component_qty",    _("Quantité composant")),
        ("component_uom",    _("Unité composant")),
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
        "barcode_mode": [
            "Code-barres (Auto/Manuel/Aucun)", "Code barres (Auto/Manuel/Aucun)",
            "Barcode Mode", "Barcode (Auto/Manual/None)"
        ],
        "barcode_manual": [
            "Code-barres (si Manuel)", "Code barres (si Manuel)",
            "Barcode Manual", "Manual Barcode"
        ],
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
    "product.finished": {
        "name": ["Nom du produit", "Nom produit fini"],
        "uom_name": ["Unité", "UoM", "Unité de mesure"],
        "list_price": ["Prix", "Prix de vente"],
        "category_name": ["Catégorie", "Catégorie produit"],
        "image_embed": ["Image (insérer PNG/JPG)", "Image", "Image URL", "Image Path"],
    },
    "mrp.bom.finished": {
        "bom_product_name": ["Produit fini (nom)", "Nom produit fini"],
        "barcode_mode": [
            "Code-barres (Auto/Manuel/Aucun)", "Code barres (Auto/Manuel/Aucun)",
            "Barcode Mode", "Barcode (Auto/Manual/None)"
        ],
        "barcode_manual": [
            "Code-barres (si Manuel)", "Code barres (si Manuel)",
            "Barcode Manual", "Manual Barcode"
        ],
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
    "product.finished": "Produits – Finis",
    "mrp.bom.finished": "Nomenclatures – Finis",
    "stock.location": "Emplacements",
}

SHEET_IMPORT_ORDER = [
    "product.category",
    "product.raw",
    "product.semi_finished",
    "mrp.bom.semi_finished",
    "product.finished",
    "mrp.bom.finished",
    "stock.location",
]
