# -*- coding: utf-8 -*-
{
    "name": "Import Wizards Caisse Manager",
    "summary": "Two-step Excel import (Test → Import) for Categories ,Products, Locations, and Bills of Materials",
    "version": "18.0.2.0.0",
    "author": "Rida Louchachha",
    "category": "Manufacturing/Inventory",
    "license": "LGPL-3",
    "depends": [
        "product",
        "stock",
        "mrp"
    ],
    "external_dependencies": {
        "python": ["openpyxl"]
    },
    "data": [
        "data/product_category_data.xml",
        "data/barcode_sequences.xml",
        "security/ir.model.access.csv",
        "views/import_wizard_views.xml",
        "views/menus_actions.xml",
        "views/import_center_views.xml",
    ],
    'installable': True,
    'auto_install': False,
    'application': False,
}