# models/cm_import_center.py
from odoo import fields, models

class CmImportCenter(models.TransientModel):
    _name = "cm.import.center"
    _description = "CM Import Center"

    note = fields.Text(
        default=(
            "1) Import Categories\n"
            "2) Import Raw Materials (products)\n"
            "3) Import Semi-finished / Fabricated Products\n"
            "4) Import Finished Products\n"
            "5) Import BOMs for Fabricated Products\n"
        ),
        readonly=True,
    )

    def _open_action(self, xmlid, ctx=None):
        action = self.env.ref(xmlid).read()[0]
        if ctx:
            action_ctx = action.get("context") or {}
            action["context"] = {**action_ctx, **ctx}
        return action

    def action_open_categories(self):
        return self._open_action(
            "import_wizards_caisse_manager_18.action_cm_import_wizard_categories",
            {"default_target_model": "product.category"},
        )

    def action_open_products_raw(self):
        return self._open_action(
            "import_wizards_caisse_manager_18.action_cm_import_wizard_product_raw",
            {"default_target_model": "product.raw"},
        )

    # NEW: semi-finished (fabricated) products
    def action_open_products_semi_finished(self):
        return self._open_action(
            "import_wizards_caisse_manager_18.action_cm_import_wizard_product_semi_finished",
            {"default_target_model": "product.semi_finished"},
        )

    # Keep: finished products
    def action_open_products_finished(self):
        return self._open_action(
            "import_wizards_caisse_manager_18.action_cm_import_wizard_product_finished",
            {"default_target_model": "product.finished"},
        )

    def action_open_locations(self):
        return self._open_action(
            "import_wizards_caisse_manager_18.action_cm_import_wizard_location",
            {"default_target_model": "stock.location"},
        )

    # BOMs tied only to fabricated (semi-finished) products
    def action_open_boms_finished(self):
        return self._open_action(
            "import_wizards_caisse_manager_18.action_cm_import_wizard_bom_finished",
            {"default_target_model": "mrp.bom.finished"},
        )
