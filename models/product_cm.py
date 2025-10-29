from odoo import api, fields, models, _

class ProductTemplate(models.Model):
    _inherit = "product.template"

    product_type_cm = fields.Selection(
        selection=[
            ("raw", _("Matière première")),
            ("semi", _("Semi-fini (fabriqué)")),
            ("finished", _("Produit fini")),
        ],
        string="CM Product Type",
        default="raw",
        help="Type métier utilisé par l’import: matière première / semi-fini / fini.",
    )

class ProductProduct(models.Model):
    _inherit = "product.product"

    product_type_cm = fields.Selection(
        related="product_tmpl_id.product_type_cm",
        store=True,
        readonly=False,
        string="CM Product Type",
    )
