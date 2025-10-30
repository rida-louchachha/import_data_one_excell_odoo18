from odoo import models, fields, api
import uuid

class CmMasterRepo(models.Model):
    _name = "cm.master.repo"
    _description = "Excel Image Repository"

    name = fields.Char(required=True, default=lambda s: "Excel Repo")
    token = fields.Char(required=True, copy=False, index=True, default=lambda s: uuid.uuid4().hex)
    # optional: who created it
    user_id = fields.Many2one("res.users", default=lambda s: s.env.user)

    def content_url(self, attachment):
        # canonical Odoo URL for an attachment
        return f"/web/content/{attachment.id}"