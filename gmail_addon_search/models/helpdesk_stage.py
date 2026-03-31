from odoo import fields, models


class HelpdeskStage(models.Model):
    _inherit = 'helpdesk.stage'

    gmail_hide_in_search = fields.Boolean(
        string='Hide in Gmail Add-on',
        default=False,
        help='When enabled, tickets in this stage will not appear in Gmail add-on search results.',
    )
