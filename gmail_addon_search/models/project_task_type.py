from odoo import fields, models


class ProjectTaskType(models.Model):
    _inherit = 'project.task.type'

    gmail_hide_in_search = fields.Boolean(
        string='Hide in Gmail Add-on',
        default=False,
        help='When enabled, tasks in this stage will not appear in Gmail add-on search results.',
    )