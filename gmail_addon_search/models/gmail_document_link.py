from odoo import fields, models


class GmailDocumentLink(models.Model):
    _name = 'gmail.document.link'
    _description = 'Google Document/Sheet -> Odoo Record Link'
    _rec_name = 'record_name'

    document_id = fields.Char(required=True, index=True)
    document_title = fields.Char(help='Cached title of the linked Google Doc/Sheet.')
    document_url = fields.Char(help='Direct URL back to the linked Google Doc/Sheet.')
    host_app = fields.Selection(
        selection=[('docs', 'Google Docs'), ('sheets', 'Google Sheets')],
        required=True,
        index=True,
    )
    res_model = fields.Char(required=True, index=True)  # project.task / helpdesk.ticket / crm.lead
    res_id = fields.Integer(required=True, index=True)
    record_name = fields.Char()
    company_id = fields.Many2one(
        'res.company', index=True, default=lambda self: self.env.company,
        help='Company of the linked record; scopes multi-company visibility.')

    # Odoo 19 dropped the _sql_constraints tuple API in favour of declarative
    # models.Constraint; the old form was silently ignored (no DB constraint).
    _document_link_unique = models.Constraint(
        'unique(document_id, host_app, res_model, res_id)',
        'This document is already linked to this record.',
    )
