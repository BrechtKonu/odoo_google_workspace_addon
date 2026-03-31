from odoo import fields, models


class GmailDocumentLink(models.Model):
    _name = 'gmail.document.link'
    _description = 'Google Document/Sheet -> Odoo Record Link'
    _rec_name = 'record_name'

    document_id = fields.Char(required=True, index=True)
    host_app = fields.Selection(
        selection=[('docs', 'Google Docs'), ('sheets', 'Google Sheets')],
        required=True,
        index=True,
    )
    res_model = fields.Char(required=True, index=True)  # 'project.task' or 'helpdesk.ticket'
    res_id = fields.Integer(required=True, index=True)
    record_name = fields.Char()

    _sql_constraints = [
        (
            'gmail_document_link_unique',
            'unique(document_id, host_app, res_model, res_id)',
            'This document is already linked to this record.',
        ),
    ]
