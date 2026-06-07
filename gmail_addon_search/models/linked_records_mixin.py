from odoo import _, fields, models


class GmailLinkedRecordsMixin(models.AbstractModel):
    """Expose linked Gmail/Outlook emails and Google Docs/Sheets on a record.

    Injected into project.task, helpdesk.ticket and crm.lead so the add-on's
    one-directional ``email/doc -> record`` links become visible the other way
    round, inside Odoo, via smart buttons on the form view.
    """
    _name = 'gmail.linked.records.mixin'
    _description = 'Linked Gmail Emails / Google Documents Mixin'

    gmail_email_link_count = fields.Integer(
        string='Linked Emails', compute='_compute_gmail_link_counts')
    gmail_document_link_count = fields.Integer(
        string='Linked Documents', compute='_compute_gmail_link_counts')

    def _compute_gmail_link_counts(self):
        email_counts = {}
        doc_counts = {}
        if self.ids:
            for res_id, count in self.env['gmail.email.link']._read_group(
                [('res_model', '=', self._name), ('res_id', 'in', self.ids)],
                groupby=['res_id'], aggregates=['__count'],
            ):
                email_counts[res_id] = count
            for res_id, count in self.env['gmail.document.link']._read_group(
                [('res_model', '=', self._name), ('res_id', 'in', self.ids)],
                groupby=['res_id'], aggregates=['__count'],
            ):
                doc_counts[res_id] = count
        for record in self:
            record.gmail_email_link_count = email_counts.get(record.id, 0)
            record.gmail_document_link_count = doc_counts.get(record.id, 0)

    def _gmail_link_action(self, link_model, name):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': name,
            'res_model': link_model,
            'view_mode': 'list,form',
            'domain': [('res_model', '=', self._name), ('res_id', '=', self.id)],
            'context': {
                'default_res_model': self._name,
                'default_res_id': self.id,
                'create': False,
            },
        }

    def action_view_gmail_email_links(self):
        return self._gmail_link_action('gmail.email.link', _('Linked Emails'))

    def action_view_gmail_document_links(self):
        return self._gmail_link_action('gmail.document.link', _('Linked Documents'))
