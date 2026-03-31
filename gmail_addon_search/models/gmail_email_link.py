from odoo import api, fields, models


class GmailEmailLink(models.Model):
    _name = 'gmail.email.link'
    _description = 'Gmail Email → Odoo Record Link'
    _rec_name = 'gmail_message_id'

    rfc_message_id = fields.Char(index=True)         # RFC 5322 Message-ID (legacy)
    gmail_message_id = fields.Char(index=True)       # Gmail internal message ID
    gmail_thread_id = fields.Char(index=True)        # Gmail thread ID
    res_model = fields.Char(required=True, index=True)  # 'project.task' or 'helpdesk.ticket'
    res_id = fields.Integer(required=True, index=True)
    record_name = fields.Char()  # cached display name

    @api.model
    def backfill_from_mail_messages(self, limit=2000):
        """
        Backfill gmail.email.link from historical mail.message records.
        Useful after deploying the add-on on existing databases.
        """
        Model = self.sudo()
        msg_rows = self.env['mail.message'].sudo().search_read(
            [('message_id', '!=', False), ('model', 'in', ['project.task', 'helpdesk.ticket']), ('res_id', '!=', False)],
            fields=['model', 'res_id', 'message_id', 'record_name'],
            order='id desc',
            limit=int(limit or 2000),
        )
        valid_rows = [
            row for row in msg_rows
            if row.get('model') and row.get('res_id') and row.get('message_id')
        ]

        candidate_msgids = [row['message_id'] for row in valid_rows]
        existing_rows = Model.search_read(
            [('rfc_message_id', 'in', candidate_msgids)],
            fields=['rfc_message_id', 'res_model', 'res_id'],
        )
        existing_set = {
            (r['rfc_message_id'], r['res_model'], r['res_id'])
            for r in existing_rows
        }

        new_vals = []
        for row in valid_rows:
            model = row['model']
            res_id = int(row['res_id'])
            msgid = row['message_id']
            if (msgid, model, res_id) not in existing_set:
                new_vals.append({
                    'rfc_message_id': msgid,
                    'gmail_message_id': '',
                    'gmail_thread_id': '',
                    'res_model': model,
                    'res_id': res_id,
                    'record_name': row.get('record_name') or '',
                })

        if new_vals:
            Model.create(new_vals)

        created = len(new_vals)
        return {'processed': len(msg_rows), 'created': created}
