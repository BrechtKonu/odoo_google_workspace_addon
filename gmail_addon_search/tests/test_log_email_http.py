import base64

from odoo.tests import tagged
from odoo.addons.mail_plugin.tests.common import (
    TestMailPluginControllerCommon,
    mock_auth_method_outlook,
)

# Valid 1x1 transparent PNG (Odoo runs image_fix_orientation on image
# attachments at create, so the bytes must be a complete image).
_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAen63NgAAAAASUVORK5CYII='
_TXT = base64.b64encode(b'hello world').decode()


@tagged('post_install', '-at_install')
class TestLogEmailAttachmentsHttp(TestMailPluginControllerCommon):
    """End-to-end cover of /gmail_addon/log_email -> message_post(attachment_ids),
    the wiring the _attach_email_files unit tests can't reach (needs the
    auth='outlook' route + a real message_post)."""

    @mock_auth_method_outlook('admin')
    def test_log_email_rehosts_inline_image_and_file(self):
        project = self.env['project.project'].create({'name': 'GWS HTTP Project'})
        task = self.env['project.task'].create(
            {'name': 'GWS HTTP Task', 'project_id': project.id})

        result = self.make_jsonrpc_request('/gmail_addon/log_email', {
            'res_model': 'project.task',
            'res_id': task.id,
            'email_body': '<p>hello</p><img src="cid:konu-img-0">',
            'email_subject': 'Logged subject',
            'attachments': [
                {'name': 'shot.png', 'mimetype': 'image/png', 'data': _PNG, 'cid': 'konu-img-0'},
                {'name': 'note.txt', 'mimetype': 'text/plain', 'data': _TXT, 'cid': None},
            ],
        })

        self.assertTrue(result.get('success'), msg=result)
        msg = self.env['mail.message'].browse(result['message_id'])

        # Both files linked to the posted message (render in the chatter message).
        self.assertEqual(len(msg.attachment_ids), 2)

        # Both scoped to the record -> show in the record's attachment box.
        record_atts = self.env['ir.attachment'].search(
            [('res_model', '=', 'project.task'), ('res_id', '=', task.id)])
        self.assertEqual(len(record_atts), 2)

        # Inline image rewritten to a /web/image URL; no cid sentinel left.
        self.assertIn('/web/image/', msg.body)
        self.assertNotIn('cid:konu-img-0', msg.body)

    @mock_auth_method_outlook('admin')
    def test_log_email_without_attachments_still_posts(self):
        project = self.env['project.project'].create({'name': 'GWS HTTP Project 2'})
        task = self.env['project.task'].create(
            {'name': 'GWS HTTP Task 2', 'project_id': project.id})

        result = self.make_jsonrpc_request('/gmail_addon/log_email', {
            'res_model': 'project.task',
            'res_id': task.id,
            'email_body': '<p>plain note</p>',
            'email_subject': 'No attachments',
        })

        self.assertTrue(result.get('success'), msg=result)
        msg = self.env['mail.message'].browse(result['message_id'])
        self.assertFalse(msg.attachment_ids)
        self.assertIn('plain note', msg.body)
