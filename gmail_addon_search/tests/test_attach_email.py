import base64

from odoo.tests.common import TransactionCase
from odoo.tools import html_sanitize

from odoo.addons.gmail_addon_search.controllers.main import GmailAddonController


# Valid 1x1 transparent PNG (Odoo runs image_fix_orientation on image
# attachments at create time, so the bytes must be a complete image).
_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAen63NgAAAAASUVORK5CYII='


class TestAttachEmail(TransactionCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.project = cls.env['project.project'].create({'name': 'GWS Attach Project'})
        cls.task = cls.env['project.task'].create({
            'name': 'GWS Attach Task', 'project_id': cls.project.id})
        cls.ctrl = GmailAddonController()

    def test_inline_image_becomes_attachment_and_rewrites_body(self):
        body = '<p>hi</p><img src="cid:konu-img-0">'
        attachment_ids, new_body = self.ctrl._attach_email_files(
            self.task,
            [{'name': 'shot.png', 'mimetype': 'image/png', 'data': _PNG, 'cid': 'konu-img-0'}],
            body,
        )
        self.assertEqual(len(attachment_ids), 1)
        att = self.env['ir.attachment'].browse(attachment_ids[0])
        # scoped to the record so it shows in the record's attachment list
        self.assertEqual(att.res_model, 'project.task')
        self.assertEqual(att.res_id, self.task.id)
        # cid sentinel rewritten to a /web/image access-token URL
        self.assertNotIn('cid:konu-img-0', new_body)
        self.assertIn('/web/image/%s?access_token=' % att.id, new_body)

    def test_file_attachment_no_cid_not_rewritten(self):
        body = '<p>see file</p>'
        attachment_ids, new_body = self.ctrl._attach_email_files(
            self.task,
            [{'name': 'doc.pdf', 'mimetype': 'application/pdf', 'data': _PNG, 'cid': None}],
            body,
        )
        self.assertEqual(len(attachment_ids), 1)
        self.assertEqual(new_body, body)  # no cid -> body untouched
        att = self.env['ir.attachment'].browse(attachment_ids[0])
        self.assertEqual(att.res_id, self.task.id)

    def test_empty_attachments_noop(self):
        body = '<p>x</p>'
        attachment_ids, new_body = self.ctrl._attach_email_files(self.task, None, body)
        self.assertEqual(attachment_ids, [])
        self.assertEqual(new_body, body)

    def test_malformed_item_skipped_not_raised(self):
        body = '<p>x</p>'
        attachment_ids, new_body = self.ctrl._attach_email_files(
            self.task, [{'name': 'empty', 'data': ''}], body)
        self.assertEqual(attachment_ids, [])  # blank data skipped

    def test_size_guard_stops_oversize_payload(self):
        big = base64.b64encode(b'a' * (self.ctrl._MAX_ATTACH_TOTAL_BYTES + 1024)).decode()
        attachment_ids, _new = self.ctrl._attach_email_files(
            self.task,
            [{'name': 'big1.bin', 'data': big}, {'name': 'big2.bin', 'data': big}],
            '<p>x</p>',
        )
        # first item alone exceeds the guard -> nothing attached
        self.assertEqual(attachment_ids, [])

    def test_sanitizer_preserves_web_image_src(self):
        # Quick win: confirm html_sanitize keeps a /web/image inline src so the
        # rehosted image actually renders in chatter.
        html = '<p>hi</p><img src="/web/image/42?access_token=abc">'
        cleaned = html_sanitize(html)
        self.assertIn('/web/image/42?access_token=abc', cleaned)
