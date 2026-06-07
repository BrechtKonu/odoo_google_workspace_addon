from psycopg2 import IntegrityError

from odoo.tests.common import TransactionCase
from odoo.tools import mute_logger


class TestCrossLink(TransactionCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.project = cls.env['project.project'].create({'name': 'GWS Test Project'})
        cls.task = cls.env['project.task'].create({
            'name': 'GWS Test Task', 'project_id': cls.project.id})

    def _email_link(self, **vals):
        base = {'res_model': 'project.task', 'res_id': self.task.id}
        base.update(vals)
        return self.env['gmail.email.link'].create(base)

    def test_email_link_count(self):
        self.assertEqual(self.task.gmail_email_link_count, 0)
        self._email_link(gmail_message_id='m1', record_name='GWS Test Task')
        self.task.invalidate_recordset(['gmail_email_link_count'])
        self.assertEqual(self.task.gmail_email_link_count, 1)

    def test_document_link_count(self):
        self.env['gmail.document.link'].create({
            'document_id': 'doc-1', 'host_app': 'docs',
            'res_model': 'project.task', 'res_id': self.task.id})
        self.task.invalidate_recordset(['gmail_document_link_count'])
        self.assertEqual(self.task.gmail_document_link_count, 1)

    def test_document_link_unique_constraint(self):
        vals = {'document_id': 'doc-dup', 'host_app': 'docs',
                'res_model': 'project.task', 'res_id': self.task.id}
        self.env['gmail.document.link'].create(vals)
        with self.assertRaises(IntegrityError), mute_logger('odoo.sql_db'):
            with self.cr.savepoint():
                self.env['gmail.document.link'].create(vals)
                self.env.flush_all()  # force the INSERT so the unique constraint fires here

    def test_company_id_default(self):
        link = self._email_link(gmail_message_id='m-company')
        self.assertEqual(link.company_id, self.env.company)

    def test_email_link_display_name(self):
        link = self._email_link(gmail_message_id='m-name', record_name='Readable Name')
        self.assertEqual(link.display_name, 'Readable Name')
        # Falls back to model,id when no cached name.
        bare = self._email_link(gmail_message_id='m-bare', record_name=False)
        self.assertEqual(bare.display_name, 'project.task,%s' % self.task.id)

    def test_action_view_email_links_domain(self):
        action = self.task.action_view_gmail_email_links()
        self.assertEqual(action['res_model'], 'gmail.email.link')
        self.assertIn(('res_model', '=', 'project.task'), action['domain'])
        self.assertIn(('res_id', '=', self.task.id), action['domain'])
