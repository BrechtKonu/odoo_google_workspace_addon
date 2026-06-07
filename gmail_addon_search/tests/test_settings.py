from odoo.tests.common import TransactionCase


class TestGmailSettingsValues(TransactionCase):
    """`res.config.settings.get_values()` must never hand the Settings form an
    integer 0 for an unset Many2one. Odoo 19's web_read cleanup() reads a falsy
    id as a NewId and does `(0).origin`, raising
    "'int' object has no attribute 'origin'" and crashing the whole Settings
    page on a fresh install (no reference field configured)."""

    def _icp(self):
        return self.env['ir.config_parameter'].sudo()

    def test_unset_reference_field_is_false_not_zero(self):
        for rtype in ('task', 'ticket', 'lead'):
            self._icp().set_param(f'gmail_addon_search.{rtype}.reference_field_id', '')
        vals = self.env['res.config.settings'].get_values()
        for fname in ('gmail_task_reference_field_id',
                      'gmail_ticket_reference_field_id',
                      'gmail_lead_reference_field_id'):
            self.assertIs(vals[fname], False,
                          "%s must be False (not 0) when unset" % fname)

    def test_stale_reference_field_is_false(self):
        # A param pointing at a deleted ir.model.fields id must degrade to False,
        # never a phantom browse(<id>) the form can't render.
        self._icp().set_param('gmail_addon_search.task.reference_field_id', '999999999')
        vals = self.env['res.config.settings'].get_values()
        self.assertIs(vals['gmail_task_reference_field_id'], False)

    def test_valid_reference_field_round_trips(self):
        field = self.env['ir.model.fields'].search(
            [('model', '=', 'project.task'), ('ttype', '=', 'char')], limit=1)
        self.assertTrue(field, "expected a char field on project.task")
        self._icp().set_param('gmail_addon_search.task.reference_field_id', str(field.id))
        vals = self.env['res.config.settings'].get_values()
        self.assertEqual(vals['gmail_task_reference_field_id'], field.id)
