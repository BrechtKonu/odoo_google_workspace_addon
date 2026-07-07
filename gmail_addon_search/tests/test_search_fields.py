from odoo.tests import tagged
from odoo.addons.mail_plugin.tests.common import (
    TestMailPluginControllerCommon,
    mock_auth_method_outlook,
)


@tagged('post_install', '-at_install')
class TestSearchKanbanFields(TestMailPluginControllerCommon):
    """/task and /ticket search must expose the fields the Google Chat cards
    render: priority, deadline, tags (plus kanban_state for tickets)."""

    @mock_auth_method_outlook('admin')
    def test_task_search_exposes_kanban_fields(self):
        project = self.env['project.project'].create({'name': 'GWS Search Project'})
        tag = self.env['project.tags'].create({'name': 'search-tag'})
        task = self.env['project.task'].create({
            'name': 'GWS Search Task',
            'project_id': project.id,
            'priority': '2',
            'date_deadline': '2030-01-15 00:00:00',
            'tag_ids': [(6, 0, tag.ids)],
        })

        result = self.make_jsonrpc_request(
            '/gmail_addon/task/search',
            {'search_term': 'GWS Search Task', 'limit': 5},
        )

        rows = [t for t in result.get('tasks', []) if t['id'] == task.id]
        self.assertEqual(len(rows), 1, msg=result)
        row = rows[0]
        self.assertEqual(row['priority'], 'High')
        self.assertEqual(row['priority_level'], '2')
        self.assertEqual(row['tag_names'], ['search-tag'])
        self.assertTrue(row['deadline'].startswith('2030-01-15'), msg=row['deadline'])

    @mock_auth_method_outlook('admin')
    def test_ticket_search_exposes_kanban_fields(self):
        if 'helpdesk.ticket' not in self.env:
            self.skipTest('helpdesk not installed')
        team = self.env['helpdesk.team'].create({'name': 'GWS Search Team'})
        tag = self.env['helpdesk.tag'].create({'name': 'ticket-tag'})
        ticket = self.env['helpdesk.ticket'].create({
            'name': 'GWS Search Ticket',
            'team_id': team.id,
            'priority': '3',
            'kanban_state': 'blocked',
            'tag_ids': [(6, 0, tag.ids)],
        })

        result = self.make_jsonrpc_request(
            '/gmail_addon/ticket/search',
            {'search_term': 'GWS Search Ticket', 'limit': 5},
        )

        rows = [t for t in result.get('tickets', []) if t['id'] == ticket.id]
        self.assertEqual(len(rows), 1, msg=result)
        row = rows[0]
        self.assertEqual(row['priority'], 'Urgent')
        self.assertEqual(row['kanban_state'], 'blocked')
        self.assertEqual(row['kanban_state_label'], 'Blocked')
        self.assertEqual(row['tag_names'], ['ticket-tag'])
