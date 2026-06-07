{
    'name': 'Google Workspace Add-on',
    'version': '19.0.1.7.2',
    'category': 'Productivity/Mail Plugins',
    'summary': 'Google Workspace add-on companion: search tasks/tickets, create records, log emails, cross-link docs',
    'description': """
Google Workspace Add-on companion module.

Provides API endpoints for the Google Workspace Add-on to:
- Search project tasks and helpdesk tickets
- Create tasks and tickets from email context
- Log emails to record chatter
- Autocomplete partners, projects, stages and teams
- Cross-link emails and Google Docs/Sheets to tasks/tickets/leads,
  with bi-directional visibility via smart buttons in Odoo

Authentication uses the mail_plugin Bearer token (Odoo API key) flow.
    """,
    'depends': ['mail_plugin', 'project', 'helpdesk', 'crm'],
    'data': [
        'security/ir.model.access.csv',
        'security/gmail_link_rules.xml',
        'data/ir_cron.xml',
        'data/server_actions.xml',
        'views/res_config_settings_views.xml',
        'views/project_task_type_views.xml',
        'views/helpdesk_stage_views.xml',
        'views/helpdesk_ticket_views.xml',
        'views/project_task_views.xml',
        'views/gmail_link_views.xml',
    ],
    'installable': True,
    'auto_install': False,
    'author': 'Konu',
    'license': 'LGPL-3',
}
