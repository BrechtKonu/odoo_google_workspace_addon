{
    'name': 'Google Workspace Add-on',
    'version': '19.0.1.4.0',
    'category': 'Productivity/Mail Plugins',
    'summary': 'Google Workspace add-on companion: search tasks/tickets, create records, log emails',
    'description': """
Google Workspace Add-on companion module.

Provides API endpoints for the Google Workspace Add-on to:
- Search project tasks and helpdesk tickets
- Create tasks and tickets from email context
- Log emails to record chatter
- Autocomplete partners, projects, stages and teams

Authentication uses the mail_plugin Bearer token (Odoo API key) flow.
Helpdesk support is optional and gracefully degraded when not installed.
    """,
    'depends': ['mail_plugin', 'project'],
    'data': [
        'security/ir.model.access.csv',
        'data/ir_cron.xml',
        'data/server_actions.xml',
        'views/project_task_type_views.xml',
        'views/helpdesk_stage_views.xml',
    ],
    'installable': True,
    'auto_install': False,
    'author': 'Konu',
    'license': 'LGPL-3',
}
