from . import gmail_email_link
from . import gmail_document_link
from . import gmail_addon_settings
from . import project_task_type
from . import keep_warm
try:
    import odoo.addons.helpdesk  # noqa
    from . import helpdesk_stage
except ImportError:
    pass
