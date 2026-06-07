from . import gmail_email_link
from . import gmail_document_link
from . import gmail_addon_settings
from . import linked_records_mixin
from . import project_task
from . import project_task_type
from . import keep_warm
try:
    import odoo.addons.helpdesk  # noqa
    from . import helpdesk_stage
    from . import helpdesk_ticket
except ImportError:
    pass
try:
    import odoo.addons.crm  # noqa
    from . import crm_lead
except ImportError:
    pass
