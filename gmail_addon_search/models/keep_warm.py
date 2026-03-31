import logging

from odoo import models

_logger = logging.getLogger(__name__)


class GmailAddonKeepWarm(models.AbstractModel):
    _name = 'gmail.addon.keepwarm'
    _description = 'Gmail Add-on Keep Warm'

    def cron_keep_warm(self):
        """Lightweight read-only cron to keep workers warm."""
        env = self.env
        env['ir.config_parameter'].sudo().get_param('web.base.url')
        env['project.project'].search([], limit=1).read(['id', 'name'])
        if 'helpdesk.team' in env:
            env['helpdesk.team'].search([], limit=1).read(['id', 'name'])
        _logger.debug('gmail.addon.keepwarm cron executed')

