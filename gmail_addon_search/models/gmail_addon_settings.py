import json

from odoo import api, fields, models


SAFE_CREATE_FIELD_TYPES = (
    'boolean', 'char', 'date', 'datetime', 'float', 'html',
    'integer', 'many2many', 'many2one', 'selection', 'text',
)
SAFE_REFERENCE_FIELD_TYPES = ('char', 'integer', 'selection', 'text')

MODEL_SPECS = {
    'task': {
        'model': 'project.task',
        'label': 'Task',
        'default_reference': 'task_number',
        'extra_exclude': {'name', 'project_id', 'partner_id', 'user_ids', 'description', 'tag_ids'},
    },
    'ticket': {
        'model': 'helpdesk.ticket',
        'label': 'Ticket',
        'default_reference': 'ticket_ref',
        'extra_exclude': {'name', 'team_id', 'partner_id', 'priority', 'description'},
    },
    'lead': {
        'model': 'crm.lead',
        'label': 'Lead',
        'default_reference': '',
        'extra_exclude': {'name', 'type', 'team_id', 'partner_id', 'contact_name', 'partner_name', 'email_from', 'description'},
    },
}


class GmailAddonConfig(models.AbstractModel):
    _name = 'gmail.addon.config'
    _description = 'Google Workspace Add-on Configuration'

    def _config_key(self, record_type, suffix):
        return f'gmail_addon_search.{record_type}.{suffix}'

    def _get_spec(self, record_type):
        spec = MODEL_SPECS.get(record_type)
        if not spec:
            raise ValueError(f'Unsupported record type: {record_type}')
        return spec

    def _field_record(self, field_id):
        if not field_id:
            return self.env['ir.model.fields']
        return self.env['ir.model.fields'].sudo().browse(int(field_id))

    def _core_excluded_fields(self, record_type):
        return {
            'id', 'display_name', 'create_uid', 'create_date', 'write_uid', 'write_date',
            'message_follower_ids', 'message_ids', 'message_partner_ids', 'activity_ids',
            '__last_update',
        } | set(self._get_spec(record_type)['extra_exclude'])

    def _is_valid_field(self, field_rec, record_type, role='extra'):
        if not field_rec or not field_rec.exists():
            return False

        spec = self._get_spec(record_type)
        allowed_types = SAFE_REFERENCE_FIELD_TYPES if role == 'reference' else SAFE_CREATE_FIELD_TYPES

        if field_rec.model != spec['model']:
            return False
        if field_rec.name in self._core_excluded_fields(record_type):
            return False
        if field_rec.ttype not in allowed_types:
            return False
        if field_rec.ttype in ('many2one', 'many2many') and not field_rec.relation:
            return False
        if field_rec.readonly:
            return False
        if field_rec.compute and not field_rec.store:
            return False
        return True

    def _selection_options(self, record_type, field_rec):
        model = self.env[self._get_spec(record_type)['model']]
        field_def = model._fields.get(field_rec.name)
        if not field_def:
            return []

        selection = field_def.selection
        if callable(selection):
            selection = selection(model)
        return [{'value': str(value), 'label': str(label)} for value, label in (selection or [])]

    def _relation_options(self, field_rec, limit=50):
        if not field_rec.relation or field_rec.relation not in self.env:
            return []
        records = self.env[field_rec.relation].sudo().name_search('', limit=limit)
        return [{'value': str(res_id), 'label': name} for res_id, name in records]

    def _serialize_field(self, record_type, field_rec):
        data = {
            'name': field_rec.name,
            'label': field_rec.field_description or field_rec.name,
            'type': field_rec.ttype,
            'help': field_rec.help or '',
            'required': bool(field_rec.required),
            'options': [],
        }
        if field_rec.ttype == 'selection':
            data['options'] = self._selection_options(record_type, field_rec)
        elif field_rec.ttype in ('many2one', 'many2many'):
            data['relation'] = field_rec.relation
            data['options'] = self._relation_options(field_rec)
        return data

    def _coerce_scalar_value(self, field_rec, raw_value):
        if raw_value in (None, '', [], False):
            return False

        if field_rec.ttype in ('char', 'text', 'html'):
            return str(raw_value)
        if field_rec.ttype == 'selection':
            return str(raw_value)
        if field_rec.ttype == 'boolean':
            if isinstance(raw_value, bool):
                return raw_value
            return str(raw_value).lower() in ('1', 'true', 'on', 'yes')
        if field_rec.ttype == 'integer':
            return int(raw_value)
        if field_rec.ttype == 'float':
            return float(raw_value)
        if field_rec.ttype in ('date', 'datetime'):
            return str(raw_value)
        if field_rec.ttype == 'many2one':
            record_id = int(raw_value)
            if record_id and field_rec.relation in self.env:
                if not self.env[field_rec.relation].sudo().browse(record_id).exists():
                    raise ValueError(f'Unknown value for {field_rec.field_description}')
            return record_id or False
        if field_rec.ttype == 'many2many':
            values = raw_value if isinstance(raw_value, list) else [v for v in str(raw_value).split(',') if v]
            ids = [int(v) for v in values]
            if ids and field_rec.relation in self.env:
                existing = self.env[field_rec.relation].sudo().browse(ids).exists().ids
                if len(existing) != len(set(ids)):
                    raise ValueError(f'Unknown value for {field_rec.field_description}')
            return [(6, 0, ids)]
        return raw_value

    def get_reference_field(self, record_type):
        icp = self.env['ir.config_parameter'].sudo()
        field_id = icp.get_param(self._config_key(record_type, 'reference_field_id'))
        field_rec = self._field_record(field_id)
        if self._is_valid_field(field_rec, record_type, role='reference'):
            return field_rec
        return self.env['ir.model.fields']

    def get_extra_fields(self, record_type):
        icp = self.env['ir.config_parameter'].sudo()
        raw_ids = icp.get_param(self._config_key(record_type, 'extra_field_ids')) or ''
        field_ids = [int(part) for part in raw_ids.split(',') if part.strip().isdigit()]
        fields_rec = self.env['ir.model.fields'].sudo().browse(field_ids)
        return fields_rec.filtered(lambda f: self._is_valid_field(f, record_type, role='extra'))

    def get_form_schema(self, record_type):
        reference = self.get_reference_field(record_type)
        reference_info = {
            'name': reference.name if reference else self._get_spec(record_type)['default_reference'],
            'label': reference.field_description if reference else 'Reference',
        }
        return {
            'record_type': record_type,
            'reference_field': reference_info,
            'extra_fields': [self._serialize_field(record_type, field_rec) for field_rec in self.get_extra_fields(record_type)],
        }

    def get_reference_field_name(self, record_type):
        field_rec = self.get_reference_field(record_type)
        if field_rec:
            return field_rec.name
        return self._get_spec(record_type)['default_reference']

    def get_reference_display_value(self, record_type, data):
        field_name = self.get_reference_field_name(record_type)
        if not field_name:
            return ''

        if isinstance(data, models.BaseModel):
            value = data[field_name] if field_name in data._fields else False
        else:
            value = data.get(field_name)

        if not value:
            return ''
        if isinstance(value, tuple):
            return value[1]
        if isinstance(value, list) and len(value) == 2 and isinstance(value[0], int):
            return value[1]
        return str(value)

    def build_reference_search_domain(self, record_type, search_term):
        field_name = self.get_reference_field_name(record_type)
        if not field_name:
            return []

        if field_name == 'id':
            numeric = ''.join(ch for ch in str(search_term or '') if ch.isdigit())
            return [('id', '=', int(numeric))] if numeric else []

        field_rec = self.env['ir.model.fields'].sudo().search([
            ('model', '=', self._get_spec(record_type)['model']),
            ('name', '=', field_name),
        ], limit=1)
        if not field_rec:
            return []

        if field_rec.ttype in ('char', 'text', 'selection'):
            return [(field_name, 'ilike', search_term)]
        if field_rec.ttype == 'integer':
            numeric = ''.join(ch for ch in str(search_term or '') if ch.isdigit())
            return [(field_name, '=', int(numeric))] if numeric else []
        return []

    def apply_extra_values(self, record_type, extra_values):
        if not extra_values:
            return {}
        if isinstance(extra_values, str):
            extra_values = json.loads(extra_values)

        allowed = {field_rec.name: field_rec for field_rec in self.get_extra_fields(record_type)}
        vals = {}
        for field_name, raw_value in (extra_values or {}).items():
            field_rec = allowed.get(field_name)
            if not field_rec:
                continue
            vals[field_name] = self._coerce_scalar_value(field_rec, raw_value)
        return vals


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    gmail_outlook_manifest_url = fields.Char(
        string='Outlook manifest download URL',
        compute='_compute_gmail_outlook_manifest_url',
        readonly=True,
    )

    gmail_task_reference_field_id = fields.Many2one(
        'ir.model.fields',
        string='Task reference field',
        domain="[('model', '=', 'project.task'), ('ttype', 'in', ('char', 'integer', 'selection', 'text'))]",
    )
    gmail_ticket_reference_field_id = fields.Many2one(
        'ir.model.fields',
        string='Ticket reference field',
        domain="[('model', '=', 'helpdesk.ticket'), ('ttype', 'in', ('char', 'integer', 'selection', 'text'))]",
    )
    gmail_lead_reference_field_id = fields.Many2one(
        'ir.model.fields',
        string='Lead reference field',
        domain="[('model', '=', 'crm.lead'), ('ttype', 'in', ('char', 'integer', 'selection', 'text'))]",
    )

    gmail_task_extra_field_ids = fields.Many2many(
        'ir.model.fields',
        'gmail_task_extra_fields_rel',
        'settings_id',
        'field_id',
        string='Task extra create fields',
        domain="[('model', '=', 'project.task'), ('ttype', 'in', ('boolean', 'char', 'date', 'datetime', 'float', 'html', 'integer', 'many2many', 'many2one', 'selection', 'text'))]",
    )
    gmail_ticket_extra_field_ids = fields.Many2many(
        'ir.model.fields',
        'gmail_ticket_extra_fields_rel',
        'settings_id',
        'field_id',
        string='Ticket extra create fields',
        domain="[('model', '=', 'helpdesk.ticket'), ('ttype', 'in', ('boolean', 'char', 'date', 'datetime', 'float', 'html', 'integer', 'many2many', 'many2one', 'selection', 'text'))]",
    )
    gmail_lead_extra_field_ids = fields.Many2many(
        'ir.model.fields',
        'gmail_lead_extra_fields_rel',
        'settings_id',
        'field_id',
        string='Lead extra create fields',
        domain="[('model', '=', 'crm.lead'), ('ttype', 'in', ('boolean', 'char', 'date', 'datetime', 'float', 'html', 'integer', 'many2many', 'many2one', 'selection', 'text'))]",
    )

    @api.depends_context('uid')
    def _compute_gmail_outlook_manifest_url(self):
        base_url = (self.env['ir.config_parameter'].sudo().get_param('web.base.url') or '').rstrip('/')
        manifest_url = f'{base_url}/gmail_addon/outlook/manifest.xml' if base_url else ''
        for settings in self:
            settings.gmail_outlook_manifest_url = manifest_url

    @api.model
    def get_values(self):
        res = super().get_values()
        icp = self.env['ir.config_parameter'].sudo()
        fields_model = self.env['ir.model.fields'].sudo()

        def _many2one(record_type):
            return int(icp.get_param(f'gmail_addon_search.{record_type}.reference_field_id') or 0)

        def _many2many(record_type):
            raw = icp.get_param(f'gmail_addon_search.{record_type}.extra_field_ids') or ''
            ids = [int(part) for part in raw.split(',') if part.strip().isdigit()]
            return [(6, 0, fields_model.browse(ids).exists().ids)]

        res.update({
            'gmail_task_reference_field_id': _many2one('task'),
            'gmail_ticket_reference_field_id': _many2one('ticket'),
            'gmail_lead_reference_field_id': _many2one('lead'),
            'gmail_task_extra_field_ids': _many2many('task'),
            'gmail_ticket_extra_field_ids': _many2many('ticket'),
            'gmail_lead_extra_field_ids': _many2many('lead'),
        })
        return res

    def set_values(self):
        super().set_values()
        icp = self.env['ir.config_parameter'].sudo()
        config = self.env['gmail.addon.config']

        def _write_reference(record_type, field_rec):
            icp.set_param(
                config._config_key(record_type, 'reference_field_id'),
                str(field_rec.id if config._is_valid_field(field_rec, record_type, role='reference') else ''),
            )

        def _write_extra(record_type, fields_rec):
            valid_ids = fields_rec.filtered(lambda f: config._is_valid_field(f, record_type, role='extra')).ids
            icp.set_param(config._config_key(record_type, 'extra_field_ids'), ','.join(str(fid) for fid in valid_ids))

        for settings in self:
            _write_reference('task', settings.gmail_task_reference_field_id)
            _write_reference('ticket', settings.gmail_ticket_reference_field_id)
            _write_reference('lead', settings.gmail_lead_reference_field_id)

            _write_extra('task', settings.gmail_task_extra_field_ids)
            _write_extra('ticket', settings.gmail_ticket_extra_field_ids)
            _write_extra('lead', settings.gmail_lead_extra_field_ids)
