import re
import logging
from email.utils import getaddresses
from xml.sax.saxutils import escape

from markupsafe import Markup

from odoo import http
from odoo.http import request
from odoo.osv import expression
from odoo.tools import html_sanitize
from odoo.tools.mail import email_normalize

_logger = logging.getLogger(__name__)


_TASK_FIELDS_BASE = ['id', 'task_number', 'name', 'project_id', 'stage_id',
                     'partner_id', 'user_ids', 'write_date']
_TICKET_FIELDS_BASE = ['id', 'name', 'team_id', 'stage_id', 'partner_id',
                       'priority', 'ticket_ref', 'user_id', 'write_date']
_LEAD_FIELDS_BASE = ['id', 'name', 'type', 'team_id', 'stage_id', 'partner_id',
                     'user_id', 'email_from', 'contact_name', 'partner_name', 'write_date']
_PRIORITY_MAP = {'0': 'Low', '1': 'Normal', '2': 'High', '3': 'Urgent'}


class GmailAddonController(http.Controller):

    # ─── HELPERS ─────────────────────────────────────────────────────────────

    def _get_base_url(self):
        if not hasattr(self, '_base_url_cache'):
            self._base_url_cache = request.httprequest.url_root.rstrip('/')
        return self._base_url_cache

    def _config(self):
        return request.env['gmail.addon.config'].sudo()

    def _record_fields(self, record_type):
        base_fields = {
            'task': list(_TASK_FIELDS_BASE),
            'ticket': list(_TICKET_FIELDS_BASE),
            'lead': list(_LEAD_FIELDS_BASE),
        }[record_type]
        ref_field_name = self._config().get_reference_field_name(record_type)
        if ref_field_name and ref_field_name != 'id' and ref_field_name not in base_fields:
            base_fields.append(ref_field_name)
        return base_fields

    def _record_reference(self, record_type, data):
        return self._config().get_reference_display_value(record_type, data)

    def _search_ref_or_name_domain(self, record_type, search_term, extra_domains=None):
        domains = [[('name', 'ilike', search_term)]]
        ref_domain = self._config().build_reference_search_domain(record_type, search_term)
        if ref_domain:
            domains.append(ref_domain)

        numeric = re.sub(r'[^0-9]', '', search_term or '')
        if numeric and numeric.isdigit():
            domains.append([('id', '=', int(numeric))])

        for domain in extra_domains or []:
            if domain:
                domains.append(domain)
        return expression.OR(domains)

    def _task_url(self, task):
        base = self._get_base_url()
        return f"{base}/odoo/all-tasks/{task.id}"

    def _ticket_url(self, ticket):
        base = self._get_base_url()
        return f"{base}/odoo/all-tickets/{ticket.id}"

    def _lead_url(self, lead):
        base = self._get_base_url()
        return f"{base}/web#id={lead.id}&model=crm.lead&view_type=form"

    def _outlook_asset_url(self, relative_path):
        base = self._get_base_url()
        clean_path = str(relative_path or '').lstrip('/')
        return f"{base}/gmail_addon_search/static/outlook_addin/{clean_path}"

    def _build_task_domain(self, search_term='', project_id=None, stage_id=None, partner_id=None, user_id=None):
        domain = [('project_id', '!=', False), ('stage_id.gmail_hide_in_search', '!=', True)]
        if search_term:
            domain += self._search_ref_or_name_domain('task', search_term)
        if project_id:
            domain += [('project_id', '=', int(project_id))]
        if stage_id:
            domain += [('stage_id', '=', int(stage_id))]
        if partner_id:
            domain += [('partner_id', '=', int(partner_id))]
        if user_id:
            domain += [('user_ids', 'in', [int(user_id)])]
        return domain

    def _build_ticket_domain(self, search_term='', team_id=None, stage_id=None, partner_id=None, user_id=None):
        domain = []
        try:
            if 'gmail_hide_in_search' in request.env['helpdesk.stage']._fields:
                domain += [('stage_id.gmail_hide_in_search', '!=', True)]
        except Exception:
            pass
        if search_term:
            domain += self._search_ref_or_name_domain('ticket', search_term)
        if team_id:
            domain += [('team_id', '=', int(team_id))]
        if stage_id:
            domain += [('stage_id', '=', int(stage_id))]
        if partner_id:
            domain += [('partner_id', '=', int(partner_id))]
        if user_id:
            domain += [('user_id', '=', int(user_id))]
        return domain

    def _build_lead_domain(self, search_term='', lead_type='all', team_id=None, stage_id=None, user_id=None):
        domain = []
        if search_term:
            extra_domains = [
                [('partner_name', 'ilike', search_term)],
                [('contact_name', 'ilike', search_term)],
                [('email_from', 'ilike', search_term)],
            ]
            domain += self._search_ref_or_name_domain('lead', search_term, extra_domains=extra_domains)
        if lead_type in ('lead', 'opportunity'):
            domain += [('type', '=', lead_type)]
        if team_id:
            domain += [('team_id', '=', int(team_id))]
        if stage_id:
            domain += [('stage_id', '=', int(stage_id))]
        if user_id:
            domain += [('user_id', '=', int(user_id))]
        return domain

    def _format_task_dict(self, t, base_url, user_names_map=None):
        user_ids = t.get('user_ids') or []
        user_name = ', '.join(
            (user_names_map or {}).get(uid, '') for uid in user_ids
        )
        task_ref = self._record_reference('task', t)
        return {
            'id': t['id'],
            'task_number': task_ref or '',
            'reference': task_ref or '',
            'name': t['name'],
            'project_name': t['project_id'][1] if t.get('project_id') else '',
            'stage_name': t['stage_id'][1] if t.get('stage_id') else '',
            'partner_name': t['partner_id'][1] if t.get('partner_id') else '',
            'user_name': user_name,
            'url': f"{base_url}/odoo/all-tasks/{t['id']}",
            'write_date': str(t['write_date']) if t.get('write_date') else '',
        }

    def _format_ticket_dict(self, t, base_url):
        ticket_ref = self._record_reference('ticket', t)
        return {
            'id': t['id'],
            'name': t['name'],
            'team_name': t['team_id'][1] if t.get('team_id') else '',
            'stage_name': t['stage_id'][1] if t.get('stage_id') else '',
            'partner_name': t['partner_id'][1] if t.get('partner_id') else '',
            'priority': _PRIORITY_MAP.get(t.get('priority', '1'), 'Normal'),
            'ticket_ref': ticket_ref or '',
            'reference': ticket_ref or '',
            'user_name': t['user_id'][1] if t.get('user_id') else '',
            'url': f"{base_url}/odoo/all-tickets/{t['id']}",
            'write_date': str(t['write_date']) if t.get('write_date') else '',
        }

    def _format_lead_dict(self, t, base_url):
        lead_ref = self._record_reference('lead', t)
        lead_type = t.get('type') or 'lead'
        return {
            'id': t['id'],
            'name': t['name'],
            'type': lead_type,
            'type_label': 'Opportunity' if lead_type == 'opportunity' else 'Lead',
            'team_name': t['team_id'][1] if t.get('team_id') else '',
            'stage_name': t['stage_id'][1] if t.get('stage_id') else '',
            'partner_name': t['partner_id'][1] if t.get('partner_id') else (t.get('partner_name') or ''),
            'contact_name': t.get('contact_name') or '',
            'email_from': t.get('email_from') or '',
            'lead_ref': lead_ref or '',
            'reference': lead_ref or '',
            'user_name': t['user_id'][1] if t.get('user_id') else '',
            'url': f"{base_url}/web#id={t['id']}&model=crm.lead&view_type=form",
            'write_date': str(t['write_date']) if t.get('write_date') else '',
        }

    def _user_names_map(self, env, task_data_list):
        all_ids = list({uid for t in task_data_list for uid in (t.get('user_ids') or [])})
        if not all_ids:
            return {}
        return {u['id']: u['name'] for u in env['res.users'].browse(all_ids).read(['name'])}

    def _find_partner_by_email(self, email):
        """
        Find the best matching partner for a given email address.
        When multiple partners share the same normalized email, prefer the one
        that has sales orders or invoices (i.e. is a real customer), then fall
        back to the first match (lowest ID).
        """
        env = request.env
        normalized = email_normalize(email) or email
        partners = env['res.partner'].search([('email_normalized', '=', normalized)])
        if not partners:
            return env['res.partner'].browse()
        if len(partners) == 1:
            return partners[0]

        for partner in partners:
            commercial_id = partner.commercial_partner_id.id or partner.id
            has_docs = False
            try:
                has_docs = bool(env['sale.order'].search_count(
                    [('partner_id', 'child_of', commercial_id)], limit=1
                ))
            except Exception:
                pass
            if not has_docs:
                try:
                    has_docs = bool(env['account.move'].search_count([
                        ('partner_id', 'child_of', commercial_id),
                        ('move_type', 'in', ['out_invoice', 'out_refund']),
                    ], limit=1))
                except Exception:
                    pass
            if has_docs:
                return partner

        return partners[0]

    def _cc_to_followers(self, record, cc_addresses):
        if not cc_addresses:
            return
        parsed = [
            (dn, addr.lower())
            for dn, addr in getaddresses([cc_addresses])
            if addr and re.match(r'[^@]+@[^@]+\.[^@]+', addr)
        ]
        if not parsed:
            return
        emails = [addr for _, addr in parsed]
        Partner = request.env['res.partner'].sudo()
        existing = Partner.search([('email_normalized', 'in', emails)])
        by_email = {p.email_normalized: p for p in existing}
        partner_ids = []
        for display_name, email in parsed:
            partner = by_email.get(email)
            if not partner:
                try:
                    partner = Partner.create({'name': display_name.strip() or email, 'email': email})
                except Exception:
                    _logger.warning('_cc_to_followers: skipping %s', email, exc_info=True)
                    continue
            partner_ids.append(partner.id)
        if partner_ids:
            record.message_subscribe(partner_ids=partner_ids)

    def _sanitize_email_body(self, email_body):
        if not email_body:
            return ''
        try:
            return html_sanitize(email_body)
        except Exception:
            _logger.warning('_sanitize_email_body: failed to sanitize body', exc_info=True)
            return ''

    def _contact_partner_ids(self, partner):
        commercial = partner.commercial_partner_id or partner
        return [pid for pid in {partner.id, commercial.id} if pid]

    # ─── EMAIL LINK HELPERS ──────────────────────────────────────────────────

    def _store_email_link(self, rfc_message_id, res_model, res_id, record_name='',
                          gmail_message_id='', gmail_thread_id='',
                          outlook_item_id='', outlook_conversation_id=''):
        if not any([rfc_message_id, gmail_message_id, gmail_thread_id, outlook_item_id, outlook_conversation_id]):
            return
        domain = [('res_model', '=', res_model), ('res_id', '=', res_id)]
        identifier_domains = []
        if gmail_message_id:
            identifier_domains.append([('gmail_message_id', '=', gmail_message_id)])
        if outlook_item_id:
            identifier_domains.append([('outlook_item_id', '=', outlook_item_id)])
        if rfc_message_id:
            identifier_domains.append([('rfc_message_id', '=', rfc_message_id)])
        if gmail_thread_id:
            identifier_domains.append([('gmail_thread_id', '=', gmail_thread_id)])
        if outlook_conversation_id:
            identifier_domains.append([('outlook_conversation_id', '=', outlook_conversation_id)])
        if identifier_domains:
            domain = expression.AND([domain, expression.OR(identifier_domains)])

        if not request.env['gmail.email.link'].search(domain, limit=1):
            request.env['gmail.email.link'].create({
                'rfc_message_id': rfc_message_id or '',
                'gmail_message_id': gmail_message_id or '',
                'gmail_thread_id': gmail_thread_id or '',
                'outlook_item_id': outlook_item_id or '',
                'outlook_conversation_id': outlook_conversation_id or '',
                'res_model': res_model,
                'res_id': res_id,
                'record_name': record_name,
            })

    def _rfc_message_id_variants(self, rfc_message_id):
        """
        Gmail/SMTP message IDs may appear with or without angle brackets.
        Return both variants for resilient matching.
        """
        mid = (rfc_message_id or '').strip()
        if not mid:
            return []
        variants = {mid}
        if mid.startswith('<') and mid.endswith('>'):
            raw = mid[1:-1].strip()
            if raw:
                variants.add(raw)
        else:
            variants.add(f'<{mid}>')
        return [v for v in variants if v]

    def _format_linked_records(self, env, links):
        task_ids = [l.res_id for l in links if l.res_model == 'project.task']
        ticket_ids = [l.res_id for l in links if l.res_model == 'helpdesk.ticket']
        lead_ids = [l.res_id for l in links if l.res_model == 'crm.lead']

        task_by_id = {}
        if task_ids:
            for t in env['project.task'].search_read(
                [('id', 'in', task_ids)],
                fields=self._record_fields('task')
            ):
                task_by_id[t['id']] = t

        ticket_by_id = {}
        if ticket_ids:
            try:
                for t in env['helpdesk.ticket'].search_read(
                    [('id', 'in', ticket_ids)],
                    fields=self._record_fields('ticket')
                ):
                    ticket_by_id[t['id']] = t
            except Exception:
                pass

        lead_by_id = {}
        if lead_ids and 'crm.lead' in env:
            try:
                for t in env['crm.lead'].search_read(
                    [('id', 'in', lead_ids)],
                    fields=self._record_fields('lead')
                ):
                    lead_by_id[t['id']] = t
            except Exception:
                pass

        base_url = self._get_base_url()
        task_user_map = self._user_names_map(env, list(task_by_id.values()))
        records = []
        seen = set()
        for link in links:
            key = (link.res_model, link.res_id)
            if key in seen:
                continue
            seen.add(key)
            if link.res_model == 'project.task' and link.res_id in task_by_id:
                t = task_by_id[link.res_id]
                user_ids = t.get('user_ids') or []
                task_ref = self._record_reference('task', t)
                records.append({
                    'type': 'task',
                    'id': t['id'],
                    'name': t['name'],
                    'task_number': task_ref or '',
                    'reference': task_ref or '',
                    'user_name': ', '.join(task_user_map.get(uid, '') for uid in user_ids),
                    'stage': t['stage_id'][1] if t.get('stage_id') else '',
                    'url': f"{base_url}/odoo/all-tasks/{t['id']}",
                })
            elif link.res_model == 'helpdesk.ticket' and link.res_id in ticket_by_id:
                t = ticket_by_id[link.res_id]
                ticket_ref = self._record_reference('ticket', t)
                records.append({
                    'type': 'ticket',
                    'id': t['id'],
                    'name': t['name'],
                    'ticket_ref': ticket_ref or '',
                    'reference': ticket_ref or '',
                    'user_name': t['user_id'][1] if t.get('user_id') else '',
                    'stage': t['stage_id'][1] if t.get('stage_id') else '',
                    'url': f"{base_url}/odoo/all-tickets/{t['id']}",
                })
            elif link.res_model == 'crm.lead' and link.res_id in lead_by_id:
                t = lead_by_id[link.res_id]
                lead_ref = self._record_reference('lead', t)
                records.append({
                    'type': 'lead',
                    'id': t['id'],
                    'name': t['name'],
                    'lead_ref': lead_ref or '',
                    'reference': lead_ref or '',
                    'lead_type': t.get('type') or 'lead',
                    'user_name': t['user_id'][1] if t.get('user_id') else '',
                    'stage': t['stage_id'][1] if t.get('stage_id') else '',
                    'url': f"{base_url}/web#id={t['id']}&model=crm.lead&view_type=form",
                })
        return records

    def _outlook_manifest_xml(self):
        urls = {
            'icon16': escape(self._outlook_asset_url('assets/icon-16.png')),
            'icon32': escape(self._outlook_asset_url('assets/icon-32.png')),
            'icon80': escape(self._outlook_asset_url('assets/icon-80.png')),
            'commands': escape(self._outlook_asset_url('commands.html')),
            'taskpane': escape(self._outlook_asset_url('taskpane.html')),
        }
        return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"
  xsi:type="MailApp">
  <Id>5ce26c96-72d1-47d0-bd0f-3ef1a62dfcb2</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Konu</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Odoo Mail Workspace"/>
  <Description DefaultValue="Search, create, and link Odoo tasks, tickets, and leads from Outlook."/>
  <IconUrl DefaultValue="{urls['icon16']}"/>
  <HighResolutionIconUrl DefaultValue="{urls['icon80']}"/>
  <SupportUrl DefaultValue="{urls['taskpane']}"/>
  <AppDomains>
    <AppDomain>{escape(self._get_base_url())}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.10"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="{urls['taskpane']}"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="{urls['taskpane']}"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.10">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="Read.Group">
                <Label resid="Group.Label"/>
                <Control xsi:type="Button" id="Read.OpenPane">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url"/>
                    <SupportsPinning>true</SupportsPinning>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="MessageComposeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="Compose.Group">
                <Label resid="Group.Label"/>
                <Control xsi:type="Button" id="Compose.OpenPane">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url"/>
                    <SupportsPinning>true</SupportsPinning>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="{urls['icon16']}"/>
        <bt:Image id="Icon.32" DefaultValue="{urls['icon32']}"/>
        <bt:Image id="Icon.80" DefaultValue="{urls['icon80']}"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="{urls['commands']}"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{urls['taskpane']}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="Odoo"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Open Odoo Workspace"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Search, create, and link Odoo records from the current Outlook email."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""

    # ─── LINKED RECORDS ENDPOINT ─────────────────────────────────────────────

    @http.route('/gmail_addon/outlook/manifest.xml', type='http', auth='public', methods=['GET'], csrf=False)
    def outlook_manifest(self, **kwargs):
        xml_body = self._outlook_manifest_xml()
        headers = [
            ('Content-Type', 'application/xml; charset=utf-8'),
            ('Content-Disposition', 'attachment; filename="konu_outlook_manifest.xml"'),
        ]
        return request.make_response(xml_body, headers=headers)

    @http.route('/gmail_addon/email/linked_records', type='jsonrpc', auth='outlook')
    def email_linked_records(self, rfc_message_id='', gmail_message_id='', gmail_thread_id='',
                             outlook_item_id='', outlook_conversation_id='', **kwargs):
        if not any([rfc_message_id, gmail_message_id, gmail_thread_id, outlook_item_id, outlook_conversation_id]):
            return {'records': []}

        rfc_ids = self._rfc_message_id_variants(rfc_message_id)

        # Build OR domain across all provided identifiers
        clauses = []
        if gmail_thread_id:
            clauses += [('gmail_thread_id', '=', gmail_thread_id)]
        if gmail_message_id:
            clauses += [('gmail_message_id', '=', gmail_message_id)]
        if outlook_conversation_id:
            clauses += [('outlook_conversation_id', '=', outlook_conversation_id)]
        if outlook_item_id:
            clauses += [('outlook_item_id', '=', outlook_item_id)]
        if rfc_ids:
            clauses += [('rfc_message_id', 'in', rfc_ids)]

        domain = ['|'] * (len(clauses) - 1) + clauses

        env = request.env
        links = env['gmail.email.link'].search(domain)

        # Fallback for old chatter/emails: resolve via mail.message.message_id and backfill links.
        if not links and rfc_ids:
            try:
                mm_rows = env['mail.message'].sudo().search_read(
                    [('message_id', 'in', rfc_ids), ('model', 'in', ['project.task', 'helpdesk.ticket', 'crm.lead']), ('res_id', '!=', False)],
                    fields=['model', 'res_id', 'message_id'],
                    limit=200,
                )
                for mm in mm_rows:
                    model = mm.get('model')
                    res_id = mm.get('res_id')
                    if model in ('project.task', 'helpdesk.ticket', 'crm.lead') and res_id:
                        self._store_email_link(
                            mm.get('message_id') or rfc_message_id,
                            model,
                            int(res_id),
                            gmail_message_id=gmail_message_id,
                            gmail_thread_id=gmail_thread_id,
                            outlook_item_id=outlook_item_id,
                            outlook_conversation_id=outlook_conversation_id,
                        )
                if mm_rows:
                    links = env['gmail.email.link'].search(domain)
            except Exception:
                _logger.warning('email_linked_records: mail.message fallback failed', exc_info=True)

        return {'records': self._format_linked_records(env, links)}

    @http.route('/gmail_addon/document/linked_records', type='jsonrpc', auth='outlook')
    def document_linked_records(self, document_id='', host_app='', limit=20, **kwargs):
        host = (host_app or '').strip().lower()
        if not document_id or host not in ('docs', 'sheets'):
            return {'records': []}

        env = request.env
        links = env['gmail.document.link'].search(
            [('document_id', '=', document_id), ('host_app', '=', host)],
            order='write_date desc, id desc',
            limit=int(limit or 20),
        )
        return {'records': self._format_linked_records(env, links)}

    @http.route('/gmail_addon/document/link_record', type='jsonrpc', auth='outlook')
    def document_link_record(self, document_id='', host_app='', res_model='',
                             res_id=None, record_name='', **kwargs):
        host = (host_app or '').strip().lower()
        if not document_id or host not in ('docs', 'sheets'):
            return {'error': 'Invalid document context'}
        if res_model not in ('project.task', 'helpdesk.ticket'):
            return {'error': 'Invalid model'}
        if not res_id:
            return {'error': 'Missing record id'}

        env = request.env
        try:
            rid = int(res_id)
        except Exception:
            return {'error': 'Invalid record id'}

        try:
            record = env[res_model].browse(rid)
        except Exception:
            return {'error': 'Model unavailable'}
        if not record.exists():
            return {'error': 'Record not found'}

        Link = env['gmail.document.link']
        domain = [
            ('document_id', '=', document_id),
            ('host_app', '=', host),
            ('res_model', '=', res_model),
            ('res_id', '=', rid),
        ]
        link = Link.search(domain, limit=1)
        if link:
            vals = {}
            if record_name and record_name != link.record_name:
                vals['record_name'] = record_name
            if vals:
                link.write(vals)
        else:
            link = Link.create({
                'document_id': document_id,
                'host_app': host,
                'res_model': res_model,
                'res_id': rid,
                'record_name': record_name or record.display_name,
            })
        return {'success': True, 'link_id': link.id}

    # ─── CONTEXT SUGGESTION ──────────────────────────────────────────────────

    @http.route('/gmail_addon/suggest_context', type='jsonrpc', auth='outlook')
    def suggest_context(self, sender_email='', filter_mine=False, **kwargs):
        """
        Given sender email, suggest project + team and return recent tasks/tickets.
        """
        env = request.env
        result = {
            'partner_id': None,
            'partner_name': '',
            'partner_email': sender_email,
            'suggested_project_id': None,
            'suggested_project_name': '',
            'suggested_team_id': None,
            'suggested_team_name': '',
            'suggested_crm_team_id': None,
            'suggested_crm_team_name': '',
            'recent_tasks': [],
            'recent_tickets': [],
            'recent_leads': [],
            'company_partner_id': None,
            'company_partner_name': '',
            'company_tasks': [],
            'company_tickets': [],
            'company_leads': [],
        }

        if not sender_email:
            return result

        partner = self._find_partner_by_email(sender_email)
        if not partner:
            return result

        commercial = partner.commercial_partner_id or partner
        contact_partner_ids = self._contact_partner_ids(partner)
        result['partner_id'] = partner.id
        result['partner_name'] = partner.name
        result['partner_email'] = partner.email or sender_email

        # Suggest project: first by partner_id on the project, then by most recent task
        project = env['project.project'].search(
            [('partner_id', 'in', [partner.id, commercial.id])],
            order='write_date desc', limit=1
        )
        if not project:
            task = env['project.task'].search(
                [('partner_id', 'in', [partner.id, commercial.id]),
                 ('project_id', '!=', False)],
                order='write_date desc', limit=1
            )
            project = task.project_id if task else env['project.project'].browse()
        if project:
            result['suggested_project_id'] = project.id
            result['suggested_project_name'] = project.name

        # Suggest helpdesk team (optional)
        try:
            team = env['helpdesk.team'].search(
                [('partner_ids', 'in', [partner.id, commercial.id])],
                order='write_date desc', limit=1
            )
            if not team:
                # Fallback: find team via existing tickets for this partner
                ticket = env['helpdesk.ticket'].search(
                    [('partner_id', 'in', [partner.id, commercial.id])],
                    order='write_date desc', limit=1
                )
                team = ticket.team_id if ticket else env['helpdesk.team'].browse()
            if team:
                result['suggested_team_id'] = team.id
                result['suggested_team_name'] = team.name
        except Exception:
            pass

        try:
            lead = env['crm.lead'].search(
                [('partner_id', 'in', [partner.id, commercial.id])],
                order='write_date desc', limit=1
            )
            if lead and lead.team_id:
                result['suggested_crm_team_id'] = lead.team_id.id
                result['suggested_crm_team_name'] = lead.team_id.name
        except Exception:
            pass

        # Recent tasks for this contact (contact only, not company)
        base_url = self._get_base_url()
        contact_ids = [partner.id]

        # Optional "only mine" filter: records where env.user is assigned or follower
        mine_task_domain = []
        mine_ticket_domain = []
        mine_lead_domain = []
        if filter_mine:
            me = env.user
            me_partner_id = me.partner_id.id
            mine_task_domain = ['|',
                ('user_ids', 'in', [me.id]),
                ('message_partner_ids', 'in', [me_partner_id])]
            mine_ticket_domain = ['|',
                ('user_id', '=', me.id),
                ('message_partner_ids', 'in', [me_partner_id])]
            mine_lead_domain = ['|',
                ('user_id', '=', me.id),
                ('message_partner_ids', 'in', [me_partner_id])]

        def _task_relation_domain(ids):
            clauses = [
                ('partner_id', 'in', ids),
                ('message_partner_ids', 'in', ids),
            ]
            if 'stakeholder_ids' in env['project.task']._fields:
                clauses.append(('stakeholder_ids', 'in', ids))
            return ['|'] * (len(clauses) - 1) + clauses

        task_data = env['project.task'].search_read(
            _task_relation_domain(contact_ids) + mine_task_domain + [
                ('project_id', '!=', False),
                ('stage_id.gmail_hide_in_search', '!=', True),
            ],
            fields=self._record_fields('task'), order='write_date desc', limit=10
        )
        user_map = self._user_names_map(env, task_data)
        result['recent_tasks'] = [self._format_task_dict(t, base_url, user_map) for t in task_data]

        # Recent tickets for this contact (contact only, not company)
        try:
            ticket_domain = ['|',
                             ('partner_id', 'in', contact_ids),
                             ('message_partner_ids', 'in', contact_ids)]
            ticket_domain += mine_ticket_domain
            if 'gmail_hide_in_search' in env['helpdesk.stage']._fields:
                ticket_domain += [('stage_id.gmail_hide_in_search', '!=', True)]
            ticket_data = env['helpdesk.ticket'].search_read(
                ticket_domain, fields=self._record_fields('ticket'), order='write_date desc', limit=10
            )
            result['recent_tickets'] = [self._format_ticket_dict(t, base_url) for t in ticket_data]
        except Exception:
            pass

        try:
            lead_domain = ['|',
                           ('partner_id', 'in', contact_ids),
                           ('message_partner_ids', 'in', contact_ids)]
            lead_domain += mine_lead_domain
            lead_data = env['crm.lead'].search_read(
                lead_domain, fields=self._record_fields('lead'), order='write_date desc', limit=10
            )
            result['recent_leads'] = [self._format_lead_dict(t, base_url) for t in lead_data]
        except Exception:
            pass

        # Company records — only when the contact belongs to a parent company
        if commercial.id != partner.id:
            result['company_partner_id'] = commercial.id
            result['company_partner_name'] = commercial.name
            company_ids = [commercial.id]

            company_task_data = env['project.task'].search_read(
                _task_relation_domain(company_ids) + mine_task_domain + [
                    ('project_id', '!=', False),
                    ('stage_id.gmail_hide_in_search', '!=', True),
                ],
                fields=self._record_fields('task'), order='write_date desc', limit=10
            )
            company_user_map = self._user_names_map(env, company_task_data)
            result['company_tasks'] = [self._format_task_dict(t, base_url, company_user_map) for t in company_task_data]

            try:
                company_ticket_domain = ['|',
                                         ('partner_id', 'in', company_ids),
                                         ('message_partner_ids', 'in', company_ids)]
                company_ticket_domain += mine_ticket_domain
                if 'gmail_hide_in_search' in env['helpdesk.stage']._fields:
                    company_ticket_domain += [('stage_id.gmail_hide_in_search', '!=', True)]
                company_ticket_data = env['helpdesk.ticket'].search_read(
                    company_ticket_domain, fields=self._record_fields('ticket'), order='write_date desc', limit=10
                )
                result['company_tickets'] = [self._format_ticket_dict(t, base_url) for t in company_ticket_data]
            except Exception:
                pass

            try:
                company_lead_domain = ['|',
                                       ('partner_id', 'in', company_ids),
                                       ('message_partner_ids', 'in', company_ids)]
                company_lead_domain += mine_lead_domain
                company_lead_data = env['crm.lead'].search_read(
                    company_lead_domain, fields=self._record_fields('lead'), order='write_date desc', limit=10
                )
                result['company_leads'] = [self._format_lead_dict(t, base_url) for t in company_lead_data]
            except Exception:
                pass

        return result

    # ─── PARTNER CREATE ──────────────────────────────────────────────────────

    @http.route('/gmail_addon/partner/create', type='jsonrpc', auth='outlook')
    def partner_create(self, name='', email='', company_name='', **kwargs):
        if not name or not email:
            return {'error': 'Name and email are required'}
        env = request.env
        existing = self._find_partner_by_email(email)
        if existing:
            return {'partner_id': existing.id, 'partner_name': existing.name, 'already_exists': True}
        vals = {'name': name.strip(), 'email': email.strip()}
        if company_name:
            company = env['res.partner'].sudo().search(
                [('name', 'ilike', company_name.strip()), ('is_company', '=', True)],
                limit=1
            )
            if not company:
                company = env['res.partner'].sudo().create({
                    'name': company_name.strip(), 'is_company': True
                })
            vals['parent_id'] = company.id
        partner = env['res.partner'].sudo().create(vals)
        base_url = self._get_base_url()
        return {
            'partner_id': partner.id,
            'partner_name': partner.name,
            'partner_url': f"{base_url}/odoo/contacts/{partner.id}",
        }

    # ─── TASK SEARCH ─────────────────────────────────────────────────────────

    @http.route('/gmail_addon/task/search', type='jsonrpc', auth='outlook')
    def task_search(self, search_term='', project_id=None, stage_id=None,
                    partner_id=None, user_id=None, limit=10, offset=0, **kwargs):
        env = request.env
        domain = self._build_task_domain(search_term, project_id, stage_id, partner_id, user_id)
        tasks_data = env['project.task'].search_read(
            domain, fields=self._record_fields('task'),
            limit=int(limit), offset=int(offset), order='write_date desc'
        )
        total = env['project.task'].search_count(domain)
        base_url = self._get_base_url()
        user_map = self._user_names_map(env, tasks_data)
        return {
            'tasks': [self._format_task_dict(t, base_url, user_map) for t in tasks_data],
            'total': total,
        }

    # ─── TICKET SEARCH ───────────────────────────────────────────────────────

    @http.route('/gmail_addon/ticket/search', type='jsonrpc', auth='outlook')
    def ticket_search(self, search_term='', team_id=None, stage_id=None,
                      partner_id=None, user_id=None, limit=10, offset=0, **kwargs):
        try:
            env = request.env
            domain = self._build_ticket_domain(search_term, team_id, stage_id, partner_id, user_id)
            tickets_data = env['helpdesk.ticket'].search_read(
                domain, fields=self._record_fields('ticket'),
                limit=int(limit), offset=int(offset), order='write_date desc'
            )
            total = env['helpdesk.ticket'].search_count(domain)
            base_url = self._get_base_url()
            return {
                'tickets': [self._format_ticket_dict(t, base_url) for t in tickets_data],
                'total': total,
            }
        except Exception:
            return {'tickets': [], 'total': 0, 'error': 'Helpdesk module not installed'}

    # ─── PROJECT SEARCH ──────────────────────────────────────────────────────

    @http.route('/gmail_addon/project/search', type='jsonrpc', auth='outlook')
    def project_search(self, search_term='', limit=20, **kwargs):
        env = request.env
        domain = [('name', 'ilike', search_term)] if search_term else []
        projects = env['project.project'].search(domain, limit=int(limit), order='name asc')
        return {
            'projects': [{
                'id': p.id,
                'name': p.name,
                'task_count': p.task_count,
                'partner_name': p.partner_id.name if p.partner_id else '',
            } for p in projects]
        }

    # ─── DROPDOWNS ───────────────────────────────────────────────────────────

    @http.route('/gmail_addon/ping', type='jsonrpc', auth='outlook')
    def ping(self, **kwargs):
        return {'ok': True}

    @http.route('/gmail_addon/form/schema', type='jsonrpc', auth='outlook')
    def form_schema(self, record_type='task', **kwargs):
        try:
            return self._config().get_form_schema((record_type or 'task').strip().lower())
        except Exception as exc:
            return {'record_type': record_type, 'reference_field': {'name': '', 'label': 'Reference'}, 'extra_fields': [], 'error': str(exc)}

    @http.route('/gmail_addon/project/dropdown', type='jsonrpc', auth='outlook')
    def project_dropdown(self, limit=200, **kwargs):
        projects = request.env['project.project'].search_read(
            [],
            fields=['id', 'name'],
            order='name asc',
            limit=int(limit or 200),
        )
        return {'projects': projects}

    @http.route('/gmail_addon/stage/dropdown', type='jsonrpc', auth='outlook')
    def stage_dropdown(self, project_id=None, team_id=None, record_type='task', limit=300, **kwargs):
        env = request.env
        stages = []
        stage_limit = int(limit or 300)
        record_type = (record_type or 'task').strip().lower()

        if record_type == 'ticket':
            try:
                stage_domain = []
                if team_id:
                    stage_domain += [('team_ids', 'in', [int(team_id)])]
                if 'gmail_hide_in_search' in env['helpdesk.stage']._fields:
                    stage_domain += [('gmail_hide_in_search', '=', False)]
                stages = env['helpdesk.stage'].search_read(
                    stage_domain,
                    fields=['id', 'name'],
                    order='sequence asc',
                    limit=stage_limit,
                )
            except Exception:
                return {'stages': [], 'error': 'Helpdesk module not installed'}
        elif project_id:
            stages = env['project.task.type'].search_read(
                [('project_ids', 'in', [int(project_id)]), ('gmail_hide_in_search', '=', False)],
                fields=['id', 'name'],
                order='sequence asc',
                limit=stage_limit,
            )
        else:
            # Return task stages (all, non-hidden) when no project filter
            stages = env['project.task.type'].search_read(
                [('gmail_hide_in_search', '=', False)],
                fields=['id', 'name'],
                order='sequence asc',
                limit=stage_limit,
            )

        return {'stages': stages}

    @http.route('/gmail_addon/team/dropdown', type='jsonrpc', auth='outlook')
    def team_dropdown(self, limit=200, **kwargs):
        try:
            teams = request.env['helpdesk.team'].search_read(
                [],
                fields=['id', 'name'],
                order='name asc',
                limit=int(limit or 200),
            )
            return {'teams': teams}
        except Exception:
            return {'teams': [], 'error': 'Helpdesk module not installed'}

    @http.route('/gmail_addon/user/dropdown', type='jsonrpc', auth='outlook')
    def user_dropdown(self, limit=200, **kwargs):
        users = request.env['res.users'].search_read(
            [('share', '=', False), ('active', '=', True)],
            fields=['id', 'name'],
            order='name asc',
            limit=int(limit or 200),
        )
        return {'users': users}

    @http.route('/gmail_addon/partner/autocomplete', type='jsonrpc', auth='outlook')
    def partner_autocomplete(self, search_term='', limit=10, companies_only=False, **kwargs):
        if companies_only:
            domain = [('is_company', '=', True)]
            if search_term:
                domain += [('name', 'ilike', search_term)]
            partners = request.env['res.partner'].search(domain, limit=int(limit), order='write_date desc')
        else:
            if not search_term:
                return {'partners': []}
            domain = ['|', ('name', 'ilike', search_term), ('email', 'ilike', search_term)]
            partners = request.env['res.partner'].search(domain, limit=int(limit), order='name asc')
        return {
            'partners': [{'id': p.id, 'name': p.name, 'email': p.email or '', 'is_company': p.is_company} for p in partners]
        }

    @http.route('/gmail_addon/crm/team/dropdown', type='jsonrpc', auth='outlook')
    def crm_team_dropdown(self, limit=200, **kwargs):
        try:
            teams = request.env['crm.team'].search_read(
                [],
                fields=['id', 'name'],
                order='name asc',
                limit=int(limit or 200),
            )
            return {'teams': teams}
        except Exception:
            return {'teams': [], 'error': 'CRM module not installed'}

    @http.route('/gmail_addon/crm/stage/dropdown', type='jsonrpc', auth='outlook')
    def crm_stage_dropdown(self, team_id=None, limit=300, **kwargs):
        try:
            stage_domain = []
            if team_id:
                stage_domain = ['|', ('team_id', '=', False), ('team_id', '=', int(team_id))]
            stages = request.env['crm.stage'].search_read(
                stage_domain,
                fields=['id', 'name'],
                order='sequence asc',
                limit=int(limit or 300),
            )
            return {'stages': stages}
        except Exception:
            return {'stages': [], 'error': 'CRM module not installed'}

    # ─── CREATE TASK ─────────────────────────────────────────────────────────

    @http.route('/gmail_addon/task/create', type='jsonrpc', auth='outlook')
    def task_create(self, project_id, name, partner_id=None, user_id=None,
                    tag_ids=None, description='', cc_addresses='',
                    extra_values=None,
                    email_body='', email_subject='', author_email='',
                    rfc_message_id='', gmail_message_id='', gmail_thread_id='',
                    outlook_item_id='', outlook_conversation_id='', **kwargs):
        env = request.env
        vals = {
            'name': name,
            'project_id': int(project_id),
            'description': description or '',
        }
        if partner_id:
            vals['partner_id'] = int(partner_id)
        if user_id:
            vals['user_ids'] = [(4, int(user_id))]
        elif env.user:
            vals['user_ids'] = [(4, env.user.id)]
        if tag_ids:
            vals['tag_ids'] = [(6, 0, [int(t) for t in tag_ids])]
        vals.update(self._config().apply_extra_values('task', extra_values))

        task = env['project.task'].create(vals)
        self._cc_to_followers(task, cc_addresses)

        if email_body:
            sanitized_body = self._sanitize_email_body(email_body)
            author = None
            if author_email:
                normalized = email_normalize(author_email) or author_email
                author = env['res.partner'].search(
                    [('email_normalized', '=', normalized)], limit=1
                )
            task.message_post(
                body=Markup(sanitized_body),
                subject=email_subject or 'Logged from Gmail',
                message_type='comment',
                subtype_xmlid='mail.mt_note',
                author_id=author.id if author else None,
            )

        self._store_email_link(rfc_message_id, 'project.task', task.id, task.name,
                               gmail_message_id=gmail_message_id,
                               gmail_thread_id=gmail_thread_id,
                               outlook_item_id=outlook_item_id,
                               outlook_conversation_id=outlook_conversation_id)
        return {'task_id': task.id, 'task_url': self._task_url(task), 'task_number': self._record_reference('task', task)}

    # ─── CREATE TICKET ───────────────────────────────────────────────────────

    @http.route('/gmail_addon/ticket/create', type='jsonrpc', auth='outlook')
    def ticket_create(self, team_id, name, partner_id=None, priority='1',
                      extra_values=None,
                      description='', cc_addresses='',
                      email_body='', email_subject='', author_email='',
                      rfc_message_id='', gmail_message_id='', gmail_thread_id='',
                      outlook_item_id='', outlook_conversation_id='', **kwargs):
        try:
            env = request.env
            vals = {
                'name': name,
                'team_id': int(team_id),
                'description': description or '',
                'priority': str(priority),
            }
            if partner_id:
                vals['partner_id'] = int(partner_id)
            vals.update(self._config().apply_extra_values('ticket', extra_values))

            ticket = env['helpdesk.ticket'].create(vals)
            self._cc_to_followers(ticket, cc_addresses)

            if email_body:
                sanitized_body = self._sanitize_email_body(email_body)
                author = None
                if author_email:
                    normalized = email_normalize(author_email) or author_email
                    author = env['res.partner'].search(
                        [('email_normalized', '=', normalized)], limit=1
                    )
                ticket.message_post(
                    body=Markup(sanitized_body),
                    subject=email_subject or 'Logged from Gmail',
                    message_type='comment',
                    subtype_xmlid='mail.mt_note',
                    author_id=author.id if author else None,
                )

            self._store_email_link(rfc_message_id, 'helpdesk.ticket', ticket.id, ticket.name,
                                   gmail_message_id=gmail_message_id,
                                   gmail_thread_id=gmail_thread_id,
                                   outlook_item_id=outlook_item_id,
                                   outlook_conversation_id=outlook_conversation_id)
            return {'ticket_id': ticket.id, 'ticket_url': self._ticket_url(ticket), 'ticket_ref': self._record_reference('ticket', ticket)}
        except Exception as e:
            return {'error': str(e)}

    @http.route('/gmail_addon/lead/search', type='jsonrpc', auth='outlook')
    def lead_search(self, search_term='', lead_type='all', team_id=None, stage_id=None,
                    user_id=None, limit=10, offset=0, **kwargs):
        try:
            env = request.env
            domain = self._build_lead_domain(search_term, lead_type, team_id, stage_id, user_id)
            leads_data = env['crm.lead'].search_read(
                domain,
                fields=self._record_fields('lead'),
                limit=int(limit),
                offset=int(offset),
                order='write_date desc',
            )
            total = env['crm.lead'].search_count(domain)
            base_url = self._get_base_url()
            return {
                'leads': [self._format_lead_dict(t, base_url) for t in leads_data],
                'total': total,
            }
        except Exception:
            return {'leads': [], 'total': 0, 'error': 'CRM module not installed'}

    @http.route('/gmail_addon/lead/create', type='jsonrpc', auth='outlook')
    def lead_create(self, name, lead_type='lead', team_id=None, partner_id=None, contact_name='',
                    partner_name='', email_from='', description='', cc_addresses='',
                    extra_values=None,
                    email_body='', email_subject='', author_email='',
                    rfc_message_id='', gmail_message_id='', gmail_thread_id='',
                    outlook_item_id='', outlook_conversation_id='', **kwargs):
        try:
            env = request.env
            vals = {
                'name': name,
                'type': lead_type if lead_type in ('lead', 'opportunity') else 'lead',
                'description': description or '',
            }
            if team_id:
                vals['team_id'] = int(team_id)
            if partner_id:
                vals['partner_id'] = int(partner_id)
            else:
                if contact_name:
                    vals['contact_name'] = contact_name
                if partner_name:
                    vals['partner_name'] = partner_name
                if email_from:
                    vals['email_from'] = email_from
            vals.update(self._config().apply_extra_values('lead', extra_values))

            lead = env['crm.lead'].create(vals)
            self._cc_to_followers(lead, cc_addresses)

            if email_body:
                sanitized_body = self._sanitize_email_body(email_body)
                author = None
                if author_email:
                    normalized = email_normalize(author_email) or author_email
                    author = env['res.partner'].search(
                        [('email_normalized', '=', normalized)], limit=1
                    )
                lead.message_post(
                    body=Markup(sanitized_body),
                    subject=email_subject or 'Logged from Gmail',
                    message_type='comment',
                    subtype_xmlid='mail.mt_note',
                    author_id=author.id if author else None,
                )

            self._store_email_link(rfc_message_id, 'crm.lead', lead.id, lead.name,
                                   gmail_message_id=gmail_message_id,
                                   gmail_thread_id=gmail_thread_id,
                                   outlook_item_id=outlook_item_id,
                                   outlook_conversation_id=outlook_conversation_id)
            return {
                'lead_id': lead.id,
                'lead_url': self._lead_url(lead),
                'lead_ref': self._record_reference('lead', lead),
            }
        except Exception as exc:
            return {'error': str(exc)}

    # ─── LOG EMAIL ───────────────────────────────────────────────────────────

    @http.route('/gmail_addon/log_email', type='jsonrpc', auth='outlook')
    def log_email(self, res_model, res_id, email_body, email_subject='',
                  author_email='', rfc_message_id='', gmail_message_id='', gmail_thread_id='',
                  outlook_item_id='', outlook_conversation_id='', **kwargs):
        env = request.env

        if res_model not in ('project.task', 'helpdesk.ticket', 'crm.lead'):
            return {'error': 'Invalid model'}

        try:
            record = env[res_model].browse(int(res_id))
            if not record.exists():
                return {'error': 'Record not found'}

            author = None
            if author_email:
                normalized = email_normalize(author_email) or author_email
                author = env['res.partner'].search(
                    [('email_normalized', '=', normalized)], limit=1
                )

            msg = record.message_post(
                body=Markup(self._sanitize_email_body(email_body)),
                subject=email_subject or 'Logged from Gmail',
                message_type='comment',
                subtype_xmlid='mail.mt_note',
                author_id=author.id if author else None,
            )
            self._store_email_link(rfc_message_id, res_model, int(res_id), record.display_name,
                                   gmail_message_id=gmail_message_id,
                                   gmail_thread_id=gmail_thread_id,
                                   outlook_item_id=outlook_item_id,
                                   outlook_conversation_id=outlook_conversation_id)
            return {'success': True, 'message_id': msg.id}
        except Exception as e:
            _logger.exception("log_email failed for %s/%s", res_model, res_id)
            return {'error': str(e)}
