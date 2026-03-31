import re
import logging
from email.utils import getaddresses

from markupsafe import Markup

from odoo import http
from odoo.http import request
from odoo.tools import html_sanitize
from odoo.tools.mail import email_normalize

_logger = logging.getLogger(__name__)


_TASK_FIELDS = ['id', 'task_number', 'name', 'project_id', 'stage_id',
                'partner_id', 'user_ids', 'write_date']
_TICKET_FIELDS = ['id', 'name', 'team_id', 'stage_id', 'partner_id',
                  'priority', 'ticket_ref', 'user_id', 'write_date']
_PRIORITY_MAP = {'0': 'Low', '1': 'Normal', '2': 'High', '3': 'Urgent'}


class GmailAddonController(http.Controller):

    # ─── HELPERS ─────────────────────────────────────────────────────────────

    def _get_base_url(self):
        if not hasattr(self, '_base_url_cache'):
            self._base_url_cache = request.httprequest.url_root.rstrip('/')
        return self._base_url_cache

    def _task_url(self, task):
        base = self._get_base_url()
        return f"{base}/odoo/all-tasks/{task.id}"

    def _ticket_url(self, ticket):
        base = self._get_base_url()
        return f"{base}/odoo/all-tickets/{ticket.id}"

    def _build_task_domain(self, search_term='', project_id=None, stage_id=None, partner_id=None, user_id=None):
        domain = [('project_id', '!=', False), ('stage_id.gmail_hide_in_search', '!=', True)]
        if search_term:
            numeric = re.sub(r'[^0-9]', '', search_term)
            if numeric and numeric.isdigit():
                name_domain = ['|', '|',
                    ('task_number', 'ilike', search_term),
                    ('name', 'ilike', search_term),
                    ('id', '=', int(numeric))]
            else:
                name_domain = ['|',
                    ('task_number', 'ilike', search_term),
                    ('name', 'ilike', search_term)]
            domain += name_domain
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
            numeric = re.sub(r'[^0-9]', '', search_term)
            if numeric and numeric.isdigit():
                name_domain = ['|',
                    ('name', 'ilike', search_term),
                    ('id', '=', int(numeric))]
            else:
                name_domain = [('name', 'ilike', search_term)]
            domain += name_domain
        if team_id:
            domain += [('team_id', '=', int(team_id))]
        if stage_id:
            domain += [('stage_id', '=', int(stage_id))]
        if partner_id:
            domain += [('partner_id', '=', int(partner_id))]
        if user_id:
            domain += [('user_id', '=', int(user_id))]
        return domain

    def _format_task_dict(self, t, base_url, user_names_map=None):
        user_ids = t.get('user_ids') or []
        user_name = ', '.join(
            (user_names_map or {}).get(uid, '') for uid in user_ids
        )
        return {
            'id': t['id'],
            'task_number': t.get('task_number') or '',
            'name': t['name'],
            'project_name': t['project_id'][1] if t.get('project_id') else '',
            'stage_name': t['stage_id'][1] if t.get('stage_id') else '',
            'partner_name': t['partner_id'][1] if t.get('partner_id') else '',
            'user_name': user_name,
            'url': f"{base_url}/odoo/all-tasks/{t['id']}",
            'write_date': str(t['write_date']) if t.get('write_date') else '',
        }

    def _format_ticket_dict(self, t, base_url):
        return {
            'id': t['id'],
            'name': t['name'],
            'team_name': t['team_id'][1] if t.get('team_id') else '',
            'stage_name': t['stage_id'][1] if t.get('stage_id') else '',
            'partner_name': t['partner_id'][1] if t.get('partner_id') else '',
            'priority': _PRIORITY_MAP.get(t.get('priority', '1'), 'Normal'),
            'ticket_ref': t.get('ticket_ref') or '',
            'user_name': t['user_id'][1] if t.get('user_id') else '',
            'url': f"{base_url}/odoo/all-tickets/{t['id']}",
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
                          gmail_message_id='', gmail_thread_id=''):
        if not rfc_message_id and not gmail_message_id:
            return
        domain = [('res_model', '=', res_model), ('res_id', '=', res_id)]
        if gmail_message_id:
            domain += [('gmail_message_id', '=', gmail_message_id)]
        elif rfc_message_id:
            domain += [('rfc_message_id', '=', rfc_message_id)]
        if not request.env['gmail.email.link'].search(domain, limit=1):
            request.env['gmail.email.link'].create({
                'rfc_message_id': rfc_message_id or '',
                'gmail_message_id': gmail_message_id or '',
                'gmail_thread_id': gmail_thread_id or '',
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

        task_by_id = {}
        if task_ids:
            for t in env['project.task'].search_read(
                [('id', 'in', task_ids)],
                fields=['id', 'name', 'stage_id', 'task_number', 'user_ids']
            ):
                task_by_id[t['id']] = t

        ticket_by_id = {}
        if ticket_ids:
            try:
                for t in env['helpdesk.ticket'].search_read(
                    [('id', 'in', ticket_ids)],
                    fields=['id', 'name', 'stage_id', 'ticket_ref', 'user_id']
                ):
                    ticket_by_id[t['id']] = t
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
                records.append({
                    'type': 'task',
                    'id': t['id'],
                    'name': t['name'],
                    'task_number': t.get('task_number') or '',
                    'user_name': ', '.join(task_user_map.get(uid, '') for uid in user_ids),
                    'stage': t['stage_id'][1] if t.get('stage_id') else '',
                    'url': f"{base_url}/odoo/all-tasks/{t['id']}",
                })
            elif link.res_model == 'helpdesk.ticket' and link.res_id in ticket_by_id:
                t = ticket_by_id[link.res_id]
                records.append({
                    'type': 'ticket',
                    'id': t['id'],
                    'name': t['name'],
                    'ticket_ref': t.get('ticket_ref') or '',
                    'user_name': t['user_id'][1] if t.get('user_id') else '',
                    'stage': t['stage_id'][1] if t.get('stage_id') else '',
                    'url': f"{base_url}/odoo/all-tickets/{t['id']}",
                })
        return records

    # ─── LINKED RECORDS ENDPOINT ─────────────────────────────────────────────

    @http.route('/gmail_addon/email/linked_records', type='jsonrpc', auth='outlook')
    def email_linked_records(self, rfc_message_id='', gmail_message_id='', gmail_thread_id='', **kwargs):
        if not rfc_message_id and not gmail_message_id and not gmail_thread_id:
            return {'records': []}

        rfc_ids = self._rfc_message_id_variants(rfc_message_id)

        # Build OR domain across all provided identifiers
        clauses = []
        if gmail_thread_id:
            clauses += [('gmail_thread_id', '=', gmail_thread_id)]
        if gmail_message_id:
            clauses += [('gmail_message_id', '=', gmail_message_id)]
        if rfc_ids:
            clauses += [('rfc_message_id', 'in', rfc_ids)]

        domain = ['|'] * (len(clauses) - 1) + clauses

        env = request.env
        links = env['gmail.email.link'].search(domain)

        # Fallback for old chatter/emails: resolve via mail.message.message_id and backfill links.
        if not links and rfc_ids:
            try:
                mm_rows = env['mail.message'].sudo().search_read(
                    [('message_id', 'in', rfc_ids), ('model', 'in', ['project.task', 'helpdesk.ticket']), ('res_id', '!=', False)],
                    fields=['model', 'res_id', 'message_id'],
                    limit=200,
                )
                for mm in mm_rows:
                    model = mm.get('model')
                    res_id = mm.get('res_id')
                    if model in ('project.task', 'helpdesk.ticket') and res_id:
                        self._store_email_link(
                            mm.get('message_id') or rfc_message_id,
                            model,
                            int(res_id),
                            gmail_message_id=gmail_message_id,
                            gmail_thread_id=gmail_thread_id,
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
            'recent_tasks': [],
            'recent_tickets': [],
            'company_partner_id': None,
            'company_partner_name': '',
            'company_tasks': [],
            'company_tickets': [],
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

        # Recent tasks for this contact (contact only, not company)
        base_url = self._get_base_url()
        contact_ids = [partner.id]

        # Optional "only mine" filter: records where env.user is assigned or follower
        mine_task_domain = []
        mine_ticket_domain = []
        if filter_mine:
            me = env.user
            me_partner_id = me.partner_id.id
            mine_task_domain = ['|',
                ('user_ids', 'in', [me.id]),
                ('message_partner_ids', 'in', [me_partner_id])]
            mine_ticket_domain = ['|',
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
            fields=_TASK_FIELDS, order='write_date desc', limit=10
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
                ticket_domain, fields=_TICKET_FIELDS, order='write_date desc', limit=10
            )
            result['recent_tickets'] = [self._format_ticket_dict(t, base_url) for t in ticket_data]
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
                fields=_TASK_FIELDS, order='write_date desc', limit=10
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
                    company_ticket_domain, fields=_TICKET_FIELDS, order='write_date desc', limit=10
                )
                result['company_tickets'] = [self._format_ticket_dict(t, base_url) for t in company_ticket_data]
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
            domain, fields=_TASK_FIELDS,
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
                domain, fields=_TICKET_FIELDS,
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

    # ─── CREATE TASK ─────────────────────────────────────────────────────────

    @http.route('/gmail_addon/task/create', type='jsonrpc', auth='outlook')
    def task_create(self, project_id, name, partner_id=None, user_id=None,
                    tag_ids=None, description='', cc_addresses='',
                    email_body='', email_subject='', author_email='',
                    rfc_message_id='', gmail_message_id='', gmail_thread_id='', **kwargs):
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
                               gmail_message_id=gmail_message_id, gmail_thread_id=gmail_thread_id)
        return {'task_id': task.id, 'task_url': self._task_url(task)}

    # ─── CREATE TICKET ───────────────────────────────────────────────────────

    @http.route('/gmail_addon/ticket/create', type='jsonrpc', auth='outlook')
    def ticket_create(self, team_id, name, partner_id=None, priority='1',
                      description='', cc_addresses='',
                      email_body='', email_subject='', author_email='',
                      rfc_message_id='', gmail_message_id='', gmail_thread_id='', **kwargs):
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
                                   gmail_message_id=gmail_message_id, gmail_thread_id=gmail_thread_id)
            return {'ticket_id': ticket.id, 'ticket_url': self._ticket_url(ticket)}
        except Exception as e:
            return {'error': str(e)}

    # ─── LOG EMAIL ───────────────────────────────────────────────────────────

    @http.route('/gmail_addon/log_email', type='jsonrpc', auth='outlook')
    def log_email(self, res_model, res_id, email_body, email_subject='',
                  author_email='', rfc_message_id='', gmail_message_id='', gmail_thread_id='', **kwargs):
        env = request.env

        if res_model not in ('project.task', 'helpdesk.ticket'):
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
                                   gmail_message_id=gmail_message_id, gmail_thread_id=gmail_thread_id)
            return {'success': True, 'message_id': msg.id}
        except Exception as e:
            _logger.exception("log_email failed for %s/%s", res_model, res_id)
            return {'error': str(e)}
