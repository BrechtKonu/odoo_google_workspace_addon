import logging

_logger = logging.getLogger(__name__)


def migrate(cr, version):
    """Backfill company_id on existing email/document links from their target
    record, so the new multi-company record rule scopes pre-existing links
    correctly. Rows whose target has no company (or is gone) keep NULL =
    shared/visible to all, which matches the prior behaviour.
    """
    for table, in (('gmail_email_link',), ('gmail_document_link',)):
        cr.execute(
            "SELECT DISTINCT res_model FROM %s "
            "WHERE company_id IS NULL AND res_model IS NOT NULL" % table
        )
        models = [row[0] for row in cr.fetchall()]
        for model in models:
            target_table = model.replace('.', '_')
            # Only models that actually have a company_id column can be backfilled.
            cr.execute(
                "SELECT 1 FROM information_schema.columns "
                "WHERE table_name = %s AND column_name = 'company_id' LIMIT 1",
                (target_table,),
            )
            if not cr.fetchone():
                continue
            cr.execute(
                "UPDATE {link} l SET company_id = t.company_id "
                "FROM {target} t "
                "WHERE l.res_model = %s AND l.res_id = t.id "
                "AND l.company_id IS NULL AND t.company_id IS NOT NULL".format(
                    link=table, target=target_table
                ),
                (model,),
            )
            _logger.info("gmail_addon_search: backfilled %s company_id for %s (%s rows)",
                         table, model, cr.rowcount)
