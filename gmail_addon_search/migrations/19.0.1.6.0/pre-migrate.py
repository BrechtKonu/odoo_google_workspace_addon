import logging

_logger = logging.getLogger(__name__)


def migrate(cr, version):
    """Drop duplicate gmail.document.link rows before the (newly enforced)
    unique constraint is created. The constraint existed only as the legacy
    _sql_constraints tuple, which Odoo 19 silently ignored, so duplicates may
    have accumulated. Keep the lowest id per (document_id, host_app, res_model,
    res_id).
    """
    cr.execute("""
        DELETE FROM gmail_document_link a
        USING gmail_document_link b
        WHERE a.id > b.id
          AND a.document_id = b.document_id
          AND a.host_app   = b.host_app
          AND a.res_model  = b.res_model
          AND a.res_id     = b.res_id
    """)
    if cr.rowcount:
        _logger.info("gmail_addon_search: removed %s duplicate gmail.document.link rows", cr.rowcount)
