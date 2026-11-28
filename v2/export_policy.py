"""Bind export consent to the exact approved snapshot, configuration and request."""
import hashlib
import hmac
import json


def preview_token(app, query, fields):
    approved = app.approved()
    version = app.state / 'approved' / 'versions' / approved['run_id']
    files = [app.state / 'approved' / 'current.json']
    files += [version / name for name in ('records.jsonl', 'manifest.json')]
    files += sorted(p for p in app.config.glob('*') if p.is_file())
    identity = {
        'run': approved['run_id'],
        'request': {k: v for k, v in query.items()
                    if k not in ('preview', 'run', 'literal_links', 'literal_dates', 'partition')},
        'columns': fields,
        'files': [(str(p.relative_to(app.root)), hashlib.sha256(p.read_bytes()).hexdigest())
                  for p in files if p.exists()],
    }
    body = json.dumps(identity, sort_keys=True, ensure_ascii=False).encode()
    return hmac.new(app.token.encode(), body, hashlib.sha256).hexdigest()


def link_exceptions(rows, fields, target):
    exceptions = []
    for index, row in enumerate(rows, 1):
        for field in ('Document Link', 'Source Workbook Path'):
            value = row.get(field)
            if field in fields and value and target(value) is None:
                exceptions.append({
                    'export_row': index, 'field': field, 'target': value,
                    'original_target': row.get('Original_Hyperlink_Target', ''),
                    'event_key': row.get('Event_Key', ''),
                    'source': row.get('Source File', ''), 'worksheet': row.get('Source Sheet', ''),
                    'row': row.get('Source Row', ''), 'cell': row.get('Source Cell', ''),
                    'reason': 'Target is unsupported, unsafe or exceeds native hyperlink limits; retained literally.',
                })
    return exceptions


def date_exceptions(rows, fields):
    from datetime import datetime
    exceptions=[]
    for index,row in enumerate(rows,1):
        for field in ('Event Date','Latest Event Date','Source Modified Date'):
            value=row.get(field)
            if field not in fields or not isinstance(value,str) or len(value)!=10 or value>='1900-01-01':continue
            try:datetime.strptime(value,'%Y-%m-%d')
            except ValueError:continue
            exceptions.append({'export_row':index,'field':field,'value':value,
                               'source':row.get('Source File',''),'worksheet':row.get('Source Sheet',''),
                               'row':row.get('Source Row',''),'cell':row.get('Source Cell',''),
                               'reason':'Date precedes the Excel 1900 date system; preserved exactly as ISO text.'})
    return exceptions
