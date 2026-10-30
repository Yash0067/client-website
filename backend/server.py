#!/usr/bin/env python3
"""Adarsh Boutique appointment site. Store requests in SQLite; export Excel.
The web server serves only the routes defined below, never arbitrary files.
"""
import base64
import binascii
import datetime as dt
import hmac
import html
import io
import json
import os
import re
import sqlite3
import sys
import threading
import time
from collections import defaultdict, deque
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit
from zoneinfo import ZoneInfo

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    sys.exit('Install the dependency first: pip install -r requirements.txt')

ROOT = Path(__file__).resolve().parent
default_data_dir = '/tmp/adarsh-boutique-data' if os.environ.get('VERCEL') else str(ROOT / 'data')
DATA = Path(os.environ.get('APP_DATA_DIR', default_data_dir)).resolve()
DATA.mkdir(mode=0o700, parents=True, exist_ok=True)
DB = DATA / 'appointments.sqlite3'
XLSX_PATH = Path(os.environ.get('APPOINTMENTS_XLSX', str(DATA / 'Adarsh_Appointments.xlsx')).strip()).expanduser().resolve()
SERVICES = {'Blouse stitching', 'Fancy dresses', 'Gowns & western wear', 'Custom co-ord sets', 'Uniform stitching', 'Salwar suits & kurtis', 'Lehenga & bridal wear', 'Petticoat & saree fall', 'Alterations & fittings', 'Something else'}
COLUMNS = ['Request ID', 'Received (India time)', 'Customer name', 'Phone', 'Service', 'Preferred date', 'Preferred time', 'Customer query', 'Status']
rate_lock = threading.Lock()
rate_events = defaultdict(deque)
xlsx_lock = threading.Lock()

ADMIN_USER = os.environ.get('ADMIN_USER', 'admin')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')
FRONTEND_URL = os.environ.get('FRONTEND_URL', 'https://client-website-frontend.vercel.app').rstrip('/')


def db_connect():
    conn = sqlite3.connect(DB, timeout=25)
    conn.execute('PRAGMA busy_timeout=25000')
    return conn


def init_db():
    with db_connect() as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS appointments (
            request_id TEXT PRIMARY KEY, received_at TEXT NOT NULL, name TEXT NOT NULL,
            phone TEXT NOT NULL, service TEXT NOT NULL, date TEXT NOT NULL,
            time TEXT NOT NULL, notes TEXT NOT NULL, status TEXT NOT NULL)''')
    try:
        os.chmod(DB, 0o600)
    except OSError:
        pass


def bounded_rate(key, maximum, period):
    now = time.monotonic()
    with rate_lock:
        q = rate_events[key]
        while q and q[0] <= now - period:
            q.popleft()
        if len(q) >= maximum:
            return False
        q.append(now)
        return True


def message_page(title, message, success=False):
    return ('''<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>''' + html.escape(title) + '''</title>
<style>body{background:#f8f3eb;color:#352725;font:16px/1.6 Arial,sans-serif;margin:0;padding:30px 18px}.card{max-width:510px;margin:7vh auto;background:#fffdf8;padding:34px;border:1px solid #e7d9cb;border-radius:6px}h1{font:32px/1.15 Georgia,serif;color:#632c35;margin-top:12px}a{color:#632c35;font-weight:600}.tag{font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#957857}.btn{display:inline-block;margin-top:16px;background:#632c35;color:#fff!important;text-decoration:none;padding:10px 20px;border-radius:4px;font-size:14px;font-weight:600}</style>
</head><body><main class="card"><div class="tag">Adarsh Boutique · Kandivali West</div><h1>''' + html.escape(title) + '</h1><p>' + html.escape(message) + '''</p><p><a class="btn" href="''' + html.escape(FRONTEND_URL) + '''">Return to the website</a></p><p><a href="https://wa.me/919967004638">Contact us on WhatsApp</a></p></main></body></html>''').encode('utf-8')


def admin_page(rows):
    rows_html = []
    for r in rows:
        rows_html.append('<tr>' + ''.join('<td>' + html.escape(str(col)) + '</td>' for col in r) + '</tr>')
    table_body = ''.join(rows_html) if rows_html else '<tr><td colspan="9" style="text-align:center;padding:24px;color:#7c7066;">No appointments recorded yet.</td></tr>'
    
    return ('''<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Adarsh Boutique · Admin Appointments</title>
<style>
body{background:#f8f3eb;color:#352725;font:15px/1.5 Arial,sans-serif;margin:0;padding:24px}
.header{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:16px;max-width:1200px;margin:0 auto 20px}
h1{font:28px Georgia,serif;color:#632c35;margin:0}
.actions{display:flex;gap:12px}
.btn{display:inline-block;background:#632c35;color:#fff;text-decoration:none;padding:10px 18px;border-radius:4px;font-size:14px;font-weight:600}
.btn:hover{background:#472029}
.btn-secondary{background:#fffdf8;color:#632c35;border:1px solid #632c35}
.card{max-width:1200px;margin:0 auto;background:#fffdf8;border:1px solid #e7d9cb;border-radius:6px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.04)}
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:14px;text-align:left}
th{background:#632c35;color:#fff;padding:12px 14px;font-weight:600;white-space:nowrap}
td{padding:12px 14px;border-bottom:1px solid #eee5dc;vertical-align:top}
tr:nth-child(even){background:#faf6f0}
tr:hover{background:#f3ece2}
.count{color:#7c7066;font-size:14px;margin-top:4px}
</style>
</head><body>
<div class="header">
  <div>
    <h1>Adarsh Boutique Appointments</h1>
    <div class="count">Total requests: ''' + str(len(rows)) + '''</div>
  </div>
  <div class="actions">
    <a class="btn" href="/admin/appointments.xlsx">📥 Download Excel Sheet (.xlsx)</a>
    <a class="btn btn-secondary" href="''' + html.escape(FRONTEND_URL) + '''" target="_blank">🌐 Open Website</a>
  </div>
</div>
<div class="card">
  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Request ID</th><th>Received (IST)</th><th>Name</th><th>Phone</th><th>Service</th><th>Date</th><th>Time</th><th>Notes</th><th>Status</th>
        </tr>
      </thead>
      <tbody>
        ''' + table_body + '''
      </tbody>
    </table>
  </div>
</div>
</body></html>''').encode('utf-8')


def valid_submission(fields):
    def get(name, minlen, maxlen):
        values = fields.get(name, [])
        if len(values) != 1:
            raise ValueError('Please check your details and try again.')
        value = values[0].strip()
        if len(value) < minlen or len(value) > maxlen or any(ord(c) < 32 and c not in '\n\r\t' for c in value):
            raise ValueError('Please check your details and try again.')
        return value

    if fields.get('website', [''])[0].strip():
        raise ValueError('Please contact the boutique directly.')
    request_id = get('request_id', 16, 80)
    if not re.fullmatch(r'[A-Za-z0-9-]+', request_id):
        raise ValueError('Please check your details and try again.')
    name = get('name', 2, 80)
    phone = get('phone', 7, 30)
    if not re.fullmatch(r'[+()0-9 .-]+', phone) or not 7 <= len(re.sub(r'[^0-9]', '', phone)) <= 15:
        raise ValueError('Please enter a valid phone number.')
    service = get('service', 2, 80)
    if service not in SERVICES:
        raise ValueError('Please choose a valid service.')
    date = get('date', 10, 10)
    try:
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', date):
            raise ValueError()
        chosen_date = dt.date.fromisoformat(date)
        if chosen_date < dt.datetime.now(ZoneInfo('Asia/Kolkata')).date():
            raise ValueError()
    except ValueError:
        raise ValueError('Please choose a valid date that is not in the past.')
    appointment_time = get('time', 5, 5)
    if not re.fullmatch(r'([01]\d|2[0-3]):[0-5]\d', appointment_time):
        raise ValueError('Please choose a valid time.')
    notes = get('notes', 0, 1000) if 'notes' in fields else ''
    if get('consent', 1, 10) != 'yes':
        raise ValueError('Please agree to the contact details being used for your request.')
    return request_id, name, phone, service, date, appointment_time, notes


def xlsx_bytes(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Appointments'
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='632C35')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = 'A1:I' + str(max(len(rows) + 1, 1))
    for row in rows:
        ws.append([str(value) for value in row])
        # Force literal strings, never formulas, including customer-supplied values.
        for cell in ws[ws.max_row]:
            cell.data_type = 's'
            cell.number_format = '@'
    for index, width in enumerate([39, 23, 25, 20, 25, 19, 19, 42, 16], 1):
        ws.column_dimensions[chr(64 + index)].width = width
    stream = io.BytesIO()
    wb.save(stream)
    return stream.getvalue()


def append_to_xlsx(row):
    with xlsx_lock:
        if XLSX_PATH.exists():
            try:
                workbook = load_workbook(XLSX_PATH)
                worksheet = workbook.active
            except Exception:
                workbook = Workbook()
                worksheet = workbook.active
        else:
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = 'Appointments'
            worksheet.append(COLUMNS)

        headers = [cell.value for cell in worksheet[1]]
        if not headers or all(value is None for value in headers):
            worksheet.delete_rows(1, worksheet.max_row)
            worksheet.append(COLUMNS)
            headers = COLUMNS
        elif len(headers) < len(COLUMNS):
            for index, value in enumerate(COLUMNS[len(headers):], len(headers) + 1):
                worksheet.cell(row=1, column=index, value=value)
            headers = [cell.value for cell in worksheet[1]]
        if len(headers) >= 8 and headers[7] in ('Details', 'Query', 'Notes'):
            worksheet.cell(row=1, column=8, value='Customer query')

        worksheet.append([str(value) for value in row])
        for cell in worksheet[worksheet.max_row]:
            cell.data_type = 's'
            cell.number_format = '@'
        worksheet.auto_filter.ref = 'A1:I' + str(worksheet.max_row)
        workbook.save(XLSX_PATH)


def xlsx_file_bytes():
    with xlsx_lock:
        if XLSX_PATH.exists():
            try:
                return XLSX_PATH.read_bytes()
            except Exception:
                pass
        with db_connect() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT request_id, received_at, name, phone, service, date, time, notes, status FROM appointments ORDER BY received_at ASC')
            rows = cursor.fetchall()
        content = xlsx_bytes(rows)
        try:
            XLSX_PATH.write_bytes(content)
        except Exception:
            pass
        return content


class Handler(BaseHTTPRequestHandler):
    server_version = 'AdarshWebsite'
    sys_version = ''

    def log_message(self, fmt, *args):
        # Avoid logging submitted data or query strings.
        sys.stderr.write('%s - %s\n' % (self.client_address[0], fmt % args))

    def reply(self, status, body, content_type='text/html; charset=utf-8', headers=None):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Referrer-Policy', 'strict-origin-when-cross-origin')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With')
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With')
        self.send_header('Access-Control-Max-Age', '86400')
        self.end_headers()

    def authorized(self):
        url_parts = urlsplit(self.path)
        query = parse_qs(url_parts.query)
        token = query.get('key', [''])[0] or query.get('token', [''])[0]
        if token and hmac.compare_digest(token.encode(), ADMIN_PASSWORD.encode()):
            return True

        header = self.headers.get('Authorization', '')
        valid = False
        if header.startswith('Basic '):
            try:
                raw = base64.b64decode(header[6:], validate=True).decode('utf-8')
                user, password = raw.split(':', 1)
                valid = hmac.compare_digest(user.encode(), ADMIN_USER.encode()) and hmac.compare_digest(password.encode(), ADMIN_PASSWORD.encode())
            except (binascii.Error, UnicodeError, ValueError):
                pass
        if valid:
            return True
        # A proxy may make visitors share an IP; this simple limiter is not a substitute
        # for a production reverse proxy's access and rate-limiting controls.
        if header and not bounded_rate(('admin', self.client_address[0]), 12, 900):
            self.reply(429, b'Too many attempts. Try again later.', 'text/plain; charset=utf-8', {'Cache-Control':'no-store'})
        else:
            self.reply(401, b'Sign in with your boutique admin credentials.', 'text/plain; charset=utf-8', {'WWW-Authenticate':'Basic realm="Adarsh Boutique Admin", charset="UTF-8"', 'Cache-Control':'no-store'})
        return False

    def do_GET(self):
        url_parts = urlsplit(self.path)
        path = url_parts.path.rstrip('/')
        if not path:
            path = '/'

        if path == '/healthz':
            self.reply(200, b'ok', 'text/plain; charset=utf-8')
            return

        if path in ('/', '/index.html'):
            self.send_response(302)
            self.send_header('Location', FRONTEND_URL)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            return

        if path in ('/admin/appointments.xlsx', '/export/appointments.xlsx', '/download-excel'):
            if not self.authorized():
                return
            data = xlsx_file_bytes()
            self.reply(200, data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={
                'Content-Disposition': 'attachment; filename="Adarsh_Appointments.xlsx"',
                'Cache-Control': 'no-store, no-cache, must-revalidate'
            })
            return

        if path == '/admin':
            if not self.authorized():
                return
            with db_connect() as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT request_id, received_at, name, phone, service, date, time, notes, status FROM appointments ORDER BY received_at DESC')
                rows = cursor.fetchall()
            self.reply(200, admin_page(rows), 'text/html; charset=utf-8', headers={'Cache-Control':'no-store'})
            return

        self.reply(404, b'Not found', 'text/plain; charset=utf-8')

    def do_POST(self):
        if urlsplit(self.path).path != '/api/appointments':
            self.reply(404, b'Not found', 'text/plain; charset=utf-8')
            return
        if not bounded_rate(('request', self.client_address[0]), 40, 3600):
            self.reply(429, message_page('Request could not be saved', 'Too many requests. Please try again later or contact the boutique.'))
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if length < 1 or length > 16384:
                raise ValueError('Please check your details and try again.')
            raw = self.rfile.read(length).decode('utf-8', errors='strict')
            content_type = self.headers.get_content_type()
            if content_type == 'application/json':
                data = json.loads(raw)
                fields = {k: [str(v)] for k, v in data.items()}
            elif content_type == 'application/x-www-form-urlencoded':
                fields = parse_qs(raw, keep_blank_values=True, max_num_fields=30, encoding='utf-8', errors='strict')
            else:
                raise ValueError('Please check your details and try again.')

            request_id, name, phone, service, date, appointment_time, notes = valid_submission(fields)
            now = dt.datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%Y-%m-%d %H:%M:%S')
            with db_connect() as conn:
                conn.execute('BEGIN IMMEDIATE')
                result = conn.execute('INSERT OR IGNORE INTO appointments VALUES (?,?,?,?,?,?,?,?,?)', (request_id, now, name, phone, service, date, appointment_time, notes, 'Pending'))
                duplicate = result.rowcount == 0
                conn.commit()  # Only report success after the database commit succeeds.
            if not duplicate:
                append_to_xlsx((request_id, now, name, phone, service, date, appointment_time, notes, 'Pending'))
            
            message = ('This request was already received. ' if duplicate else 'Your request was saved. ') + 'The boutique will contact you to confirm the date and time. Your appointment is not confirmed yet.'
            
            accept = self.headers.get('Accept', '')
            if 'application/json' in accept or content_type == 'application/json':
                resp_json = json.dumps({'success': True, 'duplicate': duplicate, 'message': message}).encode('utf-8')
                self.reply(200, resp_json, 'application/json; charset=utf-8', headers={'Cache-Control':'no-store'})
            else:
                self.reply(200, message_page('Appointment request received', message, True), headers={'Cache-Control':'no-store'})
        except (ValueError, UnicodeError, json.JSONDecodeError) as exc:
            err_msg = str(exc) if str(exc) and not str(exc).startswith('Expecting') else 'Please check the form details and try again, or contact the boutique on WhatsApp.'
            accept = self.headers.get('Accept', '')
            if 'application/json' in accept or self.headers.get_content_type() == 'application/json':
                self.reply(400, json.dumps({'success': False, 'error': err_msg}).encode('utf-8'), 'application/json; charset=utf-8', headers={'Cache-Control':'no-store'})
            else:
                self.reply(400, message_page('Request could not be saved', err_msg), headers={'Cache-Control':'no-store'})
        except Exception as exc:
            print('Appointment save failed:', type(exc).__name__, file=sys.stderr)
            accept = self.headers.get('Accept', '')
            if 'application/json' in accept:
                self.reply(500, json.dumps({'success': False, 'error': 'Please try again later or contact the boutique on WhatsApp.'}).encode('utf-8'), 'application/json; charset=utf-8', headers={'Cache-Control':'no-store'})
            else:
                self.reply(500, message_page('Request could not be saved', 'Please try again later or contact the boutique on WhatsApp.'), headers={'Cache-Control':'no-store'})


if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', '8000'))
    host = os.environ.get('HOST', '0.0.0.0')
    server = ThreadingHTTPServer((host, port), Handler)
    print('Adarsh Boutique listening on port', port, flush=True)
    server.serve_forever()
