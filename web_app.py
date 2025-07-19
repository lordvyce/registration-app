import http.server
import socketserver
import json
import os
from urllib.parse import parse_qs

PORT = 8000
DATA_FILE = 'appointments.json'

HTML_HEAD = """
<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>{title}</title>
<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css'>
</head>
<body class='bg-light'>
<div class='container my-4'>
<h1 class='mb-4'>{title}</h1>
"""

HTML_FOOT = """
</div>
</body>
</html>
"""

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return []
    return []

def save_data(data):
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def next_id(data):
    if not data:
        return 1
    return max(item.get('id', 0) for item in data) + 1

class Handler(http.server.SimpleHTTPRequestHandler):
    def render(self, title, body):
        html = HTML_HEAD.format(title=title) + body + HTML_FOOT
        self.send_response(200)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(html.encode('utf-8'))

    def do_GET(self):
        if self.path.startswith('/add'):
            self.show_add_form()
        elif self.path.startswith('/delete'):
            self.handle_delete()
        else:
            self.list_appointments()

    def do_POST(self):
        if self.path.startswith('/add'):
            length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(length).decode('utf-8')
            fields = parse_qs(post_data)
            data = load_data()
            new_entry = {
                'id': next_id(data),
                'patient_name': fields.get('patient_name', [''])[0],
                'procedure': fields.get('procedure', [''])[0],
                'phone_number': fields.get('phone_number', [''])[0],
                'email': fields.get('email', [''])[0],
                'appointment_date': fields.get('appointment_date', [''])[0],
                'appointment_time': fields.get('appointment_time', [''])[0]
            }
            data.append(new_entry)
            save_data(data)
            self.send_response(303)
            self.send_header('Location', '/')
            self.end_headers()
        else:
            self.send_error(404)

    def list_appointments(self):
        data = load_data()
        rows = ''.join(
            f"<tr><td>{d.get('patient_name')}</td><td>{d.get('procedure')}</td>"\
            f"<td>{d.get('appointment_date')} {d.get('appointment_time')}</td>"\
            f"<td><a href='/delete?id={d.get('id')}' class='btn btn-danger btn-sm'>Delete</a></td></tr>"
            for d in data
        ) or '<tr><td colspan="4">No appointments</td></tr>'
        body = f"""
        <a href='/add' class='btn btn-primary mb-3'>Add Appointment</a>
        <table class='table table-bordered table-striped'>
            <thead><tr><th>Patient</th><th>Procedure</th><th>Date &amp; Time</th><th>Actions</th></tr></thead>
            <tbody>{rows}</tbody>
        </table>
        """
        self.render('Appointments', body)

    def show_add_form(self):
        body = """
        <form method='post'>
            <div class='mb-3'>
                <label class='form-label'>Patient Name</label>
                <input name='patient_name' class='form-control' required>
            </div>
            <div class='mb-3'>
                <label class='form-label'>Procedure</label>
                <input name='procedure' class='form-control' required>
            </div>
            <div class='mb-3'>
                <label class='form-label'>Phone</label>
                <input name='phone_number' class='form-control'>
            </div>
            <div class='mb-3'>
                <label class='form-label'>Email</label>
                <input name='email' class='form-control'>
            </div>
            <div class='mb-3'>
                <label class='form-label'>Date</label>
                <input type='date' name='appointment_date' class='form-control' required>
            </div>
            <div class='mb-3'>
                <label class='form-label'>Time</label>
                <input type='time' name='appointment_time' class='form-control' required>
            </div>
            <button type='submit' class='btn btn-success'>Save</button>
            <a href='/' class='btn btn-secondary'>Cancel</a>
        </form>
        """
        self.render('Add Appointment', body)

    def handle_delete(self):
        query = self.path.partition('?')[2]
        params = parse_qs(query)
        try:
            del_id = int(params.get('id', [0])[0])
        except (ValueError, TypeError):
            del_id = 0
        data = load_data()
        data = [d for d in data if d.get('id') != del_id]
        save_data(data)
        self.send_response(303)
        self.send_header('Location', '/')
        self.end_headers()

if __name__ == '__main__':
    with socketserver.TCPServer(('', PORT), Handler) as httpd:
        print(f'Serving on port {PORT}...')
        httpd.serve_forever()
