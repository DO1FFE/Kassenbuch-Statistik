from flask import Flask, request, send_file, render_template_string, redirect, url_for, session
import pandas as pd
import os
from datetime import datetime
import glob

app = Flask(__name__)
app.secret_key = 'IhrSehrGeheimerSchlüssel'

VORGESEHENES_PASSWORT = 'IhrPasswort'

@app.route('/', methods=['GET', 'POST'])
def password_form():
    if request.method == 'POST':
        password = request.form.get('password')
        if password == VORGESEHENES_PASSWORT:
            session['authenticated'] = True
            return redirect(url_for('upload_file_form'))
        else:
            return '<p>Falsches Passwort. Bitte versuchen Sie es erneut.</p>' + render_password_form()

    return render_password_form()

def render_password_form():
    return '''
    <html>
       <body>
          <h1>Passwort eingeben</h1>
          <form action="/" method="post">
             Passwort: <input type="password" name="password" />
             <input type="submit" />
          </form>
       </body>
    </html>
    '''

@app.route('/upload', methods=['GET', 'POST'])
def upload_file_form():
    if 'authenticated' not in session:
        return redirect(url_for('password_form'))

    if request.method == 'POST':
        f = request.files['file']
        if not f:
            return 'Keine Datei hochgeladen', 400

        df = pd.read_excel(f, skiprows=3)
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0]).dt.strftime('%d.%m.%Y')
        df.rename(columns={df.columns[0]: 'Datum', df.columns[7]: 'Tickettyp'}, inplace=True)

        ticket_types = ['TT', 'MT', 'JT', 'V', 'R']
        df['Tickettyp'] = df['Tickettyp'].str[:2]
        df = df[df['Tickettyp'].isin(ticket_types)]
        grouped_data = df.groupby(['Datum', 'Tickettyp']).size().unstack(fill_value=0)

        for ticket_type in ticket_types:
            if ticket_type not in grouped_data:
                grouped_data[ticket_type] = 0

        grouped_data = grouped_data[ticket_types]
        grouped_data.loc['Gesamt'] = grouped_data.sum()

        csv_file_path = 'Statistiken/statistik.csv'
        if os.path.exists(csv_file_path):
            existing_data = pd.read_csv(csv_file_path)
        else:
            existing_data = pd.DataFrame()

        combined_data = pd.concat([existing_data, grouped_data.reset_index()])
        combined_data = combined_data.drop_duplicates(subset=['Datum'], keep='last')
        combined_data.to_csv(csv_file_path, index=False, date_format='%d.%m.%Y')

        current_date = datetime.now().strftime('%y%m%d')
        filename = f"{current_date}-Kassenbuch-Statistik.xlsx"
        os.makedirs('Statistiken', exist_ok=True)
        local_path = os.path.join('Statistiken', filename)

        with pd.ExcelWriter(local_path, engine='openpyxl') as writer:
            combined_data.to_excel(writer, index=False)

    # Pfad zur neuesten Statistik-Datei finden
    list_of_files = glob.glob('Statistiken/*.xlsx') 
    latest_file = max(list_of_files, key=os.path.getctime) if list_of_files else None
    latest_file_name = os.path.basename(latest_file) if latest_file else "Keine Statistik verfügbar"

    return render_template_string('''
    <html>
       <body>
          <h1>Kassenbuch hochladen</h1>
          <form action="/upload" method="post" enctype="multipart/form-data">
             <input type="file" name="file" />
             <input type="submit" />
          </form>
          <br>
          <a href="/download/{{ latest_file_name }}">{{ latest_file_name }}</a>
       </body>
    </html>
    ''', latest_file_name=latest_file_name)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join('Statistiken', filename), as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8098, debug=True)
