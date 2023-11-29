from flask import Flask, request, send_file, render_template_string, redirect, url_for, session
import pandas as pd
import os
from datetime import datetime
import glob
import calendar

app = Flask(__name__)
app.secret_key = 'IhrSehrGeheimerSchlüssel'

VORGESEHENES_PASSWORT = 'IhrPasswort'

def render_header():
    return '<h1>Kassenbuch zu Statistik Konverter</h1>'

def render_footer():
    return '<p>&copy; Copyright by Erik Schauer, 2023, do1ffe@darc.de</p>'

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
          ''' + render_header() + '''
          <form action="/" method="post">
             Passwort: <input type="password" name="password" />
             <input type="submit" />
          </form>
          ''' + render_footer() + '''
       </body>
    </html>
    '''

@app.route('/upload', methods=['GET', 'POST'])
def upload_file_form():
    if 'authenticated' not in session:
        return redirect(url_for('password_form'))

    html_table_daily = ""
    html_table_monthly = ""
    current_date = datetime.now().strftime('%y%m%d')
    current_year = datetime.now().year
    filename = f"{current_date}-Kassenbuch-Statistik.xlsx"
    yearly_filename = f"{current_year}-Statistik.xlsx"
    csv_file_path = 'Statistiken/statistik.csv'

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

        if os.path.exists(csv_file_path):
            existing_data = pd.read_csv(csv_file_path)
            existing_data['Datum'] = pd.to_datetime(existing_data['Datum'], format='%d.%m.%Y')
        else:
            existing_data = pd.DataFrame()

        combined_data = pd.concat([existing_data, grouped_data.reset_index()])
        combined_data = combined_data.drop_duplicates(subset=['Datum'], keep='last')
        combined_data.to_csv(csv_file_path, index=False, date_format='%d.%m.%Y')

        os.makedirs('Statistiken', exist_ok=True)
        local_path = os.path.join('Statistiken', filename)

        with pd.ExcelWriter(local_path, engine='openpyxl') as writer:
            combined_data.to_excel(writer, index=False)

        html_table_daily = combined_data.to_html(classes='table table-striped', index=False)

        # Monatliche Statistik erstellen
        df['Monat'] = pd.to_datetime(df['Datum'], format='%d.%m.%Y').dt.to_period('M')
        monthly_data = df.groupby(['Monat', 'Tickettyp']).size().unstack(fill_value=0)
        monthly_data.index = monthly_data.index.strftime('%B')  # Monatsnamen anzeigen
        monthly_data.loc['Gesamt'] = monthly_data.sum()
        html_table_monthly = monthly_data.to_html(classes='table table-striped')

    header = render_header()
    footer = render_footer()

    return render_template_string('''
    <html>
       <body>
          ''' + header + '''
          <div style="display:flex;">
            <div style="margin-right: 50px;">
              <h2>Tägliche Statistik</h2>
              ''' + html_table_daily + '''
            </div>
            <div>
              <h2>Monatliche Statistik</h2>
              ''' + html_table_monthly + '''
            </div>
          </div>
          <br>
          <a href="/download/{{ filename }}">Aktuelle Excel-Datei herunterladen</a>
          <br><br>
          <a href="/download/{{ yearly_filename }}">Monatliche Statistik herunterladen</a>
          <br><br>
          <a href="/">Zurück zur Hauptseite</a>
          ''' + footer + '''
       </body>
    </html>
    ''', filename=filename, yearly_filename=yearly_filename)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join('Statistiken', filename), as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8098, debug=True)
