from flask import Flask, request, send_file, render_template_string, redirect, url_for, session
import pandas as pd
import os
from datetime import datetime
import glob

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

    current_year = datetime.now().year
    csv_file_path = f'Statistiken/{current_year}-statistik.csv'

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
            existing_data['Datum'] = pd.to_datetime(existing_data['Datum'], errors='coerce').dt.strftime('%d.%m.%Y')
            existing_data.set_index('Datum', inplace=True)
        else:
            existing_data = pd.DataFrame()

        combined_data = pd.concat([existing_data, grouped_data])
        combined_data = combined_data[~combined_data.index.duplicated(keep='last')]
        combined_data.reset_index().to_csv(csv_file_path, index=False, date_format='%d.%m.%Y')

    # Tägliche und monatliche Statistik anzeigen
    if os.path.exists(csv_file_path):
        existing_data = pd.read_csv(csv_file_path)
        existing_data['Datum'] = pd.to_datetime(existing_data['Datum'], errors='coerce').dt.strftime('%d.%m.%Y')
        existing_data.set_index('Datum', inplace=True)
        html_table_daily = existing_data.to_html(classes='table table-striped')

        is_date = pd.to_datetime(existing_data.index, errors='coerce').notna()
        existing_data.loc[is_date, 'Monat'] = pd.to_datetime(existing_data[is_date].index).strftime('%B')
        monthly_data = existing_data[is_date].groupby('Monat').sum()
        html_table_monthly = monthly_data.to_html(classes='table table-striped')
    else:
        html_table_daily = "<p>Keine täglichen Daten vorhanden.</p>"
        html_table_monthly = "<p>Keine monatlichen Daten vorhanden.</p>"

    header = render_header()
    footer = render_footer()

    return render_template_string('''
    <html>
       <body>
          ''' + header + '''
          <form action="/upload" method="post" enctype="multipart/form-data">
             <input type="file" name="file" />
             <input type="submit" value="Kassenbuch hochladen"/>
          </form>
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
          <a href="/generate_excel">Excel-Datei generieren</a>
          <br>
          <a href="/list_excel_files">Verfügbare Excel-Dateien anzeigen</a>
          <br>
          <a href="/clear_statistics">Statistikdaten löschen</a>
          <br><br>
          <a href="/">Zurück zur Hauptseite</a>
          ''' + footer + '''
       </body>
    </html>
    ''')

@app.route('/generate_excel', methods=['GET', 'POST'])
def generate_excel():
    if 'authenticated' not in session:
        return redirect(url_for('password_form'))

    current_year = datetime.now().year
    csv_file_path = f'Statistiken/{current_year}-statistik.csv'
    excel_file_path = f'Statistiken/{current_year}-statistik.xlsx'

    if os.path.exists(csv_file_path):
        df = pd.read_csv(csv_file_path)
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce').dt.strftime('%d.%m.%Y')
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

    return redirect(url_for('upload_file_form'))

@app.route('/list_excel_files')
def list_excel_files():
    if 'authenticated' not in session:
        return redirect(url_for('password_form'))

    files = glob.glob('Statistiken/*.xlsx')
    file_list = [os.path.basename(file) for file in files]

    header = render_header()
    footer = render_footer()

    return render_template_string('''
    <html>
       <body>
          ''' + header + '''
          <h2>Verfügbare Excel-Dateien:</h2>
          <ul>
            {% for file in file_list %}
              <li><a href="/download/{{ file }}">{{ file }}</a></li>
            {% endfor %}
          </ul>
          <br>
          <a href="/upload">Zurück zur Upload-Seite</a>
          ''' + footer + '''
       </body>
    </html>
    ''', file_list=file_list)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join('Statistiken', filename), as_attachment=True)

@app.route('/clear_statistics', methods=['GET', 'POST'])
def clear_statistics():
    if 'authenticated' not in session:
        return redirect(url_for('password_form'))

    if request.method == 'POST':
        for f in glob.glob('Statistiken/*'):
            os.remove(f)
        return redirect(url_for('upload_file_form'))

    footer = render_footer()

    return render_template_string('''
    <html>
       <body>
          ''' + render_header() + '''
          <p>Sind Sie sicher, dass Sie alle Statistikdateien löschen möchten?</p>
          <form action="/clear_statistics" method="post">
             <input type="submit" value="Statistiken löschen"/>
          </form>
          <br>
          <a href="/">Zurück zur Hauptseite</a>
          ''' + render_footer() + '''
       </body>
    </html>
    ''')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8098, debug=True)
