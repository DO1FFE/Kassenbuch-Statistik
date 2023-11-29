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

    html_table_daily = ""
    html_table_monthly = ""
    current_year = datetime.now().year
    filename = f"{current_year}-Kassenbuch-Statistik.xlsx"
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
            existing_data = pd.read_csv(csv_file_path, index_col='Datum')
        else:
            existing_data = pd.DataFrame()

        combined_data = pd.concat([existing_data, grouped_data])
        combined_data = combined_data[~combined_data.index.duplicated(keep='last')]
        combined_data.to_csv(csv_file_path)

        html_table_daily = combined_data.to_html(classes='table table-striped')

        # Monatliche Statistik erstellen
        is_date = pd.to_datetime(combined_data.index, errors='coerce').notna()
        combined_data.loc[is_date, 'Monat'] = pd.to_datetime(combined_data[is_date].index).strftime('%B')
        monthly_data = combined_data[is_date].groupby('Monat').sum()
        html_table_monthly = monthly_data.to_html(classes='table table-striped')

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
          <a href="/download/{{ filename }}">Jährliche Statistik herunterladen</a>
          <br><br>
          <a href="/clear_statistics">Statistikdaten löschen</a>
          <br><br>
          <a href="/">Zurück zur Hauptseite</a>
          ''' + footer + '''
       </body>
    </html>
    ''', filename=filename)

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
