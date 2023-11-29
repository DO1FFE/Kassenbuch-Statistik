from flask import Flask, request, send_file, render_template_string
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl.utils import get_column_letter

app = Flask(__name__)

@app.route('/')
def upload_file_form():
    current_year = datetime.now().year
    copyright_year = "2023" if current_year == 2023 else f"2023-{current_year}"

    return '''
    <html>
       <body>
          <h1>Konvertierung Kassenbuch in Statistik</h1>
          <p>Bitte laden Sie das aktuelle Kassenbuch im Excel-Format hoch.</p>
          <form action="/upload" method="post" enctype="multipart/form-data">
             <input type="file" name="file" />
             <input type="submit" />
          </form>
          <p>&copy; Copyright by Erik Schauer, ''' + copyright_year + ''', do1ffe@darc.de</p>
       </body>
    </html>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        if not f:
            return 'Keine Datei hochgeladen', 400

        df = pd.read_excel(f)
        df['Datum'] = pd.to_datetime(df['Datum'], errors='coerce')
        df = df.dropna(subset=['Datum'])
        df['Tickettyp'] = df['Quittung Text'].str[:2]
        ticket_types = ['TT', 'MT', 'JT', 'V', 'R']
        df = df[df['Tickettyp'].isin(ticket_types)]

        grouped_data = df.groupby([df['Datum'].dt.strftime('%d.%m.%Y'), 'Tickettyp']).size().unstack(fill_value=0)
        for ticket_type in ['TT', 'MT', 'JT', 'V', 'R']:
            if ticket_type not in grouped_data:
                grouped_data[ticket_type] = 0
        grouped_data = grouped_data[['TT', 'MT', 'JT', 'V', 'R']]
        grouped_data = grouped_data.rename(index={'Total': 'Gesamt'})

        current_date = datetime.now().strftime('%y%m%d')
        filename = f"{current_date}-Kassenbuch-Statistik.xlsx"
        os.makedirs('Statistiken', exist_ok=True)
        local_path = os.path.join('Statistiken', filename)

        with pd.ExcelWriter(local_path, engine='openpyxl') as writer:
            grouped_data.to_excel(writer)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            for column_cells in worksheet.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length

        html_table = grouped_data.to_html(classes='table table-striped')

        return render_template_string(f'''
        <html>
            <head>
                <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
                <style>
                    .table { width: auto !important; margin: auto; }
                </style>
            </head>
            <body>
                <h2>Statistik-Tabelle</h2>
                {html_table}
                <br>
                <a href="/download/{filename}">Excel-Datei herunterladen</a>
                <br><br>
                <a href="/">Zurück zur Hauptseite</a>
            </body>
        </html>
        ''')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join('Statistiken', filename), as_attachment=True)

if __name__ == '__main__':
   app.run(host='0.0.0.0', port=8098, debug=True)
