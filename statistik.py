from flask import Flask, request, send_file, render_template_string
import pandas as pd
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

        # Einlesen der Excel-Datei ab Zeile 4
        df = pd.read_excel(f, skiprows=3)

        # Datumskonvertierung
        df['Datum'] = pd.to_datetime(df['Datum']).dt.strftime('%d.%m.%Y')

        df['Tickettyp'] = df['Quittung Text'].str[:2]
        ticket_types = ['TT', 'MT', 'JT', 'V', 'R']
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
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            for col in worksheet.columns:
                max_length = max(len(str(cell.value)) for cell in col)
                worksheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

        html_table = combined_data.to_html(classes='table table-striped', index=False)

        return render_template_string(f'''
        <html>
            <head>
                <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
                <style>
                    .table {{ width: auto !important; margin: auto; }}
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
