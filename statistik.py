from flask import Flask, request, send_file
import pandas as pd
import io
import os
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def upload_file_form():
    current_year = datetime.now().year
    if current_year == 2023:
        copyright_year = "2023"
    else:
        copyright_year = f"2023-{current_year}"

    return f'''
    <html>
       <body>
          <h1>Konvertierung Kassenbuch in Statistik</h1>
          <p>Bitte laden Sie das aktuelle Kassenbuch im Excel-Format hoch.</p>
          <form action="/upload" method="post" enctype="multipart/form-data">
             <input type="file" name="file" />
             <input type="submit" />
          </form>
          <p>&copy; Copyright by Erik Schauer, {copyright_year}, do1ffe@darc.de</p>
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

        # Extrahieren der ersten zwei Buchstaben der Tickettypen
        df['Tickettyp'] = df['Quittung Text'].str[:2]
        ticket_types = ['TT', 'MT', 'JT', 'V', 'R']
        df = df[df['Tickettyp'].isin(ticket_types)]

        # Gruppierung der Daten
        grouped_data = df.groupby([df['Datum'].dt.strftime('%d.%m.%Y'), 'Tickettyp']).size().unstack(fill_value=0)

        # Anordnung der Spalten in der gewünschten Reihenfolge
        desired_order = ['TT', 'MT', 'JT', 'V', 'R']
        grouped_data = grouped_data[desired_order]

        # Gesamtsumme am Ende hinzufügen
        grouped_data.loc['Total'] = grouped_data.sum()

        current_date = datetime.now().strftime('%y%m%d')
        filename = f"{current_date}-Kassenbuch-Statistik.xlsx"

        os.makedirs('Statistiken', exist_ok=True)
        local_path = os.path.join('Statistiken', filename)
        with pd.ExcelWriter(local_path, engine='openpyxl') as writer:
            grouped_data.to_excel(writer)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            grouped_data.to_excel(writer)
        output.seek(0)

        return send_file(output, attachment_filename=filename, as_attachment=True)

if __name__ == '__main__':
   app.run(host='0.0.0.0', port=8098, debug=True)
