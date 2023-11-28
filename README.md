# Kassenbuch-Statistik-Konverter

Dieses Flask-Projekt ermöglicht das Hochladen von Kassenbuchdaten im Excel-Format und generiert daraus eine übersichtliche Statistik in Form einer Excel-Datei. Es identifiziert und gruppiert verschiedene Tickettypen (Tagestickets, Monatstickets, Jahrestickets, Verleih, Reparatur) und stellt diese Informationen in einer formatierten Excel-Datei bereit.

## Installation

Um dieses Projekt zu verwenden, stellen Sie sicher, dass Sie Python und Flask auf Ihrem System installiert haben. Folgen Sie diesen Schritten:

1. Klonen Sie dieses Repository auf Ihr lokales System.
2. Installieren Sie die erforderlichen Abhängigkeiten:

    ```bash
    pip install flask pandas openpyxl
    ```

## Verwendung

Starten Sie den Flask-Server mit dem folgenden Befehl:

    ```bash
    python app.py
    ```

Öffnen Sie einen Webbrowser und navigieren Sie zu `http://0.0.0.0:8098`. Laden Sie Ihre Kassenbuch-Excel-Datei hoch, und der Server wird eine formatierte Statistik-Excel-Datei generieren, die Sie herunterladen können.

## Funktionen

- Hochladen von Kassenbuchdaten im Excel-Format.
- Automatische Identifikation und Gruppierung von Tickettypen.
- Erstellung einer detaillierten Statistik als Excel-Datei.
- Speichern der generierten Datei sowohl zum Herunterladen als auch lokal im Verzeichnis 'Statistiken'.

## Autor

Erik Schauer - [do1ffe@darc.de](mailto:do1ffe@darc.de)
