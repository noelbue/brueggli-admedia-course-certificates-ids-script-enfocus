# InDesign Data Merge Script für Kurszertifikate

## Übersicht

Dieses InDesign Server Skript automatisiert die Erstellung von personalisierten PDF-Zertifikaten mittels InDesign Data Merge. Das Skript wird über Enfocus Switch und InDesign Server ausgeführt und verarbeitet CSV-Daten mit zugehörigen Bildern.

## Funktionsweise

Das Skript führt folgende Schritte aus:

1. **Template laden**: Öffnet das InDesign-Template-Dokument
2. **CSV einlesen**: Liest die CSV-Datei mit Teilnehmerdaten
3. **Feldextraktion**: Extrahiert das `name`-Feld für die PDF-Benennung
4. **Data Merge Setup**: Konfiguriert die Datenzusammenführung
5. **PDF-Generierung**: Erstellt für jeden Datensatz ein individuelles PDF
6. **Export**: Speichert PDFs mit dem Namen aus der CSV-Spalte `name`

## Parameter

Das Skript erwartet zwei Parameter von Enfocus Switch:

- **`$arg1`**: `archiveName` - Name des entpackten Archiv-Ordners
- **`$arg2`**: `csvName` - Name der CSV-Datei (z.B. `my.csv`)

## Verzeichnisstruktur

```
C:\Enfocus_Switch_Flowdaten\Kurszertifikate\
├── CSV_IMAGES\
│   └── [archiveName]\
│       ├── my.csv
│       ├── A.png
│       ├── B.png
│       └── C.png
├── IN\
├── LOG\
│   └── Logs.log
├── OUT\
│   ├── Anna.pdf
│   ├── Ben.pdf
│   └── Carla.pdf
├── SCRIPTS\
│   └── ba-id-datamerge-enfocus.js
├── TEMPLATES\
│   └── Testfile_Noel.indd
└── TEST_DATA\
    └── Archive.zip
```

## CSV-Struktur

Die CSV-Datei muss mindestens folgende Spalten enthalten:

```csv
name,@image_url
Anna,A.png
Ben,B.png
Carla,C.png
```

- **`name`**: Wird als PDF-Dateiname verwendet (PFLICHTFELD)
- **`image_url`**: Referenz zum Bild (muss im InDesign-Template als Data Merge Feld eingebunden sein)

## Logging

Alle Aktivitäten werden in `C:\Enfocus_Switch_Flowdaten\Kurszertifikate\LOG\Logs.log` protokolliert.

**Log-Level:**

- `INFO` - Allgemeine Informationen zum Workflow
- `DEBUG` - Detaillierte technische Informationen
- `ERROR` - Fehlermeldungen

### Beispiel Log-Ausgabe

```
2025-11-18 13:37:55 CEST	INFO  | Starte mit main() ...
2025-11-18 13:37:55 CEST	DEBUG | Dokument-Anzahl: 1
2025-11-18 13:37:55 CEST	DEBUG | Dokument-Name: _00O21_Testfile_Noel.indd
2025-11-18 13:37:55 CEST	DEBUG | archiveName = Archive
2025-11-18 13:37:55 CEST	DEBUG | sourceFilePath = C:\Enfocus_Switch_Flowdaten\Kurszertifikate\CSV_IMAGES\Archive\my.csv
2025-11-18 13:37:55 CEST	INFO  | Lese das CSV my.csv ...
2025-11-18 13:37:55 CEST	INFO  | CSV Inhalt: name,@image_url,Anna,A.png,Ben,B.png,Carla,C.png
2025-11-18 13:37:55 CEST	DEBUG | Total Records = 3
2025-11-18 13:37:55 CEST	DEBUG | Extrahierte Namen: Anna,Ben,Carla
2025-11-18 13:37:55 CEST	INFO  | Starte Merge-Schleife ...
2025-11-18 13:37:55 CEST	INFO  | Merging record 1
2025-11-18 13:37:55 CEST	INFO  | Merging record 2
2025-11-18 13:37:55 CEST	INFO  | Merging record 3
2025-11-18 13:37:56 CEST	INFO  | Template geschlossen.
```

## PDF-Export-Einstellungen

PDFs werden mit dem Preset **"High Quality Print"** exportiert.
