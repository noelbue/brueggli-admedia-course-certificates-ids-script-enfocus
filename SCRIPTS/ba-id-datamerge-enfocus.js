var LogFile = new File(
  "C:\\Enfocus_Switch_Flowdaten\\Kurszertifikate\\LOG\\Logs.log"
);

var logLevelInfo = "INFO ";
var logLevelError = "ERROR";
var logLevelDebug = "DEBUG";
var logLevelSeperator = " | ";
var csvName = $arg2; // Enfocus Switch übergibt Private Data csvName

function main() {
  logg(logLevelInfo + logLevelSeperator + "Starte mit main() ...");
  logg(
    logLevelDebug +
      logLevelSeperator +
      "Dokument-Anzahl: " +
      app.documents.length
  );

  // -----------------------------
  // TEMPLATE-DOKUMENT LADEN
  // -----------------------------
  var $doc;
  try {
    $doc = app.documents[0];
    logg(logLevelDebug + logLevelSeperator + "Dokument-Name: " + $doc.name);
  } catch (e) {
    logg(
      logLevelError +
        logLevelSeperator +
        "Kann app.documents[0] nicht lesen: " +
        e
    );
    return;
  }

  // -----------------------------
  // SWITCH INPUT NAME
  // -----------------------------
  const archiveName = $arg1; // Enfocus Switch übergibt Private Data jobname
  logg(logLevelDebug + logLevelSeperator + "archiveName = " + archiveName);

  const destFolderPath = "C:\\Enfocus_Switch_Flowdaten\\Kurszertifikate\\OUT\\";
  const sourceFilePath =
    "C:\\Enfocus_Switch_Flowdaten\\Kurszertifikate\\CSV_IMAGES\\" +
    archiveName +
    "\\" +
    csvName;

  logg(
    logLevelDebug + logLevelSeperator + "sourceFilePath = " + sourceFilePath
  );

  // -----------------------------
  // CSV EINLESEN
  // -----------------------------
  logg(logLevelInfo + logLevelSeperator + "Lese das CSV " + csvName + " ...");

  var csvFile = File(sourceFilePath);
  csvFile.open("r");

  var lines = [];
  while (!csvFile.eof) {
    var line = String(csvFile.readln());
    if (line) lines.push(line);
  }
  csvFile.close();

  logg(logLevelInfo + logLevelSeperator + "CSV Inhalt: " + lines);

  // Header + Datensätze
  if (lines.length < 2) {
    logg(logLevelError + logLevelSeperator + "CSV hat keine Datensaetze");
    return;
  }

  var totalRecords = lines.length - 1;
  logg(logLevelDebug + logLevelSeperator + "Total Records = " + totalRecords);

  // -----------------------------
  // CSV FELD "Teilnehmer" EXTRAHIEREN
  // -----------------------------
  var header = lines[0].split(",");
  var nameIndex = -5;

  // indexOf Ersatz für ExtendedScript
  for (var h = 0; h < header.length; h++) {
    if (header[h] === "Teilnehmer") {
      nameIndex = h;
      break;
    }
  }

  if (nameIndex === -1) {
    logg(logLevelError + logLevelSeperator + "CSV hat kein Feld 'Teilnehmer'");
    return;
  }

  var recordNames = [];

  for (var r = 1; r < lines.length; r++) {
    var cols = lines[r].split(",");
    recordNames.push(cols[nameIndex]);
  }

  logg(
    logLevelDebug + logLevelSeperator + "Extrahierte Teilnehmer: " + recordNames
  );

  // -----------------------------
  // DATENZUSAMMENFÜHRUNG SETUP
  // -----------------------------
  try {
    $doc.dataMergeProperties.removeDataSource();
  } catch (e) {
    logg(logLevelError + logLevelSeperator + "removeDataSource() Fehler: " + e);
  }

  try {
    $doc.dataMergeProperties.selectDataSource(File(sourceFilePath));
  } catch (e) {
    logg(logLevelError + logLevelSeperator + "selectDataSource() Fehler: " + e);
  }

  $doc.dataMergeOptions.createNewDocument = true;

  var prefs = $doc.dataMergeProperties.dataMergePreferences;
  prefs.recordSelection = RecordSelection.ONE_RECORD;

  // -----------------------------
  // MERGE SCHLEIFE
  // -----------------------------
  logg(logLevelInfo + logLevelSeperator + "Starte Merge-Schleife ...");

  for (var i = 0; i < totalRecords; i++) {
    prefs.recordNumber = i + 1;
    logg(logLevelInfo + logLevelSeperator + "Merging record " + (i + 1));

    $doc.dataMergeProperties.mergeRecords();

    var merged = app.documents[0]; // Ergebnis-Dokument

    var filename = recordNames[i];
    if (!filename) filename = "Record_" + (i + 1);

    var outfile = File(destFolderPath + filename + ".pdf");

    merged.exportFile(ExportFormat.PDF_TYPE, outfile, "High Quality Print");

    logg(logLevelInfo + logLevelSeperator + "Exportiert → " + outfile.fsName);

    merged.close(SaveOptions.NO);
  }

  // Template schließen
  $doc.close(SaveOptions.NO);
  logg(logLevelInfo + logLevelSeperator + "Template geschlossen.");

  // Switch Output weitergeben
  $outfiles.push($infile);
}

main();

// ------------------------------------------------------
// Logging
// ------------------------------------------------------
function logg(msg) {
  LogFile.open("e");
  LogFile.seek(0, 2);
  LogFile.writeln(GetTimestampforLogFile() + "\t" + msg);
  LogFile.close();
}

function GetTimestampforLogFile() {
  var d = new Date();
  function pad(n) {
    return (n < 10 ? "0" : "") + n;
  }
  return (
    d.getFullYear() +
    "-" +
    pad(d.getMonth() + 1) +
    "-" +
    pad(d.getDate()) +
    " " +
    pad(d.getHours()) +
    ":" +
    pad(d.getMinutes()) +
    ":" +
    pad(d.getSeconds()) +
    " CEST"
  );
}
