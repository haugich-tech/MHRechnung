// =============================================================================
// MHRECHNUNG — Empfänger für die Datensicherung (Google Apps Script)
// -----------------------------------------------------------------------------
// Nimmt Datenbank- und PDF-Sicherungen von MHRechnung (DriveBackup.vb) entgegen und
// legt sie in Google Drive unter  Meine Ablage/Sicherung/MHRechnung/<Ordner>/...  ab.
// <Ordner> wird von MHRechnung automatisch aus der Steuernummer gebildet, nicht frei
// eingetragen - damit landen zwei Betriebe mit unterschiedlicher Steuernummer nie
// versehentlich im selben Ordner, auch nicht bei einer separaten Programminstanz.
//
// Getrennt vom Herdenmanager-Empfänger: eigenes Skript, eigener Ordnerpfad, eigenes
// Geheimwort. Beide Programme sichern unabhängig voneinander.
//
// Zwei Sicherungsarten mit unterschiedlicher Aufbewahrung:
//   - Datenbank (Aktion "sichern_db"): rohe .sqlite-Datei, nur die neuesten
//     BEHALTEN_DB Versionen bleiben, ältere wandern in den Papierkorb.
//   - Rechnungs-PDFs (Aktion "sichern_pdf"): einzeln, für immer aufbewahrt, keine
//     Rotation - MHRechnung schickt ohnehin nur neu hinzugekommene Rechnungen.
//
// Schutz vor Vermischung zweier Datenbanken im selben Ordner: jede Anfrage schickt
// eine unsichtbare Kennung (GUID) mit. Beim ersten Schreiben in einen Ordner wird sie
// dort in einer Markierungsdatei "_kennung.txt" hinterlegt; bei jeder weiteren Anfrage
// wird geprüft, ob die mitgeschickte Kennung dazu passt. Falls nicht: Ablehnung statt
// stillem Vermischen (siehe pruefeKennung).
//
// Einrichtung (einmalig):
//   1. script.google.com → Neues Projekt, diesen Text einfügen, Geheimwort eintragen
//   2. Bereitstellen → Neue Bereitstellung → Typ „Web-App“
//      Ausführen als: Ich   ·   Zugriff: Jeder
//   3. Die Web-App-URL + das Geheimwort in MHRechnung unter Einstellungen →
//      "7. Google-Drive-Sicherung" eintragen
//
// Nach jeder Änderung am Code: Bereitstellen → Bereitstellungen verwalten →
// Bearbeiten (Stift) → Version „Neue Version“ → Bereitstellen. Sonst läuft unter
// der alten Adresse weiter der alte Code.
//
// Diese Datei liegt versioniert im MHRechnung-Repository. Dort steht absichtlich ein
// Platzhalter statt des Geheimworts, weil die Datei über HTTP ausgeliefert wird. Das
// Geheimwort steht nur hier im Skript und in den MHRechnung-Einstellungen.
// =============================================================================

const GEHEIMWORT = 'HIER_GEHEIMWORT_EINTRAGEN';
const ORDNER_BASIS = ['Sicherung', 'MHRechnung'];   // ab „Meine Ablage“
const BEHALTEN_DB  = 10;
const PRAEFIX_DB    = 'mhrechnung-db-';
const KENNUNG_DATEI = '_kennung.txt';

function doPost(e) {
  try {
    const anfrage = JSON.parse(e.postData.contents);
    if (GEHEIMWORT.indexOf('HIER_') === 0 || !anfrage || anfrage.geheim !== GEHEIMWORT) {
      return antwort({ ok: false, fehler: 'Geheimwort stimmt nicht' });
    }

    const betriebOrdnerName = String(anfrage.ordner || '');
    if (!/^[A-Za-z0-9-]+$/.test(betriebOrdnerName)) {
      return antwort({ ok: false, fehler: 'Ungültiger Ordnername: ' + betriebOrdnerName });
    }
    const betriebOrdner = holeOrdner(ORDNER_BASIS.concat([betriebOrdnerName]));

    const kennungFehler = pruefeKennung(betriebOrdner, String(anfrage.kennung || ''));
    if (kennungFehler) {
      return antwort({ ok: false, fehler: kennungFehler });
    }

    if (anfrage.aktion === 'test') {
      return antwort({ ok: true, ordner: ORDNER_BASIS.concat([betriebOrdnerName]).join('/') });
    }
    if (anfrage.aktion === 'sichern_db') {
      return sichereDatenbank(betriebOrdner, anfrage);
    }
    if (anfrage.aktion === 'sichern_pdf') {
      return sicherePdf(betriebOrdner, anfrage);
    }
    return antwort({ ok: false, fehler: 'Unbekannte Aktion' });
  } catch (err) {
    return antwort({ ok: false, fehler: String(err && err.message ? err.message : err) });
  }
}

// Aufruf der Adresse im Browser: zeigt nur, dass der Empfänger läuft
function doGet() {
  return ContentService.createTextOutput('MHRechnung-Sicherung: Empfänger ist bereit.');
}

function sichereDatenbank(betriebOrdner, anfrage) {
  const name = String(anfrage.name || '');
  if (!new RegExp('^' + PRAEFIX_DB + '\\d{4}-\\d{2}-\\d{2}_\\d{4}\\.sqlite$').test(name)) {
    return antwort({ ok: false, fehler: 'Ungültiger Dateiname: ' + name });
  }

  const bytes = pruefeUndDekodiere(anfrage);
  if (bytes.fehler) return antwort({ ok: false, fehler: bytes.fehler });

  const ordner = holeUnterordner(betriebOrdner, 'Datenbank');
  ersetzeGleichnamige(ordner, name);
  const datei = ordner.createFile(Utilities.newBlob(bytes.werte, 'application/x-sqlite3', name));

  // Aufräumen: die neuesten BEHALTEN_DB bleiben. Die Namen enthalten Datum und Uhrzeit,
  // alphabetisch absteigend ist deshalb auch zeitlich absteigend.
  const alle = dateienMitPraefix(ordner, PRAEFIX_DB).sort((a, b) => (a.getName() < b.getName() ? 1 : -1));
  let geloescht = 0;
  for (const f of alle.slice(BEHALTEN_DB)) {
    f.setTrashed(true);
    geloescht++;
  }

  return antwort({ ok: true, datei: datei.getName(), groesse: datei.getSize(), anzahl: Math.min(alle.length, BEHALTEN_DB), geloescht: geloescht });
}

function sicherePdf(betriebOrdner, anfrage) {
  const name = String(anfrage.name || '');
  if (!/^\d{7}(_zugf)?\.pdf$/.test(name)) {
    return antwort({ ok: false, fehler: 'Ungültiger Dateiname: ' + name });
  }

  const bytes = pruefeUndDekodiere(anfrage);
  if (bytes.fehler) return antwort({ ok: false, fehler: bytes.fehler });

  // Keine Rotation: Rechnungen bleiben für immer. Ein gleichnamiger, erneuter Upload
  // (z.B. nach einem abgebrochenen vorherigen Versuch) ersetzt lediglich die Datei,
  // statt einen Fehler zu werfen - MHRechnung schickt nie eine bereits geänderte
  // Rechnung erneut (Rechnungs-PDFs werden nach dem Erstellen nie mehr verändert).
  const ordner = holeUnterordner(betriebOrdner, 'Rechnungen');
  ersetzeGleichnamige(ordner, name);
  const datei = ordner.createFile(Utilities.newBlob(bytes.werte, 'application/pdf', name));

  return antwort({ ok: true, datei: datei.getName(), groesse: datei.getSize() });
}

// Prüft Größe und Prüfsumme (vollständig angekommen?) und liefert die dekodierten Bytes.
function pruefeUndDekodiere(anfrage) {
  const bytes = Utilities.base64Decode(String(anfrage.daten || ''));
  if (bytes.length !== Number(anfrage.groesse) || md5Hex(bytes) !== String(anfrage.md5 || '').toLowerCase()) {
    return { fehler: 'Datei kam unvollständig an – bitte erneut sichern' };
  }
  return { werte: bytes };
}

// Schutz vor Vermischung zweier Datenbanken im selben Ordner: legt beim ersten
// Schreiben/Testen die Kennung fest, lehnt bei einer abweichenden Kennung ab.
function pruefeKennung(betriebOrdner, kennung) {
  if (!kennung) return 'Keine Kennung mitgeschickt';

  const treffer = betriebOrdner.getFilesByName(KENNUNG_DATEI);
  if (treffer.hasNext()) {
    const bestehend = treffer.next().getBlob().getDataAsString().trim();
    if (bestehend !== kennung) {
      return 'Dieser Ordner gehört bereits zu einer anderen Datenbank (Kennung stimmt nicht überein)';
    }
    return null;
  }

  betriebOrdner.createFile(KENNUNG_DATEI, kennung, MimeType.PLAIN_TEXT);
  return null;
}

function ersetzeGleichnamige(ordner, name) {
  const gleich = ordner.getFilesByName(name);
  while (gleich.hasNext()) {
    const f = gleich.next();
    if (!f.isTrashed()) f.setTrashed(true);
  }
}

function dateienMitPraefix(ordner, praefix) {
  const liste = [];
  const it = ordner.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed() && f.getName().indexOf(praefix) === 0) liste.push(f);
  }
  return liste;
}

// Holt einen Unterordner (Datenbank/Rechnungen) im Betriebsordner; legt ihn bei Bedarf an.
function holeUnterordner(basisOrdner, name) {
  const treffer = [];
  const it = basisOrdner.getFoldersByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) treffer.push(f);
  }
  if (treffer.length > 1) {
    throw new Error('Es gibt mehrere Ordner „' + name + '“ – bitte einen davon umbenennen');
  }
  return treffer.length === 1 ? treffer[0] : basisOrdner.createFolder(name);
}

// Ordner-Pfad (Liste von Namen ab „Meine Ablage“) holen; fehlende Ebenen werden angelegt.
function holeOrdner(pfadTeile) {
  let ordner = DriveApp.getRootFolder();
  for (const teil of pfadTeile) {
    ordner = holeUnterordner(ordner, teil);
  }
  return ordner;
}

function md5Hex(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');
}

function antwort(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Zum Ausprobieren im Editor (▶ Ausführen): legt nichts an außer fehlenden Ordnern und
// zeigt im Protokoll, wie viele Datenbank- bzw. PDF-Sicherungen im Testordner liegen.
function testeOrdner() {
  const ordner = holeOrdner(ORDNER_BASIS.concat(['TEST']));
  const db = holeUnterordner(ordner, 'Datenbank');
  const pdf = holeUnterordner(ordner, 'Rechnungen');
  Logger.log('Datenbank: ' + dateienMitPraefix(db, PRAEFIX_DB).length + ' Sicherungen, PDFs: ' + pdf.getFiles().hasNext());
}
