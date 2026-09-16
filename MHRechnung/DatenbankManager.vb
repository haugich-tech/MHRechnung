Imports System.Data.SQLite
Imports System.IO

Public Class DatenbankManager
    ' Absoluter Pfad zur .sqlite-Datei. Wird beim Programmstart von Form1 gesetzt
    ' (siehe DbKonfiguration) - der Speicherort ist frei wählbar, deshalb kein Const mehr.
    Public Shared Property DatenbankPfad As String = ""

    Private Shared ReadOnly Property ConnectionString As String
        Get
            Return "Data Source=" & DatenbankPfad & ";Version=3;"
        End Get
    End Property

    Public Shared Sub InitialisiereDatenbank()
        If String.IsNullOrWhiteSpace(DatenbankPfad) Then
            Throw New Exception("Kein Datenbank-Pfad gesetzt (DatenbankManager.DatenbankPfad ist leer).")
        End If

        ' 1. Ordner und Datei erstellen, falls sie komplett fehlen
        Dim ordner As String = Path.GetDirectoryName(DatenbankPfad)
        If Not String.IsNullOrWhiteSpace(ordner) AndAlso Not Directory.Exists(ordner) Then
            Directory.CreateDirectory(ordner)
        End If
        If Not File.Exists(DatenbankPfad) Then
            SQLiteConnection.CreateFile(DatenbankPfad)
        End If

        ' 2. Verbindung öffnen und Tabellen IMMER prüfen/erstellen
        ' parseViaFramework:=True - wichtig bei Netzwerkpfaden (\\Server\Freigabe\...):
        ' Ohne diesen Parameter parst SQLite den Pfad selbst und verschluckt sich an
        ' UNC-Pfaden ("unable to open database file"). Mit True übernimmt .NET die
        ' Pfadauflösung, die UNC-Pfade korrekt versteht.
        Using conn As New SQLiteConnection(ConnectionString, True)
            conn.Open()
            Dim sqlCmd As New SQLiteCommand(conn)

            ' IF NOT EXISTS sorgt dafür, dass SQLite nur fehlende Tabellen anlegt
            ' und bestehende Tabellen (mit deinen echten Daten) in Ruhe lässt.
            ' mwst_satz ist REAL statt INTEGER, weil die Durchschnittssätze nach §24 UStG
            ' (z.B. 7,8% / 5,5%) keine ganzen Zahlen sind. Keine IBAN/BIC/Mandat/SEPA-Typ
            ' mehr bei den Mitgliedern - es wird nur noch per Überweisung bezahlt.
            sqlCmd.CommandText = "
                CREATE TABLE IF NOT EXISTS mitglieder (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mitgliedsnummer TEXT UNIQUE,
                    name TEXT NOT NULL,
                    strasse TEXT,
                    plz TEXT,
                    ort TEXT,
                    land_code TEXT DEFAULT 'DE',
                    steuernummer TEXT,
                    betriebsnummer TEXT,
                    email TEXT,
                    versandart TEXT
                );

                CREATE TABLE IF NOT EXISTS artikel (id INTEGER PRIMARY KEY AUTOINCREMENT, artikelnummer TEXT, bezeichnung TEXT, einzelpreis_netto REAL, mwst_satz REAL, einheit TEXT);
                CREATE TABLE IF NOT EXISTS rechnungen (id INTEGER PRIMARY KEY AUTOINCREMENT, rechnungsnummer TEXT, datum TEXT, lieferdatum TEXT, mitglied_id INTEGER, status TEXT);
                CREATE TABLE IF NOT EXISTS rechnungspositionen (id INTEGER PRIMARY KEY AUTOINCREMENT, rechnung_id INTEGER, artikel_bezeichnung TEXT, anzahl REAL, einzelpreis REAL, mwst_satz REAL);
                CREATE TABLE IF NOT EXISTS einstellungen (schluessel TEXT PRIMARY KEY, wert TEXT);
            "
            sqlCmd.ExecuteNonQuery()

            ' 2b. Migration für bestehende Datenbanken: CREATE TABLE IF NOT EXISTS legt bei schon
            ' vorhandenen Tabellen keine neuen Spalten nach. Für den Netto/Brutto-Umbau werden
            ' hier gezielt fehlende Spalten nachgerüstet, ohne bestehende Daten anzufassen.
            ' - rechnungen.preisart: legt pro Rechnung fest, ob die gedruckte Einzelpreis-Spalte
            '   Netto oder Brutto zeigt.
            ' - rechnungspositionen.einzelpreis_brutto: der exakt eingegebene Bruttopreis pro
            '   Position (falls Brutto eingetippt wurde) - wird beim Drucken unverändert
            '   übernommen, damit keine Rundungs-Rückrechnung krumme Zahlen erzeugt.
            ' - mitglieder.standard_preisart: Vorbelegung für neue Rechnungen dieses Kunden.
            ' - artikel.einzelpreis_brutto: exakter Standard-Bruttopreis des Artikels (analog).
            FuegeSpalteHinzuFallsFehlt(conn, "rechnungen", "preisart", "TEXT DEFAULT 'Netto'")
            FuegeSpalteHinzuFallsFehlt(conn, "rechnungspositionen", "einzelpreis_brutto", "REAL")
            FuegeSpalteHinzuFallsFehlt(conn, "mitglieder", "standard_preisart", "TEXT DEFAULT 'Netto'")
            FuegeSpalteHinzuFallsFehlt(conn, "artikel", "einzelpreis_brutto", "REAL")

            ' 3. Grundeinstellungen NUR einfügen, wenn die Einstellungen-Tabelle noch absolut leer ist.
            ' Die Firmen-/Bankdaten sind hier einmalig aus deiner hochgeladenen Beispielrechnung
            ' vorbelegt, damit du nicht bei null anfangen musst - änderbar bleibt alles jederzeit
            ' im Tab "Einstellungen". Die beiden MwSt-Sätze, die Startnummer und die Zahlungstexte
            ' bekommen ebenfalls einen sinnvollen Standardwert, damit das Programm sofort nutzbar
            ' ist und der gesetzlich nötige §24-UStG-Hinweis nicht vergessen wird.
            sqlCmd.CommandText = "SELECT COUNT(*) FROM einstellungen"
            Dim anzahl As Integer = Convert.ToInt32(sqlCmd.ExecuteScalar())

            If anzahl = 0 Then
                sqlCmd.CommandText = "
                    INSERT INTO einstellungen (schluessel, wert) VALUES
                        ('laufende_rechnungsnummer', '2026001'),
                        ('mwst_satz_1', '7.8'),
                        ('mwst_satz_2', '5.5'),
                        ('firma_name', 'Max Haug'),
                        ('firma_strasse', 'Augsburger Str. 32'),
                        ('firma_plz', '86842'),
                        ('firma_ort', 'Türkheim'),
                        ('firma_tel', '08245/2362'),
                        ('firma_email', 'mail@maxhaug.de'),
                        ('firma_steuer', '138/191/61423'),
                        ('firma_iban', 'DE68701695750000013650'),
                        ('firma_bic', 'GENODEF1TRH'),
                        ('firma_bank', 'Raiffeisenbank Türkheim');

                    INSERT INTO artikel (artikelnummer, bezeichnung, einzelpreis_netto, einzelpreis_brutto, mwst_satz, einheit) VALUES
                        ('001', 'Ballen Weizenstroh 90er', 30.00, 32.34, 7.8, 'C62'),
                        ('002', 'Hackschnitzel Fichte, lose (m³)', 55.00, 58.03, 5.5, 'MTQ');
                "
                sqlCmd.ExecuteNonQuery()

                Dim textZahlungStandard As String =
                    "Bitte überweisen Sie den Rechnungsbetrag innerhalb von 14 Tagen ohne Abzug auf oben genanntes Konto (Rechnungsnummer " &
                    "[RE-nummer] als Verwendungszweck angeben). Die Umsatzsteuer ist gemäß § 24 UStG pauschaliert; der Leistungsempfänger kann " &
                    "die ausgewiesene Steuer als Vorsteuer abziehen."
                Dim textGutschriftStandard As String =
                    "Der oben genannte Betrag wird Ihnen gutgeschrieben. Bitte teilen Sie uns mit, falls keine Verrechnung mit einer künftigen " &
                    "Rechnung gewünscht ist."

                Dim cmdText As New SQLiteCommand("INSERT INTO einstellungen (schluessel, wert) VALUES (@k, @v)", conn)
                cmdText.Parameters.AddWithValue("@k", "text_zahlung")
                cmdText.Parameters.AddWithValue("@v", textZahlungStandard)
                cmdText.ExecuteNonQuery()

                cmdText.Parameters.Clear()
                cmdText.Parameters.AddWithValue("@k", "text_gutschrift")
                cmdText.Parameters.AddWithValue("@v", textGutschriftStandard)
                cmdText.ExecuteNonQuery()
            End If
        End Using
    End Sub

    Public Shared Function HoleVerbindung() As SQLiteConnection
        ' parseViaFramework:=True - siehe Kommentar in InitialisiereDatenbank (UNC-Pfade).
        Dim conn As New SQLiteConnection(ConnectionString, True)
        conn.Open()
        Return conn
    End Function

    ''' <summary>Fügt einer bestehenden Tabelle eine Spalte hinzu, falls sie noch fehlt (per
    ''' PRAGMA table_info geprüft). SQLite kennt kein "ADD COLUMN IF NOT EXISTS" - das hier
    ''' macht denselben Effekt manuell, ohne bestehende Daten anzufassen.</summary>
    Private Shared Sub FuegeSpalteHinzuFallsFehlt(conn As SQLiteConnection, tabelle As String, spalte As String, spaltenTyp As String)
        Dim cmdCheck As New SQLiteCommand($"PRAGMA table_info({tabelle})", conn)
        Using reader = cmdCheck.ExecuteReader()
            While reader.Read()
                If reader("name").ToString().Equals(spalte, StringComparison.OrdinalIgnoreCase) Then
                    Return ' Spalte existiert bereits
                End If
            End While
        End Using

        Dim cmdAlter As New SQLiteCommand($"ALTER TABLE {tabelle} ADD COLUMN {spalte} {spaltenTyp}", conn)
        cmdAlter.ExecuteNonQuery()
    End Sub
End Class
