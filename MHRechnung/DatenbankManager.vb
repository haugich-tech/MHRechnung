Imports System.Data.SQLite
Imports System.IO

Public Class DatenbankManager
    Private Const dbName As String = "MHRechnung_Daten.sqlite"
    Private Const connectionString As String = "Data Source=" & dbName & ";Version=3;"

    Public Shared Sub InitialisiereDatenbank()
        ' 1. Datei erstellen, falls sie komplett fehlt
        If Not File.Exists(dbName) Then
            SQLiteConnection.CreateFile(dbName)
        End If

        ' 2. Verbindung öffnen und Tabellen IMMER prüfen/erstellen
        Using conn As New SQLiteConnection(connectionString)
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

            ' 3. Grundeinstellungen NUR einfügen, wenn die Einstellungen-Tabelle noch absolut leer ist.
            ' Firmenname, Adresse, IBAN usw. werden bewusst NICHT vorbelegt - die trägst du
            ' selbst im Tab "Einstellungen" ein. Die beiden MwSt-Sätze, die Startnummer und die
            ' Zahlungstexte bekommen einen sinnvollen Standardwert, damit das Programm sofort
            ' nutzbar ist und der gesetzlich nötige §24-UStG-Hinweis nicht vergessen wird.
            sqlCmd.CommandText = "SELECT COUNT(*) FROM einstellungen"
            Dim anzahl As Integer = Convert.ToInt32(sqlCmd.ExecuteScalar())

            If anzahl = 0 Then
                sqlCmd.CommandText = "
                    INSERT INTO einstellungen (schluessel, wert) VALUES
                        ('laufende_rechnungsnummer', '2026001'),
                        ('mwst_satz_1', '7.8'),
                        ('mwst_satz_2', '5.5');

                    INSERT INTO artikel (artikelnummer, bezeichnung, einzelpreis_netto, mwst_satz, einheit) VALUES
                        ('001', 'Ballen Weizenstroh 90er', 30.00, 7.8, 'C62'),
                        ('002', 'Hackschnitzel Fichte, lose (m³)', 55.00, 5.5, 'MTQ');
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
        Dim conn As New SQLiteConnection(connectionString)
        conn.Open()
        Return conn
    End Function
End Class
