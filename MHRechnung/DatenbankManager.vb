Imports System.Data.SQLite
Imports System.IO

Public Class DatenbankManager
    Private Const dbName As String = "LEG_Daten.sqlite"
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
                    iban TEXT, 
                    bic TEXT,
                    bankname TEXT,
                    mandat_datum TEXT, 
                    sepa_typ TEXT,
                    betriebsnummer TEXT,
                    email TEXT, 
                    versandart TEXT
                );
                
                CREATE TABLE IF NOT EXISTS artikel (id INTEGER PRIMARY KEY AUTOINCREMENT, artikelnummer TEXT, bezeichnung TEXT, einzelpreis_netto REAL, mwst_satz INTEGER, einheit TEXT);
                CREATE TABLE IF NOT EXISTS rechnungen (id INTEGER PRIMARY KEY AUTOINCREMENT, rechnungsnummer TEXT, datum TEXT, mitglied_id INTEGER, status TEXT);
                CREATE TABLE IF NOT EXISTS rechnungspositionen (id INTEGER PRIMARY KEY AUTOINCREMENT, rechnung_id INTEGER, artikel_bezeichnung TEXT, anzahl REAL, einzelpreis REAL, mwst_satz INTEGER);
                CREATE TABLE IF NOT EXISTS einstellungen (schluessel TEXT PRIMARY KEY, wert TEXT);
            "
            sqlCmd.ExecuteNonQuery()

            ' 3. Dummy-Daten NUR einfügen, wenn die Einstellungen-Tabelle noch absolut leer ist
            sqlCmd.CommandText = "SELECT COUNT(*) FROM einstellungen"
            Dim anzahl As Integer = Convert.ToInt32(sqlCmd.ExecuteScalar())

            If anzahl = 0 Then
                sqlCmd.CommandText = "
                    INSERT INTO einstellungen (schluessel, wert) VALUES ('firmenname', 'LEG Wertachtal GBR'), ('glaeubiger_id', 'DE98ZZZ09999999999'), ('laufende_rechnungsnummer', '2026016');
                    
                    INSERT INTO mitglieder (mitgliedsnummer, name, strasse, plz, ort, iban, sepa_typ, versandart) 
                    VALUES ('M-001', 'Michael Schwarzenbach', 'Musterweg 1', '86842', 'Türkheim', 'DE1573...', 'FRST', 'Beides'), 
                           ('M-002', 'Bettina Mayer', 'Dorfstraße 5', '86842', 'Türkheim', 'DE1234...', 'FRST', 'E-Mail');
                           
                    INSERT INTO artikel (artikelnummer, bezeichnung, einzelpreis_netto, mwst_satz, einheit) VALUES ('016', 'Vario 720 Felgen', 900.00, 7, 'C62'), ('001', 'saltz', 0.12, 7, 'KGM'), ('022', 'Dieselzusatz', 15.50, 19, 'LTR');
                "
                sqlCmd.ExecuteNonQuery()
            End If
        End Using
    End Sub

    Public Shared Function HoleVerbindung() As SQLiteConnection
        Dim conn As New SQLiteConnection(connectionString)
        conn.Open()
        Return conn
    End Function
End Class