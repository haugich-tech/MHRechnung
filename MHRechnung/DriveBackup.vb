Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Security.Cryptography
Imports System.Data.SQLite

''' <summary>Sichert Datenbank und Rechnungs-PDFs auf Google Drive, über ein selbst gehostetes
''' Google-Apps-Script als Empfänger (siehe google-apps-script/mhrechnung-backup-empfaenger.gs).
''' Getrennt von GitHub (das sichert nur den Programmcode) und getrennt vom bestehenden
''' E-Mail-Backup (SendeDatenbankBackup in Form1.vb) - beide laufen unabhängig nebeneinander.
'''
''' Zwei grundlegend verschiedene Sicherungsarten:
''' - Datenbank: rohe .sqlite-Datei (kein ZIP nötig, ist schon eine einzelne Datei), auf
'''   Drive werden nur die neuesten 10 Versionen behalten (siehe Skript, Aktion "sichern_db").
''' - Rechnungs-PDFs: einzeln, unverändert, für immer aufbewahrt (Aktion "sichern_pdf").
'''   Damit nicht bei jedem Lauf alle PDFs erneut hochgeladen werden, merkt sich das Programm
'''   in "backup_letzte_rechnung_id" die zuletzt gesicherte Rechnungs-ID und sichert beim
'''   nächsten Mal nur, was seither neu dazugekommen ist.
'''
''' Der Zielordner auf Drive wird nicht frei eingetragen, sondern automatisch aus der
''' Steuernummer gebildet (siehe SicherungsOrdnername) - damit landen zwei Betriebe mit
''' unterschiedlicher Steuernummer nie versehentlich im selben Ordner. Zusätzlich schickt
''' jede Sicherung eine einmalig erzeugte, unsichtbare Kennung mit (siehe HoleOderErzeugeKennung);
''' das Google-Skript lehnt eine Sicherung ab, falls der Zielordner bereits einer anderen
''' Kennung gehört - Schutz davor, dass zwei verschiedene Datenbanken sich unbemerkt einen
''' Drive-Ordner teilen (z.B. weil eine alte Datenbank-Kopie parallel im Einsatz ist).</summary>
Public Class DriveBackup

    Private Shared Function LiesEinstellung(conn As SQLiteConnection, schluessel As String, Optional fallback As String = "") As String
        Dim cmd As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = @k", conn)
        cmd.Parameters.AddWithValue("@k", schluessel)
        Dim res = cmd.ExecuteScalar()
        If res Is Nothing OrElse String.IsNullOrWhiteSpace(res.ToString()) Then Return fallback
        Return res.ToString()
    End Function

    Private Shared Sub SchreibeEinstellung(conn As SQLiteConnection, schluessel As String, wert As String)
        Dim cmdCheck As New SQLiteCommand("SELECT COUNT(*) FROM einstellungen WHERE schluessel = @k", conn)
        cmdCheck.Parameters.AddWithValue("@k", schluessel)
        Dim vorhanden As Boolean = CInt(cmdCheck.ExecuteScalar()) > 0
        Dim cmd As New SQLiteCommand(conn)
        cmd.CommandText = If(vorhanden, "UPDATE einstellungen SET wert = @v WHERE schluessel = @k", "INSERT INTO einstellungen (schluessel, wert) VALUES (@k, @v)")
        cmd.Parameters.AddWithValue("@k", schluessel)
        cmd.Parameters.AddWithValue("@v", wert)
        cmd.ExecuteNonQuery()
    End Sub

    ''' <summary>Ordnername auf Drive, automatisch aus der Steuernummer gebildet (z.B.
    ''' "138/191/61423" -> "138-191-61423"), damit nichts manuell gepflegt werden muss und
    ''' nichts vergessen werden kann.</summary>
    Private Shared Function SicherungsOrdnername(conn As SQLiteConnection) As String
        Dim steuer As String = LiesEinstellung(conn, "firma_steuer", "")
        Dim bereinigt As New StringBuilder()
        For Each c As Char In steuer
            bereinigt.Append(If(Char.IsLetterOrDigit(c), c, "-"c))
        Next
        Dim ergebnis As String = bereinigt.ToString().Trim("-"c)
        Return If(String.IsNullOrWhiteSpace(ergebnis), "unbekannt", ergebnis)
    End Function

    Private Shared Function HoleOderErzeugeKennung(conn As SQLiteConnection) As String
        Dim kennung As String = LiesEinstellung(conn, "backup_drive_kennung", "")
        If String.IsNullOrWhiteSpace(kennung) Then
            kennung = Guid.NewGuid().ToString()
            SchreibeEinstellung(conn, "backup_drive_kennung", kennung)
        End If
        Return kennung
    End Function

    Private Shared Function Md5Hex(bytes As Byte()) As String
        ' Variablenname bewusst nicht "md5" - VB.NET ist nicht case-sensitiv, das würde mit
        ' dem Typnamen "MD5" kollidieren und die Typinferenz von "Using" zum Scheitern bringen.
        Using hasher As MD5 = MD5.Create()
            Dim hash = hasher.ComputeHash(bytes)
            Dim sb As New StringBuilder()
            For Each b In hash
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()
        End Using
    End Function

    Private Shared Function JsonEscape(text As String) As String
        If text Is Nothing Then Return ""
        Dim sb As New StringBuilder()
        For Each c As Char In text
            Select Case c
                Case Chr(34) : sb.Append("\").Append(Chr(34))
                Case "\"c : sb.Append("\\")
                Case vbCr : sb.Append("\r")
                Case vbLf : sb.Append("\n")
                Case vbTab : sb.Append("\t")
                Case Else
                    If AscW(c) < &H20 Then
                        sb.Append("\u" & AscW(c).ToString("x4"))
                    Else
                        sb.Append(c)
                    End If
            End Select
        Next
        Return sb.ToString()
    End Function

    Private Shared Function JsonWertOk(antwortJson As String) As Boolean
        Return antwortJson.Replace(" ", "").Contains(Chr(34) & "ok" & Chr(34) & ":true")
    End Function

    Private Shared Function JsonFehlertext(antwortJson As String) As String
        Dim marker As String = Chr(34) & "fehler" & Chr(34) & ":" & Chr(34)
        Dim start As Integer = antwortJson.IndexOf(marker)
        If start < 0 Then Return ""
        start += marker.Length
        Dim ende As Integer = antwortJson.IndexOf(Chr(34), start)
        If ende < 0 Then Return ""
        Return antwortJson.Substring(start, ende - start)
    End Function

    ''' <summary>Schickt einen fertigen JSON-Körper an die Web-App-URL und liefert die
    ''' Rohantwort zurück. httpFehler bleibt leer, falls die Anfrage technisch geklappt hat
    ''' (auch wenn das Skript inhaltlich "ok: false" antwortet - das wird über JsonWertOk
    ''' geprüft, nicht hier).</summary>
    Private Shared Function PosteAnDrive(url As String, jsonKoerper As String, ByRef httpFehler As String) As String
        httpFehler = ""
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromMinutes(3)
                Dim inhalt As New StringContent(jsonKoerper, Encoding.UTF8, "application/json")
                Dim antwort = client.PostAsync(url, inhalt).GetAwaiter().GetResult()
                Dim antwortText As String = antwort.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                If Not antwort.IsSuccessStatusCode Then
                    httpFehler = $"HTTP-Fehler {CInt(antwort.StatusCode)}"
                End If
                Return antwortText
            End Using
        Catch ex As Exception
            httpFehler = ex.Message
            Return ""
        End Try
    End Function

    Private Shared Function BaueJson(felder As List(Of (Schluessel As String, Wert As String)), zahlenFelder As List(Of (Schluessel As String, Wert As String))) As String
        Dim sb As New StringBuilder("{")
        Dim erstes As Boolean = True
        For Each f In felder
            If Not erstes Then sb.Append(",")
            sb.Append(Chr(34)).Append(f.Schluessel).Append(Chr(34)).Append(":").Append(Chr(34)).Append(JsonEscape(f.Wert)).Append(Chr(34))
            erstes = False
        Next
        For Each f In zahlenFelder
            If Not erstes Then sb.Append(",")
            sb.Append(Chr(34)).Append(f.Schluessel).Append(Chr(34)).Append(":").Append(f.Wert)
            erstes = False
        Next
        sb.Append("}")
        Return sb.ToString()
    End Function

    ''' <summary>Ob die Google-Drive-Sicherung in den Einstellungen aktiviert ist - für den
    ''' "Jetzt sichern"-Button, damit der bei deaktivierter Sicherung nicht fälschlich meldet,
    ''' alle PDFs seien "auf dem aktuellen Stand" (SichereNeuePdfs schweigt in dem Fall
    ''' bewusst, damit der automatische Aufruf nach der Stapelverarbeitung nicht ständig
    ''' unnötig nachfragt).</summary>
    Public Shared Function IstAktiviert() As Boolean
        Using conn = DatenbankManager.HoleVerbindung()
            Return LiesEinstellung(conn, "backup_drive_aktiv", "False") = "True"
        End Using
    End Function

    ''' <summary>Prüft Web-App-URL und Geheimwort, bevor sie gespeichert werden (der
    ''' "Verbindung testen"-Button in den Einstellungen). Liefert bei Erfolg zusätzlich den
    ''' Ordnernamen zurück, der auf Drive verwendet würde.</summary>
    Public Shared Function VerbindungTesten(url As String, geheim As String, ByRef fehler As String, ByRef ordnerName As String) As Boolean
        fehler = ""
        ordnerName = ""
        Try
            Using conn = DatenbankManager.HoleVerbindung()
                ordnerName = SicherungsOrdnername(conn)
                Dim kennung As String = HoleOderErzeugeKennung(conn)
                Dim json As String = BaueJson(New List(Of (String, String)) From {
                    ("geheim", geheim), ("aktion", "test"), ("ordner", ordnerName), ("kennung", kennung)
                }, New List(Of (String, String))())

                Dim httpFehler As String = ""
                Dim antwort As String = PosteAnDrive(url, json, httpFehler)
                If Not String.IsNullOrEmpty(httpFehler) Then
                    fehler = httpFehler
                    Return False
                End If
                If Not JsonWertOk(antwort) Then
                    fehler = JsonFehlertext(antwort)
                    If String.IsNullOrWhiteSpace(fehler) Then fehler = "Verbindung fehlgeschlagen."
                    Return False
                End If
                Return True
            End Using
        Catch ex As Exception
            fehler = ex.Message
            Return False
        End Try
    End Function

    ''' <summary>Liest eine Datei robust ein: im Moment des Sicherns (beim Beenden des
    ''' Programms) kann SQLite die Datenbankdatei noch ganz kurz offen halten, und über ein
    ''' Netzlaufwerk (SMB) wird so eine Freigabe manchmal spürbar verzögert wirksam - beides
    ''' führt zu "Datei wird von einem anderen Prozess verwendet". Öffnet deshalb mit
    ''' möglichst freizügiger Freigabe (FileShare.ReadWrite, wir wollen ja nur lesen) und
    ''' versucht es bei einer Sharing-Violation mit kurzer, steigender Pause erneut.</summary>
    Private Shared Function LiesDateiMitWiederholung(pfad As String) As Byte()
        Const versucheMax As Integer = 5
        Dim letzterFehler As IOException = Nothing
        For versuch As Integer = 1 To versucheMax
            Try
                Using fs As New FileStream(pfad, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim bytes(CInt(fs.Length) - 1) As Byte
                    Dim gelesen As Integer = 0
                    While gelesen < bytes.Length
                        Dim n As Integer = fs.Read(bytes, gelesen, bytes.Length - gelesen)
                        If n = 0 Then Exit While
                        gelesen += n
                    End While
                    Return bytes
                End Using
            Catch ex As IOException
                letzterFehler = ex
                If versuch < versucheMax Then System.Threading.Thread.Sleep(500 * versuch)
            End Try
        Next
        Throw letzterFehler
    End Function

    ''' <summary>Sichert die Datenbank als rohe .sqlite-Datei (Rotation auf die neuesten 10
    ''' Versionen erledigt das Google-Skript). Läuft automatisch (manuell:=False) höchstens
    ''' 1x/Tag, damit bei häufigem Öffnen/Schließen nicht sofort alle 10 Plätze verbraucht
    ''' sind. Gibt bei Erfolg "" zurück, sonst einen Fehlertext.</summary>
    Public Shared Function SichereDatenbank(manuell As Boolean) As String
        Try
            Dim url As String = "", geheim As String = "", ordner As String = "", kennung As String = ""

            ' Verbindung nur für das Lesen der Einstellungen offen halten und danach wieder
            ' schließen - sonst konkurriert unsere eigene, noch offene SQLiteConnection mit
            ' dem gleich folgenden direkten Lesen derselben Datei über FileStream und
            ' erzeugt selbst genau die "Datei wird von einem anderen Prozess verwendet"-
            ' Meldung, die eigentlich vermieden werden soll.
            Using conn = DatenbankManager.HoleVerbindung()
                If LiesEinstellung(conn, "backup_drive_aktiv", "False") <> "True" Then
                    Return If(manuell, "Die Google-Drive-Sicherung ist nicht aktiviert.", "")
                End If

                url = LiesEinstellung(conn, "backup_drive_url", "")
                geheim = LiesEinstellung(conn, "backup_drive_geheimwort", "")
                If String.IsNullOrWhiteSpace(url) OrElse String.IsNullOrWhiteSpace(geheim) Then
                    Return If(manuell, "Bitte trage zuerst Web-App-URL und Geheimwort ein.", "")
                End If

                If Not manuell Then
                    Dim letzte As Date
                    If Date.TryParse(LiesEinstellung(conn, "backup_drive_letzte_db_sicherung", ""), letzte) AndAlso letzte.Date = Date.Today Then
                        Return ""
                    End If
                End If

                ordner = SicherungsOrdnername(conn)
                kennung = HoleOderErzeugeKennung(conn)
            End Using

            Dim dbBytes As Byte() = LiesDateiMitWiederholung(DatenbankManager.DatenbankPfad)
            Dim dateiname As String = $"mhrechnung-db-{DateTime.Now:yyyy-MM-dd_HHmm}.sqlite"

            Dim json As String = BaueJson(New List(Of (String, String)) From {
                ("geheim", geheim), ("aktion", "sichern_db"), ("ordner", ordner), ("kennung", kennung),
                ("name", dateiname), ("daten", Convert.ToBase64String(dbBytes)), ("md5", Md5Hex(dbBytes))
            }, New List(Of (String, String)) From {("groesse", dbBytes.Length.ToString())})

            Dim httpFehler As String = ""
            Dim antwort As String = PosteAnDrive(url, json, httpFehler)
            If Not String.IsNullOrEmpty(httpFehler) Then Return httpFehler
            If Not JsonWertOk(antwort) Then
                Dim f As String = JsonFehlertext(antwort)
                Return If(String.IsNullOrWhiteSpace(f), "Unbekannter Fehler bei der Drive-Sicherung.", f)
            End If

            Using conn = DatenbankManager.HoleVerbindung()
                SchreibeEinstellung(conn, "backup_drive_letzte_db_sicherung", DateTime.Now.ToString("yyyy-MM-dd"))
            End Using
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>Lädt alle seit der letzten Sicherung neu hinzugekommenen Rechnungs-PDFs
    ''' (normale Variante + e-Rechnung) einzeln und dauerhaft auf Drive hoch - keine Rotation,
    ''' keine erneute Sicherung bereits gesicherter Rechnungen. baseDir ist der
    ''' Haupt-Speicherpfad (Einstellung "speicherpfad"), wie ihn VerarbeiteStapel schon kennt.
    ''' Gibt eine Liste von Fehlertexten zurück (leer = alles erfolgreich); bei einem Fehler
    ''' wird ab dieser Rechnung abgebrochen, damit nichts übersprungen wird - der nächste
    ''' Lauf setzt automatisch wieder an der gleichen Stelle an.</summary>
    Public Shared Function SichereNeuePdfs(baseDir As String) As List(Of String)
        Dim fehlerListe As New List(Of String)
        Try
            Using conn = DatenbankManager.HoleVerbindung()
                If LiesEinstellung(conn, "backup_drive_aktiv", "False") <> "True" Then Return fehlerListe

                Dim url As String = LiesEinstellung(conn, "backup_drive_url", "")
                Dim geheim As String = LiesEinstellung(conn, "backup_drive_geheimwort", "")
                If String.IsNullOrWhiteSpace(url) OrElse String.IsNullOrWhiteSpace(geheim) Then Return fehlerListe

                Dim ordner As String = SicherungsOrdnername(conn)
                Dim kennung As String = HoleOderErzeugeKennung(conn)
                Dim letzteId As Integer = 0
                Integer.TryParse(LiesEinstellung(conn, "backup_letzte_rechnung_id", "0"), letzteId)

                Dim cmd As New SQLiteCommand("SELECT id, rechnungsnummer FROM rechnungen WHERE id > @letzte AND status = 'Verarbeitet & Exportiert' ORDER BY id", conn)
                cmd.Parameters.AddWithValue("@letzte", letzteId)
                Dim ausstehend As New List(Of (Id As Integer, ReNr As String))
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        ausstehend.Add((CInt(reader("id")), reader("rechnungsnummer").ToString()))
                    End While
                End Using

                For Each re In ausstehend
                    ' Die Rechnungsnummer beginnt mit dem Jahr, in dem sie vergeben wurde
                    ' (Format JJJJXXX) - das ist fast immer auch der Jahresordner, in dem das
                    ' PDF liegt (Stapelverarbeitung erstellt es meist am selben Tag). Als
                    ' Rückfallebene wird zusätzlich das aktuelle Jahr geprüft, falls eine
                    ' Rechnung über den Jahreswechsel hinweg unverarbeitet liegen blieb.
                    Dim jahresKandidaten As New List(Of String) From {DateTime.Now.Year.ToString()}
                    If re.ReNr.Length >= 4 AndAlso Not jahresKandidaten.Contains(re.ReNr.Substring(0, 4)) Then
                        jahresKandidaten.Insert(0, re.ReNr.Substring(0, 4))
                    End If

                    Dim normalesPdf As String = Nothing
                    Dim eRechnungPdf As String = Nothing
                    For Each jahr In jahresKandidaten
                        Dim kandidatNormal = Path.Combine(baseDir, jahr, "erstellt", "pdf", re.ReNr & ".pdf")
                        Dim kandidatE = Path.Combine(baseDir, jahr, "erstellt", "e-rechnungen", re.ReNr & "_zugf.pdf")
                        If normalesPdf Is Nothing AndAlso File.Exists(kandidatNormal) Then normalesPdf = kandidatNormal
                        If eRechnungPdf Is Nothing AndAlso File.Exists(kandidatE) Then eRechnungPdf = kandidatE
                    Next

                    Dim vorgangOk As Boolean = True
                    Dim f As String = ""
                    If normalesPdf IsNot Nothing Then
                        If Not LadeDateiHoch(url, geheim, ordner, kennung, normalesPdf, re.ReNr & ".pdf", f) Then
                            fehlerListe.Add($"Re-{re.ReNr}: {f}")
                            vorgangOk = False
                        End If
                    End If
                    If vorgangOk AndAlso eRechnungPdf IsNot Nothing Then
                        If Not LadeDateiHoch(url, geheim, ordner, kennung, eRechnungPdf, re.ReNr & "_zugf.pdf", f) Then
                            fehlerListe.Add($"Re-{re.ReNr}: {f}")
                            vorgangOk = False
                        End If
                    End If

                    If Not vorgangOk Then Exit For

                    letzteId = re.Id
                    SchreibeEinstellung(conn, "backup_letzte_rechnung_id", letzteId.ToString())
                Next
            End Using
        Catch ex As Exception
            fehlerListe.Add("Unerwarteter Fehler bei der PDF-Sicherung: " & ex.Message)
        End Try
        Return fehlerListe
    End Function

    Private Shared Function LadeDateiHoch(url As String, geheim As String, ordner As String, kennung As String, dateipfad As String, zielname As String, ByRef fehler As String) As Boolean
        fehler = ""
        Dim bytes As Byte() = File.ReadAllBytes(dateipfad)
        Dim json As String = BaueJson(New List(Of (String, String)) From {
            ("geheim", geheim), ("aktion", "sichern_pdf"), ("ordner", ordner), ("kennung", kennung),
            ("name", zielname), ("daten", Convert.ToBase64String(bytes)), ("md5", Md5Hex(bytes))
        }, New List(Of (String, String)) From {("groesse", bytes.Length.ToString())})

        Dim httpFehler As String = ""
        Dim antwort As String = PosteAnDrive(url, json, httpFehler)
        If Not String.IsNullOrEmpty(httpFehler) Then
            fehler = httpFehler
            Return False
        End If
        If Not JsonWertOk(antwort) Then
            fehler = JsonFehlertext(antwort)
            If String.IsNullOrWhiteSpace(fehler) Then fehler = "Unbekannter Fehler beim Hochladen."
            Return False
        End If
        Return True
    End Function

End Class
