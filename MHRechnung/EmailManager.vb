Imports System.Net
Imports System.Net.Mail
Imports System.IO
Imports System.Data.SQLite

Public Class EmailManager

    Private Shared Function LadeEinstellungen() As Dictionary(Of String, String)
        Dim settings As New Dictionary(Of String, String)()
        Using conn = DatenbankManager.HoleVerbindung()
            Dim cmd As New SQLiteCommand("SELECT schluessel, wert FROM einstellungen", conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    settings(reader("schluessel").ToString()) = reader("wert").ToString()
                End While
            End Using
        End Using
        Return settings
    End Function

    Private Shared Function GetSetting(settings As Dictionary(Of String, String), key As String, Optional fallback As String = "") As String
        If settings.ContainsKey(key) AndAlso Not String.IsNullOrWhiteSpace(settings(key)) Then
            Return settings(key)
        End If
        Return fallback
    End Function

    Public Shared Sub SendeRechnung(reNr As String, empfaengerEmail As String, pdfPfad As String)
        Dim s = LadeEinstellungen()

        ' 1. SMTP-Daten aus den Einstellungen laden
        Dim smtpServer = GetSetting(s, "smtp_server", "")
        Dim smtpPortStr = GetSetting(s, "smtp_port", "587")
        Dim smtpUser = GetSetting(s, "smtp_user", "")
        Dim smtpPass = GetSetting(s, "smtp_pass", "")

        Dim kontaktMail = GetSetting(s, "firma_email", "")
        Dim firmenName = GetSetting(s, "firma_name", "LEG Wertachtal GBR")
        Dim baseDir = GetSetting(s, "speicherpfad", "C:\LEG_Rechnungen")

        If String.IsNullOrWhiteSpace(smtpServer) OrElse String.IsNullOrWhiteSpace(smtpUser) OrElse String.IsNullOrWhiteSpace(smtpPass) Then
            Throw New Exception("Die E-Mail-Server-Daten (SMTP) sind nicht vollständig im Tab 'Einstellungen' hinterlegt.")
        End If

        Dim smtpPort As Integer = 587
        Integer.TryParse(smtpPortStr, smtpPort)
        If smtpPort = 465 Then smtpPort = 587

        ' 2. E-Mail Text & Betreff
        Dim rohText As String = GetSetting(s, "text_email", "")
        Dim textMitNummer As String = rohText.Replace("[RE-nummer]", "re-" & reNr)
        Dim subject As String = $"Rechnung re-{reNr} von {firmenName}"

        ' --- NEU: HTML FORMATIERUNG UND LOGO ---
        ' Zeilenumbrüche aus deinem Textfeld in echte HTML-Zeilenumbrüche umwandeln
        Dim htmlBody As String = textMitNummer.Replace(vbCrLf, "<br>").Replace(vbLf, "<br>")

        ' Den Text in ein schönes Arial-Layout verpacken (Schriftart, Größe, Farbe)
        htmlBody = $"<div style='font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #333333; line-height: 1.5;'>" &
                   htmlBody &
                   "<br><br>"

        ' Prüfen, ob ein Logo im Speicherpfad existiert
        Dim logoPfad As String = ToolPfade.LogoPfad ' Wenn es ein PNG ist, hier auf logo.png ändern!
        Dim hatLogo As Boolean = File.Exists(logoPfad)

        If hatLogo Then
            ' Das 'cid:' sagt dem E-Mail-Programm, dass das Bild im Text eingebettet ist
            htmlBody &= "<img src='cid:FirmenLogo' alt='LEG Wertachtal Logo' style='max-width: 250px;'><br>"
        End If

        htmlBody &= "</div>"
        ' ----------------------------------------

        ' 3. TLS 1.2 ERZWINGEN
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        ' 4. E-Mail zusammenbauen und senden
        Using mail As New MailMessage()
            mail.From = New MailAddress(smtpUser, firmenName)
            mail.To.Add(empfaengerEmail)
            mail.Bcc.Add(smtpUser)

            If Not String.IsNullOrWhiteSpace(kontaktMail) Then
                mail.ReplyToList.Add(New MailAddress(kontaktMail, firmenName))
            End If

            mail.Subject = subject

            ' WICHTIG: mail.Body und mail.IsBodyHtml lassen wir hier komplett weg!
            ' Wir nutzen stattdessen die sauberen AlternateViews.

            ' Version A: Die einfache Text-Version (falls jemand ein uraltes Mail-Programm hat)
            Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(textMitNummer, Nothing, "text/plain")
            mail.AlternateViews.Add(plainView)

            ' Version B: Die schöne HTML-Version inkl. eingebettetem Logo
            Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(htmlBody, Nothing, "text/html")

            If hatLogo Then
                ' Logo laden und als "image/jpeg" definieren, damit das Mail-Programm es richtig versteht
                Dim logoResource As New LinkedResource(logoPfad)
                logoResource.ContentId = "FirmenLogo"
                logoResource.ContentType.MediaType = "image/jpeg" ' Sehr wichtig für die Anzeige!
                htmlView.LinkedResources.Add(logoResource)
            End If

            mail.AlternateViews.Add(htmlView)

            ' PDF anhängen
            If File.Exists(pdfPfad) Then
                mail.Attachments.Add(New Attachment(pdfPfad))
            Else
                Throw New Exception("Die PDF-Datei zum Anhängen wurde nicht gefunden.")
            End If

            ' Senden
            Using client As New SmtpClient(smtpServer, smtpPort)
                client.Credentials = New NetworkCredential(smtpUser, smtpPass)
                client.EnableSsl = True
                client.Timeout = 10000

                client.Send(mail)
            End Using
        End Using
    End Sub

End Class