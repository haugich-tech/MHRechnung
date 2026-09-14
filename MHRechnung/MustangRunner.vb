Imports System.IO
Imports System.Diagnostics
Imports System.Data.SQLite

Public Class MustangRunner

    ' Hilfsfunktion: Holt den Speicherpfad aus der Datenbank
    Private Shared Function HoleSpeicherpfad() As String
        Dim pfad As String = "C:\MHRechnung"
        Try
            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmd As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'speicherpfad'", conn)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(result.ToString()) Then
                    pfad = result.ToString()
                End If
            End Using
        Catch
            ' Fallback auf Standard, falls DB-Zugriff scheitert
        End Try
        Return pfad
    End Function

    Public Shared Sub ErstelleZUGFeRDPdf(reNr As String)
        Dim baseDir As String = HoleSpeicherpfad()
        Dim jahr As String = DateTime.Now.Year.ToString()

        ' --- NEU: Sicherheits-Check für die Ordnerstruktur ---
        Dim tempPath As String = Path.Combine(baseDir, "temp")

        ' Falls der temp-Ordner fehlt, erstelle ihn einfach (macht nichts, wenn er schon da ist)
        If Not Directory.Exists(tempPath) Then
            Directory.CreateDirectory(tempPath)
        End If

        ' Auch den Ziel-Ordner für die E-Rechnungen sicherheitshalber prüfen
        Dim zielPfadOrdner As String = Path.Combine(baseDir, jahr, "erstellt", "e-rechnungen")
        If Not Directory.Exists(zielPfadOrdner) Then
            Directory.CreateDirectory(zielPfadOrdner)
        End If

        ' 1. Pfade zu den Dateien definieren
        ' Rechnungsnummer in den Temp-Namen: verhindert Verwechslungen zwischen Rechnungen
        Dim quellPdf As String = Path.Combine(baseDir, jahr, "erstellt", "pdf", reNr & ".pdf")
        Dim tempPdfA As String = Path.Combine(baseDir, "temp", "temp_pdfa_" & reNr & ".pdf")
        Dim quellXml As String = Path.Combine(baseDir, "temp", "factur-x_" & reNr & ".xml")
        Dim zielPdf As String = Path.Combine(baseDir, jahr, "erstellt", "e-rechnungen", reNr & "_zugf.pdf")

        ' 2. Pfade zu den Tools (basierend auf deinen Tests)
        Dim gsExe As String = ToolPfade.GhostscriptExe
        Dim javaExe As String = ToolPfade.JavaExe
        Dim mustangJar As String = ToolPfade.MustangJar

        ' Timeout pro Prozess (2 Minuten) - verhindert dauerhaft eingefrorene UI
        Const PROZESS_TIMEOUT_MS As Integer = 120000

        ' 3. Vorab-Checks
        If Not File.Exists(quellPdf) Then Throw New FileNotFoundException("Original-PDF fehlt: " & quellPdf)
        If Not File.Exists(quellXml) Then Throw New FileNotFoundException("XML fehlt: " & quellXml)
        If Not File.Exists(gsExe) Then Throw New FileNotFoundException("Ghostscript fehlt unter: " & gsExe)
        If Not File.Exists(javaExe) Then Throw New FileNotFoundException("Java fehlt unter: " & javaExe)
        If Not File.Exists(mustangJar) Then Throw New FileNotFoundException("Mustang-CLI fehlt unter: " & mustangJar)

        Try
            ' --- SCHRITT 1: GHOSTSCRIPT (Normales PDF -> PDF/A-3 konvertieren) ---
            ' Wir nutzen exakt deine funktionierenden Parameter:
            Dim gsArgs As String = $"-dPDFA=3 -dBATCH -dNOPAUSE -sProcessColorModel=DeviceRGB -sColorConversionStrategy=RGB -sDEVICE=pdfwrite -dPDFACompatibilityPolicy=1 -sOutputFile=""{tempPdfA}"" ""{quellPdf}"""

            Dim gsStartInfo As New ProcessStartInfo(gsExe, gsArgs) With {
                .CreateNoWindow = True,
                .UseShellExecute = False
            }
            Using gsProc = Process.Start(gsStartInfo)
                If Not gsProc.WaitForExit(PROZESS_TIMEOUT_MS) Then
                    gsProc.Kill()
                    Throw New Exception("Ghostscript hat nicht innerhalb von 2 Minuten geantwortet und wurde abgebrochen.")
                End If
                If gsProc.ExitCode <> 0 Then
                    Throw New Exception($"Ghostscript hat einen Fehler gemeldet (ExitCode {gsProc.ExitCode}). Die PDF/A-Konvertierung ist fehlgeschlagen.")
                End If
            End Using

            ' --- SCHRITT 2: MUSTANG (XML in das neue PDF/A einbetten) ---
            If File.Exists(tempPdfA) Then
                ' Wir nutzen den bewährten 'combine' Befehl aus deinem Excel-Makro:
                Dim mustangArgs As String = $"-jar ""{mustangJar}"" --action combine --no-additional-attachments --source ""{tempPdfA}"" --source-xml ""{quellXml}"" --out ""{zielPdf}"" --format fx --version 1 --profile E"

                Dim mStartInfo As New ProcessStartInfo(javaExe, mustangArgs) With {
                    .CreateNoWindow = True,
                    .UseShellExecute = False
                }
                Using mProc = Process.Start(mStartInfo)
                    If Not mProc.WaitForExit(PROZESS_TIMEOUT_MS) Then
                        mProc.Kill()
                        Throw New Exception("Mustang (Java) hat nicht innerhalb von 2 Minuten geantwortet und wurde abgebrochen.")
                    End If
                    If mProc.ExitCode <> 0 Then
                        Throw New Exception($"Mustang hat einen Fehler gemeldet (ExitCode {mProc.ExitCode}). Das XML konnte nicht eingebettet werden.")
                    End If
                End Using

                ' Endkontrolle: Wurde die E-Rechnung wirklich erzeugt?
                If Not File.Exists(zielPdf) Then
                    Throw New Exception("Die ZUGFeRD-PDF wurde nicht erzeugt, obwohl Mustang keinen Fehler gemeldet hat: " & zielPdf)
                End If
            Else
                Throw New Exception("Ghostscript konnte die PDF/A-Datei nicht erzeugen.")
            End If

        Catch ex As Exception
            Throw New Exception("Fehler in der ZUGFeRD-Kette (Rechnung " & reNr & "): " & ex.Message)
        Finally
            ' Aufräumen: Temp-Dateien immer löschen, auch im Fehlerfall
            Try
                If File.Exists(tempPdfA) Then File.Delete(tempPdfA)
                If File.Exists(quellXml) Then File.Delete(quellXml)
            Catch
                ' Aufräumfehler sind unkritisch
            End Try
        End Try
    End Sub

End Class