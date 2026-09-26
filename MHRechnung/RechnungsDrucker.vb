Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports PdfSharp.Fonts
Imports System.IO
Imports System.Data.SQLite
Imports QRCoder

Public Class RechnungsDrucker

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

    ' Zeichnet den Text und gibt die neue Y-Position zurück
    Private Shared Function DrawWrappedText(gfx As XGraphics, text As String, font As XFont,
                                             brush As XBrush, x As Double, y As Double,
                                             maxWidth As Double, lineHeight As Double) As Double
        If String.IsNullOrWhiteSpace(text) Then Return y
        Dim lines As String() = text.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
        For Each rawLine As String In lines
            Dim words As String() = rawLine.Split(" "c)
            Dim currentLine As String = ""
            For Each word As String In words
                Dim testLine As String = If(currentLine = "", word, currentLine & " " & word)
                If gfx.MeasureString(testLine, font).Width > maxWidth AndAlso currentLine <> "" Then
                    gfx.DrawString(currentLine, font, brush, x, y)
                    y += lineHeight
                    currentLine = word
                Else
                    currentLine = testLine
                End If
            Next
            If currentLine <> "" Then
                gfx.DrawString(currentLine, font, brush, x, y)
                y += lineHeight
            End If
        Next
        Return y
    End Function

    ' Berechnet VORHER die Höhe des Textes für den Seitenumbruch
    Private Shared Function MeasureWrappedTextHeight(gfx As XGraphics, text As String, font As XFont,
                                                     maxWidth As Double, lineHeight As Double) As Double
        If String.IsNullOrWhiteSpace(text) Then Return 0
        Dim lines As String() = text.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
        Dim totalHeight As Double = 0
        For Each rawLine As String In lines
            Dim words As String() = rawLine.Split(" "c)
            Dim currentLine As String = ""
            For Each word As String In words
                Dim testLine As String = If(currentLine = "", word, currentLine & " " & word)
                If gfx.MeasureString(testLine, font).Width > maxWidth AndAlso currentLine <> "" Then
                    totalHeight += lineHeight
                    currentLine = word
                Else
                    currentLine = testLine
                End If
            Next
            If currentLine <> "" Then
                totalHeight += lineHeight
            End If
        Next
        Return totalHeight
    End Function

    Private Shared Sub DrawRight(gfx As XGraphics, text As String, font As XFont,
                                  brush As XBrush, rightEdge As Double, y As Double)
        Dim w As Double = gfx.MeasureString(text, font).Width
        gfx.DrawString(text, font, brush, rightEdge - w, y)
    End Sub

    Public Shared Sub ErstelleRechnung(reID As Integer, reNr As String)

        ' ── 0. Einstellungen ──────────────────────────────────────────────────
        Dim s As Dictionary(Of String, String) = LadeEinstellungen()

        Dim speicherPfad As String = GetSetting(s, "speicherpfad", "C:\MHRechnung")
        Dim firmaName As String = GetSetting(s, "firma_name", "")
        Dim firmaStrasse As String = GetSetting(s, "firma_strasse", "")
        Dim firmaPLZ As String = GetSetting(s, "firma_plz", "")
        Dim firmaOrt As String = GetSetting(s, "firma_ort", "")
        Dim firmaEmail As String = GetSetting(s, "firma_email", "")
        Dim firmaTel As String = GetSetting(s, "firma_tel", "")
        Dim firmaHandy As String = GetSetting(s, "firma_handy", "")
        Dim firmaSteuer As String = GetSetting(s, "firma_steuer", "")
        Dim firmaIBAN As String = GetSetting(s, "firma_iban", "")
        Dim firmaBIC As String = GetSetting(s, "firma_bic", "")
        Dim firmaBank As String = GetSetting(s, "firma_bank", "")
        ' Die MwSt-Sätze selbst müssen hier nicht geladen werden - jede Position trägt
        ' ihren eigenen Satz bereits aus der Datenbank (rechnungspositionen.mwst_satz).

        Dim textZahlung As String = GetSetting(s, "text_zahlung", "")
        Dim textGutschrift As String = GetSetting(s, "text_gutschrift", "")

        ' ── 1. Ordner vorbereiten ─────────────────────────────────────────────
        Dim jahr As String = DateTime.Now.Year.ToString()
        Dim pdfPath As String = Path.Combine(speicherPfad, jahr, "erstellt", "pdf")
        If Not Directory.Exists(pdfPath) Then Directory.CreateDirectory(pdfPath)
        Dim dateiName As String = Path.Combine(pdfPath, reNr & ".pdf")

        ' ── 2. Daten aus DB laden ─────────────────────────────────────────────
        Dim kundenName As String = ""
        Dim kundenStrasse As String = ""
        Dim kundenPLZ As String = ""
        Dim kundenOrt As String = ""
        Dim kundenSteuer As String = ""
        Dim kundenBetrieb As String = ""
        Dim kundenEmail As String = ""
        Dim rechnungsdatum As String = ""
        Dim lieferdatum As String = ""
        Dim preisart As String = "Netto"
        Dim positionen As New List(Of (bez As String, anz As Decimal, prs As Decimal, prsBrutto As Decimal, mwst As Decimal))

        Using conn = DatenbankManager.HoleVerbindung()
            Dim sqlK = "SELECT r.datum, r.lieferdatum, r.preisart, m.name, m.strasse, m.plz, m.ort, m.steuernummer, m.betriebsnummer, m.email " &
                       "FROM rechnungen r JOIN mitglieder m ON r.mitglied_id = m.id WHERE r.id = @id"
            Dim cmdK As New SQLiteCommand(sqlK, conn)
            cmdK.Parameters.AddWithValue("@id", reID)
            Using r = cmdK.ExecuteReader()
                If r.Read() Then
                    rechnungsdatum = r("datum").ToString()
                    lieferdatum = r("lieferdatum").ToString()
                    If r("preisart").ToString() = "Brutto" Then preisart = "Brutto"
                    kundenName = r("name").ToString()
                    kundenStrasse = r("strasse").ToString()
                    kundenPLZ = r("plz").ToString()
                    kundenOrt = r("ort").ToString()
                    kundenSteuer = r("steuernummer").ToString()
                    kundenBetrieb = r("betriebsnummer").ToString()
                    kundenEmail = r("email").ToString().Trim()
                End If
            End Using
            If String.IsNullOrWhiteSpace(rechnungsdatum) Then rechnungsdatum = DateTime.Now.ToString("dd.MM.yyyy")
            If String.IsNullOrWhiteSpace(lieferdatum) Then lieferdatum = rechnungsdatum

            Dim cmdP As New SQLiteCommand("SELECT artikel_bezeichnung, anzahl, einzelpreis, einzelpreis_brutto, mwst_satz FROM rechnungspositionen WHERE rechnung_id = @id", conn)
            cmdP.Parameters.AddWithValue("@id", reID)
            Using r = cmdP.ExecuteReader()
                While r.Read()
                    Dim posNetto As Decimal = CDec(r("einzelpreis"))
                    Dim posMwst As Decimal = CDec(r("mwst_satz"))
                    ' Exakt gespeicherten Bruttopreis übernehmen (wichtig für glatte, vereinbarte
                    ' Beträge wie "2.000,00 €") - fehlt er (NULL, z.B. Alt-/Importdaten), aus
                    ' Netto berechnen.
                    Dim posBrutto As Decimal
                    If IsDBNull(r("einzelpreis_brutto")) OrElse CDec(r("einzelpreis_brutto")) <= 0 Then
                        posBrutto = posNetto * (1 + posMwst / 100D)
                    Else
                        posBrutto = CDec(r("einzelpreis_brutto"))
                    End If
                    positionen.Add((r("artikel_bezeichnung").ToString(), CDec(r("anzahl")), posNetto, posBrutto, posMwst))
                End While
            End Using
        End Using

        Dim nettoGesamt As Decimal = 0
        Dim mwstGesamt As Decimal = 0
        Dim bruttoGesamt As Decimal = 0
        Dim nettoBasisProSatz As New Dictionary(Of Decimal, Decimal)
        Dim mwstBetragProSatz As New Dictionary(Of Decimal, Decimal)

        If preisart = "Brutto" Then
            ' Bei einer Brutto-Rechnung ist der eingegebene Bruttopreis die vereinbarte, "echte"
            ' Zahl (z.B. glatt 2.000,00 €) - deshalb hier als Ausgangspunkt nehmen und Netto/
            ' Steuer je Satz davon zurückrechnen (Steuer = Brutto - Netto, exakte Subtraktion),
            ' statt wie im Netto-Zweig unten erst Netto zu runden und Brutto daraus
            ' hochzurechnen. Genau das hätte sonst wieder die Cent-Abweichung erzeugt, die der
            ' ganze Netto/Brutto-Umbau vermeiden sollte (z.B. 4.000,01 € statt der vereinbarten
            ' glatten 4.000,00 €).
            Dim bruttoBasisProSatz As New Dictionary(Of Decimal, Decimal)
            For Each pos In positionen
                Dim zb As Decimal = Math.Round(pos.anz * pos.prsBrutto, 2, MidpointRounding.AwayFromZero)
                If Not bruttoBasisProSatz.ContainsKey(pos.mwst) Then bruttoBasisProSatz(pos.mwst) = 0
                bruttoBasisProSatz(pos.mwst) += zb
            Next
            For Each kv In bruttoBasisProSatz
                Dim nettoBucket As Decimal = Math.Round(kv.Value / (1 + kv.Key / 100D), 2, MidpointRounding.AwayFromZero)
                Dim steuerBucket As Decimal = kv.Value - nettoBucket
                nettoBasisProSatz(kv.Key) = nettoBucket
                mwstBetragProSatz(kv.Key) = steuerBucket
                nettoGesamt += nettoBucket
                mwstGesamt += steuerBucket
                bruttoGesamt += kv.Value
            Next
        Else
            For Each pos In positionen
                ' Zeilennetto auf 2 Stellen runden - identische Logik wie im ZUGFeRD-XML
                Dim zn As Decimal = Math.Round(pos.anz * pos.prs, 2, MidpointRounding.AwayFromZero)
                nettoGesamt += zn
                If Not nettoBasisProSatz.ContainsKey(pos.mwst) Then nettoBasisProSatz(pos.mwst) = 0
                nettoBasisProSatz(pos.mwst) += zn
            Next
            ' Steuer auf die gerundete Basis berechnen - PDF stimmt damit exakt mit XML überein
            For Each kv In nettoBasisProSatz
                Dim betrag As Decimal = Math.Round(kv.Value * (kv.Key / 100D), 2, MidpointRounding.AwayFromZero)
                mwstBetragProSatz(kv.Key) = betrag
                mwstGesamt += betrag
            Next
            bruttoGesamt = nettoGesamt + mwstGesamt
        End If

        ' --- GUTSCHRIFTEN-LOGIK ---
        Dim titelText As String = "RECHNUNG"
        Dim summenLabel As String = "Rechnungsbetrag"
        Dim basisText As String = textZahlung

        If bruttoGesamt < 0 Then
            titelText = "GUTSCHRIFT / KORREKTUR"
            summenLabel = "Guthaben / Auszahlungsbetrag"
            basisText = textGutschrift
        ElseIf bruttoGesamt = 0 Then
            titelText = "RECHNUNG"
            summenLabel = "Rechnungsbetrag"
            basisText = "Der Rechnungsbetrag beläuft sich auf 0,00 €. Diese Rechnung dient lediglich zur Information/Korrektur, es ist keine Zahlung erforderlich."
        End If

        Dim zahlungsText As String = basisText.Replace("[RE-nummer]", "re-" & reNr)

        ' ── 3. PDF ZEICHNEN (MULTI-PAGE LOGIK) ────────────────────────────────
        ' Temp-Datei für den GiroCode (falls erzeugt) - wird erst nach document.Save
        ' aufgeräumt, siehe ganz unten.
        Dim qrTempPfad As String = Nothing
        Using document As New PdfDocument()
            Dim pages As New List(Of PdfPage)
            Dim gfxList As New List(Of XGraphics)

            Dim p1 As PdfPage = document.AddPage()
            p1.Size = PdfSharp.PageSize.A4
            pages.Add(p1)

            Dim pH As Double = p1.Height.Point
            Dim mL As Double = 40
            Dim mR As Double = 555
            Dim cW As Double = mR - mL

            Dim fNorm As New XFont("Arial", 9, XFontStyleEx.Regular)
            Dim fBold As New XFont("Arial", 9, XFontStyleEx.Bold)
            Dim fSmall As New XFont("Arial", 8, XFontStyleEx.Regular)
            Dim fSmallB As New XFont("Arial", 8, XFontStyleEx.Bold)
            Dim fTiny As New XFont("Arial", 7, XFontStyleEx.Regular)
            Dim fTitle As New XFont("Arial", 28, XFontStyleEx.Bold)

            Dim bBlack As XBrush = XBrushes.Black
            Dim bGray As XBrush = New XSolidBrush(XColor.FromArgb(100, 100, 100))
            Dim penBlack As XPen = New XPen(XColor.FromArgb(0, 0, 0), 0.6)
            Dim penGray As XPen = New XPen(XColor.FromArgb(180, 180, 180), 0.4)
            Dim penThick As XPen = New XPen(XColor.FromArgb(0, 0, 0), 1.0)

            Dim logoPfad As String = ToolPfade.LogoPfad
            If Not File.Exists(logoPfad) Then logoPfad = Path.Combine(speicherPfad, "logo.png")

            ' --- HILFSFUNKTIONEN FÜR KOPF/FUSS/TABELLE ---
            ' Als Einzelunternehmer gibt es keine Pflichtangaben für Geschäftsbriefe
            ' (die gelten nur für im Handelsregister eingetragene Rechtsformen) und
            ' auch keinen Geschäftsführer oder eine vom Firmensitz getrennte Geschäfts-
            ' stelle - deshalb nur eine schlanke, einzeilige Kontakt-Fußzeile.
            Dim fussZeile As String = String.Join("   ·   ",
                {firmaName, firmaStrasse, (firmaPLZ & " " & firmaOrt).Trim(),
                 If(String.IsNullOrEmpty(firmaTel), "", "Tel " & firmaTel),
                 If(String.IsNullOrEmpty(firmaHandy), "", "Handy/WhatsApp " & firmaHandy),
                 If(String.IsNullOrEmpty(firmaEmail), "", firmaEmail)}.Where(Function(t) Not String.IsNullOrWhiteSpace(t)))

            Dim drawFooter = Sub(g As XGraphics)
                                 Dim footY As Double = pH - 30
                                 g.DrawLine(penGray, mL, footY - 8, mR, footY - 8)
                                 Dim fw As Double = g.MeasureString(fussZeile, fTiny).Width
                                 g.DrawString(fussZeile, fTiny, bGray, mL + (cW - fw) / 2, footY)
                             End Sub

            Dim cBez As Double = mL + 2
            Dim cAnz As Double = mL + 280
            Dim cPre As Double = mL + 360
            Dim cMwSt As Double = mL + 440
            Dim maxDescWidth As Double = cAnz - cBez - 20

            Dim drawTableHeader = Function(g As XGraphics, topY As Double) As Double
                                      g.DrawLine(penThick, mL, topY, mR, topY)
                                      Dim hdrY As Double = topY + 13
                                      g.DrawString("Beschreibung", fBold, bBlack, cBez, hdrY)
                                      DrawRight(g, "Menge", fBold, bBlack, cAnz + 20, hdrY)
                                      DrawRight(g, "Einzelpreis", fBold, bBlack, cPre + 20, hdrY)
                                      DrawRight(g, "MwSt", fBold, bBlack, cMwSt + 15, hdrY)
                                      DrawRight(g, "Gesamt", fBold, bBlack, mR - 2, hdrY)

                                      Dim hdrLineY As Double = hdrY + 8
                                      g.DrawLine(penThick, mL, hdrLineY, mR, hdrLineY)
                                      Return hdrLineY + 15
                                  End Function

            ' Kornfeld-Akzentlinie oben auf jeder Seite
            Dim bKornblumenblau As XBrush = New XSolidBrush(XColor.FromArgb(61, 90, 128))
            Dim drawAkzentlinie = Sub(g As XGraphics)
                                       g.DrawRectangle(bKornblumenblau, 0, 0, p1.Width.Point, 5)
                                   End Sub

            Dim drawPage2Header = Sub(g As XGraphics)
                                      drawAkzentlinie(g)
                                      If File.Exists(logoPfad) Then
                                          Try
                                              Dim logoImg As XImage = XImage.FromFile(logoPfad)
                                              Dim maxWidth As Double = 140 ' 50% von Seite 1
                                              Dim maxHeight As Double = 47.5
                                              Dim imgW As Double = logoImg.PixelWidth
                                              Dim imgH As Double = logoImg.PixelHeight

                                              If imgW > maxWidth Then : imgH = imgH * (maxWidth / imgW) : imgW = maxWidth : End If
                                              If imgH > maxHeight Then : imgW = imgW * (maxHeight / imgH) : imgH = maxHeight : End If

                                              g.DrawImage(logoImg, mL, 22, imgW, imgH)
                                          Catch
                                          End Try
                                      Else
                                          Dim fErsatz As New XFont("Arial", 11, XFontStyleEx.Bold)
                                          g.DrawString(firmaName, fErsatz, bBlack, mL, 40)
                                      End If

                                      Dim centerX As Double = mL + cW / 2
                                      Dim drawC = Sub(txt As String, fnt As XFont, yy As Double)
                                                      Dim w = g.MeasureString(txt, fnt).Width
                                                      g.DrawString(txt, fnt, bBlack, centerX - w / 2, yy)
                                                  End Sub

                                      drawC("Rechnungsnr.: " & reNr, fNorm, 30)
                                      drawC("Rechnungsdatum: " & rechnungsdatum, fNorm, 42)
                                      drawC("Empfänger: " & kundenName, fBold, 54)
                                  End Sub

            ' --- SEITE 1 AUFBAU ---
            Dim currentGfx As XGraphics = XGraphics.FromPdfPage(p1)
            gfxList.Add(currentGfx)
            drawAkzentlinie(currentGfx)

            Dim logoX As Double = mL
            Dim logoY As Double = 22

            If File.Exists(logoPfad) Then
                Try
                    Dim logoImg As XImage = XImage.FromFile(logoPfad)
                    Dim maxW As Double = 280, maxH As Double = 95
                    Dim imgW As Double = logoImg.PixelWidth, imgH As Double = logoImg.PixelHeight
                    If imgW > maxW Then : imgH = imgH * (maxW / imgW) : imgW = maxW : End If
                    If imgH > maxH Then : imgW = imgW * (maxH / imgH) : imgH = maxH : End If
                    currentGfx.DrawImage(logoImg, logoX, logoY, imgW, imgH)
                Catch
                End Try
            Else
                Dim kornblumenblau As XColor = XColor.FromArgb(61, 90, 128)
                Dim bBlau As XBrush = New XSolidBrush(kornblumenblau)
                Dim fMonogramm As New XFont("Arial", 30, XFontStyleEx.Bold)
                Dim fFirmaGross As New XFont("Arial", 14, XFontStyleEx.Bold)
                currentGfx.DrawString("MH", fMonogramm, bBlau, logoX, logoY)
                If Not String.IsNullOrWhiteSpace(firmaName) Then
                    currentGfx.DrawString(firmaName, fFirmaGross, bBlack, logoX, logoY + 46)
                End If
            End If

            Dim infoX As Double = 340
            Dim infoY As Double = 20

            ' --- HIER WIRD DER NEUE TITEL GEZEICHNET (Mit automatischer Größenanpassung) ---
            Dim passenderFont As XFont = If(titelText = "RECHNUNG", fTitle, New XFont("Arial", 16, XFontStyleEx.Bold))
            currentGfx.DrawString(titelText, passenderFont, bBlack, infoX, infoY + 20)

            Dim lineY As Double = infoY + 30
            currentGfx.DrawLine(penBlack, infoX, lineY, mR, lineY)

            Dim rowHInfo As Double = 14
            Dim iY As Double = lineY + 12

            Dim drawInfoRow = Sub(lbl As String, val As String, bold As Boolean)
                                  currentGfx.DrawString(lbl, fSmallB, bBlack, infoX, iY)
                                  Dim vFont As XFont = If(bold, fSmallB, fSmall)
                                  DrawRight(currentGfx, val, vFont, bBlack, mR, iY)
                                  iY += rowHInfo
                              End Sub

            drawInfoRow("Rechnungsdatum:", rechnungsdatum, True)
            ' Lieferdatum ist eine Pflichtangabe nach §14 Abs. 4 Nr. 6 UStG, sofern es vom
            ' Rechnungsdatum abweicht - deshalb immer separat ausgewiesen.
            drawInfoRow("Lieferdatum:", lieferdatum, True)
            drawInfoRow("Rechnungsnr.:", reNr, True)
            If Not String.IsNullOrEmpty(kundenSteuer) Then drawInfoRow("Kd.-Steuernr.:", kundenSteuer, False)
            If Not String.IsNullOrEmpty(kundenBetrieb) Then drawInfoRow("Kd.-Betriebsnr.:", kundenBetrieb, False)
            If Not String.IsNullOrEmpty(firmaSteuer) Then drawInfoRow("Steuernr.:", firmaSteuer, False)
            If Not String.IsNullOrEmpty(firmaEmail) Then drawInfoRow("E-Mail:", firmaEmail, False)
            If Not String.IsNullOrEmpty(firmaIBAN) Then drawInfoRow("IBAN:", firmaIBAN, True)
            If Not String.IsNullOrEmpty(firmaBIC) Then drawInfoRow("BIC:", firmaBIC, False)
            If Not String.IsNullOrEmpty(firmaBank) Then drawInfoRow("Bank:", firmaBank, False)

            ' Kunden-E-Mail nur anzeigen, wenn eine gültige Adresse hinterlegt ist - mit einer
            ' Leerzeile von den eigenen (Rechnungs-)Daten abgesetzt, damit klar ist, dass es sich
            ' um den Versandweg an den Kunden handelt, nicht um eine weitere eigene Angabe.
            Dim emailRegex As New System.Text.RegularExpressions.Regex("^[^@\s]+@[^@\s]+\.[^@\s]+$")
            If Not String.IsNullOrEmpty(kundenEmail) AndAlso emailRegex.IsMatch(kundenEmail) Then
                iY += rowHInfo
                drawInfoRow("Versand an E-Mail:", kundenEmail, False)
            End If

            Dim fensterLeft As Double = 56
            Dim absenderY As Double = 135
            Dim addrY As Double = 150
            Dim absender As String = firmaName & " - " & firmaStrasse & " - " & firmaPLZ & " " & firmaOrt
            currentGfx.DrawString(absender, fTiny, bGray, fensterLeft, absenderY)
            currentGfx.DrawString(kundenName, fBold, bBlack, fensterLeft, addrY)
            currentGfx.DrawString(kundenStrasse, fNorm, bBlack, fensterLeft, addrY + 13)
            currentGfx.DrawString(kundenPLZ & " " & kundenOrt, fBold, bBlack, fensterLeft, addrY + 26)

            ' --- ARTIKEL SCHLEIFE MIT SEITENUMBRUCH ---
            Dim rowY As Double = drawTableHeader(currentGfx, 260)
            ' nettoGesamtLaufend bleibt für die interne Logik unangetastet (Netto ist und
            ' bleibt die Bezugsgröße für die Steuerberechnung im Summenblock unten).
            ' anzeigeGesamtLaufend ist der separate, rein optische "Übertrag"-Wert - zeigt je
            ' nach preisart Netto oder Brutto, passend zur Einzelpreis-/Gesamt-Spalte der Zeile.
            Dim nettoGesamtLaufend As Decimal = 0
            Dim anzeigeGesamtLaufend As Decimal = 0

            For i As Integer = 0 To positionen.Count - 1
                Dim pos = positionen(i)
                Dim zeileNetto As Decimal = pos.anz * pos.prs
                Dim einzelpreisAnzeige As Decimal = If(preisart = "Brutto", pos.prsBrutto, pos.prs)
                Dim zeileAnzeige As Decimal = If(preisart = "Brutto", pos.anz * pos.prsBrutto, zeileNetto)

                ' Prüfen, ob der nächste Artikel noch auf die Seite passt
                Dim neededHeight As Double = MeasureWrappedTextHeight(currentGfx, pos.bez, fNorm, maxDescWidth, 12) + 8
                If neededHeight < 15 Then neededHeight = 15

                If rowY + neededHeight > pH - 130 Then
                    ' 1. Tabelle beenden & Übertrag schreiben
                    currentGfx.DrawString("Übertrag:", fBold, bBlack, cMwSt - 20, rowY)
                    DrawRight(currentGfx, anzeigeGesamtLaufend.ToString("N2") & " €", fBold, bBlack, mR - 2, rowY)
                    Dim tBot As Double = rowY + 10
                    currentGfx.DrawLine(penThick, mL, tBot, mR, tBot)
                    drawFooter(currentGfx)

                    ' 2. Neue Seite anlegen
                    Dim pNew As PdfPage = document.AddPage()
                    pNew.Size = PdfSharp.PageSize.A4
                    pages.Add(pNew)
                    currentGfx = XGraphics.FromPdfPage(pNew)
                    gfxList.Add(currentGfx)

                    ' 3. Kopfzeile für Folgeseite
                    drawPage2Header(currentGfx)
                    rowY = drawTableHeader(currentGfx, 110) ' 3 cm Abstand (110)

                    ' 4. Übertrag oben reinschreiben
                    currentGfx.DrawString("Übertrag von Seite " & (pages.Count - 1), fNorm, bBlack, cBez, rowY)
                    DrawRight(currentGfx, anzeigeGesamtLaufend.ToString("N2") & " €", fNorm, bBlack, mR - 2, rowY)
                    rowY += 15
                End If

                ' Artikel zeichnen
                If pos.anz <> 0 AndAlso pos.prs <> 0 Then
                    ' "N0" rundete Bruchmengen (z.B. 2,5) auf ganze Zahlen und zeigte "3" an,
                    ' obwohl intern korrekt mit 2,5 gerechnet wurde - reiner Anzeigefehler.
                    ' Wie FormatMwSt: InvariantCulture + manuelles Komma, damit die Ausgabe nicht
                    ' vom Thread-Culture zur Druckzeit abhängt.
                    DrawRight(currentGfx, pos.anz.ToString("0.##", Globalization.CultureInfo.InvariantCulture).Replace(".", ","), fNorm, bBlack, cAnz + 20, rowY)
                    DrawRight(currentGfx, einzelpreisAnzeige.ToString("N2") & " €", fNorm, bBlack, cPre + 20, rowY)
                    DrawRight(currentGfx, FormatMwSt(pos.mwst), fNorm, bBlack, cMwSt + 15, rowY)
                    DrawRight(currentGfx, zeileAnzeige.ToString("N2") & " €", fNorm, bBlack, mR - 2, rowY)
                End If

                Dim nextY As Double = DrawWrappedText(currentGfx, pos.bez, fNorm, bBlack, cBez, rowY, maxDescWidth, 12)
                rowY = nextY + 8
                nettoGesamtLaufend += zeileNetto
                anzeigeGesamtLaufend += zeileAnzeige
            Next

            ' --- PRÜFEN OB SUMMENBLOCK NOCH PASST ---
            If rowY + 160 > pH - 110 Then
                ' Reicht nicht mehr für den Summenblock -> neue Seite nur für die Summe
                currentGfx.DrawString("Übertrag:", fBold, bBlack, cMwSt - 20, rowY)
                DrawRight(currentGfx, anzeigeGesamtLaufend.ToString("N2") & " €", fBold, bBlack, mR - 2, rowY)
                Dim tBot As Double = rowY + 10
                currentGfx.DrawLine(penThick, mL, tBot, mR, tBot)
                drawFooter(currentGfx)

                Dim pFinal As PdfPage = document.AddPage()
                pFinal.Size = PdfSharp.PageSize.A4
                pages.Add(pFinal)
                currentGfx = XGraphics.FromPdfPage(pFinal)
                gfxList.Add(currentGfx)

                drawPage2Header(currentGfx)
                rowY = drawTableHeader(currentGfx, 110)

                currentGfx.DrawString("Übertrag von Seite " & (pages.Count - 1), fNorm, bBlack, cBez, rowY)
                DrawRight(currentGfx, anzeigeGesamtLaufend.ToString("N2") & " €", fNorm, bBlack, mR - 2, rowY)
                rowY += 15
            End If

            ' Tabellen-Abschlusslinie
            Dim tableBottom As Double = rowY + 5
            currentGfx.DrawLine(penThick, mL, tableBottom, mR, tableBottom)

            ' --- SUMMENBLOCK (auf der letzten Seite) ---
            Dim sY As Double = tableBottom + 15
            Dim sLblX As Double = mL + 200
            Dim sEqX As Double = mL + 418
            Dim sGesX As Double = mR

            ' NEU: Eine feste, unsichtbare rechte Kante für die Wörter (kurz nach dem "=" Zeichen)
            Dim labelAlignX As Double = sEqX + 15

            DrawRight(currentGfx, "Nettosumme", fBold, bBlack, labelAlignX, sY)
            DrawRight(currentGfx, nettoGesamt.ToString("N2") & " €", fBold, bBlack, sGesX, sY)
            sY += 16

            Dim sMwstLblX As Double = sLblX + 50

            ' Steuern werden für jeden tatsächlich vorkommenden Satz angezeigt (nicht nur 7,8%/5,5%,
            ' falls z.B. eine ältere Rechnung noch einen anderen Satz enthält)
            For Each kv In nettoBasisProSatz.OrderByDescending(Function(x) x.Key)
                If kv.Value <> 0 Then
                    currentGfx.DrawString(FormatMwSt(kv.Key) & " MwSt auf", fSmall, bBlack, sMwstLblX, sY)
                    DrawRight(currentGfx, kv.Value.ToString("N2") & " €", fSmall, bBlack, sEqX - 5, sY)
                    currentGfx.DrawString("=", fSmall, bBlack, sEqX, sY)
                    DrawRight(currentGfx, mwstBetragProSatz(kv.Key).ToString("N2") & " €", fSmall, bBlack, sGesX, sY)
                    sY += 13
                End If
            Next

            currentGfx.DrawLine(penBlack, sLblX, sY - 2, mR, sY - 2)
            sY += 8

            ' --- GIROCODE (SEPA-Überweisungs-QR-Code, EPC-Standard) + Zahlungsdaten als Text ---
            ' Beginnt erst auf Höhe "Bruttosumme" (sY an dieser Stelle), nicht ganz oben beim
            ' Summenblock, und liegt links davon im freien Bereich bis sLblX (mL+200). QR-Code
            ' links, Zahlungsdaten (Kontoinhaber/IBAN/BIC/Betrag) rechts daneben, Bildunterschrift
            ' darunter über die volle Breite. qrBlockBottomY wird weiter unten mit der Endposition
            ' der rechten Spalte verglichen (Math.Max), damit der folgende Zahlungstext nie mit
            ' diesem Block kollidiert.
            ' Nur bei echten, positiven Rechnungen sinnvoll (nicht bei Gutschriften/0€-Rechnungen -
            ' der QR-Code fordert den KUNDEN zum Zahlen auf). Scheitert die Erzeugung aus
            ' irgendeinem Grund, darf das die restliche Rechnung nicht verhindern - deshalb
            ' Try/Catch ohne Weitergabe, analog zum Logo weiter oben.
            Dim qrBlockBottomY As Double = sY
            If bruttoGesamt > 0 AndAlso Not String.IsNullOrWhiteSpace(firmaIBAN) AndAlso Not String.IsNullOrWhiteSpace(firmaBIC) AndAlso Not String.IsNullOrWhiteSpace(firmaName) Then
                Try
                    Dim girocode As New PayloadGenerator.Girocode(
                        iban:=firmaIBAN.Replace(" ", ""),
                        bic:=firmaBIC.Replace(" ", ""),
                        name:=firmaName,
                        amount:=bruttoGesamt,
                        remittanceInformation:="re-" & reNr)

                    Using qrData = QRCodeGenerator.GenerateQrCode(girocode)
                        ' PngByteQRCode (QRCoders eigener, minimaler PNG-Encoder) erzeugt eine
                        ' PNG-Variante, die PdfSharps XImage.FromFile nicht als gültiges
                        ' Bildformat erkennt ("Unsupported image format"). Über den
                        ' System.Drawing.Bitmap-Renderer + GDI+ läuft es über den ganz normalen
                        ' Windows-PNG-Encoder, den PdfSharp garantiert lesen kann - genau wie
                        ' beim Firmenlogo weiter oben, das denselben Weg (Datei -> XImage.FromFile)
                        ' nimmt.
                        Using qrBitmap As System.Drawing.Bitmap = New QRCode(qrData).GetGraphic(20)
                            qrTempPfad = Path.Combine(Path.GetTempPath(), $"mhrechnung_qr_{reNr}.png")
                            qrBitmap.Save(qrTempPfad, System.Drawing.Imaging.ImageFormat.Png)
                        End Using

                        Dim qrBlockTopY As Double = sY
                        Dim qrGroesse As Double = 77
                        Dim rahmenPolster As Double = 7
                        ' Bis mL+300 ist noch reichlich Platz frei - "Bruttosumme"/"Rechnungsbetrag"
                        ' stehen rechtsbündig erst ab ca. mL+433, ihre Beschriftung beginnt erst gut
                        ' 90-100pt davor. Deutlich breiter als vorher (mL+200), damit z.B. die IBAN
                        ' in einer Zeile Platz hat, ohne mit den Summenzeilen zu kollidieren.
                        Dim rahmenRechts As Double = mL + 300
                        Dim qrX As Double = mL + rahmenPolster
                        ' Etwas mehr Luft oben zwischen Rahmen und Inhalt (ca. eine Leerzeile)
                        Dim qrY As Double = qrBlockTopY + rahmenPolster + 8
                        Dim qrImg As XImage = XImage.FromFile(qrTempPfad)
                        ' QR-Codes haben zwingend einen weißen Ruhebereich ("Quiet Zone") um das
                        ' eigentliche Muster, der Teil der Bilddatei ist - dadurch wirkt das
                        ' schwarze Muster optisch eingerückt, obwohl die Bildkante oben bündig mit
                        ' dem Text ist. Das Bild selbst etwas nach oben schieben, um das
                        ' auszugleichen; die Textspalte bleibt bei qrY (unverändert).
                        Dim qrBildY As Double = qrY - 10
                        currentGfx.DrawImage(qrImg, qrX, qrBildY, qrGroesse, qrGroesse)

                        ' Zahlungsdaten rechts neben dem QR-Code. Schriftgröße 8pt statt 7,7pt -
                        ' entspricht damit der Größe, die im restlichen Dokument für Zusatzinfos
                        ' (nicht Hauptinhalt) verwendet wird (fSmall/fSmallB), z.B. Kd.-Steuernr.,
                        ' BIC, Bank im Kopf-Info-Block. Zeilenabstand dazu passend auf 9pt.
                        Dim zdX As Double = qrX + qrGroesse + 6
                        Dim zdLineHeight As Double = 9
                        ' Eine Leerzeile vor "Kontoinhaber", damit der (kürzere) Textblock ungefähr
                        ' mittig neben dem höheren QR-Code steht, statt oben bündig damit.
                        Dim zdY As Double = qrY + zdLineHeight
                        Dim zdMaxWidth As Double = rahmenRechts - rahmenPolster - zdX
                        Dim fZahldaten As New XFont("Arial", 8, XFontStyleEx.Regular)
                        zdY = DrawWrappedText(currentGfx, "Kontoinhaber: " & firmaName, fZahldaten, bBlack, zdX, zdY, zdMaxWidth, zdLineHeight)
                        zdY = DrawWrappedText(currentGfx, "IBAN: " & firmaIBAN, fZahldaten, bBlack, zdX, zdY, zdMaxWidth, zdLineHeight)
                        zdY = DrawWrappedText(currentGfx, "BIC: " & firmaBIC, fZahldaten, bBlack, zdX, zdY, zdMaxWidth, zdLineHeight)
                        zdY = DrawWrappedText(currentGfx, "Betrag: " & bruttoGesamt.ToString("N2") & " €", fZahldaten, bBlack, zdX, zdY, zdMaxWidth, zdLineHeight)

                        ' Bildunterschrift über die volle Breite unter QR-Code + Zahlungsdaten
                        Dim capY As Double = Math.Max(qrY + qrGroesse, zdY) + 6
                        capY = DrawWrappedText(currentGfx, "Einfach mit der Banking-App scannen", fZahldaten, bGray, qrX, capY, rahmenRechts - rahmenPolster - qrX, zdLineHeight)

                        Dim rahmenUnten As Double = capY + rahmenPolster - 4
                        ' Abgerundete Ecken statt eckigem Rahmen, etwas dicker als die übrigen
                        ' dünnen Linien der Rechnung (penThick statt penBlack).
                        currentGfx.DrawRoundedRectangle(penThick, mL, qrBlockTopY, rahmenRechts - mL, rahmenUnten - qrBlockTopY, 10, 10)

                        qrBlockBottomY = rahmenUnten
                    End Using
                Catch ex As Exception
                    ' Temporäre Diagnose: der eigentliche Fehler wird sonst hier lautlos
                    ' verschluckt (wie beim Logo weiter oben) - bis der GiroCode zuverlässig
                    ' läuft, wird er zusätzlich in eine Textdatei geschrieben, die man einfach
                    ' nachschauen kann, statt raten zu müssen.
                    Try
                        File.WriteAllText(Path.Combine(Path.GetTempPath(), "mhrechnung_qr_fehler.txt"),
                            $"{DateTime.Now:dd.MM.yyyy HH:mm:ss} - Rechnung {reNr}:" & vbCrLf & ex.ToString())
                    Catch
                    End Try
                End Try
            End If

            ' Bündig zeichnen:
            DrawRight(currentGfx, "Bruttosumme", fBold, bBlack, labelAlignX, sY + 5)
            DrawRight(currentGfx, bruttoGesamt.ToString("N2") & " €", fBold, bBlack, sGesX, sY + 5)
            sY += 25

            ' Bündig zeichnen:
            DrawRight(currentGfx, summenLabel, fBold, bBlack, labelAlignX, sY)
            DrawRight(currentGfx, bruttoGesamt.ToString("N2") & " €", fBold, bBlack, sGesX, sY)
            sY += 25

            ' Verhindert eine Überlappung mit dem GiroCode-Block links, falls der (z.B. mit
            ' Zahlungsdaten-Text) höher wird als der Summenblock rechts.
            sY = Math.Max(sY, qrBlockBottomY + 10)

            ' --- ZAHLUNGS-/ GUTSCHRIFTS-TEXT ---
            If Not String.IsNullOrWhiteSpace(zahlungsText) Then
                Dim fSepa As New XFont("Arial", 8, XFontStyleEx.Bold)
                Dim words As String() = zahlungsText.Split(" "c)
                Dim currentLine As String = ""
                Dim textZeilen As New List(Of String)

                For Each word As String In words
                    Dim testLine As String = If(currentLine = "", word, currentLine & " " & word)
                    If currentGfx.MeasureString(testLine, fSepa).Width > cW Then
                        textZeilen.Add(currentLine)
                        currentLine = word
                    Else
                        currentLine = testLine
                    End If
                Next
                If currentLine <> "" Then textZeilen.Add(currentLine)

                For Each line As String In textZeilen
                    Dim lw As Double = currentGfx.MeasureString(line, fSepa).Width
                    currentGfx.DrawString(line, fSepa, bBlack, mL + (cW - lw) / 2, sY)
                    sY += 12
                Next
            End If

            ' Fußzeile auf der letzten Seite zeichnen
            drawFooter(currentGfx)

            ' --- FINALER DURCHLAUF: SEITENZAHLEN EINTRAGEN & DISPOSEN ---
            For i As Integer = 0 To pages.Count - 1
                Dim gfx As XGraphics = gfxList(i)
                If pages.Count > 1 Then
                    Dim seiteText As String = $"Seite {i + 1} von {pages.Count}"
                    If i = 0 Then
                        ' Seite 1: Rechts über dem Info-Block
                        DrawRight(gfx, seiteText, fNorm, bBlack, mR, 20)
                    Else
                        ' Seite 2+: Rechts oben in der Ecke
                        DrawRight(gfx, seiteText, fNorm, bBlack, mR, 22)
                    End If
                End If
                gfx.Dispose()
            Next

            document.Save(dateiName)
        End Using

        If qrTempPfad IsNot Nothing Then
            Try : File.Delete(qrTempPfad) : Catch : End Try
        End If
    End Sub

    ' Formatiert einen MwSt-Satz für die PDF-Anzeige, z.B. 7,8 -> "7,8%"
    Private Shared Function FormatMwSt(satz As Decimal) As String
        Return satz.ToString("0.0###", Globalization.CultureInfo.InvariantCulture).Replace(".", ",") & "%"
    End Function
End Class