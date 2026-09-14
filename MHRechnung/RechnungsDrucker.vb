Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports PdfSharp.Fonts
Imports System.IO
Imports System.Data.SQLite

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

        Dim speicherPfad As String = GetSetting(s, "speicherpfad", "C:\LEG_Rechnungen")
        Dim firmaName As String = GetSetting(s, "firma_name", "LEG Wertachtal GBR")
        Dim firmaStrasse As String = GetSetting(s, "firma_strasse", "Augsburger Str. 40")
        Dim firmaPLZ As String = GetSetting(s, "firma_plz", "86842")
        Dim firmaOrt As String = GetSetting(s, "firma_ort", "Türkheim")
        Dim firmaEmail As String = GetSetting(s, "firma_email", "")
        Dim firmaTel As String = GetSetting(s, "firma_tel", "")
        Dim firmaSteuer As String = GetSetting(s, "firma_steuer", "")
        Dim firmaIBAN As String = GetSetting(s, "firma_iban", "")
        Dim firmaBIC As String = GetSetting(s, "firma_bic", "")
        Dim firmaBank As String = GetSetting(s, "firma_bank", "")
        Dim firmaGlaeubiger As String = GetSetting(s, "firma_glaeubiger", "")

        Dim textZahlung As String = GetSetting(s, "text_zahlung", "")
        Dim textGutschrift As String = GetSetting(s, "text_gutschrift", "")

        Dim fussGsstName As String = GetSetting(s, "fuss_gst_name", "Büro")
        Dim fussSitz As String = GetSetting(s, "fuss_sitz", "Sitz der Gesellschaft " & firmaOrt)
        Dim fussGericht As String = GetSetting(s, "fuss_gericht", "Gerichtsstand Memmingen")
        Dim fussGF1 As String = GetSetting(s, "fuss_gf1_name", "")
        Dim fussGF1Tel As String = GetSetting(s, "fuss_gf1_tel", "")
        Dim fussGF2 As String = GetSetting(s, "fuss_gf2_name", "")
        Dim fussGF2Tel As String = GetSetting(s, "fuss_gf2_tel", "")

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
        Dim kundenIBAN As String = ""
        Dim kundenMandat As String = ""
        Dim kundenSteuer As String = ""
        Dim kundenBetrieb As String = ""
        Dim positionen As New List(Of (bez As String, anz As Decimal, prs As Decimal, mwst As Integer))

        Using conn = DatenbankManager.HoleVerbindung()
            Dim sqlK = "SELECT m.name, m.strasse, m.plz, m.ort, m.iban, m.mitgliedsnummer, m.steuernummer, m.betriebsnummer " &
                       "FROM rechnungen r JOIN mitglieder m ON r.mitglied_id = m.id WHERE r.id = @id"
            Dim cmdK As New SQLiteCommand(sqlK, conn)
            cmdK.Parameters.AddWithValue("@id", reID)
            Using r = cmdK.ExecuteReader()
                If r.Read() Then
                    kundenName = r("name").ToString()
                    kundenStrasse = r("strasse").ToString()
                    kundenPLZ = r("plz").ToString()
                    kundenOrt = r("ort").ToString()
                    kundenIBAN = r("iban").ToString()
                    kundenMandat = r("mitgliedsnummer").ToString()
                    kundenSteuer = r("steuernummer").ToString()
                    kundenBetrieb = r("betriebsnummer").ToString()
                End If
            End Using

            Dim cmdP As New SQLiteCommand("SELECT artikel_bezeichnung, anzahl, einzelpreis, mwst_satz FROM rechnungspositionen WHERE rechnung_id = @id", conn)
            cmdP.Parameters.AddWithValue("@id", reID)
            Using r = cmdP.ExecuteReader()
                While r.Read()
                    positionen.Add((r("artikel_bezeichnung").ToString(), CDec(r("anzahl")), CDec(r("einzelpreis")), CInt(r("mwst_satz"))))
                End While
            End Using
        End Using

        Dim nettoGesamt As Decimal = 0, nettoBasis7 As Decimal = 0, nettoBasis19 As Decimal = 0
        Dim mwst7 As Decimal = 0, mwst19 As Decimal = 0

        For Each pos In positionen
            ' Zeilennetto auf 2 Stellen runden - identische Logik wie im ZUGFeRD-XML
            Dim zn As Decimal = Math.Round(pos.anz * pos.prs, 2, MidpointRounding.AwayFromZero)
            nettoGesamt += zn
            If pos.mwst = 7 Then
                nettoBasis7 += zn
            ElseIf pos.mwst = 19 Then
                nettoBasis19 += zn
            End If
            ' Andere Sätze (z.B. 0% steuerfrei) bekommen keine MwSt aufgeschlagen
        Next
        ' Steuer auf die gerundete Basis berechnen - PDF stimmt damit exakt mit XML und SEPA überein
        mwst7 = Math.Round(nettoBasis7 * 0.07D, 2, MidpointRounding.AwayFromZero)
        mwst19 = Math.Round(nettoBasis19 * 0.19D, 2, MidpointRounding.AwayFromZero)
        Dim bruttoGesamt As Decimal = nettoGesamt + mwst7 + mwst19

        ' --- NEUE GUTSCHRIFTEN LOGIK ---
        Dim titelText As String = "RECHNUNG"
        Dim summenLabel As String = "Abbuchungsbetrag"
        Dim basisText As String = textZahlung

        If bruttoGesamt < 0 Then
            titelText = "GUTSCHRIFT / KORREKTUR"
            summenLabel = "Guthaben / Auszahlungsbetrag"
            basisText = textGutschrift
        ElseIf bruttoGesamt = 0 Then
            titelText = "RECHNUNG"
            summenLabel = "Rechnungsbetrag"
            basisText = "Der Rechnungsbetrag beläuft sich auf 0,00 €. Diese Rechnung dient lediglich zur Information/Korrektur, es ist keine Zahlung oder Abbuchung erforderlich."
        End If

        Dim sepaText As String = basisText _
            .Replace("[RE-nummer]", "re-" & reNr) _
            .Replace("[Kunden-IBAN]", If(kundenIBAN = "", "—", kundenIBAN)) _
            .Replace("[Kunden-Mandat]", If(kundenMandat = "", "—", kundenMandat)) _
            .Replace("[Gläubiger-ID]", If(firmaGlaeubiger = "", "—", firmaGlaeubiger))

        ' ── 3. PDF ZEICHNEN (MULTI-PAGE LOGIK) ────────────────────────────────
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
            Dim fTinyB As New XFont("Arial", 7, XFontStyleEx.Bold)
            Dim fTitle As New XFont("Arial", 28, XFontStyleEx.Bold)

            Dim bBlack As XBrush = XBrushes.Black
            Dim bGray As XBrush = New XSolidBrush(XColor.FromArgb(100, 100, 100))
            Dim penBlack As XPen = New XPen(XColor.FromArgb(0, 0, 0), 0.6)
            Dim penGray As XPen = New XPen(XColor.FromArgb(180, 180, 180), 0.4)
            Dim penThick As XPen = New XPen(XColor.FromArgb(0, 0, 0), 1.0)

            Dim logoPfad As String = ToolPfade.LogoPfad
            If Not File.Exists(logoPfad) Then logoPfad = Path.Combine(speicherPfad, "logo.png")

            ' --- HILFSFUNKTIONEN FÜR KOPF/FUSS/TABELLE ---
            Dim drawFooter = Sub(g As XGraphics)
                                 Dim footY As Double = pH - 52
                                 Dim textYOffset As Double = 6
                                 Dim col1 As Double = mL
                                 Dim col2 As Double = mL + cW / 3
                                 Dim col3 As Double = mL + cW / 3 * 2

                                 g.DrawLine(penGray, mL, footY - 4, mR, footY - 4)
                                 g.DrawString("Geschäftsstelle", fTinyB, bBlack, col1, footY + textYOffset)
                                 g.DrawString(fussGsstName, fTiny, bBlack, col1, footY + 9 + textYOffset)
                                 g.DrawString(firmaStrasse, fTiny, bBlack, col1, footY + 18 + textYOffset)
                                 g.DrawString(firmaPLZ & " " & firmaOrt, fTiny, bBlack, col1, footY + 27 + textYOffset)
                                 If Not String.IsNullOrEmpty(firmaTel) Then g.DrawString("Tel: " & firmaTel, fTiny, bBlack, col1, footY + 36 + textYOffset)

                                 g.DrawString(fussSitz, fTiny, bBlack, col2, footY + textYOffset)
                                 g.DrawString(fussGericht, fTiny, bBlack, col2, footY + 9 + textYOffset)

                                 If Not String.IsNullOrEmpty(fussGF1) Then
                                     g.DrawString("1. Geschäftsführer " & fussGF1, fTiny, bBlack, col3, footY + textYOffset)
                                     If Not String.IsNullOrEmpty(fussGF1Tel) Then g.DrawString("Tel: " & fussGF1Tel, fTiny, bBlack, col3, footY + 9 + textYOffset)
                                 End If
                                 If Not String.IsNullOrEmpty(fussGF2) Then
                                     g.DrawString("2. Geschäftsführer " & fussGF2, fTiny, bBlack, col3, footY + 18 + textYOffset)
                                     If Not String.IsNullOrEmpty(fussGF2Tel) Then g.DrawString("Tel: " & fussGF2Tel, fTiny, bBlack, col3, footY + 27 + textYOffset)
                                 End If
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

            Dim drawPage2Header = Sub(g As XGraphics)
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
                                      drawC("Rechnungsdatum: " & DateTime.Now.ToString("dd.MM.yyyy"), fNorm, 42)
                                      drawC("Empfänger: " & kundenName, fBold, 54)
                                  End Sub

            ' --- SEITE 1 AUFBAU ---
            Dim currentGfx As XGraphics = XGraphics.FromPdfPage(p1)
            gfxList.Add(currentGfx)

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
                Dim darkGreen As XColor = XColor.FromArgb(30, 110, 30)
                Dim bGreen As XBrush = New XSolidBrush(darkGreen)
                Dim fLEGBig As New XFont("Arial", 13, XFontStyleEx.Bold)
                Dim fLEGSub As New XFont("Arial", 7, XFontStyleEx.Regular)
                Dim fWert As New XFont("Arial", 22, XFontStyleEx.Bold)
                currentGfx.DrawString("L", fLEGBig, bGreen, logoX, logoY)
                currentGfx.DrawString("andwirtschaftliche", fLEGSub, bBlack, logoX + 14, logoY - 2)
                currentGfx.DrawString("E", fLEGBig, bGreen, logoX, logoY + 14)
                currentGfx.DrawString("inkaufs", fLEGSub, bBlack, logoX + 14, logoY + 12)
                currentGfx.DrawString("G", fLEGBig, bGreen, logoX, logoY + 28)
                currentGfx.DrawString("emeinschaft", fLEGSub, bBlack, logoX + 14, logoY + 26)
                currentGfx.DrawString("Wertachtal", fWert, bGreen, logoX - 2, logoY + 52)
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

            drawInfoRow("Rechnungsdatum:", DateTime.Now.ToString("dd.MM.yyyy"), True)
            drawInfoRow("Rechnungsnr.:", reNr, True)
            If Not String.IsNullOrEmpty(kundenSteuer) Then drawInfoRow("Kd.-Steuernr.:", kundenSteuer, False)
            If Not String.IsNullOrEmpty(kundenBetrieb) Then drawInfoRow("Kd.-Betriebsnr.:", kundenBetrieb, False)
            drawInfoRow("LEG-Steuernr.:", firmaSteuer, False)
            drawInfoRow("LEG-Email:", firmaEmail, False)
            drawInfoRow("IBAN:", firmaIBAN, True)
            drawInfoRow("BIC:", firmaBIC, False)
            If Not String.IsNullOrEmpty(firmaBank) Then drawInfoRow("Bank:", firmaBank, False)

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
            Dim nettoGesamtLaufend As Decimal = 0

            For i As Integer = 0 To positionen.Count - 1
                Dim pos = positionen(i)
                Dim zeileNetto As Decimal = pos.anz * pos.prs

                ' Prüfen, ob der nächste Artikel noch auf die Seite passt
                Dim neededHeight As Double = MeasureWrappedTextHeight(currentGfx, pos.bez, fNorm, maxDescWidth, 12) + 8
                If neededHeight < 15 Then neededHeight = 15

                If rowY + neededHeight > pH - 130 Then
                    ' 1. Tabelle beenden & Übertrag schreiben
                    currentGfx.DrawString("Übertrag:", fBold, bBlack, cMwSt - 20, rowY)
                    DrawRight(currentGfx, nettoGesamtLaufend.ToString("N2") & " €", fBold, bBlack, mR - 2, rowY)
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
                    DrawRight(currentGfx, nettoGesamtLaufend.ToString("N2") & " €", fNorm, bBlack, mR - 2, rowY)
                    rowY += 15
                End If

                ' Artikel zeichnen
                If pos.anz <> 0 AndAlso pos.prs <> 0 Then
                    DrawRight(currentGfx, pos.anz.ToString("N0"), fNorm, bBlack, cAnz + 20, rowY)
                    DrawRight(currentGfx, pos.prs.ToString("N2") & " €", fNorm, bBlack, cPre + 20, rowY)
                    DrawRight(currentGfx, pos.mwst.ToString() & "%", fNorm, bBlack, cMwSt + 15, rowY)
                    DrawRight(currentGfx, zeileNetto.ToString("N2") & " €", fNorm, bBlack, mR - 2, rowY)
                End If

                Dim nextY As Double = DrawWrappedText(currentGfx, pos.bez, fNorm, bBlack, cBez, rowY, maxDescWidth, 12)
                rowY = nextY + 8
                nettoGesamtLaufend += zeileNetto
            Next

            ' --- PRÜFEN OB SUMMENBLOCK NOCH PASST ---
            If rowY + 160 > pH - 110 Then
                ' Reicht nicht mehr für den Summenblock -> neue Seite nur für die Summe
                currentGfx.DrawString("Übertrag:", fBold, bBlack, cMwSt - 20, rowY)
                DrawRight(currentGfx, nettoGesamtLaufend.ToString("N2") & " €", fBold, bBlack, mR - 2, rowY)
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
                DrawRight(currentGfx, nettoGesamtLaufend.ToString("N2") & " €", fNorm, bBlack, mR - 2, rowY)
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

            ' --- KORREKTUR: STEUERN WERDEN IMMER ANGEZEIGT WENN UNGLEICH 0 ---
            If mwst7 <> 0 Then
                currentGfx.DrawString("7% MwSt auf", fSmall, bBlack, sMwstLblX, sY)
                DrawRight(currentGfx, nettoBasis7.ToString("N2") & " €", fSmall, bBlack, sEqX - 5, sY)
                currentGfx.DrawString("=", fSmall, bBlack, sEqX, sY)
                DrawRight(currentGfx, mwst7.ToString("N2") & " €", fSmall, bBlack, sGesX, sY)
                sY += 13
            End If
            If mwst19 <> 0 Then
                currentGfx.DrawString("19% MwSt auf", fSmall, bBlack, sMwstLblX, sY)
                DrawRight(currentGfx, nettoBasis19.ToString("N2") & " €", fSmall, bBlack, sEqX - 5, sY)
                currentGfx.DrawString("=", fSmall, bBlack, sEqX, sY)
                DrawRight(currentGfx, mwst19.ToString("N2") & " €", fSmall, bBlack, sGesX, sY)
                sY += 13
            End If

            currentGfx.DrawLine(penBlack, sLblX, sY - 2, mR, sY - 2)
            sY += 8

            ' Bündig zeichnen:
            DrawRight(currentGfx, "Bruttosumme", fBold, bBlack, labelAlignX, sY + 5)
            DrawRight(currentGfx, bruttoGesamt.ToString("N2") & " €", fBold, bBlack, sGesX, sY + 5)
            sY += 25

            ' Bündig zeichnen:
            DrawRight(currentGfx, summenLabel, fBold, bBlack, labelAlignX, sY)
            DrawRight(currentGfx, bruttoGesamt.ToString("N2") & " €", fBold, bBlack, sGesX, sY)
            sY += 25

            ' --- SEPA-TEXT / GUTSCHRIFTS-TEXT ---
            If Not String.IsNullOrWhiteSpace(sepaText) Then
                Dim fSepa As New XFont("Arial", 8, XFontStyleEx.Bold)
                Dim words As String() = sepaText.Split(" "c)
                Dim currentLine As String = ""
                Dim sepaLines As New List(Of String)

                For Each word As String In words
                    Dim testLine As String = If(currentLine = "", word, currentLine & " " & word)
                    If currentGfx.MeasureString(testLine, fSepa).Width > cW Then
                        sepaLines.Add(currentLine)
                        currentLine = word
                    Else
                        currentLine = testLine
                    End If
                Next
                If currentLine <> "" Then sepaLines.Add(currentLine)

                For Each line As String In sepaLines
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
    End Sub
End Class