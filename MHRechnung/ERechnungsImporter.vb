Imports System.IO
Imports System.Xml
Imports System.Windows.Forms
Imports System.Diagnostics

Public Class ERechnungsImporter

    ' =========================================================
    ' DATENSTRUKTUREN (Die "Pakete", die wir zurückgeben)
    ' =========================================================
    Public Class ImportedPosition
        Public Property Bezeichnung As String = ""
        Public Property EinzelpreisNetto As Decimal = 0
        Public Property MwStSatz As Decimal = 19D
    End Class

    Public Class ImportedInvoice
        Public Property GesamtBrutto As Decimal = 0
        Public Property Positionen As New List(Of ImportedPosition)
    End Class

    ' =========================================================
    ' HAUPTFUNKTION ZUM AUSLESEN DER DATEI
    ' =========================================================
    Public Shared Function LeseDatei(dateipfad As String) As ImportedInvoice
        Dim result As New ImportedInvoice()

        Try
            Dim xmlInhalt As String = ""

            ' ========================================================
            ' PDF-DATEIEN (Briefumschlag öffnen mit Mustang)
            ' ========================================================
            If dateipfad.ToLower().EndsWith(".pdf") Then

                ' 1. Pfade direkt aus deinem ToolPfade-Modul holen!
                Dim javaExe As String = ToolPfade.JavaExe
                Dim mustangJar As String = ToolPfade.MustangJar

                ' Sicherheitscheck: Sind die Dateien wirklich da?
                If Not File.Exists(javaExe) OrElse Not File.Exists(mustangJar) Then
                    MessageBox.Show("Das Mustang-Tool oder Java wurde nicht gefunden." & vbCrLf & "Bitte prüfe deine Tool-Ordnerstruktur.", "Systemfehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Nothing
                End If

                ' 2. Temporäre Datei für die extrahierte XML erzeugen
                Dim tempXmlPfad As String = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() & ".xml")

                ' 3. Java & Mustang unsichtbar im Hintergrund starten
                Dim psi As New ProcessStartInfo()
                psi.FileName = javaExe
                ' Wichtig: --action extract ist der Befehl für Mustang, um die XML aus der PDF zu holen
                psi.Arguments = $"-jar ""{mustangJar}"" --action extract --source ""{dateipfad}"" --out ""{tempXmlPfad}"""
                psi.WindowStyle = ProcessWindowStyle.Hidden
                psi.CreateNoWindow = True
                psi.UseShellExecute = False

                Using proc As Process = Process.Start(psi)
                    ' Timeout 60 Sekunden: verhindert, dass die UI dauerhaft einfriert wenn Java hängt
                    If Not proc.WaitForExit(60000) Then
                        proc.Kill()
                        MessageBox.Show("Die XML-Extraktion hat zu lange gedauert und wurde abgebrochen.", "Timeout", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return Nothing
                    End If
                End Using

                ' 4. Prüfen, ob eine XML gefunden und extrahiert wurde (leere Datei = Fehlschlag)
                If File.Exists(tempXmlPfad) AndAlso New FileInfo(tempXmlPfad).Length > 0 Then
                    xmlInhalt = File.ReadAllText(tempXmlPfad)
                    File.Delete(tempXmlPfad) ' Direkt wieder aufräumen
                Else
                    If File.Exists(tempXmlPfad) Then File.Delete(tempXmlPfad)
                    MessageBox.Show("In dieser PDF wurde keine gültige elektronische Rechnung (ZUGFeRD/XRechnung) gefunden.",
                                    "Keine E-Rechnung", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return Nothing
                End If

                ' ========================================================
                ' REINE XML-DATEIEN (Direkt einlesen)
                ' ========================================================
            ElseIf dateipfad.ToLower().EndsWith(".xml") Then
                xmlInhalt = File.ReadAllText(dateipfad)

                ' Falsches Format
            Else
                MessageBox.Show("Bitte nur XML- oder PDF-Dateien in diesen Bereich ziehen.", "Falsches Format", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return Nothing
            End If

            ' ========================================================
            ' XML AUSWERTEN (Für XRechnung und extrahierte ZUGFeRD)
            ' ========================================================
            Dim doc As New XmlDocument()
            doc.LoadXml(xmlInhalt)

            ' 1. Gesamtbetrag (Brutto) auslesen
            ' Wir suchen mit XPath nach dem Element, egal welchen Namespace es hat
            Dim bruttoNode As XmlNode = doc.SelectSingleNode("//*[local-name()='GrandTotalAmount']")
            If bruttoNode Is Nothing Then bruttoNode = doc.SelectSingleNode("//*[local-name()='PayableAmount']") ' Fallback UBL Format

            If bruttoNode IsNot Nothing Then
                Dim bruttoStr As String = bruttoNode.InnerText.Replace(".", ",")
                Decimal.TryParse(bruttoStr, result.GesamtBrutto)
            End If

            ' 2. Alle Artikel-Positionen auslesen
            ' UN/CEFACT Format = IncludedSupplyChainTradeLineItem | UBL Format = InvoiceLine
            Dim lineItems As XmlNodeList = doc.SelectNodes("//*[local-name()='IncludedSupplyChainTradeLineItem' or local-name()='InvoiceLine']")

            For Each node As XmlNode In lineItems
                Dim pos As New ImportedPosition()

                ' Name und Beschreibung holen
                Dim name As String = HoleNodeText(node, "Name")
                Dim desc As String = HoleNodeText(node, "Description")

                ' Titel und Beschreibung intelligent mit Zeilenumbruch zusammenführen
                If Not String.IsNullOrWhiteSpace(name) AndAlso Not String.IsNullOrWhiteSpace(desc) Then
                    pos.Bezeichnung = name & vbCrLf & desc
                ElseIf Not String.IsNullOrWhiteSpace(name) Then
                    pos.Bezeichnung = name
                Else
                    pos.Bezeichnung = desc
                End If

                ' Einzelpreis Netto holen
                Dim preisStr As String = HoleNodeText(node, "ChargeAmount")
                If String.IsNullOrEmpty(preisStr) Then preisStr = HoleNodeText(node, "PriceAmount") ' Fallback
                If Not String.IsNullOrEmpty(preisStr) Then
                    Decimal.TryParse(preisStr.Replace(".", ","), pos.EinzelpreisNetto)
                End If

                ' Mehrwertsteuer holen (Standard ist 19%, falls nichts gefunden wird)
                Dim mwstStr As String = HoleNodeText(node, "RateApplicablePercent")
                If String.IsNullOrEmpty(mwstStr) Then mwstStr = HoleNodeText(node, "Percent") ' Fallback
                If Not String.IsNullOrEmpty(mwstStr) Then
                    Dim mwstDecimal As Decimal = 19
                    Decimal.TryParse(mwstStr.Replace(".", ","), mwstDecimal)
                    pos.MwStSatz = mwstDecimal
                End If

                result.Positionen.Add(pos)
            Next

        Catch ex As Exception
            MessageBox.Show("Fehler beim Auslesen der E-Rechnung: " & ex.Message, "Import-Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

        Return result
    End Function

    ' Hilfsfunktion, um Werte tief im XML ohne Namensraum-Stress zu finden
    Private Shared Function HoleNodeText(parentNode As XmlNode, tagName As String) As String
        Dim node As XmlNode = parentNode.SelectSingleNode(".//*[local-name()='" & tagName & "']")
        If node IsNot Nothing Then
            Return node.InnerText.Trim()
        End If
        Return ""
    End Function

End Class