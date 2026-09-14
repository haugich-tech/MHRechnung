Imports System.IO
Imports System.Xml
Imports System.Data.SQLite

Public Class SepaGenerator

    Private Shared Function GetSetting(conn As SQLiteConnection, key As String) As String
        Dim cmd As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = @k", conn)
        cmd.Parameters.AddWithValue("@k", key)
        Dim result = cmd.ExecuteScalar()
        Return If(result IsNot Nothing, result.ToString(), "")
    End Function

    ' Diese Methode bekommt die Liste der Rechnungen jetzt direkt aus der Form1 übergeben!
    Public Shared Sub ErstelleXML_Direkt(rechnungen As List(Of Dictionary(Of String, String)), sepaDir As String, dateiName As String)
        If rechnungen.Count = 0 Then Return

        Dim anzahlTx As Integer = 0
        Dim summeGesamt As Decimal = 0
        Dim msgId As String = "LEG-" & DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim einzugsDatum As String = DateTime.Now.AddDays(4).ToString("yyyy-MM-dd")

        Using conn = DatenbankManager.HoleVerbindung()
            Dim myName = GetSetting(conn, "firma_name")
            Dim myIban = GetSetting(conn, "firma_iban")
            Dim myBic = GetSetting(conn, "firma_bic")
            Dim myCreditorId = GetSetting(conn, "firma_glaeubiger")

            If String.IsNullOrWhiteSpace(myIban) OrElse String.IsNullOrWhiteSpace(myCreditorId) Then
                Throw New Exception("Für die SEPA-Datei fehlen deine eigene IBAN oder die Gläubiger-ID in den Einstellungen!")
            End If

            ' Zuerst die Summen und gültigen Transaktionen filtern (nur Betrag > 0 und mit IBAN)
            Dim sepaRechnungen As New List(Of Dictionary(Of String, String))
            For Each r In rechnungen
                Dim brutto As Decimal = CDec(r("brutto"))
                Dim iban As String = r("iban").ToString()
                If brutto > 0 AndAlso Not String.IsNullOrWhiteSpace(iban) Then
                    sepaRechnungen.Add(r)
                    summeGesamt += brutto
                    anzahlTx += 1
                End If
            Next

            If anzahlTx = 0 Then
                MessageBox.Show("Achtung: Es wurde keine SEPA-Datei erstellt, da bei den verarbeiteten Rechnungen entweder keine IBAN hinterlegt war oder der Betrag 0,00 € betrug.", "Kein SEPA-Export", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim fileName As String = Path.Combine(sepaDir, dateiName & ".xml")
            Dim settings As New XmlWriterSettings() With {.Indent = True, .Encoding = System.Text.Encoding.UTF8}

            Using writer As XmlWriter = XmlWriter.Create(fileName, settings)
                writer.WriteStartDocument()
                writer.WriteStartElement("Document", "urn:iso:std:iso:20022:tech:xsd:pain.008.001.02")
                writer.WriteStartElement("CstmrDrctDbtInitn")

                ' --- Group Header ---
                writer.WriteStartElement("GrpHdr")
                writer.WriteElementString("MsgId", msgId)
                writer.WriteElementString("CreDtTm", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                writer.WriteElementString("NbOfTxs", anzahlTx.ToString())
                writer.WriteElementString("CtrlSum", summeGesamt.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture))
                writer.WriteStartElement("InitgPty")
                writer.WriteElementString("Nm", myName)
                writer.WriteEndElement()
                writer.WriteEndElement()

                ' --- Payment Information (Gruppiert nach SEPA-Typ) ---
                Dim sepaGruppen = sepaRechnungen.GroupBy(Function(r) r("sTyp").ToString())

                For Each gruppe In sepaGruppen
                    Dim aktuellerTyp As String = gruppe.Key
                    Dim txInGruppe = gruppe.ToList()
                    Dim summeGruppe As Decimal = txInGruppe.Sum(Function(r) CDec(r("brutto")))
                    Dim anzahlGruppe As Integer = txInGruppe.Count

                    ' --- Payment Information Block (wird für jeden Typ einmal erstellt) ---
                    writer.WriteStartElement("PmtInf")
                    writer.WriteElementString("PmtInfId", "PMT-" & msgId & "-" & aktuellerTyp) ' Eindeutige ID pro Block
                    writer.WriteElementString("PmtMtd", "DD")
                    writer.WriteElementString("NbOfTxs", anzahlGruppe.ToString())
                    writer.WriteElementString("CtrlSum", summeGruppe.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture))

                    writer.WriteStartElement("PmtTpInf")
                    writer.WriteStartElement("SvcLvl")
                    writer.WriteElementString("Cd", "SEPA")
                    writer.WriteEndElement()
                    writer.WriteStartElement("LclInstrm")
                    writer.WriteElementString("Cd", "CORE")
                    writer.WriteEndElement()

                    ' HIER WIRD DER TYP JETZT EXAKT ZUGEWIESEN!
                    writer.WriteElementString("SeqTp", aktuellerTyp)

                    writer.WriteEndElement() ' PmtTpInf

                    writer.WriteElementString("ReqdColltnDt", einzugsDatum)

                    writer.WriteStartElement("Cdtr")
                    writer.WriteElementString("Nm", myName)
                    writer.WriteEndElement()

                    writer.WriteStartElement("CdtrAcct")
                    writer.WriteStartElement("Id")
                    writer.WriteElementString("IBAN", myIban)
                    writer.WriteEndElement()
                    writer.WriteEndElement()

                    writer.WriteStartElement("CdtrAgt")
                    writer.WriteStartElement("FinInstnId")
                    If Not String.IsNullOrWhiteSpace(myBic) Then
                        writer.WriteElementString("BIC", myBic)
                    Else
                        writer.WriteElementString("Othr", "NOTPROVIDED")
                    End If
                    writer.WriteEndElement()
                    writer.WriteEndElement()

                    writer.WriteElementString("ChrgBr", "SLEV")

                    writer.WriteStartElement("CdtrSchmeId")
                    writer.WriteStartElement("Id")
                    writer.WriteStartElement("PrvtId")
                    writer.WriteStartElement("Othr")
                    writer.WriteElementString("Id", myCreditorId)
                    writer.WriteStartElement("SchmeNm")
                    writer.WriteElementString("Prtry", "SEPA")
                    writer.WriteEndElement()
                    writer.WriteEndElement()
                    writer.WriteEndElement()
                    writer.WriteEndElement()
                    writer.WriteEndElement()

                    ' --- Transaktionen für DIESE Gruppe ---
                    For Each r In txInGruppe
                        writer.WriteStartElement("DrctDbtTxInf")

                        writer.WriteStartElement("PmtId")
                        writer.WriteElementString("EndToEndId", "RE-" & r("reNr"))
                        writer.WriteEndElement()

                        Dim betragString As String = Math.Round(CDec(r("brutto")), 2).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)

                        writer.WriteStartElement("InstdAmt")
                        writer.WriteAttributeString("Ccy", "EUR")
                        writer.WriteString(betragString)
                        writer.WriteEndElement()

                        ' Mandatsdatum ins ISO-Format wandeln (SEPA verlangt yyyy-MM-dd)
                        Dim mandatDatum As String = r("mDatum")
                        Dim geparst As DateTime
                        If DateTime.TryParseExact(mandatDatum, {"dd.MM.yyyy", "yyyy-MM-dd", "d.M.yyyy"},
                                                  System.Globalization.CultureInfo.InvariantCulture,
                                                  System.Globalization.DateTimeStyles.None, geparst) Then
                            mandatDatum = geparst.ToString("yyyy-MM-dd")
                        End If

                        writer.WriteStartElement("DrctDbtTx")
                        writer.WriteStartElement("MndtRltdInf")
                        writer.WriteElementString("MndtId", r("mandat"))
                        writer.WriteElementString("DtOfSgntr", mandatDatum)
                        writer.WriteEndElement()
                        writer.WriteEndElement()

                        writer.WriteStartElement("DbtrAgt")
                        writer.WriteStartElement("FinInstnId")
                        If Not String.IsNullOrWhiteSpace(r("bic")) Then
                            writer.WriteElementString("BIC", r("bic"))
                        Else
                            writer.WriteElementString("Othr", "NOTPROVIDED")
                        End If
                        writer.WriteEndElement()
                        writer.WriteEndElement()

                        writer.WriteStartElement("Dbtr")
                        writer.WriteElementString("Nm", r("name"))
                        writer.WriteEndElement()

                        writer.WriteStartElement("DbtrAcct")
                        writer.WriteStartElement("Id")
                        writer.WriteElementString("IBAN", r("iban").Replace(" ", "").ToUpper())
                        writer.WriteEndElement()
                        writer.WriteEndElement()

                        writer.WriteStartElement("RmtInf")
                        writer.WriteElementString("Ustrd", "Rechnung re-" & r("reNr") & " LEG Wertachtal")
                        writer.WriteEndElement()

                        writer.WriteEndElement() ' DrctDbtTxInf
                    Next

                    writer.WriteEndElement() ' PmtInf (Block für diese Gruppe beenden)
                Next ' Nächste Gruppe (falls vorhanden)

                writer.WriteEndElement() ' CstmrDrctDbtInitn
                writer.WriteEndElement() ' Document
            End Using
        End Using
    End Sub

End Class