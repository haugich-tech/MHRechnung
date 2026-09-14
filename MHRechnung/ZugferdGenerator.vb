Imports System.IO
Imports System.Text
Imports System.Data.SQLite
Imports System.Globalization

Public Class ZugferdGenerator

    Public Shared Sub ErstelleXML(rechnungsId As Integer, reNr As String, mitgliedId As Integer)
        Dim baseDir As String = "C:\LEG_Rechnungen"

        ' Firmen-Einstellungen (Verkäufer)
        Dim fName As String = "", fStrasse As String = "", fPlz As String = "", fOrt As String = "", fLand As String = "DE"
        Dim fIban As String = "", fBic As String = "", fSteuer As String = "", fGlaeubiger As String = ""
        Dim fEmail As String = "", fTel As String = ""
        Dim textZahlung As String = ""
        Dim textGutschrift As String = ""

        ' Kundendaten (Käufer)
        Dim kName As String = "", kStrasse As String = "", kPlz As String = "", kOrt As String = "", kLand As String = "DE"
        Dim kIban As String = "", kMandat As String = "", kMitgliedsNr As String = "", kEmail As String = ""

        Dim rDatum As Date = DateTime.Now
        Dim positionen As New List(Of Dictionary(Of String, Object))
        Dim summeNetto As Decimal = 0
        Dim summeBrutto As Decimal = 0

        Dim steuerBasis7 As Decimal = 0, steuerBetrag7 As Decimal = 0
        Dim steuerBasis19 As Decimal = 0, steuerBetrag19 As Decimal = 0

        ' Schalter für die Steuerblöcke (wichtig für Null-Rechnungen)
        Dim hatMwst7 As Boolean = False
        Dim hatMwst19 As Boolean = False

        Using conn = DatenbankManager.HoleVerbindung()
            ' --- Einstellungen laden ---
            Dim cmdEinst As New SQLiteCommand("SELECT schluessel, wert FROM einstellungen", conn)
            Using reader = cmdEinst.ExecuteReader()
                While reader.Read()
                    Dim key As String = reader("schluessel").ToString()
                    Dim val As String = reader("wert").ToString()
                    Select Case key
                        Case "speicherpfad" : If Not String.IsNullOrWhiteSpace(val) Then baseDir = val
                        Case "firma_name" : fName = val
                        Case "firma_strasse" : fStrasse = val
                        Case "firma_plz" : fPlz = val
                        Case "firma_ort" : fOrt = val
                        Case "firma_iban" : fIban = val
                        Case "firma_bic" : fBic = val
                        Case "firma_steuer" : fSteuer = val
                        Case "firma_glaeubiger" : fGlaeubiger = val
                        Case "firma_email" : fEmail = val
                        Case "firma_tel" : fTel = val
                        Case "text_zahlung" : textZahlung = val
                        Case "text_gutschrift" : textGutschrift = val
                    End Select
                End While
            End Using

            ' --- Rechnungsdatum ---
            Dim cmdReDatum As New SQLiteCommand("SELECT datum FROM rechnungen WHERE id = @id", conn)
            cmdReDatum.Parameters.AddWithValue("@id", rechnungsId)
            Dim datumStr = cmdReDatum.ExecuteScalar()?.ToString()
            If Not String.IsNullOrWhiteSpace(datumStr) Then
                Date.TryParseExact(datumStr, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, rDatum)
            End If

            ' --- Kundendaten ---
            Dim cmdKd As New SQLiteCommand("SELECT name, strasse, plz, ort, land_code, iban, mitgliedsnummer, mandat_datum, email FROM mitglieder WHERE id = @mid", conn)
            cmdKd.Parameters.AddWithValue("@mid", mitgliedId)
            Using reader = cmdKd.ExecuteReader()
                If reader.Read() Then
                    kName = reader("name").ToString()
                    kStrasse = reader("strasse").ToString()
                    kPlz = reader("plz").ToString()
                    kOrt = reader("ort").ToString()
                    Dim lCode = reader("land_code").ToString()
                    If Not String.IsNullOrWhiteSpace(lCode) Then kLand = lCode
                    kIban = reader("iban").ToString().Replace(" ", "")
                    kMitgliedsNr = reader("mitgliedsnummer").ToString()
                    kMandat = reader("mandat_datum").ToString()
                    kEmail = reader("email")?.ToString()
                End If
            End Using

            ' --- Positionen laden ---
            Dim cmdPos As New SQLiteCommand("SELECT artikel_bezeichnung, anzahl, einzelpreis, mwst_satz FROM rechnungspositionen WHERE rechnung_id = @id", conn)
            cmdPos.Parameters.AddWithValue("@id", rechnungsId)
            Using reader = cmdPos.ExecuteReader()
                While reader.Read()
                    Dim bez As String = reader("artikel_bezeichnung").ToString()
                    Dim anz As Decimal = CDec(reader("anzahl"))
                    Dim prs As Decimal = CDec(reader("einzelpreis"))
                    Dim mwst As Integer = CInt(reader("mwst_satz"))
                    ' Zeilennetto auf 2 Stellen runden - Grundlage für alle weiteren Summen
                    Dim posNetto As Decimal = Math.Round(anz * prs, 2, MidpointRounding.AwayFromZero)
                    summeNetto += posNetto

                    If mwst = 7 Then
                        hatMwst7 = True
                        steuerBasis7 += posNetto
                    ElseIf mwst = 19 Then
                        hatMwst19 = True
                        steuerBasis19 += posNetto
                    End If

                    Dim posDict As New Dictionary(Of String, Object)
                    posDict("bez") = bez
                    posDict("anz") = anz
                    posDict("prs") = prs
                    posDict("netto") = posNetto
                    posDict("mwst") = mwst
                    positionen.Add(posDict)
                End While
            End Using
            ' Steuer auf die gerundete Basis berechnen und runden (EN 16931: BR-CO-14/15)
            ' Damit gilt exakt: Basis + Steuer = Gesamtbetrag
            steuerBetrag7 = Math.Round(steuerBasis7 * 0.07D, 2, MidpointRounding.AwayFromZero)
            steuerBetrag19 = Math.Round(steuerBasis19 * 0.19D, 2, MidpointRounding.AwayFromZero)
            summeBrutto = summeNetto + steuerBetrag7 + steuerBetrag19
        End Using

        ' --- LOGIK FÜR GUTSCHRIFTEN (Multiplikator) ---
        Dim isGutschrift As Boolean = (summeBrutto < 0)
        Dim docTypeCode As String = If(isGutschrift, "381", "380")
        Dim signMult As Decimal = If(isGutschrift, -1D, 1D)

        ' --- Pfade vorbereiten ---
        Dim xmlOrdner As String = Path.Combine(baseDir, rDatum.Year.ToString(), "erstellt", "xml")
        If Not Directory.Exists(xmlOrdner) Then Directory.CreateDirectory(xmlOrdner)
        Dim tempOrdner As String = Path.Combine(baseDir, "temp")
        If Not Directory.Exists(tempOrdner) Then Directory.CreateDirectory(tempOrdner)

        Dim archivXmlPfad As String = Path.Combine(xmlOrdner, reNr & ".xml")
        ' Rechnungsnummer im Dateinamen: verhindert, dass bei mehreren Rechnungen
        ' das XML der einen in das PDF der anderen eingebettet wird
        Dim tempXmlPfad As String = Path.Combine(tempOrdner, "factur-x_" & reNr & ".xml")
        Dim formatDatum As String = rDatum.ToString("yyyyMMdd")

        ' --- XML aufbauen ---
        Dim xml As New StringBuilder()
        xml.AppendLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
        xml.AppendLine("<rsm:CrossIndustryInvoice xmlns:rsm=""urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100"" xmlns:ram=""urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100"" xmlns:udt=""urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100"">")

        ' Kontext 
        xml.AppendLine("  <rsm:ExchangedDocumentContext>")
        xml.AppendLine("    <ram:BusinessProcessSpecifiedDocumentContextParameter>")
        xml.AppendLine("      <ram:ID>urn:fdc:peppol.eu:2017:poacc:billing:01:1.0</ram:ID>")
        xml.AppendLine("    </ram:BusinessProcessSpecifiedDocumentContextParameter>")
        xml.AppendLine("    <ram:GuidelineSpecifiedDocumentContextParameter>")
        xml.AppendLine("      <ram:ID>urn:cen.eu:en16931:2017#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0</ram:ID>")
        xml.AppendLine("    </ram:GuidelineSpecifiedDocumentContextParameter>")
        xml.AppendLine("  </rsm:ExchangedDocumentContext>")

        ' Dokumenten-Info
        xml.AppendLine("  <rsm:ExchangedDocument>")
        xml.AppendLine("    <ram:ID>" & XE(reNr) & "</ram:ID>")
        xml.AppendLine("    <ram:TypeCode>" & docTypeCode & "</ram:TypeCode>")
        xml.AppendLine("    <ram:IssueDateTime><udt:DateTimeString format=""102"">" & formatDatum & "</udt:DateTimeString></ram:IssueDateTime>")
        xml.AppendLine("  </rsm:ExchangedDocument>")

        xml.AppendLine("  <rsm:SupplyChainTradeTransaction>")

        ' ── RECHNUNGSPOSITIONEN ──
        Dim lineId As Integer = 1
        Dim lineTotalSum As Decimal = 0
        Dim allowanceSum As Decimal = 0

        For Each pos In positionen
            Dim anz As Decimal = CDec(pos("anz"))
            Dim prs As Decimal = CDec(pos("prs"))
            Dim netto As Decimal = CDec(pos("netto"))
            Dim mwst As Integer = CInt(pos("mwst"))

            Dim finalNetto As Decimal = netto * signMult

            ' XRechnung BR-27: Nur positive Artikel dürfen als normale Zeilen auftauchen!
            If finalNetto >= 0 Then
                lineTotalSum += finalNetto

                xml.AppendLine("    <ram:IncludedSupplyChainTradeLineItem>")
                xml.AppendLine("      <ram:AssociatedDocumentLineDocument>")
                xml.AppendLine("        <ram:LineID>" & lineId & "</ram:LineID>")
                xml.AppendLine("      </ram:AssociatedDocumentLineDocument>")
                xml.AppendLine("      <ram:SpecifiedTradeProduct>")
                xml.AppendLine("        <ram:Name>" & XE(CStr(pos("bez"))) & "</ram:Name>")
                xml.AppendLine("      </ram:SpecifiedTradeProduct>")
                xml.AppendLine("      <ram:SpecifiedLineTradeAgreement>")
                xml.AppendLine("        <ram:NetPriceProductTradePrice>")
                xml.AppendLine("          <ram:ChargeAmount>" & FZ(Math.Abs(prs)) & "</ram:ChargeAmount>")
                xml.AppendLine("        </ram:NetPriceProductTradePrice>")
                xml.AppendLine("      </ram:SpecifiedLineTradeAgreement>")
                xml.AppendLine("      <ram:SpecifiedLineTradeDelivery>")
                xml.AppendLine("        <ram:BilledQuantity unitCode=""C62"">" & FZ(Math.Abs(anz)) & "</ram:BilledQuantity>")
                xml.AppendLine("      </ram:SpecifiedLineTradeDelivery>")
                xml.AppendLine("      <ram:SpecifiedLineTradeSettlement>")
                xml.AppendLine("        <ram:ApplicableTradeTax>")
                xml.AppendLine("          <ram:TypeCode>VAT</ram:TypeCode>")
                xml.AppendLine("          <ram:CategoryCode>S</ram:CategoryCode>")
                xml.AppendLine("          <ram:RateApplicablePercent>" & FZ(mwst) & "</ram:RateApplicablePercent>")
                xml.AppendLine("        </ram:ApplicableTradeTax>")
                xml.AppendLine("        <ram:SpecifiedTradeSettlementLineMonetarySummation>")
                xml.AppendLine("          <ram:LineTotalAmount>" & FZ(finalNetto) & "</ram:LineTotalAmount>")
                xml.AppendLine("        </ram:SpecifiedTradeSettlementLineMonetarySummation>")
                xml.AppendLine("      </ram:SpecifiedLineTradeSettlement>")
                xml.AppendLine("    </ram:IncludedSupplyChainTradeLineItem>")
                lineId += 1
            End If
        Next

        ' ── VEREINBARUNG (Verkäufer / Käufer) ──
        xml.AppendLine("    <ram:ApplicableHeaderTradeAgreement>")
        xml.AppendLine("      <ram:BuyerReference>" & XE(kMitgliedsNr) & "</ram:BuyerReference>")

        ' Verkäufer
        xml.AppendLine("      <ram:SellerTradeParty>")
        xml.AppendLine("        <ram:ID>" & XE(fName) & "</ram:ID>")
        xml.AppendLine("        <ram:Name>" & XE(fName) & "</ram:Name>")
        xml.AppendLine("        <ram:DefinedTradeContact>")
        xml.AppendLine("          <ram:PersonName>Buchhaltung</ram:PersonName>")
        If Not String.IsNullOrWhiteSpace(fTel) Then
            xml.AppendLine("          <ram:TelephoneUniversalCommunication>")
            xml.AppendLine("            <ram:CompleteNumber>" & XE(fTel) & "</ram:CompleteNumber>")
            xml.AppendLine("          </ram:TelephoneUniversalCommunication>")
        End If
        If Not String.IsNullOrWhiteSpace(fEmail) Then
            xml.AppendLine("          <ram:EmailURIUniversalCommunication>")
            xml.AppendLine("            <ram:URIID>" & XE(fEmail) & "</ram:URIID>")
            xml.AppendLine("          </ram:EmailURIUniversalCommunication>")
        End If
        xml.AppendLine("        </ram:DefinedTradeContact>")
        xml.AppendLine("        <ram:PostalTradeAddress>")
        xml.AppendLine("          <ram:PostcodeCode>" & XE(fPlz) & "</ram:PostcodeCode>")
        xml.AppendLine("          <ram:LineOne>" & XE(fStrasse) & "</ram:LineOne>")
        xml.AppendLine("          <ram:CityName>" & XE(fOrt) & "</ram:CityName>")
        xml.AppendLine("          <ram:CountryID>" & XE(fLand) & "</ram:CountryID>")
        xml.AppendLine("        </ram:PostalTradeAddress>")
        If Not String.IsNullOrWhiteSpace(fEmail) Then
            xml.AppendLine("        <ram:URIUniversalCommunication>")
            xml.AppendLine("          <ram:URIID schemeID=""EM"">" & XE(fEmail) & "</ram:URIID>")
            xml.AppendLine("        </ram:URIUniversalCommunication>")
        End If
        If Not String.IsNullOrWhiteSpace(fSteuer) Then
            xml.AppendLine("        <ram:SpecifiedTaxRegistration>")
            xml.AppendLine("          <ram:ID schemeID=""FC"">" & XE(fSteuer) & "</ram:ID>")
            xml.AppendLine("        </ram:SpecifiedTaxRegistration>")
        End If
        xml.AppendLine("      </ram:SellerTradeParty>")

        ' Käufer
        xml.AppendLine("      <ram:BuyerTradeParty>")
        xml.AppendLine("        <ram:Name>" & XE(kName) & "</ram:Name>")
        xml.AppendLine("        <ram:PostalTradeAddress>")
        xml.AppendLine("          <ram:PostcodeCode>" & XE(kPlz) & "</ram:PostcodeCode>")
        xml.AppendLine("          <ram:LineOne>" & XE(kStrasse) & "</ram:LineOne>")
        xml.AppendLine("          <ram:CityName>" & XE(kOrt) & "</ram:CityName>")
        xml.AppendLine("          <ram:CountryID>" & XE(kLand) & "</ram:CountryID>")
        xml.AppendLine("        </ram:PostalTradeAddress>")
        If Not String.IsNullOrWhiteSpace(kEmail) Then
            xml.AppendLine("        <ram:URIUniversalCommunication>")
            xml.AppendLine("          <ram:URIID schemeID=""EM"">" & XE(kEmail) & "</ram:URIID>")
            xml.AppendLine("        </ram:URIUniversalCommunication>")
        End If
        xml.AppendLine("      </ram:BuyerTradeParty>")
        xml.AppendLine("    </ram:ApplicableHeaderTradeAgreement>")

        ' ── LIEFERUNG ──
        xml.AppendLine("    <ram:ApplicableHeaderTradeDelivery>")
        xml.AppendLine("      <ram:ActualDeliverySupplyChainEvent>")
        xml.AppendLine("        <ram:OccurrenceDateTime><udt:DateTimeString format=""102"">" & formatDatum & "</udt:DateTimeString></ram:OccurrenceDateTime>")
        xml.AppendLine("      </ram:ActualDeliverySupplyChainEvent>")
        xml.AppendLine("    </ram:ApplicableHeaderTradeDelivery>")

        ' ── ZAHLUNG & SUMMEN ──
        xml.AppendLine("    <ram:ApplicableHeaderTradeSettlement>")
        If Not String.IsNullOrWhiteSpace(fGlaeubiger) Then
            xml.AppendLine("      <ram:CreditorReferenceID>" & XE(fGlaeubiger) & "</ram:CreditorReferenceID>")
        End If
        xml.AppendLine("      <ram:PaymentReference>" & XE(reNr) & "</ram:PaymentReference>")
        xml.AppendLine("      <ram:InvoiceCurrencyCode>EUR</ram:InvoiceCurrencyCode>")

        ' Zahlungsmittel (Lastschrift)
        xml.AppendLine("      <ram:SpecifiedTradeSettlementPaymentMeans>")
        xml.AppendLine("        <ram:TypeCode>59</ram:TypeCode>")
        xml.AppendLine("        <ram:Information>SEPA direct debit</ram:Information>")
        If Not String.IsNullOrWhiteSpace(kIban) Then
            xml.AppendLine("        <ram:PayerPartyDebtorFinancialAccount>")
            xml.AppendLine("          <ram:IBANID>" & XE(kIban) & "</ram:IBANID>")
            xml.AppendLine("        </ram:PayerPartyDebtorFinancialAccount>")
        End If
        xml.AppendLine("      </ram:SpecifiedTradeSettlementPaymentMeans>")

        ' Steuern - Nutzt nun die "hatMwst" Schalter, damit auch 0-Euro-Rechnungen ihre Steuern ausweisen
        If hatMwst19 Then
            xml.AppendLine("      <ram:ApplicableTradeTax>")
            xml.AppendLine("        <ram:CalculatedAmount>" & FZ(steuerBetrag19 * signMult) & "</ram:CalculatedAmount>")
            xml.AppendLine("        <ram:TypeCode>VAT</ram:TypeCode>")
            xml.AppendLine("        <ram:BasisAmount>" & FZ(steuerBasis19 * signMult) & "</ram:BasisAmount>")
            xml.AppendLine("        <ram:CategoryCode>S</ram:CategoryCode>")
            xml.AppendLine("        <ram:RateApplicablePercent>19.00</ram:RateApplicablePercent>")
            xml.AppendLine("      </ram:ApplicableTradeTax>")
        End If
        If hatMwst7 Then
            xml.AppendLine("      <ram:ApplicableTradeTax>")
            xml.AppendLine("        <ram:CalculatedAmount>" & FZ(steuerBetrag7 * signMult) & "</ram:CalculatedAmount>")
            xml.AppendLine("        <ram:TypeCode>VAT</ram:TypeCode>")
            xml.AppendLine("        <ram:BasisAmount>" & FZ(steuerBasis7 * signMult) & "</ram:BasisAmount>")
            xml.AppendLine("        <ram:CategoryCode>S</ram:CategoryCode>")
            xml.AppendLine("        <ram:RateApplicablePercent>7.00</ram:RateApplicablePercent>")
            xml.AppendLine("      </ram:ApplicableTradeTax>")
        End If

        ' Hier werden alle negativen Artikel normgerecht als "Abschlag" aufgelistet
        For Each pos In positionen
            Dim netto As Decimal = CDec(pos("netto"))
            Dim finalNetto As Decimal = netto * signMult

            If finalNetto < 0 Then
                Dim absAllowance As Decimal = Math.Abs(finalNetto)
                allowanceSum += absAllowance
                Dim mwst As Integer = CInt(pos("mwst"))

                xml.AppendLine("      <ram:SpecifiedTradeAllowanceCharge>")
                xml.AppendLine("        <ram:ChargeIndicator>")
                xml.AppendLine("          <udt:Indicator>false</udt:Indicator>")
                xml.AppendLine("        </ram:ChargeIndicator>")
                xml.AppendLine("        <ram:ActualAmount>" & FZ(absAllowance) & "</ram:ActualAmount>")
                xml.AppendLine("        <ram:Reason>" & XE(CStr(pos("bez"))) & "</ram:Reason>")
                xml.AppendLine("        <ram:CategoryTradeTax>")
                xml.AppendLine("          <ram:TypeCode>VAT</ram:TypeCode>")
                xml.AppendLine("          <ram:CategoryCode>S</ram:CategoryCode>")
                xml.AppendLine("          <ram:RateApplicablePercent>" & FZ(mwst) & "</ram:RateApplicablePercent>")
                xml.AppendLine("        </ram:CategoryTradeTax>")
                xml.AppendLine("      </ram:SpecifiedTradeAllowanceCharge>")
            End If
        Next

        ' Zahlungsbedingungen & Mandatsreferenz
        xml.AppendLine("      <ram:SpecifiedTradePaymentTerms>")

        Dim basisText As String = If(isGutschrift, textGutschrift, textZahlung)
        Dim sepaTextXml As String = basisText _
            .Replace("[RE-nummer]", "re-" & reNr) _
            .Replace("[Kunden-IBAN]", If(kIban = "", "—", kIban)) _
            .Replace("[Kunden-Mandat]", If(kMitgliedsNr = "", "—", kMitgliedsNr)) _
            .Replace("[Gläubiger-ID]", If(fGlaeubiger = "", "—", fGlaeubiger))

        If String.IsNullOrWhiteSpace(sepaTextXml) Then
            sepaTextXml = If(isGutschrift, "Gutschrift / Korrektur", "Zahlung per SEPA-Lastschrift")
        End If

        ' ---> NEU: Fängt 0-Euro-Rechnungen ab und überschreibt den Text <---
        If summeBrutto = 0 Then
            sepaTextXml = "Der Rechnungsbetrag beläuft sich auf 0,00 €. Diese Rechnung dient lediglich zur Information/Korrektur, es ist keine Zahlung oder Abbuchung erforderlich."
        End If

        xml.AppendLine("        <ram:Description>" & XE(sepaTextXml) & "</ram:Description>")
        If Not String.IsNullOrWhiteSpace(kMitgliedsNr) Then
            xml.AppendLine("        <ram:DirectDebitMandateID>" & XE(kMitgliedsNr) & "</ram:DirectDebitMandateID>")
        End If
        xml.AppendLine("      </ram:SpecifiedTradePaymentTerms>")

        ' Gesamtsummen - LineTotalAmount enthält nur die echten Artikel, AllowanceTotalAmount die Abschläge
        xml.AppendLine("      <ram:SpecifiedTradeSettlementHeaderMonetarySummation>")
        xml.AppendLine("        <ram:LineTotalAmount>" & FZ(lineTotalSum) & "</ram:LineTotalAmount>")
        xml.AppendLine("        <ram:ChargeTotalAmount>0.00</ram:ChargeTotalAmount>")
        xml.AppendLine("        <ram:AllowanceTotalAmount>" & FZ(allowanceSum) & "</ram:AllowanceTotalAmount>")
        xml.AppendLine("        <ram:TaxBasisTotalAmount>" & FZ(summeNetto * signMult) & "</ram:TaxBasisTotalAmount>")
        xml.AppendLine("        <ram:TaxTotalAmount currencyID=""EUR"">" & FZ((steuerBetrag7 + steuerBetrag19) * signMult) & "</ram:TaxTotalAmount>")
        xml.AppendLine("        <ram:GrandTotalAmount>" & FZ(summeBrutto * signMult) & "</ram:GrandTotalAmount>")
        xml.AppendLine("        <ram:TotalPrepaidAmount>0.00</ram:TotalPrepaidAmount>")
        xml.AppendLine("        <ram:DuePayableAmount>" & FZ(summeBrutto * signMult) & "</ram:DuePayableAmount>")
        xml.AppendLine("      </ram:SpecifiedTradeSettlementHeaderMonetarySummation>")
        xml.AppendLine("    </ram:ApplicableHeaderTradeSettlement>")

        xml.AppendLine("  </rsm:SupplyChainTradeTransaction>")
        xml.AppendLine("</rsm:CrossIndustryInvoice>")

        ' Speichern (ohne BOM!)
        Dim utf8OhneBom As New UTF8Encoding(False)
        File.WriteAllText(archivXmlPfad, xml.ToString(), utf8OhneBom)
        File.Copy(archivXmlPfad, tempXmlPfad, True)
    End Sub

    Private Shared Function XE(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("""", "&quot;").Replace("'", "&apos;")
    End Function

    Private Shared Function FZ(zahl As Decimal) As String
        Return zahl.ToString("0.00", CultureInfo.InvariantCulture)
    End Function

End Class