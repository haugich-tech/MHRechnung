Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.SQLite
Imports ClosedXML.Excel
Imports System.IO
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing.Printing
Imports System.Net.Mail
Imports System.Net

Public Class Form1

    ' =========================================================================
    ' DESIGN-KONSTANTEN (Zentrales Farbschema - hier ändern = überall wirkt)
    ' =========================================================================
    ' Farbschema "Kornfeld": Kornblumenblau + Weizengold auf warmem Papierweiß
    Private Shared ReadOnly CLR_BLAU_DUNKEL As Color = Color.FromArgb(45, 74, 102)      ' Kornblumenblau, dunkel (Header, Akzente)
    Private Shared ReadOnly CLR_BLAU_MITTEL As Color = Color.FromArgb(61, 90, 128)      ' Kornblumenblau (Buttons primär)
    Private Shared ReadOnly CLR_BLAU_HELL As Color = Color.FromArgb(223, 231, 240)      ' Helles Blau (Hintergründe, Hover)
    Private Shared ReadOnly CLR_GOLD_AKZENT As Color = Color.FromArgb(201, 162, 39)     ' Weizengold (Aktiv-Marker)
    Private Shared ReadOnly CLR_WEISS As Color = Color.White
    Private Shared ReadOnly CLR_HINTERGRUND As Color = Color.FromArgb(247, 242, 232)    ' Warmes Papierweiß
    Private Shared ReadOnly CLR_PANEL_BG As Color = Color.FromArgb(239, 231, 213)       ' Panel-Hintergrund
    Private Shared ReadOnly CLR_BORDER As Color = Color.FromArgb(222, 210, 184)         ' Rahmenfarbe
    Private Shared ReadOnly CLR_TEXT_DUNKEL As Color = Color.FromArgb(38, 34, 27)       ' Haupttext
    Private Shared ReadOnly CLR_TEXT_GRAU As Color = Color.FromArgb(110, 100, 85)       ' Nebentext / Labels
    Private Shared ReadOnly CLR_ROT As Color = Color.FromArgb(180, 50, 50)              ' Löschen / Gefahr
    Private Shared ReadOnly CLR_ORANGE As Color = Color.FromArgb(200, 110, 20)          ' Bearbeiten / Warnung
    Private Shared ReadOnly CLR_HEADER_BG As Color = Color.FromArgb(30, 48, 68)         ' Tiefdunkel für Kopfzeile
    Private Shared ReadOnly FONT_TITLE As New Font("Segoe UI", 10.5F, FontStyle.Bold)
    Private Shared ReadOnly FONT_NORMAL As New Font("Segoe UI", 9.5F)
    Private Shared ReadOnly FONT_KLEIN As New Font("Segoe UI", 8.5F)
    Private Shared ReadOnly FONT_GROSS As New Font("Segoe UI", 11.0F, FontStyle.Bold)

    ' =========================================================================
    ' GLOBALE UI ELEMENTE (Klassenvariablen)
    ' =========================================================================
    Private WithEvents MainTabs As New TabControl()
    ' Keine Emojis in Tab-Titeln: WinForms (GDI) kann sie nicht darstellen (nur Kästchen)
    Private TabErfassung As New TabPage("Rechnungserfassung")
    Private TabKontrolle As New TabPage("Kontrolle & Stapel")
    Private TabArtikel As New TabPage("Artikelverwaltung")
    Private TabMitglieder As New TabPage("Kunden")
    Private TabEinstellungen As New TabPage("Einstellungen")
    Private TabArchiv As New TabPage("Rechnungs-Archiv")

    ' --- Elemente für Tab 1 (Erfassung) ---
    Private pnlRows As New FlowLayoutPanel()
    Private lstArtikel As New ListBox()
    Private chkMitglieder As New CheckedListBox()
    Private pnlPlusContainer As Panel
    Private btnRechnungErstellen As Button
    Private lblSummenTab1 As New Label()
    ' Ersetzt den früheren cbEingabeModus (Preis-Eingabe NETTO/BRUTTO für die ganze Rechnung):
    ' Jede Zeile hat jetzt eigene, synchronisierte Netto-/Brutto-Felder, deshalb ist die globale
    ' Eingabeart-Wahl überflüssig. cbPreisart legt stattdessen fest, welche der beiden Spalten
    ' auf der GEDRUCKTEN Rechnung als Einzelpreis erscheint (pro Rechnung gespeichert).
    Private WithEvents cbPreisart As New ComboBox()
    Private txtLieferdatum As New TextBox()

    ' --- Elemente für Tab 2 (Kontrolle & Stapel) ---
    Private dgvRechnungen As New DataGridView()
    Private pnlDetailsContent As New FlowLayoutPanel()
    Private lblDetailsSumme As New Label()
    Private lblSummenTab2 As New Label()
    Private lblKontrolleTab2 As New Label()

    ' --- Elemente für Tab 3 (Artikel) ---
    Private dgvArtikelVerwaltung As New DataGridView()
    Private cmsArtikel As New ContextMenuStrip()
    Private dragQuelleIndex As Integer = -1
    Private dragZielIndex As Integer = -1

    ' --- Elemente für Tab 4 (Kunden) ---
    Private WithEvents dgvMitglieder As New DataGridView()
    Private aktuelleMitgliedId As Integer = 0

    Private txtM_Nr As New TextBox()
    Private txtM_Name As New TextBox()
    Private txtM_Email As New TextBox()
    Private txtM_Steuer As New TextBox()
    Private txtM_Betrieb As New TextBox()
    Private cbM_Versand As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}
    Private cbM_Preisart As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}

    Private txtM_Strasse As New TextBox()
    Private txtM_PLZ As New TextBox()
    Private txtM_Ort As New TextBox()
    Private txtM_Land As New TextBox()

    ' --- Elemente für Tab 5 (Einstellungen) ---
    Private chkE_MengenNullen As New CheckBox()
    Private chkE_AutoBackup As New CheckBox()
    Private chkE_Ausgangskopie As New CheckBox() With {.Checked = True}
    Private txtE_AusgangskopieEmail As New TextBox()
    Private cbE_DruckAnzahl As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}
    Private cbE_StandardDrucker As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}

    Private txtE_FirmaName As New TextBox()
    Private txtE_FirmaStrasse As New TextBox()
    Private txtE_FirmaPLZ As New TextBox()
    Private txtE_FirmaOrt As New TextBox()
    Private txtE_FirmaMail As New TextBox()
    Private txtE_FirmaTel As New TextBox()
    Private txtE_Steuer As New TextBox()
    Private txtE_UStId As New TextBox()
    Private txtE_IBAN As New TextBox()
    Private txtE_BIC As New TextBox()
    Private txtE_Bank As New TextBox()
    Private txtE_MwSt1 As New TextBox()
    Private txtE_MwSt2 As New TextBox()

    Private WithEvents txtStartReNr As New TextBox()
    Private WithEvents chkSicherLoeschen As New CheckBox()
    Private WithEvents btnLoeschen As New Button()

    Private txtE_TextZahlung As New TextBox() With {.Multiline = True, .ScrollBars = ScrollBars.Vertical}
    Private txtE_TextEmail As New TextBox() With {.Multiline = True, .ScrollBars = ScrollBars.Vertical}
    Private txtE_TextGutschrift As New TextBox() With {.Multiline = True, .ScrollBars = ScrollBars.Vertical}
    Private txtE_SmtpServer As New TextBox()
    Private txtE_SmtpPort As New TextBox()
    Private txtE_SmtpUser As New TextBox()
    Private txtE_SmtpPass As New TextBox() With {.PasswordChar = "*"c}

    Private txtE_Speicherpfad As New TextBox()
    Private txtE_DbPfad As New TextBox() With {.ReadOnly = True}

    ' --- Elemente für Tab 6 (Archiv) ---
    Private dgvArchiv As New DataGridView()
    Private pnlArchivDetails As New FlowLayoutPanel()
    Private lblArchivSumme As New Label()

    ' --- Globale Elemente ---
    Private lblAktuelleReNr As New Label()

    Public Sub New()
        InitializeComponent()
    End Sub

    ' =========================================================================
    ' ROBUSTES ZAHLEN-PARSING (akzeptiert Komma UND Punkt als Dezimaltrenner)
    ' Verhindert, dass "1.50" bei deutscher Windows-Einstellung als 150 gelesen wird.
    ' =========================================================================
    Private Shared Function ParseBetrag(text As String, ByRef wert As Decimal) As Boolean
        wert = 0
        If String.IsNullOrWhiteSpace(text) Then Return False
        Dim s As String = text.Trim().Replace(" ", "").Replace("€", "")
        If s.Contains(",") AndAlso s.Contains(".") Then
            ' Beide Zeichen vorhanden: das hintere ist das Dezimaltrennzeichen
            If s.LastIndexOf(","c) > s.LastIndexOf("."c) Then
                s = s.Replace(".", "").Replace(",", ".")   ' 1.234,56 -> 1234.56
            Else
                s = s.Replace(",", "")                     ' 1,234.56 -> 1234.56
            End If
        Else
            s = s.Replace(",", ".")                        ' 1,50 -> 1.50
        End If
        Return Decimal.TryParse(s, Globalization.NumberStyles.Number, Globalization.CultureInfo.InvariantCulture, wert)
    End Function

    ' =========================================================================
    ' KONFIGURIERBARE MwSt-SÄTZE (§24 UStG Durchschnittssätze - in Einstellungen änderbar)
    ' =========================================================================
    Private mwstSatz1 As Decimal = 7.8D
    Private mwstSatz2 As Decimal = 5.5D

    ' Muss VOR dem Aufbau der Tabs laufen, damit die MwSt-Dropdowns die richtigen
    ' Werte bekommen. Fehlt ein Eintrag in der DB (frisches System), bleibt der
    ' Compile-Zeit-Standardwert oben stehen.
    ' =========================================================================
    ' DATENBANK-PFAD (muss vor jedem DB-Zugriff feststehen - siehe DbKonfiguration.vb)
    ' =========================================================================
    Private Sub LegeDatenbankPfadFest()
        Dim pfad As String = DbKonfiguration.LiesDatenbankPfad()
        Dim pfadGueltig As Boolean = Not String.IsNullOrWhiteSpace(pfad) AndAlso Directory.Exists(Path.GetDirectoryName(pfad))

        If Not pfadGueltig Then
            ' Falls direkt neben der .exe schon eine Datenbank liegt (z.B. von einer
            ' Installation ohne diese Pfad-Datei), die einfach automatisch übernehmen,
            ' statt gleich einen Dialog aufzuzwingen.
            Dim standardPfad As String = Path.Combine(Application.StartupPath, "MHRechnung_Daten.sqlite")
            If File.Exists(standardPfad) Then
                pfad = standardPfad
                pfadGueltig = True
            End If
        End If

        If Not pfadGueltig Then
            pfad = ZeigeDatenbankWahlDialog(erzwingeWahl:=True)
        End If

        DbKonfiguration.SchreibeDatenbankPfad(pfad)
        DatenbankManager.DatenbankPfad = pfad
    End Sub

    ''' <summary>Zeigt den Datenbank-Auswahl-Dialog. Bei erzwingeWahl:=True kann das
    ''' Fenster nicht weggeklickt werden (das Programm braucht zwingend eine Datenbank);
    ''' bei False gibt es einen Abbrechen-Button und die Funktion kann "" zurückgeben.</summary>
    Private Function ZeigeDatenbankWahlDialog(erzwingeWahl As Boolean) As String
        Dim gewaehlterPfad As String = ""

        Dim dlg As New Form With {
            .Text = "Datenbank auswählen",
            .Size = New Size(500, 230),
            .StartPosition = FormStartPosition.CenterScreen,
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .MaximizeBox = False, .MinimizeBox = False,
            .ControlBox = Not erzwingeWahl,
            .BackColor = CLR_HINTERGRUND
        }

        Dim lbl As New Label With {
            .Text = If(erzwingeWahl,
                "Es wurde noch keine Datenbank-Datei gefunden." & vbCrLf & vbCrLf &
                "Bitte wähle eine vorhandene Datenbank aus oder lege eine neue an.",
                "Wähle eine vorhandene Datenbank aus oder lege eine neue an."),
            .Location = New Point(20, 20), .Size = New Size(450, 70),
            .Font = FONT_NORMAL, .ForeColor = CLR_TEXT_DUNKEL
        }

        Dim btnVorhanden = MachePrimaerButton("Vorhandene öffnen…", 210, 44)
        btnVorhanden.Location = New Point(20, 105)
        Dim btnNeu = MacheSekundaerButton("Neue anlegen…", 210, 44)
        btnNeu.Location = New Point(250, 105)

        AddHandler btnVorhanden.Click, Sub()
                                            Dim ofd As New OpenFileDialog With {
                                                .Filter = "SQLite-Datenbank (*.sqlite)|*.sqlite|Alle Dateien (*.*)|*.*",
                                                .Title = "Vorhandene Datenbank auswählen"
                                            }
                                            If ofd.ShowDialog() = DialogResult.OK Then
                                                gewaehlterPfad = ofd.FileName
                                                dlg.DialogResult = DialogResult.OK
                                                dlg.Close()
                                            End If
                                        End Sub

        AddHandler btnNeu.Click, Sub()
                                      Dim sfd As New SaveFileDialog With {
                                          .Filter = "SQLite-Datenbank (*.sqlite)|*.sqlite",
                                          .Title = "Neue Datenbank anlegen",
                                          .FileName = "MHRechnung_Daten.sqlite"
                                      }
                                      If sfd.ShowDialog() = DialogResult.OK Then
                                          gewaehlterPfad = sfd.FileName
                                          dlg.DialogResult = DialogResult.OK
                                          dlg.Close()
                                      End If
                                  End Sub

        dlg.Controls.AddRange({lbl, btnVorhanden, btnNeu})

        If Not erzwingeWahl Then
            Dim btnAbbrechen As New Button With {
                .Text = "Abbrechen", .Location = New Point(20, 160), .Size = New Size(100, 32),
                .DialogResult = DialogResult.Cancel, .Font = FONT_NORMAL
            }
            dlg.Controls.Add(btnAbbrechen)
            dlg.CancelButton = btnAbbrechen
        End If

        Dim result = dlg.ShowDialog()

        If result = DialogResult.OK Then
            Return gewaehlterPfad
        ElseIf erzwingeWahl Then
            MessageBox.Show("Ohne ausgewählte Datenbank kann MHRechnung nicht gestartet werden. Das Programm wird beendet.", "Abgebrochen", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Application.Exit()
            Return ""
        Else
            Return ""
        End If
    End Function

    Private Sub LadeMwstSaetzeFrueh()
        Using conn = DatenbankManager.HoleVerbindung()
            Dim cmd As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'mwst_satz_1'", conn)
            Dim v1 = cmd.ExecuteScalar()
            If v1 IsNot Nothing Then Decimal.TryParse(v1.ToString(), Globalization.NumberStyles.Number, Globalization.CultureInfo.InvariantCulture, mwstSatz1)

            cmd.CommandText = "SELECT wert FROM einstellungen WHERE schluessel = 'mwst_satz_2'"
            Dim v2 = cmd.ExecuteScalar()
            If v2 IsNot Nothing Then Decimal.TryParse(v2.ToString(), Globalization.NumberStyles.Number, Globalization.CultureInfo.InvariantCulture, mwstSatz2)
        End Using
    End Sub

    ' Formatiert einen MwSt-Satz für die Anzeige, z.B. 7,8 -> "7,8%"
    Private Shared Function FmtMwSt(satz As Decimal) As String
        Return satz.ToString("0.0###", Globalization.CultureInfo.InvariantCulture).Replace(".", ",") & "%"
    End Function

    ' Liefert den Bruttopreis als Anzeigetext - nimmt den exakt gespeicherten Wert, falls
    ' vorhanden (wichtig, damit ein glatt eingegebener Bruttopreis nicht durch eine
    ' Rückrechnung krumme Nachkommastellen bekommt). Fehlt er (NULL, z.B. bei Artikeln/
    ' Positionen aus der Zeit vor dem Netto/Brutto-Umbau), wird er aus Netto berechnet.
    Private Shared Function BruttoText(bruttoWert As Object, netto As Decimal, mwstProzent As Decimal) As String
        If bruttoWert IsNot Nothing AndAlso Not DBNull.Value.Equals(bruttoWert) Then
            Dim b As Decimal = CDec(bruttoWert)
            If b > 0 Then Return b.ToString("N2")
        End If
        Return (netto * (1 + mwstProzent / 100D)).ToString("N2")
    End Function

    ' Liest einen MwSt-Satz aus einem Anzeigetext ("7,8%", "7,8", "5,5 %" ...)
    Private Shared Function ParseMwSt(text As String) As Decimal
        Dim wert As Decimal = 0
        If String.IsNullOrWhiteSpace(text) Then Return 0
        ParseBetrag(text.Replace("%", ""), wert)
        Return wert
    End Function

    ' =========================================================================
    ' FORM LOAD & INITIALISIERUNG
    ' =========================================================================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If PdfSharp.Fonts.GlobalFontSettings.FontResolver Is Nothing Then
            PdfSharp.Fonts.GlobalFontSettings.FontResolver = New LegFontResolver()
        End If

        Me.Text = "MHRechnung — Rechnungs-Manager  v1.0.0 (2026-09-14)"
        Me.Size = New Size(1400, 950)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = FONT_NORMAL
        Me.BackColor = CLR_HINTERGRUND

        ' Datenbank-Pfad ermitteln, BEVOR irgendetwas auf die DB zugreift.
        LegeDatenbankPfadFest()

        ' DB muss vor dem Aufbau der Tabs bereitstehen, da die MwSt-Dropdowns
        ' schon beim Bauen die konfigurierten Sätze brauchen.
        DatenbankManager.InitialisiereDatenbank()
        LadeMwstSaetzeFrueh()

        BaueKopfzeile()

        MainTabs.Dock = DockStyle.Fill
        MainTabs.Padding = New Point(18, 9)
        MainTabs.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        StyleTabControl(MainTabs)

        BaueTabErfassung()
        BaueTabKontrolle()
        BaueTabArtikel()
        BaueTabMitglieder()
        BaueTabEinstellungen()
        BaueTabArchiv()

        MainTabs.TabPages.Add(TabErfassung)
        MainTabs.TabPages.Add(TabKontrolle)
        MainTabs.TabPages.Add(TabArtikel)
        MainTabs.TabPages.Add(TabMitglieder)
        MainTabs.TabPages.Add(TabEinstellungen)
        MainTabs.TabPages.Add(TabArchiv)

        Me.Controls.Add(MainTabs)
        MainTabs.BringToFront()

        LadeDaten()
        LadeEinstellungen()

        If Not String.IsNullOrWhiteSpace(txtE_FirmaName.Text) Then
            Me.Text = txtE_FirmaName.Text & " — Rechnungs-Manager  v1.0.0 (2026-09-14)"
        End If

        If Not ToolPfade.SindAlleToolsBereit() Then
            Dim fehlend As String = ""
            If Not IO.File.Exists(ToolPfade.JavaExe) Then fehlend &= "- Java" & vbCrLf
            If Not IO.File.Exists(ToolPfade.GhostscriptExe) Then fehlend &= "- Ghostscript" & vbCrLf
            If Not IO.File.Exists(ToolPfade.MustangJar) Then fehlend &= "- Mustang-CLI" & vbCrLf

            MessageBox.Show("Die folgenden ZUGFeRD-Komponenten fehlen im Tools-Ordner:" & vbCrLf & vbCrLf &
                            fehlend & vbCrLf &
                            "Die E-Rechnungs-Erstellung wird nicht funktionieren!",
                            "System-Check", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            If btnRechnungErstellen IsNot Nothing Then btnRechnungErstellen.Enabled = False
        End If

        If Not IO.File.Exists(ToolPfade.LogoPfad) Then
            MessageBox.Show("Das Firmenlogo wurde nicht gefunden unter:" & vbCrLf &
                            ToolPfade.LogoPfad & vbCrLf & vbCrLf &
                            "Bitte erstelle den Ordner 'logo' und speichere dort 'logo.jpg'!",
                            "Logo fehlt", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SendeDatenbankBackup(False)
    End Sub

    ' =========================================================================
    ' DESIGN-HILFSMETHODEN
    ' =========================================================================

    ''' <summary>Stylt einen DataGridView modern mit Blau-Akzenten.</summary>
    Private Sub StyleDgv(dgv As DataGridView)
        dgv.BackgroundColor = CLR_WEISS
        dgv.BorderStyle = BorderStyle.None
        dgv.GridColor = CLR_BORDER
        dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        dgv.RowHeadersVisible = False
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.MultiSelect = False
        dgv.ReadOnly = True
        dgv.Font = FONT_NORMAL
        dgv.RowTemplate.Height = 32
        dgv.ColumnHeadersHeight = 38
        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersDefaultCellStyle.BackColor = CLR_BLAU_DUNKEL
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = CLR_WEISS
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        dgv.ColumnHeadersDefaultCellStyle.Padding = New Padding(6, 0, 0, 0)
        dgv.DefaultCellStyle.BackColor = CLR_WEISS
        dgv.DefaultCellStyle.ForeColor = CLR_TEXT_DUNKEL
        dgv.DefaultCellStyle.SelectionBackColor = CLR_BLAU_HELL
        dgv.DefaultCellStyle.SelectionForeColor = CLR_TEXT_DUNKEL
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 250, 245)
    End Sub

    ''' <summary>Stylt einen Button als primären Blau-Button.</summary>
    Private Function MachePrimaerButton(text As String, w As Integer, h As Integer) As Button
        Dim btn As New Button With {
            .Text = text, .Size = New Size(w, h),
            .BackColor = CLR_BLAU_MITTEL, .ForeColor = CLR_WEISS,
            .FlatStyle = FlatStyle.Flat, .Font = FONT_TITLE,
            .Cursor = Cursors.Hand
        }
        btn.FlatAppearance.BorderSize = 0
        btn.FlatAppearance.MouseOverBackColor = CLR_BLAU_DUNKEL
        Return btn
    End Function

    ''' <summary>Stylt einen Button als sekundären (weiß/Rahmen) Button.</summary>
    Private Function MacheSekundaerButton(text As String, w As Integer, h As Integer) As Button
        Dim btn As New Button With {
            .Text = text, .Size = New Size(w, h),
            .BackColor = CLR_WEISS, .ForeColor = CLR_BLAU_DUNKEL,
            .FlatStyle = FlatStyle.Flat, .Font = FONT_TITLE,
            .Cursor = Cursors.Hand
        }
        btn.FlatAppearance.BorderSize = 1
        btn.FlatAppearance.BorderColor = CLR_BLAU_MITTEL
        btn.FlatAppearance.MouseOverBackColor = CLR_BLAU_HELL
        Return btn
    End Function

    ''' <summary>Stylt einen Button als Gefahr-Button (Rot).</summary>
    Private Function MacheGefahrButton(text As String, w As Integer, h As Integer) As Button
        Dim btn As New Button With {
            .Text = text, .Size = New Size(w, h),
            .BackColor = CLR_ROT, .ForeColor = CLR_WEISS,
            .FlatStyle = FlatStyle.Flat, .Font = FONT_TITLE,
            .Cursor = Cursors.Hand
        }
        btn.FlatAppearance.BorderSize = 0
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(140, 30, 30)
        Return btn
    End Function

    ''' <summary>Erstellt eine moderne Abschnitts-Überschrift.</summary>
    Private Function MacheAbschnittsLabel(text As String) As Label
        Return New Label With {
            .Text = "  " & text,
            .Dock = DockStyle.Top,
            .Height = 36,
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .BackColor = CLR_BLAU_HELL,
            .ForeColor = CLR_BLAU_DUNKEL,
            .TextAlign = ContentAlignment.MiddleLeft,
            .BorderStyle = BorderStyle.None
        }
    End Function

    ''' <summary>Erstellt eine moderne Toolbar-Leiste oben in einem Tab.</summary>
    Private Function MacheToolbar() As Panel
        Dim p As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 56,
            .BackColor = CLR_PANEL_BG,
            .Padding = New Padding(12, 10, 12, 10)
        }
        ' Untere Trennlinie simulieren via BorderStyle geht nicht direkt —
        ' Wir zeichnen sie im Paint-Event
        AddHandler p.Paint, Sub(s As Object, ev As PaintEventArgs)
                                ev.Graphics.DrawLine(New Pen(CLR_BORDER, 1), 0, p.Height - 1, p.Width, p.Height - 1)
                            End Sub
        Return p
    End Function

    ''' <summary>Stylt den TabControl mit Kornfeld-Farbschema (Kornblumenblau).</summary>
    Private Sub StyleTabControl(tc As TabControl)
        tc.DrawMode = TabDrawMode.OwnerDrawFixed
        tc.ItemSize = New Size(170, 40)
        AddHandler tc.DrawItem, Sub(s As Object, ev As DrawItemEventArgs)
                                    Dim tab As TabPage = tc.TabPages(ev.Index)
                                    Dim isSelected As Boolean = (tc.SelectedIndex = ev.Index)
                                    Dim bg As Color = If(isSelected, CLR_BLAU_MITTEL, CLR_PANEL_BG)
                                    Dim fg As Color = If(isSelected, CLR_WEISS, CLR_TEXT_GRAU)
                                    ev.Graphics.FillRectangle(New SolidBrush(bg), ev.Bounds)
                                    If isSelected Then
                                        ev.Graphics.FillRectangle(New SolidBrush(CLR_GOLD_AKZENT),
                                            ev.Bounds.Left, ev.Bounds.Bottom - 3, ev.Bounds.Width, 3)
                                    End If
                                    Dim sf As New StringFormat With {
                                        .Alignment = StringAlignment.Center,
                                        .LineAlignment = StringAlignment.Center
                                    }
                                    ev.Graphics.DrawString(tab.Text, New Font("Segoe UI", 9.0F, If(isSelected, FontStyle.Bold, FontStyle.Regular)),
                                        New SolidBrush(fg), RectangleF.op_Implicit(ev.Bounds), sf)
                                End Sub
    End Sub

    ''' <summary>Stylt einen GroupBox modern (Rahmen blau, Titel blau).</summary>
    Private Sub StyleGroupBox(gb As GroupBox)
        gb.Font = FONT_TITLE
        gb.ForeColor = CLR_BLAU_DUNKEL
    End Sub

    ''' <summary>Erstellt ein modernes TextBox-Label-Paar in einer GroupBox.</summary>
    Private Sub ErstelleFeld(gb As GroupBox, text As String, ctrl As Control, x As Integer, y As Integer, w As Integer)
        Dim lbl As New Label With {
            .Text = text, .Location = New Point(x, y),
            .AutoSize = True, .ForeColor = CLR_TEXT_GRAU,
            .Font = FONT_KLEIN
        }
        ctrl.Location = New Point(x, y + 18)
        ctrl.Width = w
        ctrl.Font = FONT_NORMAL
        If TypeOf ctrl Is TextBox Then
            DirectCast(ctrl, TextBox).BorderStyle = BorderStyle.FixedSingle
            ctrl.BackColor = CLR_WEISS
        ElseIf TypeOf ctrl Is ComboBox Then
            DirectCast(ctrl, ComboBox).FlatStyle = FlatStyle.Flat
            ctrl.BackColor = CLR_WEISS
        End If
        gb.Controls.Add(lbl)
        gb.Controls.Add(ctrl)
    End Sub

    ''' <summary>Erstellt ein mehrzeiliges TextBox-Label-Paar in einer GroupBox.</summary>
    Private Sub ErstelleMultiFeld(gb As GroupBox, text As String, ctrl As TextBox, x As Integer, y As Integer, w As Integer, h As Integer)
        Dim lbl As New Label With {
            .Text = text, .Location = New Point(x, y),
            .AutoSize = True, .ForeColor = CLR_TEXT_GRAU,
            .Font = FONT_KLEIN
        }
        ctrl.Location = New Point(x, y + 18)
        ctrl.Size = New Size(w, h)
        ctrl.Font = FONT_NORMAL
        ctrl.BorderStyle = BorderStyle.FixedSingle
        ctrl.BackColor = CLR_WEISS
        gb.Controls.Add(lbl)
        gb.Controls.Add(ctrl)
    End Sub

    ' =========================================================================
    ' DATEN LADEN
    ' =========================================================================
    Private Sub LadeDaten()
        LadeArtikelListeLinks()
        LadeMitgliederListe()
        LadeArtikelTabelle()

        Using conn = DatenbankManager.HoleVerbindung()
            Dim daMit As New SQLiteDataAdapter("SELECT id, name, standard_preisart FROM mitglieder ORDER BY name", conn)
            Dim dtMit As New DataTable()
            daMit.Fill(dtMit)
            chkMitglieder.DataSource = dtMit
            chkMitglieder.DisplayMember = "name"
            chkMitglieder.ValueMember = "id"

            Dim cmdNr As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'laufende_rechnungsnummer'", conn)
            Dim result = cmdNr.ExecuteScalar()
            If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                lblAktuelleReNr.Text = "re-" & result.ToString()
            End If
        End Using
    End Sub

    Private Sub LadeArtikelListeLinks()
        Using conn = DatenbankManager.HoleVerbindung()
            ' Artikelnummer bewusst nicht mehr im Anzeigetext (nur noch interner Ordnungs-/
            ' Kürzel-Zweck) - die Sortierung nach artikelnummer bleibt aber unverändert, weil
            ' die Nummer genau dafür da ist: häufige Artikel bekommen eine niedrige Nummer und
            ' stehen dadurch oben in der Liste.
            Dim sql = "SELECT id, REPLACE(REPLACE(CASE WHEN LENGTH(bezeichnung) > 25 THEN SUBSTR(bezeichnung, 1, 25) || '...' ELSE bezeichnung END, char(10), ' '), char(13), '') || ' | ' || printf('%.2f', einzelpreis_netto) || ' € (' || mwst_satz || '%)' AS anzeige, bezeichnung, einzelpreis_netto, einzelpreis_brutto, mwst_satz FROM artikel ORDER BY artikelnummer"
            Dim daArt As New SQLiteDataAdapter(sql, conn)
            Dim dtArt As New DataTable()
            daArt.Fill(dtArt)
            lstArtikel.DataSource = dtArt
            lstArtikel.DisplayMember = "anzeige"
            lstArtikel.ValueMember = "id"
        End Using
    End Sub

    ' Zeichnet eine Zeile der Artikel-Auswahlliste in drei festen Spalten: Beschreibung links
    ' (mit "..." abgeschnitten, falls zu lang für den verbleibenden Platz), Preis und MwSt-Satz
    ' jeweils rechtsbündig an einer festen Position - dadurch stehen Preis und MwSt bei jeder
    ' Zeile exakt untereinander, unabhängig von der Länge der Artikelbezeichnung.
    Private Sub LstArtikel_DrawItem(sender As Object, e As DrawItemEventArgs)
        If e.Index < 0 Then Return
        e.DrawBackground()

        Dim row As DataRowView = DirectCast(lstArtikel.Items(e.Index), DataRowView)
        Dim bezeichnung As String = row("bezeichnung").ToString()
        ' Brutto statt Netto (Nutzer denkt/verhandelt in Brutto) - BruttoText fängt Artikel
        ' ohne gespeicherten Bruttopreis ab (Altdaten von vor dem Netto/Brutto-Umbau).
        Dim preisText As String = BruttoText(row("einzelpreis_brutto"), CDec(row("einzelpreis_netto")), CDec(row("mwst_satz"))) & " €"
        Dim mwstText As String = FmtMwSt(CDec(row("mwst_satz")))

        Dim istAusgewaehlt As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        Dim textFarbe As Color = If(istAusgewaehlt, CLR_WEISS, CLR_TEXT_DUNKEL)

        ' Preis rückt dicht an die MwSt heran (nur noch 4px Abstand statt vorher 6px Lücke bei
        ' gleichzeitig schmalerer Preis-Spalte) - dadurch bleibt mehr Platz für die Beschreibung.
        Dim mwstRect As New Rectangle(e.Bounds.Right - 54, e.Bounds.Top, 48, e.Bounds.Height)
        Dim preisRect As New Rectangle(e.Bounds.Right - 132, e.Bounds.Top, 74, e.Bounds.Height)
        Dim beschrRect As New Rectangle(e.Bounds.Left + 6, e.Bounds.Top, e.Bounds.Width - 140, e.Bounds.Height)

        Dim flagsRechts = TextFormatFlags.Right Or TextFormatFlags.VerticalCenter Or TextFormatFlags.SingleLine
        Dim flagsLinks = TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.SingleLine

        TextRenderer.DrawText(e.Graphics, bezeichnung, e.Font, beschrRect, textFarbe, flagsLinks)
        TextRenderer.DrawText(e.Graphics, preisText, e.Font, preisRect, textFarbe, flagsRechts)
        TextRenderer.DrawText(e.Graphics, mwstText, e.Font, mwstRect, textFarbe, flagsRechts)

        e.DrawFocusRectangle()
    End Sub

    Private Sub LadeMitgliederListe()
        Using conn = DatenbankManager.HoleVerbindung()
            Dim sql = "SELECT id, mitgliedsnummer AS 'Nr', name AS 'Name', ort AS 'Ort' FROM mitglieder ORDER BY name"
            Dim da = New SQLiteDataAdapter(sql, conn)
            Dim dt = New DataTable()
            da.Fill(dt)
            dgvMitglieder.DataSource = dt
            If dgvMitglieder.Columns.Contains("id") Then dgvMitglieder.Columns("id").Visible = False
        End Using
    End Sub

    ' =========================================================================
    ' TAB 1: RECHNUNGSERFASSUNG LOGIK
    ' =========================================================================

    ' Muss exakt dieselbe Rechenmethode wie RechnungsDrucker.ErstelleRechnung verwenden -
    ' sonst weicht diese Voransicht von der später gedruckten Rechnung ab (genau das war der
    ' gemeldete Fehler: PDF zeigte nach dem Brutto-Fix den korrekten Betrag, die Voransicht
    ' hier aber noch den alten, Netto-zuerst gerundeten). Pro Satz gebündelt runden, nicht
    ' pro Zeile einzeln - und im Brutto-Modus ist Brutto die Bezugsgröße (Steuer = Brutto -
    ' Netto), nicht andersherum.
    Private Sub BerechneSummenTab1(sender As Object, e As EventArgs)
        Dim isBruttoModus As Boolean = (cbPreisart.SelectedIndex = 1)
        Dim nettoBasisA As Decimal = 0, nettoBasisB As Decimal = 0
        Dim bruttoBasisA As Decimal = 0, bruttoBasisB As Decimal = 0

        For Each ctrl As Control In pnlRows.Controls
            If ctrl.Name <> "Zeile" Then Continue For
            Dim pnl As Panel = DirectCast(ctrl, Panel)

            Dim txtAnzahl = DirectCast(pnl.Controls.Find("txtAnzahl", True)(0), TextBox)
            Dim txtPreisNetto = DirectCast(pnl.Controls.Find("txtPreisNetto", True)(0), TextBox)
            Dim txtPreisBrutto = DirectCast(pnl.Controls.Find("txtPreisBrutto", True)(0), TextBox)
            Dim cbMwSt = DirectCast(pnl.Controls.Find("cbMwSt", True)(0), ComboBox)
            Dim lblGesamt = DirectCast(pnl.Controls.Find("lblGesamt", True)(0), Label)
            Dim txtText = DirectCast(pnl.Controls.Find("txtText", True)(0), TextBox)
            Dim btnSave As Button = TryCast(pnl.Controls.Find("btnSaveArt", True).FirstOrDefault(), Button)

            Dim anzahl As Decimal = 0, preisNettoEingabe As Decimal = 0, preisBruttoEingabe As Decimal = 0
            Dim isPreisAktiv = ParseBetrag(txtPreisNetto.Text, preisNettoEingabe)
            ParseBetrag(txtPreisBrutto.Text, preisBruttoEingabe)
            ParseBetrag(txtAnzahl.Text, anzahl)

            Dim satzProzent As Decimal = ParseMwSt(cbMwSt.Text)
            Dim istSatz2 As Boolean = (Math.Abs(satzProzent - mwstSatz2) < 0.01D)
            Dim mwstSatz As Decimal = satzProzent / 100D

            If isPreisAktiv Then
                If isBruttoModus Then
                    Dim zeilenBrutto As Decimal = Math.Round(anzahl * preisBruttoEingabe, 2, MidpointRounding.AwayFromZero)
                    If istSatz2 Then bruttoBasisB += zeilenBrutto Else bruttoBasisA += zeilenBrutto
                    lblGesamt.Text = zeilenBrutto.ToString("N2") & " €"
                Else
                    Dim zeilenNetto As Decimal = Math.Round(anzahl * preisNettoEingabe, 2, MidpointRounding.AwayFromZero)
                    If istSatz2 Then nettoBasisB += zeilenNetto Else nettoBasisA += zeilenNetto
                    lblGesamt.Text = (zeilenNetto * (1 + mwstSatz)).ToString("N2") & " €"
                End If
            End If

            If btnSave IsNot Nothing Then
                Dim txtBez As String = txtText.Text.Trim()
                If Not String.IsNullOrWhiteSpace(txtBez) AndAlso isPreisAktiv AndAlso cbMwSt.SelectedIndex >= 0 Then
                    btnSave.Enabled = True
                    btnSave.BackColor = CLR_BLAU_HELL
                    btnSave.ForeColor = CLR_BLAU_DUNKEL
                Else
                    btnSave.Enabled = False
                    btnSave.BackColor = Color.FromArgb(220, 220, 220)
                    btnSave.ForeColor = CLR_TEXT_GRAU
                End If
            End If
        Next

        Dim netto As Decimal, mwstA As Decimal, mwstB As Decimal
        If isBruttoModus Then
            Dim nettoA As Decimal = Math.Round(bruttoBasisA / (1 + mwstSatz1 / 100D), 2, MidpointRounding.AwayFromZero)
            Dim nettoB As Decimal = Math.Round(bruttoBasisB / (1 + mwstSatz2 / 100D), 2, MidpointRounding.AwayFromZero)
            mwstA = bruttoBasisA - nettoA
            mwstB = bruttoBasisB - nettoB
            netto = nettoA + nettoB
        Else
            netto = nettoBasisA + nettoBasisB
            mwstA = Math.Round(nettoBasisA * (mwstSatz1 / 100D), 2, MidpointRounding.AwayFromZero)
            mwstB = Math.Round(nettoBasisB * (mwstSatz2 / 100D), 2, MidpointRounding.AwayFromZero)
        End If

        lblSummenTab1.Text = $"  Netto: {netto:N2} €     MwSt {FmtMwSt(mwstSatz1)}: {mwstA:N2} €     MwSt {FmtMwSt(mwstSatz2)}: {mwstB:N2} €     Brutto-Gesamt: {(netto + mwstA + mwstB):N2} €"
        PruefeEingaben()
    End Sub

    Private Sub PruefeEingaben()
        If btnRechnungErstellen Is Nothing Then Return

        Dim isValid As Boolean = False

        If chkMitglieder.CheckedItems.Count = 0 Then
            btnRechnungErstellen.Enabled = False
            btnRechnungErstellen.BackColor = Color.FromArgb(160, 160, 160)
            Return
        End If

        For Each ctrl As Control In pnlRows.Controls
            If ctrl.Name <> "Zeile" Then Continue For
            Dim pnl As Panel = DirectCast(ctrl, Panel)

            Dim txtAnz = DirectCast(pnl.Controls.Find("txtAnzahl", True)(0), TextBox).Text
            Dim txtBez = DirectCast(pnl.Controls.Find("txtText", True)(0), TextBox).Text
            Dim txtPrs = DirectCast(pnl.Controls.Find("txtPreisNetto", True)(0), TextBox).Text

            Dim anzahl As Decimal = 0
            ParseBetrag(txtAnz, anzahl)

            If anzahl > 0 AndAlso Not String.IsNullOrWhiteSpace(txtBez) AndAlso Not String.IsNullOrWhiteSpace(txtPrs) Then
                isValid = True
                Exit For
            End If
        Next

        btnRechnungErstellen.Enabled = isValid
        btnRechnungErstellen.BackColor = If(isValid, CLR_BLAU_MITTEL, Color.FromArgb(160, 160, 160))
    End Sub

    Private Sub TxtText_TextChanged(sender As Object, e As EventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        If txtBox.Text.StartsWith("++") AndAlso txtBox.Text.Length >= 5 Then
            Dim artNr As String = txtBox.Text.Substring(2, 3)
            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmd As New SQLiteCommand("SELECT bezeichnung, einzelpreis_netto, mwst_satz FROM artikel WHERE artikelnummer = @nr", conn)
                cmd.Parameters.AddWithValue("@nr", artNr)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim pnlZeile As Panel = DirectCast(txtBox.Parent.Parent, Panel)
                        txtBox.Text = reader("bezeichnung").ToString()
                        ' Nur Netto setzen und danach den MwSt-Satz - die Sync-Handler in
                        ' ErstelleArtikelZeile rechnen Brutto dann automatisch passend mit,
                        ' kein manuelles Setzen von txtPreisBrutto nötig.
                        DirectCast(pnlZeile.Controls.Find("txtPreisNetto", True)(0), TextBox).Text = CDec(reader("einzelpreis_netto")).ToString("N2")
                        DirectCast(pnlZeile.Controls.Find("cbMwSt", True)(0), ComboBox).Text = FmtMwSt(CDec(reader("mwst_satz")))
                        txtBox.SelectionStart = txtBox.Text.Length
                    End If
                End Using
            End Using
        End If
        PruefeEingaben()
    End Sub

    Private Sub SpeichereRechnung(sender As Object, e As EventArgs)
        Dim heutigesDatum As String = DateTime.Now.ToString("dd.MM.yyyy")
        Dim anzahlErstellt As Integer = 0

        ' --- 1. Alle Positionszeilen VOR dem Schreiben einsammeln und validieren ---
        ' Netto- und Brutto-Feld sind durch die Sync-Logik in ErstelleArtikelZeile immer
        ' beide aktuell und stimmen zusammen mit dem gewählten MwSt-Satz überein - deshalb
        ' werden hier einfach beide so übernommen, wie sie gerade angezeigt werden, ohne
        ' Rückrechnung (das vermeidet die Rundungs-Drift, um die es ja gerade ging).
        Dim positionen As New List(Of (bez As String, anzahl As Decimal, nettoPreis As Decimal, bruttoPreis As Decimal, mwst As Decimal))

        For Each ctrl As Control In pnlRows.Controls
            If ctrl.Name <> "Zeile" Then Continue For
            Dim pnl As Panel = DirectCast(ctrl, Panel)

            Dim anzahl As Decimal = 0
            ParseBetrag(DirectCast(pnl.Controls.Find("txtAnzahl", True)(0), TextBox).Text, anzahl)
            If anzahl <= 0 Then Continue For

            Dim bez As String = DirectCast(pnl.Controls.Find("txtText", True)(0), TextBox).Text
            If String.IsNullOrWhiteSpace(bez) Then Continue For

            Dim preisNettoText As String = DirectCast(pnl.Controls.Find("txtPreisNetto", True)(0), TextBox).Text
            Dim nettoPreis As Decimal = 0
            If Not ParseBetrag(preisNettoText, nettoPreis) Then
                MessageBox.Show($"Die Position '{bez}' hat keinen gültigen Preis ('{preisNettoText}')." & vbCrLf &
                                "Es wurde KEINE Rechnung erstellt. Bitte korrigiere die Eingabe.",
                                "Ungültiger Preis", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim preisBruttoText As String = DirectCast(pnl.Controls.Find("txtPreisBrutto", True)(0), TextBox).Text
            Dim bruttoPreis As Decimal = 0
            ParseBetrag(preisBruttoText, bruttoPreis)

            Dim cbMwStZeile = DirectCast(pnl.Controls.Find("cbMwSt", True)(0), ComboBox)
            Dim mwstProzent As Decimal = ParseMwSt(cbMwStZeile.Text)
            If mwstProzent <= 0 OrElse cbMwStZeile.SelectedIndex < 0 Then
                MessageBox.Show($"Die Position '{bez}' hat keinen gültigen MwSt-Satz." & vbCrLf &
                                $"Es wurde KEINE Rechnung erstellt. Bitte wähle {FmtMwSt(mwstSatz1)} oder {FmtMwSt(mwstSatz2)} aus.",
                                "Ungültige MwSt", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            positionen.Add((bez, anzahl, nettoPreis, bruttoPreis, mwstProzent))
        Next

        If positionen.Count = 0 Then
            MessageBox.Show("Keine gültigen Positionen gefunden. Es wurde keine Rechnung erstellt.",
                            "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' --- 2. Alles in EINER Transaktion schreiben (verhindert doppelte Rechnungsnummern) ---
        Using conn = DatenbankManager.HoleVerbindung()
            Using trans = conn.BeginTransaction()
                Try
                    Dim cmdGetNr As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'laufende_rechnungsnummer'", conn, trans)
                    Dim nrObj = cmdGetNr.ExecuteScalar()
                    If nrObj Is Nothing OrElse DBNull.Value.Equals(nrObj) Then
                        Throw New Exception("Die laufende Rechnungsnummer fehlt in den Einstellungen. Bitte im Tab 'Einstellungen' eine Start-Rechnungsnummer festlegen.")
                    End If
                    Dim aktuelleReNr As Integer = CInt(nrObj)

                    Dim lieferdatum As String = txtLieferdatum.Text.Trim()
                    If String.IsNullOrWhiteSpace(lieferdatum) Then lieferdatum = heutigesDatum

                    Dim preisart As String = If(cbPreisart.SelectedIndex = 1, "Brutto", "Netto")

                    For Each item As Object In chkMitglieder.CheckedItems
                        Dim mitgliedId As Integer = CInt(DirectCast(item, DataRowView)("id"))

                        Dim cmdRe As New SQLiteCommand("INSERT INTO rechnungen (rechnungsnummer, datum, lieferdatum, mitglied_id, status, preisart) VALUES (@nr, @dat, @lief, @mid, 'Erfasst', @preisart); SELECT last_insert_rowid();", conn, trans)
                        cmdRe.Parameters.AddWithValue("@nr", aktuelleReNr.ToString())
                        cmdRe.Parameters.AddWithValue("@dat", heutigesDatum)
                        cmdRe.Parameters.AddWithValue("@lief", lieferdatum)
                        cmdRe.Parameters.AddWithValue("@mid", mitgliedId)
                        cmdRe.Parameters.AddWithValue("@preisart", preisart)
                        Dim neueReId As Integer = CInt(cmdRe.ExecuteScalar())

                        For Each pos In positionen
                            Dim cmdPos As New SQLiteCommand("INSERT INTO rechnungspositionen (rechnung_id, artikel_bezeichnung, anzahl, einzelpreis, einzelpreis_brutto, mwst_satz) VALUES (@rid, @bez, @anz, @prs, @prsBrutto, @mwst)", conn, trans)
                            cmdPos.Parameters.AddWithValue("@rid", neueReId)
                            cmdPos.Parameters.AddWithValue("@bez", pos.bez)
                            cmdPos.Parameters.AddWithValue("@anz", pos.anzahl)
                            cmdPos.Parameters.AddWithValue("@prs", pos.nettoPreis)
                            cmdPos.Parameters.AddWithValue("@prsBrutto", pos.bruttoPreis)
                            cmdPos.Parameters.AddWithValue("@mwst", pos.mwst)
                            cmdPos.ExecuteNonQuery()
                        Next

                        aktuelleReNr += 1
                        anzahlErstellt += 1
                    Next

                    Dim cmdUpdateNr As New SQLiteCommand("UPDATE einstellungen SET wert = @val WHERE schluessel = 'laufende_rechnungsnummer'", conn, trans)
                    cmdUpdateNr.Parameters.AddWithValue("@val", aktuelleReNr.ToString())
                    cmdUpdateNr.ExecuteNonQuery()

                    trans.Commit()
                    lblAktuelleReNr.Text = "re-" & aktuelleReNr.ToString()
                Catch ex As Exception
                    trans.Rollback()
                    MessageBox.Show("Fehler beim Erstellen der Rechnung(en) - es wurde NICHTS gespeichert:" & vbCrLf & vbCrLf & ex.Message,
                                    "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try
            End Using

            Dim cmdMengen As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'prog_mengen_nullen'", conn)
            Dim resMengen = cmdMengen.ExecuteScalar()
            If resMengen IsNot Nothing AndAlso resMengen.ToString() = "True" Then
                For Each ctrl As Control In pnlRows.Controls
                    If ctrl.Name = "Zeile" Then
                        Dim pnl As Panel = DirectCast(ctrl, Panel)
                        Dim txtAnz = DirectCast(pnl.Controls.Find("txtAnzahl", True)(0), TextBox)
                        txtAnz.Text = ""
                    End If
                Next
            End If
        End Using

        For i As Integer = 0 To chkMitglieder.Items.Count - 1
            chkMitglieder.SetItemChecked(i, False)
        Next

        BerechneSummenTab1(Nothing, Nothing)
    End Sub

    ' =========================================================================
    ' HAUPT-TABS WECHSELN
    ' =========================================================================
    Private Sub MainTabs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MainTabs.SelectedIndexChanged
        If MainTabs.SelectedTab.Name = "TabKontrolle" Then
            LadeStapelverarbeitung()
        ElseIf MainTabs.SelectedTab.Name = "TabMitglieder" Then
            LadeMitgliederListe()
        ElseIf MainTabs.SelectedTab.Name = "TabArchiv" Then
            LadeArchiv()
        End If
    End Sub

    ' =========================================================================
    ' ZENTRALE BRUTTO-BERECHNUNG
    ' Einheitliche Rundungslogik für Anzeige, PDF und Excel:
    ' Zeilennetto auf 2 Stellen runden, Steuer auf die gerundete Basis, dann runden.
    ' Identisch mit RechnungsDrucker und ZugferdGenerator - keine Cent-Differenzen.
    ' =========================================================================
    Private Shared Function BerechneBruttoAusDb(conn As SQLiteConnection, reID As Integer) As Decimal
        Dim netto As Decimal = 0
        Dim basisProSatz As New Dictionary(Of Decimal, Decimal)
        Dim cmd As New SQLiteCommand("SELECT anzahl, einzelpreis, mwst_satz FROM rechnungspositionen WHERE rechnung_id = @id", conn)
        cmd.Parameters.AddWithValue("@id", reID)
        Using r = cmd.ExecuteReader()
            While r.Read()
                Dim zn As Decimal = Math.Round(CDec(r("anzahl")) * CDec(r("einzelpreis")), 2, MidpointRounding.AwayFromZero)
                netto += zn
                Dim mwst As Decimal = CDec(r("mwst_satz"))
                If Not basisProSatz.ContainsKey(mwst) Then basisProSatz(mwst) = 0
                basisProSatz(mwst) += zn
            End While
        End Using
        Dim steuerGesamt As Decimal = 0
        For Each kv In basisProSatz
            steuerGesamt += Math.Round(kv.Value * (kv.Key / 100D), 2, MidpointRounding.AwayFromZero)
        Next
        Return netto + steuerGesamt
    End Function

    ' =========================================================================
    ' ALLGEMEINE DETAIL-ANZEIGE
    ' =========================================================================
    Private Sub LadeDetailsAllgemein(reID As Integer, pnlContent As FlowLayoutPanel, lblSumme As Label)
        pnlContent.Controls.Clear()

        Dim gesamtNetto As Decimal = 0
        Dim basisProSatz As New Dictionary(Of Decimal, Decimal)

        Dim zeilenBreite As Integer = Math.Max(650, pnlContent.Width - 25)
        Dim bezBreite As Integer = zeilenBreite - 380

        Using conn = DatenbankManager.HoleVerbindung()
            Dim sql = "SELECT anzahl, artikel_bezeichnung, einzelpreis, mwst_satz FROM rechnungspositionen WHERE rechnung_id = @id"
            Dim cmd As New SQLiteCommand(sql, conn)
            cmd.Parameters.AddWithValue("@id", reID)

            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim anz As Decimal = CDec(reader("anzahl"))
                    Dim bez As String = reader("artikel_bezeichnung").ToString()
                    Dim epreis As Decimal = CDec(reader("einzelpreis"))
                    Dim mwst As Decimal = CDec(reader("mwst_satz"))

                    ' Gleiche Rundungslogik wie PDF/XML
                    Dim netto As Decimal = Math.Round(anz * epreis, 2, MidpointRounding.AwayFromZero)
                    gesamtNetto += netto
                    If Not basisProSatz.ContainsKey(mwst) Then basisProSatz(mwst) = 0
                    basisProSatz(mwst) += netto

                    Dim lblBez As New Label With {
                        .Text = bez,
                        .Location = New Point(70, 10),
                        .Width = bezBreite,
                        .AutoSize = False,
                        .Font = FONT_NORMAL,
                        .ForeColor = CLR_TEXT_DUNKEL
                    }

                    Dim berechneteHoehe As Integer = lblBez.GetPreferredSize(New Size(bezBreite, 0)).Height
                    lblBez.Height = berechneteHoehe + 10
                    Dim zeilenHoehe As Integer = Math.Max(45, lblBez.Height + 20)

                    Dim pnlZeile As New Panel With {
                        .Width = zeilenBreite,
                        .Height = zeilenHoehe,
                        .Margin = New Padding(0, 0, 0, 4),
                        .BackColor = CLR_WEISS
                    }
                    ' Untere Trennlinie
                    AddHandler pnlZeile.Paint, Sub(s As Object, ev As PaintEventArgs)
                                                   ev.Graphics.DrawLine(New Pen(CLR_BORDER, 1), 0, DirectCast(s, Panel).Height - 1, DirectCast(s, Panel).Width, DirectCast(s, Panel).Height - 1)
                                               End Sub

                    Dim lblAnz As New Label With {.Text = anz.ToString(), .Location = New Point(10, 10), .Width = 50, .Font = FONT_NORMAL, .ForeColor = CLR_TEXT_GRAU}
                    Dim lblEpreis As New Label With {.Text = epreis.ToString("N2") & " €", .Location = New Point(zeilenBreite - 300, 10), .Width = 80, .TextAlign = ContentAlignment.TopRight, .Font = FONT_NORMAL, .ForeColor = CLR_TEXT_GRAU}
                    Dim lblMwst As New Label With {.Text = FmtMwSt(mwst), .Location = New Point(zeilenBreite - 200, 10), .Width = 60, .TextAlign = ContentAlignment.TopRight, .Font = FONT_NORMAL, .ForeColor = CLR_TEXT_GRAU}
                    Dim lblNetto As New Label With {.Text = netto.ToString("N2") & " €", .Location = New Point(zeilenBreite - 120, 10), .Width = 80, .Font = New Font("Segoe UI", 10, FontStyle.Bold), .TextAlign = ContentAlignment.TopRight, .ForeColor = CLR_BLAU_DUNKEL}

                    pnlZeile.Controls.AddRange({lblAnz, lblBez, lblEpreis, lblMwst, lblNetto})
                    pnlContent.Controls.Add(pnlZeile)
                End While
            End Using
        End Using

        Dim gesamtSteuer As Decimal = 0
        Dim steuerZeilen As New List(Of String)
        For Each kv In basisProSatz.OrderByDescending(Function(x) x.Key)
            Dim satzSteuer As Decimal = Math.Round(kv.Value * (kv.Key / 100D), 2, MidpointRounding.AwayFromZero)
            gesamtSteuer += satzSteuer
            steuerZeilen.Add($"MwSt {FmtMwSt(kv.Key)}: {satzSteuer:N2} €")
        Next
        Dim gesamtBrutto As Decimal = gesamtNetto + gesamtSteuer

        lblSumme.Text = $"Netto: {gesamtNetto:N2} €" & vbCrLf &
                               String.Join(vbCrLf, steuerZeilen) & vbCrLf & vbCrLf &
                               $"RECHNUNGSBETRAG: {gesamtBrutto:N2} €"
    End Sub

    ' =========================================================================
    ' TAB 2: STAPELVERARBEITUNG & KONTROLLE
    ' =========================================================================
    Private Sub LadeStapelverarbeitung()
        Using conn = DatenbankManager.HoleVerbindung()
            Dim s1 As String = mwstSatz1.ToString(Globalization.CultureInfo.InvariantCulture)
            Dim s2 As String = mwstSatz2.ToString(Globalization.CultureInfo.InvariantCulture)
            Dim f1 As String = (mwstSatz1 / 100D).ToString(Globalization.CultureInfo.InvariantCulture)
            Dim f2 As String = (mwstSatz2 / 100D).ToString(Globalization.CultureInfo.InvariantCulture)
            Dim sql = "SELECT r.id, r.rechnungsnummer AS 'Re-Nr', m.name AS 'Empfänger', " &
                      "(SELECT ROUND(SUM(ROUND(anzahl*einzelpreis,2)) " &
                      " + ROUND(SUM(CASE WHEN mwst_satz=" & s1 & " THEN ROUND(anzahl*einzelpreis,2) ELSE 0 END)*" & f1 & ",2) " &
                      " + ROUND(SUM(CASE WHEN mwst_satz=" & s2 & " THEN ROUND(anzahl*einzelpreis,2) ELSE 0 END)*" & f2 & ",2),2) " &
                      " FROM rechnungspositionen WHERE rechnung_id = r.id) AS 'Brutto' " &
                      "FROM rechnungen r JOIN mitglieder m ON r.mitglied_id = m.id WHERE r.status = 'Erfasst'"
            Dim daRe As New SQLiteDataAdapter(sql, conn)
            Dim dtRe As New DataTable()
            daRe.Fill(dtRe)

            dgvRechnungen.DataSource = dtRe
            If dgvRechnungen.Columns.Contains("id") Then dgvRechnungen.Columns("id").Visible = False

            If dgvRechnungen.Columns.Contains("Brutto") Then
                dgvRechnungen.Columns("Brutto").DefaultCellStyle.Format = "N2"
                dgvRechnungen.Columns("Brutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If

            Dim sumBrutto As Decimal = 0
            For Each row As DataRow In dtRe.Rows
                If Not IsDBNull(row("Brutto")) Then sumBrutto += CDec(row("Brutto"))
            Next
            lblSummenTab2.Text = $"  STAPEL-SUMMEN  ·  Gesamt-Brutto: {sumBrutto:N2} €"
        End Using

        If dgvRechnungen.Rows.Count > 0 Then
            DgvRechnungen_SelectionChanged(Nothing, Nothing)
        Else
            pnlDetailsContent.Controls.Clear()
            lblDetailsSumme.Text = $"Netto: 0,00 €" & vbCrLf & $"MwSt {FmtMwSt(mwstSatz1)}: 0,00 €" & vbCrLf & $"MwSt {FmtMwSt(mwstSatz2)}: 0,00 €" & vbCrLf & vbCrLf & "RECHNUNGSBETRAG: 0,00 €"
        End If
    End Sub

    Private Sub PruefeSummen(sender As Object, e As EventArgs)
        Dim txtHaendler As TextBox = DirectCast(sender, TextBox)
        Dim haendlerSumme As Decimal = 0
        ParseBetrag(txtHaendler.Text, haendlerSumme)

        Dim stapelSumme As Decimal = 0
        For Each row As DataGridViewRow In dgvRechnungen.Rows
            If Not row.IsNewRow AndAlso row.Cells("Brutto").Value IsNot DBNull.Value Then
                stapelSumme += CDec(row.Cells("Brutto").Value)
            End If
        Next

        Dim diff As Decimal = stapelSumme - haendlerSumme
        lblKontrolleTab2.Text = $"Differenz: {diff:N2} €"
        lblKontrolleTab2.ForeColor = If(diff = 0 And haendlerSumme > 0, CLR_BLAU_MITTEL, CLR_ROT)
        If diff = 0 And haendlerSumme > 0 Then lblKontrolleTab2.Text &= " ✓ PASST"
    End Sub

    Private Sub DgvRechnungen_SelectionChanged(sender As Object, e As EventArgs)
        If dgvRechnungen.CurrentRow IsNot Nothing AndAlso Not dgvRechnungen.CurrentRow.IsNewRow Then
            Dim reID As Integer = 0
            If Integer.TryParse(dgvRechnungen.CurrentRow.Cells("id").Value.ToString(), reID) Then
                LadeDetailsAllgemein(reID, pnlDetailsContent, lblDetailsSumme)
            End If
        Else
            pnlDetailsContent.Controls.Clear()
        End If
    End Sub

    Private Sub BtnLoeschenTab2_Click(sender As Object, e As EventArgs)
        If dgvRechnungen.CurrentRow Is Nothing OrElse dgvRechnungen.CurrentRow.IsNewRow Then
            MessageBox.Show("Bitte wähle zuerst eine Rechnung aus der Liste aus.", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim reID As Integer = CInt(dgvRechnungen.CurrentRow.Cells("id").Value)
        Dim reNrStr As String = dgvRechnungen.CurrentRow.Cells("Re-Nr").Value.ToString()
        Dim geloeschteNr As Integer = 0
        If Not Integer.TryParse(reNrStr, geloeschteNr) Then
            MessageBox.Show($"Die Rechnungsnummer '{reNrStr}' ist keine gültige Zahl - die Rechnung kann nicht automatisch umnummeriert werden.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim antwort = MessageBox.Show($"Möchtest du die Rechnung {reNrStr} wirklich löschen?" & vbCrLf & vbCrLf &
                                      "Alle neueren Rechnungen in diesem Stapel rücken automatisch eine Nummer nach unten, damit keine Lücke entsteht!",
                                      "Rechnung löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If antwort = DialogResult.Yes Then
            Using conn = DatenbankManager.HoleVerbindung()
                Using trans = conn.BeginTransaction()
                    Try
                        Dim cmdDelPos As New SQLiteCommand("DELETE FROM rechnungspositionen WHERE rechnung_id = @id", conn, trans)
                        cmdDelPos.Parameters.AddWithValue("@id", reID)
                        cmdDelPos.ExecuteNonQuery()

                        Dim cmdDelRe As New SQLiteCommand("DELETE FROM rechnungen WHERE id = @id", conn, trans)
                        cmdDelRe.Parameters.AddWithValue("@id", reID)
                        cmdDelRe.ExecuteNonQuery()

                        Dim cmdUpdate As New SQLiteCommand("UPDATE rechnungen SET rechnungsnummer = CAST((CAST(rechnungsnummer AS INTEGER) - 1) AS TEXT) WHERE status = 'Erfasst' AND CAST(rechnungsnummer AS INTEGER) > @geloeschteNr", conn, trans)
                        cmdUpdate.Parameters.AddWithValue("@geloeschteNr", geloeschteNr)
                        cmdUpdate.ExecuteNonQuery()

                        Dim cmdZaehler As New SQLiteCommand("UPDATE einstellungen SET wert = CAST((CAST(wert AS INTEGER) - 1) AS TEXT) WHERE schluessel = 'laufende_rechnungsnummer'", conn, trans)
                        cmdZaehler.ExecuteNonQuery()

                        trans.Commit()
                    Catch ex As Exception
                        trans.Rollback()
                        MessageBox.Show("Fehler beim Löschen - es wurde NICHTS geändert:" & vbCrLf & vbCrLf & ex.Message,
                                        "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try
                End Using
            End Using

            LadeStapelverarbeitung()
            LadeDaten()
            MessageBox.Show("Die Rechnung wurde gelöscht und die nachfolgenden Nummern wurden lückenlos angepasst!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub DruckeDokument(pdfPfad As String, druckerName As String)
        Try
            If Not File.Exists(pdfPfad) Then Return
            Dim psi As New ProcessStartInfo() With {
                .FileName = pdfPfad,
                .Verb = "printto",
                .Arguments = $"""{druckerName}""",
                .CreateNoWindow = True,
                .WindowStyle = ProcessWindowStyle.Hidden
            }
            Dim prozess As Process = Process.Start(psi)
            If prozess IsNot Nothing Then
                prozess.WaitForExit(10000)
                If Not prozess.HasExited Then prozess.Kill()
            End If
        Catch ex As Exception
            MessageBox.Show($"Fehler beim Drucken der Datei {Path.GetFileName(pdfPfad)}:" & vbCrLf & ex.Message, "Druckfehler", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub VerarbeiteStapel(sender As Object, e As EventArgs)
        If dgvRechnungen.Rows.Count = 0 Then
            MessageBox.Show("Es gibt keine Rechnungen zum Verarbeiten.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim antwort = MessageBox.Show("Möchtest du den aktuellen Stapel jetzt verarbeiten und alle Aufgaben (PDFs, Mails, Druck) automatisch ausführen?", "Stapelverarbeitung starten", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If antwort = DialogResult.No Then Return

        Dim btn As Button = DirectCast(sender, Button)
        Dim originalText As String = btn.Text
        Dim anzahl As Integer = 0
        Dim anzahlMails As Integer = 0

        btn.Enabled = False
        btn.Text = "VERARBEITUNG LÄUFT …"
        Application.DoEvents()

        Dim baseDir As String = "C:\MHRechnung"
        Dim verarbeiteteRechnungen As New List(Of Dictionary(Of String, String))()
        Dim fehlerListe As New List(Of String)()
        Dim archivFehlerListe As New List(Of String)()

        Dim standardDrucker As String = ""
        Dim druckKopien As Integer = 1
        Dim druckerGefunden As Boolean = False

        Try
        Using conn = DatenbankManager.HoleVerbindung()
            Dim cmdPfad As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'speicherpfad'", conn)
            Dim pfadObj = cmdPfad.ExecuteScalar()
            If pfadObj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(pfadObj.ToString()) Then
                baseDir = pfadObj.ToString()
            End If

            Dim cmdDr = New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'prog_standard_drucker'", conn)
            Dim resDr = cmdDr.ExecuteScalar()
            If resDr IsNot Nothing Then standardDrucker = resDr.ToString()

            If Not String.IsNullOrWhiteSpace(standardDrucker) Then
                For Each p As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
                    If p = standardDrucker Then
                        druckerGefunden = True
                        Exit For
                    End If
                Next

                If Not druckerGefunden Then
                    MessageBox.Show($"Dein eingestellter Standard-Drucker '{standardDrucker}' wurde auf diesem PC nicht gefunden!" & vbCrLf & vbCrLf & "Er wurde möglicherweise umbenannt, ausgestöpselt oder gelöscht." & vbCrLf & "Keine Panik: Der automatische Druck wird für diesen Durchlauf einfach übersprungen. Bitte wähle in den Einstellungen später den Drucker neu aus.", "Drucker nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If

            Dim cmdKopien = New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'prog_druck_anzahl'", conn)
            Dim resKopien = cmdKopien.ExecuteScalar()
            If resKopien IsNot Nothing AndAlso Integer.TryParse(resKopien.ToString(), druckKopien) = False Then
                druckKopien = 1
            End If

            Dim sql = "SELECT r.id, r.rechnungsnummer, m.id as mitglied_id, m.email, m.versandart, " &
                      "m.name, " &
                      "(SELECT SUM(anzahl * einzelpreis * (1 + mwst_satz/100.0)) FROM rechnungspositionen WHERE rechnung_id = r.id) AS brutto " &
                      "FROM rechnungen r JOIN mitglieder m ON r.mitglied_id = m.id WHERE r.status = 'Erfasst'"

            ' --- 1. Erst ALLE Rechnungsdaten einlesen (Reader schließen, bevor Updates laufen) ---
            Dim cmd = New SQLiteCommand(sql, conn)
            Dim alleRechnungen As New List(Of Dictionary(Of String, String))()
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim rData As New Dictionary(Of String, String)
                    rData("id") = CInt(reader("id")).ToString()
                    rData("reNr") = reader("rechnungsnummer").ToString()
                    rData("mid") = CInt(reader("mitglied_id")).ToString()
                    rData("email") = reader("email").ToString()
                    rData("versandart") = reader("versandart").ToString()
                    rData("name") = reader("name").ToString()
                    rData("brutto") = If(IsDBNull(reader("brutto")), 0D, CDec(reader("brutto"))).ToString()

                    alleRechnungen.Add(rData)
                End While
            End Using

            ' Brutto zentral mit der einheitlichen Rundungslogik berechnen (statt SQL-Float):
            ' Damit ist der Rechnungsbetrag garantiert identisch mit PDF und XML.
            For Each rData In alleRechnungen
                rData("brutto") = BerechneBruttoAusDb(conn, CInt(rData("id"))).ToString()
            Next

            ' --- 2. Jede Rechnung einzeln verarbeiten; Fehler überspringen statt alles abzubrechen ---
            Dim jahr As String = DateTime.Now.Year.ToString()

            For Each rData In alleRechnungen
                anzahl += 1
                btn.Text = $"Verarbeite {anzahl} …"
                Application.DoEvents()

                Dim reID As Integer = CInt(rData("id"))
                Dim reNr As String = rData("reNr")
                Dim mid As Integer = CInt(rData("mid"))
                Dim mitgliedMail As String = rData("email")
                Dim versandart As String = rData("versandart")

                ' PDF + E-Rechnung erzeugen. Schlägt das fehl, bleibt der Status 'Erfasst'
                ' und die Rechnung kann nach Behebung erneut verarbeitet werden.
                Try
                    RechnungsDrucker.ErstelleRechnung(reID, reNr)
                    ZugferdGenerator.ErstelleXML(reID, reNr, mid)
                    MustangRunner.ErstelleZUGFeRDPdf(reNr)
                Catch ex As Exception
                    fehlerListe.Add($"Rechnung {reNr}: {ex.Message}")
                    Continue For
                End Try

                Dim fertigesPdf As String = Path.Combine(baseDir, jahr, "erstellt", "e-rechnungen", reNr & "_zugf.pdf")

                If versandart = "E-Mail" OrElse versandart = "Beides" Then
                    If Not String.IsNullOrWhiteSpace(mitgliedMail) Then
                        Try
                            EmailManager.SendeRechnung(reNr, mitgliedMail, fertigesPdf)
                            anzahlMails += 1
                        Catch ex As Exception
                            fehlerListe.Add($"Rechnung {reNr}: E-Mail an {mitgliedMail} fehlgeschlagen - {ex.Message}")
                        End Try
                    End If
                End If

                If druckerGefunden AndAlso (versandart = "Post" OrElse versandart = "Beides") Then
                    For i As Integer = 1 To druckKopien
                        DruckeDokument(fertigesPdf, standardDrucker)
                    Next
                End If

                ' Archiv-Kopie ("Kopie JEDER versendeten Rechnung") gilt für jede verarbeitete
                ' Rechnung, unabhängig von der Zustellart an den Kunden (Post/E-Mail/Beides) -
                ' sie dient der eigenen Dokumentation, nicht dem Kundenversand. Deshalb hier
                ' unconditional, nicht mehr an "versandart = E-Mail" gekoppelt. Ist nur ein
                ' Komfort-Feature - ein Fehler hier darf den erfolgreichen Rechnungsversand
                ' bzw. -druck nicht als fehlgeschlagen melden (kein Continue For, kein Abbruch),
                ' wird aber nicht mehr verschluckt, sondern separat gesammelt und im
                ' Abschlussbericht angezeigt.
                Try
                    EmailManager.SendeAusgangskopie(reNr, rData("name"), mitgliedMail, CDec(rData("brutto")), fertigesPdf)
                Catch exArchiv As Exception
                    archivFehlerListe.Add($"Re-{reNr}: {exArchiv.Message}")
                End Try

                Dim upCmd = New SQLiteCommand("UPDATE rechnungen SET status = 'Verarbeitet & Exportiert' WHERE id = @id", conn)
                upCmd.Parameters.AddWithValue("@id", reID)
                upCmd.ExecuteNonQuery()

                verarbeiteteRechnungen.Add(rData)
            Next
        End Using

        ' Keine Excel-Zusammenfassung mehr: Die diente ursprünglich als Begleitliste
        ' zur SEPA-Sammelüberweisung/-lastschrift für die Bank. Ohne SEPA hat sie
        ' keinen Zweck mehr - die Übersicht gibt's bei Bedarf im Rechnungs-Archiv.
        ' Stattdessen einfach den Ausgabeordner öffnen, damit die gerade erstellten
        ' PDFs/E-Rechnungen gleich sichtbar sind.
        If verarbeiteteRechnungen.Count > 0 Then
            Dim jahr As String = DateTime.Now.Year.ToString()
            Dim ausgabeOrdner As String = Path.Combine(baseDir, jahr, "erstellt")
            Try
                If Directory.Exists(ausgabeOrdner) Then Process.Start("explorer.exe", ausgabeOrdner)
            Catch
                ' Öffnen des Explorers ist nur Komfort - kein Grund, den Bericht als Fehler zu werten
            End Try
        End If

        Catch ex As Exception
            ' Unerwarteter Fehler außerhalb der Einzelrechnungs-Behandlung
            fehlerListe.Add("Unerwarteter Abbruch: " & ex.Message)
        Finally
            ' Button IMMER wiederherstellen, egal was passiert ist
            btn.Enabled = True
            btn.Text = originalText
        End Try

        LadeStapelverarbeitung()

        ' --- Abschlussbericht ---
        Dim bericht As String = $"{verarbeiteteRechnungen.Count} von {anzahl} Rechnungen wurden erfolgreich verarbeitet." & vbCrLf &
                                $"{anzahlMails} E-Mails wurden versendet."
        If archivFehlerListe.Count > 0 Then
            bericht &= vbCrLf & vbCrLf & "ℹ Archiv-Kopie(n) konnten nicht gesendet werden (die Rechnung selbst wurde trotzdem korrekt versendet):" & vbCrLf &
                       String.Join(vbCrLf, archivFehlerListe)
        End If
        If fehlerListe.Count > 0 Then
            bericht &= vbCrLf & vbCrLf & "⚠ FEHLER (diese Rechnungen bleiben auf 'Erfasst' und können erneut verarbeitet werden):" & vbCrLf &
                       String.Join(vbCrLf, fehlerListe)
            MessageBox.Show(bericht, "Workflow mit Fehlern abgeschlossen", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf archivFehlerListe.Count > 0 Then
            MessageBox.Show(bericht, "Workflow Abgeschlossen (mit Hinweis)", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show(bericht, "Workflow Abgeschlossen", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    ' =========================================================================
    ' TAB 3: ARTIKELVERWALTUNG & EXCEL
    ' =========================================================================
    Private Sub LadeArtikelTabelle()
        Using conn = DatenbankManager.HoleVerbindung()
            ' Zeigt Brutto statt Netto (Nutzer denkt/verhandelt in Brutto) - COALESCE fängt
            ' Artikel ohne gespeicherten Bruttopreis ab (z.B. noch nicht bearbeitete Altdaten
            ' von vor dem Netto/Brutto-Umbau).
            Dim da = New SQLiteDataAdapter("SELECT artikelnummer AS 'Nr', bezeichnung AS 'Beschreibung', mwst_satz AS 'MwSt (%)', COALESCE(einzelpreis_brutto, einzelpreis_netto * (1 + mwst_satz / 100.0)) AS 'Einzelpreis (Brutto)' FROM artikel ORDER BY artikelnummer", conn)
            Dim dt = New DataTable()
            da.Fill(dt)
            dgvArtikelVerwaltung.DataSource = dt
        End Using

        dgvArtikelVerwaltung.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        ' Eigene, größere Schrift nur für diese Tabelle (bisher FONT_NORMAL = 9,5pt, wirkte zu
        ' winzig) - bewusst nicht FONT_NORMAL selbst geändert, das würde das ganze Programm
        ' betreffen. Zeilenhöhe passend mitvergrößert (32 -> 42px).
        dgvArtikelVerwaltung.DefaultCellStyle.Font = New Font("Segoe UI", 12)
        dgvArtikelVerwaltung.RowTemplate.Height = 42

        ' Artikelnummer dient nur noch intern der Reihenfolge (niedrige Nummer = oben) und dem
        ' "++Nummer"-Kürzel bei der Rechnungserfassung - in der Liste selbst ist sie nur noch
        ' Ballast, deshalb hier ausgeblendet (bleibt aber als Spalte gebunden, siehe
        ' BearbeiteArtikelZeile/LoescheArtikelZeile/VerschiebeArtikel, die row.Cells("Nr")
        ' bzw. eine eigene Abfrage nutzen). Sichtbar bleibt sie im Bearbeiten-Fenster (Titel)
        ' und im Excel-Export/Import.
        If dgvArtikelVerwaltung.Columns.Contains("Nr") Then
            dgvArtikelVerwaltung.Columns("Nr").Visible = False
        End If

        If dgvArtikelVerwaltung.Columns.Contains("Beschreibung") Then
            dgvArtikelVerwaltung.Columns("Beschreibung").Width = 479
            dgvArtikelVerwaltung.Columns("Beschreibung").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        End If

        If dgvArtikelVerwaltung.Columns.Contains("MwSt (%)") Then
            dgvArtikelVerwaltung.Columns("MwSt (%)").Width = 80
            dgvArtikelVerwaltung.Columns("MwSt (%)").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End If

        If dgvArtikelVerwaltung.Columns.Contains("Einzelpreis (Brutto)") Then
            dgvArtikelVerwaltung.Columns("Einzelpreis (Brutto)").Width = 140
            dgvArtikelVerwaltung.Columns("Einzelpreis (Brutto)").DefaultCellStyle.Format = "N2"
            dgvArtikelVerwaltung.Columns("Einzelpreis (Brutto)").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If
    End Sub

    Private Sub DgvArtikel_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 OrElse dgvArtikelVerwaltung.Rows(e.RowIndex).IsNewRow Then Return
        BearbeiteArtikelZeile(dgvArtikelVerwaltung.Rows(e.RowIndex))
    End Sub

    ' Gemeinsame Bearbeiten-Logik für Doppelklick UND das Rechtsklick-Kontextmenü -
    ' unverändert übernommen, nur nicht mehr an ein bestimmtes Klick-Ereignis gebunden.
    Private Sub BearbeiteArtikelZeile(row As DataGridViewRow)
        Dim artNr = row.Cells("Nr").Value.ToString()

        ' Frisch aus der DB laden statt aus der Grid-Zeile: die Tabelle zeigt seit dem
        ' Netto/Brutto-Umbau nur noch Brutto (berechnet), der Netto-Wert steht dort gar nicht
        ' mehr drin - für die zwei Bearbeiten-Felder werden beide echten Werte gebraucht.
        Dim bezeichnungAktuell As String = "", nettoAktuell As Decimal = 0, mwstAktuell As Decimal = 0
        Dim bruttoWertAktuell As Object = Nothing
        Using conn0 = DatenbankManager.HoleVerbindung()
            Dim cmd0 As New SQLiteCommand("SELECT bezeichnung, einzelpreis_netto, einzelpreis_brutto, mwst_satz FROM artikel WHERE artikelnummer = @nr", conn0)
            cmd0.Parameters.AddWithValue("@nr", artNr)
            Using r0 = cmd0.ExecuteReader()
                If r0.Read() Then
                    bezeichnungAktuell = r0("bezeichnung").ToString()
                    nettoAktuell = CDec(r0("einzelpreis_netto"))
                    bruttoWertAktuell = r0("einzelpreis_brutto")
                    mwstAktuell = CDec(r0("mwst_satz"))
                End If
            End Using
        End Using

        Using editorForm As New Form With {
            .Text = $"Artikel {artNr} bearbeiten",
            .Size = New Size(530, 550),
            .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedToolWindow,
            .BackColor = CLR_HINTERGRUND
        }
            Dim pnlBottom As New Panel With {.Dock = DockStyle.Bottom, .Height = 60, .Width = editorForm.ClientSize.Width, .BackColor = CLR_PANEL_BG}
            Dim btnOK = MachePrimaerButton("Speichern", 110, 35)
            btnOK.DialogResult = DialogResult.OK
            btnOK.Location = New Point(pnlBottom.Width - 130, 12)
            btnOK.Anchor = AnchorStyles.Top Or AnchorStyles.Right

            Dim btnCancel = MacheSekundaerButton("Abbrechen", 110, 35)
            btnCancel.DialogResult = DialogResult.Cancel
            btnCancel.Location = New Point(pnlBottom.Width - 250, 12)
            btnCancel.Anchor = AnchorStyles.Top Or AnchorStyles.Right

            Dim btnDelete = MacheGefahrButton("Löschen", 110, 35)
            btnDelete.Location = New Point(15, 12)
            btnDelete.Anchor = AnchorStyles.Top Or AnchorStyles.Left

            pnlBottom.Controls.AddRange({btnOK, btnCancel, btnDelete})

            Dim txtDesc As New TextBox With {
                .Multiline = True, .Text = bezeichnungAktuell,
                .Font = New Font("Arial", 12), .ScrollBars = ScrollBars.Vertical,
                .Width = 479, .Height = 350, .Location = New Point(20, 10),
                .BorderStyle = BorderStyle.FixedSingle, .BackColor = CLR_WEISS
            }

            ' Zwei synchronisierte Preisfelder wie in der Rechnungserfassung - welches Feld
            ' zuletzt getippt wird, gilt als die "echte" Zahl, das andere wird live mitgerechnet.
            Dim pnlWerte As New Panel With {.Location = New Point(20, txtDesc.Bottom + 15), .Size = New Size(479, 50)}
            Dim lblNetto As New Label With {.Text = "Netto (€):", .Font = FONT_NORMAL, .Location = New Point(0, 5), .AutoSize = True, .ForeColor = CLR_TEXT_GRAU}
            Dim txtPreisNetto As New TextBox With {
                .Text = nettoAktuell.ToString("N2"), .Font = New Font("Arial", 12),
                .Location = New Point(65, 2), .Width = 75, .TextAlign = HorizontalAlignment.Right,
                .BorderStyle = BorderStyle.FixedSingle, .BackColor = CLR_WEISS
            }
            Dim lblBrutto As New Label With {.Text = "Brutto (€):", .Font = FONT_NORMAL, .Location = New Point(150, 5), .AutoSize = True, .ForeColor = CLR_TEXT_GRAU}
            Dim txtPreisBrutto As New TextBox With {
                .Text = BruttoText(bruttoWertAktuell, nettoAktuell, mwstAktuell), .Font = New Font("Arial", 12),
                .Location = New Point(215, 2), .Width = 75, .TextAlign = HorizontalAlignment.Right,
                .BorderStyle = BorderStyle.FixedSingle, .BackColor = CLR_WEISS
            }
            Dim lblMwSt As New Label With {.Text = "MwSt:", .Font = FONT_NORMAL, .Location = New Point(300, 5), .AutoSize = True, .ForeColor = CLR_TEXT_GRAU}
            Dim cbMwSt As New ComboBox With {
                .Font = New Font("Arial", 12), .Location = New Point(350, 2),
                .Width = 80, .DropDownStyle = ComboBoxStyle.DropDownList,
                .FlatStyle = FlatStyle.Flat, .BackColor = CLR_WEISS
            }
            cbMwSt.Items.AddRange({FmtMwSt(mwstSatz1), FmtMwSt(mwstSatz2)})
            cbMwSt.SelectedIndex = If(Math.Abs(mwstAktuell - mwstSatz2) < 0.01D, 1, 0)
            pnlWerte.Controls.AddRange({lblNetto, txtPreisNetto, lblBrutto, txtPreisBrutto, lblMwSt, cbMwSt})

            Dim aktualisiertGerade As Boolean = False
            AddHandler txtPreisNetto.TextChanged, Sub()
                                                       If aktualisiertGerade Then Return
                                                       Dim netto As Decimal = 0
                                                       If ParseBetrag(txtPreisNetto.Text, netto) Then
                                                           aktualisiertGerade = True
                                                           Dim satz As Decimal = ParseMwSt(cbMwSt.Text)
                                                           txtPreisBrutto.Text = (netto * (1 + satz / 100D)).ToString("N2")
                                                           aktualisiertGerade = False
                                                       End If
                                                   End Sub
            AddHandler txtPreisBrutto.TextChanged, Sub()
                                                        If aktualisiertGerade Then Return
                                                        Dim brutto As Decimal = 0
                                                        If ParseBetrag(txtPreisBrutto.Text, brutto) Then
                                                            aktualisiertGerade = True
                                                            Dim satz As Decimal = ParseMwSt(cbMwSt.Text)
                                                            txtPreisNetto.Text = (brutto / (1 + satz / 100D)).ToString("N2")
                                                            aktualisiertGerade = False
                                                        End If
                                                    End Sub
            AddHandler cbMwSt.SelectedIndexChanged, Sub()
                                                         Dim netto As Decimal = 0
                                                         If ParseBetrag(txtPreisNetto.Text, netto) Then
                                                             aktualisiertGerade = True
                                                             Dim satz As Decimal = ParseMwSt(cbMwSt.Text)
                                                             txtPreisBrutto.Text = (netto * (1 + satz / 100D)).ToString("N2")
                                                             aktualisiertGerade = False
                                                         End If
                                                     End Sub

            editorForm.Controls.AddRange({txtDesc, pnlWerte, pnlBottom})
            editorForm.CancelButton = btnCancel

            AddHandler btnDelete.Click, Sub()
                                            If MessageBox.Show("Artikel wirklich unwiderruflich löschen?", "Achtung", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                                                Using conn = DatenbankManager.HoleVerbindung()
                                                    Dim cmd As New SQLiteCommand("DELETE FROM artikel WHERE artikelnummer = @nr", conn)
                                                    cmd.Parameters.AddWithValue("@nr", artNr)
                                                    cmd.ExecuteNonQuery()
                                                End Using
                                                editorForm.DialogResult = DialogResult.Abort
                                                editorForm.Close()
                                            End If
                                        End Sub

            Dim res = editorForm.ShowDialog()
            If res = DialogResult.OK Then
                Dim neuerNetto As Decimal = 0
                ParseBetrag(txtPreisNetto.Text, neuerNetto)
                Dim neuerBrutto As Decimal = 0
                ParseBetrag(txtPreisBrutto.Text, neuerBrutto)
                Using conn = DatenbankManager.HoleVerbindung()
                    Dim cmd As New SQLiteCommand("UPDATE artikel SET bezeichnung = @bez, einzelpreis_netto = @prs, einzelpreis_brutto = @prsBrutto, mwst_satz = @mwst WHERE artikelnummer = @nr", conn)
                    cmd.Parameters.AddWithValue("@bez", txtDesc.Text)
                    cmd.Parameters.AddWithValue("@prs", neuerNetto)
                    cmd.Parameters.AddWithValue("@prsBrutto", neuerBrutto)
                    cmd.Parameters.AddWithValue("@mwst", ParseMwSt(cbMwSt.Text))
                    cmd.Parameters.AddWithValue("@nr", artNr)
                    cmd.ExecuteNonQuery()
                End Using
                LadeArtikelTabelle()
                LadeArtikelListeLinks()
            ElseIf res = DialogResult.Abort Then
                LadeArtikelTabelle()
                LadeArtikelListeLinks()
            End If
        End Using
    End Sub

    ' Löschen direkt übers Rechtsklick-Menü, ohne erst das Bearbeiten-Fenster zu öffnen -
    ' die Sicherheitsabfrage bleibt wie beim Löschen-Button im Bearbeiten-Fenster erhalten.
    Private Sub LoescheArtikelZeile(row As DataGridViewRow)
        Dim artNr = row.Cells("Nr").Value.ToString()
        If MessageBox.Show("Artikel wirklich unwiderruflich löschen?", "Achtung", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmd As New SQLiteCommand("DELETE FROM artikel WHERE artikelnummer = @nr", conn)
                cmd.Parameters.AddWithValue("@nr", artNr)
                cmd.ExecuteNonQuery()
            End Using
            LadeArtikelTabelle()
            LadeArtikelListeLinks()
        End If
    End Sub

    ' Verschiebt einen Artikel per Drag & Drop an eine neue Position. Es gibt bewusst keine
    ' separate Sortier-Spalte: Die Artikelnummer selbst ist der Ordnungs-Maßstab (niedrige
    ' Nummer = oben, siehe ORDER BY artikelnummer in LadeArtikelTabelle/LadeArtikelListeLinks) -
    ' beim Verschieben werden deshalb einfach alle Artikel anhand der neuen Reihenfolge lückenlos
    ' neu durchnummeriert (001, 002, ...), genau wie beim automatischen Vergeben einer neuen
    ' Nummer an anderer Stelle im Programm.
    Private Sub VerschiebeArtikel(vonIndex As Integer, nachIndex As Integer)
        Dim ids As New List(Of Integer)
        Using conn = DatenbankManager.HoleVerbindung()
            Dim cmd As New SQLiteCommand("SELECT id FROM artikel ORDER BY artikelnummer", conn)
            Using r = cmd.ExecuteReader()
                While r.Read()
                    ids.Add(CInt(r("id")))
                End While
            End Using
        End Using

        If vonIndex < 0 OrElse vonIndex >= ids.Count OrElse nachIndex < 0 OrElse nachIndex >= ids.Count Then Return

        Dim verschobeneId As Integer = ids(vonIndex)
        ids.RemoveAt(vonIndex)
        ids.Insert(nachIndex, verschobeneId)

        Using conn = DatenbankManager.HoleVerbindung()
            Using trans = conn.BeginTransaction()
                For i As Integer = 0 To ids.Count - 1
                    Dim cmdUp As New SQLiteCommand("UPDATE artikel SET artikelnummer = @nr WHERE id = @id", conn, trans)
                    cmdUp.Parameters.AddWithValue("@nr", (i + 1).ToString("D3"))
                    cmdUp.Parameters.AddWithValue("@id", ids(i))
                    cmdUp.ExecuteNonQuery()
                Next
                trans.Commit()
            End Using
        End Using

        LadeArtikelTabelle()
        LadeArtikelListeLinks()
    End Sub

    ' Rechtsklick: markiert die Zeile unter dem Mauszeiger, bevor das Kontextmenü aufgeht
    ' (sonst würde sich das Menü auf die vorher markierte Zeile beziehen).
    ' Linksklick: merkt sich die Startzeile für ein mögliches Drag & Drop (siehe MouseUp) -
    ' interferiert nicht mit normalem Klick/Doppelklick, weil bei einem reinen Klick ohne
    ' Bewegung Start- und Zielzeile identisch sind und MouseUp dann nichts verschiebt.
    Private Sub DgvArtikel_MouseDown(sender As Object, e As MouseEventArgs)
        Dim hit = dgvArtikelVerwaltung.HitTest(e.X, e.Y)
        If hit.RowIndex < 0 Then Return

        If e.Button = MouseButtons.Right Then
            dgvArtikelVerwaltung.ClearSelection()
            dgvArtikelVerwaltung.Rows(hit.RowIndex).Selected = True
            dgvArtikelVerwaltung.CurrentCell = dgvArtikelVerwaltung.Rows(hit.RowIndex).Cells(If(hit.ColumnIndex >= 0, hit.ColumnIndex, 0))
        ElseIf e.Button = MouseButtons.Left Then
            dragQuelleIndex = hit.RowIndex
        End If
    End Sub

    ' Zeigt während des Ziehens eine Einfüge-Linie an der Zielposition an (wie man's aus dem
    ' Windows Explorer beim Verschieben von Dateien kennt). Zieht man nach unten, landet der
    ' Artikel HINTER der Zielzeile (Linie unten), zieht man nach oben, landet er DAVOR (Linie
    ' oben) - das entspricht genau dem tatsächlichen Verhalten von VerschiebeArtikel.
    Private Sub DgvArtikel_MouseMove(sender As Object, e As MouseEventArgs)
        If dragQuelleIndex < 0 Then Return
        Dim neuesZiel = dgvArtikelVerwaltung.HitTest(e.X, e.Y).RowIndex
        If neuesZiel <> dragZielIndex Then
            dragZielIndex = neuesZiel
            dgvArtikelVerwaltung.Invalidate()
        End If
    End Sub

    Private Sub DgvArtikel_Paint(sender As Object, e As PaintEventArgs)
        If dragQuelleIndex < 0 OrElse dragZielIndex < 0 OrElse dragZielIndex = dragQuelleIndex Then Return
        If dragZielIndex >= dgvArtikelVerwaltung.Rows.Count Then Return
        Dim rect = dgvArtikelVerwaltung.GetRowDisplayRectangle(dragZielIndex, True)
        If rect.IsEmpty Then Return
        Dim y As Integer = If(dragZielIndex > dragQuelleIndex, rect.Bottom - 1, rect.Top)
        Using pen As New Pen(Color.FromArgb(61, 90, 128), 3) ' Kornblumenblau, passend zum Programmdesign
            e.Graphics.DrawLine(pen, rect.Left, y, rect.Right, y)
        End Using
    End Sub

    Private Sub DgvArtikel_MouseUp(sender As Object, e As MouseEventArgs)
        If dragQuelleIndex < 0 Then Return
        Dim zielIndex = dgvArtikelVerwaltung.HitTest(e.X, e.Y).RowIndex
        If zielIndex >= 0 AndAlso zielIndex <> dragQuelleIndex Then
            VerschiebeArtikel(dragQuelleIndex, zielIndex)
        End If
        dragQuelleIndex = -1
        dragZielIndex = -1
        dgvArtikelVerwaltung.Invalidate()
    End Sub

    Private Sub ExportiereExcel(sender As Object, e As EventArgs)
        Dim sfd As New SaveFileDialog() With {.Filter = "Excel Dateien|*.xlsx", .FileName = "LEG_Artikelstamm.xlsx"}
        If sfd.ShowDialog() = DialogResult.OK Then
            Using wb As New XLWorkbook()
                Dim ws = wb.Worksheets.Add("Artikel")
                ws.Cell(1, 1).Value = "Nr"
                ws.Cell(1, 2).Value = "Beschreibung"
                ws.Cell(1, 3).Value = "MwSt (%)"
                ws.Cell(1, 4).Value = "Einzelpreis (Brutto, €)"
                ws.Range("A1:D1").Style.Font.Bold = True

                Dim rowIdx As Integer = 2

                For Each dgvRow As DataGridViewRow In dgvArtikelVerwaltung.Rows
                    If dgvRow.IsNewRow Then Continue For

                    ws.Cell(rowIdx, 1).Value = "'" & Convert.ToString(dgvRow.Cells(0).Value)
                    ws.Cell(rowIdx, 2).Value = Convert.ToString(dgvRow.Cells(1).Value)

                    Dim mwst As Double = 0
                    Double.TryParse(Convert.ToString(dgvRow.Cells(2).Value), mwst)
                    ws.Cell(rowIdx, 3).Value = mwst

                    Dim preis As Double = 0
                    Double.TryParse(Convert.ToString(dgvRow.Cells(3).Value), preis)
                    ws.Cell(rowIdx, 4).Value = preis

                    rowIdx += 1
                Next

                ws.Columns().AdjustToContents()
                ws.Column(2).Style.Alignment.WrapText = True
                ws.Column(2).Width = 50

                wb.SaveAs(sfd.FileName)
            End Using
            MessageBox.Show("Export erfolgreich abgeschlossen!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ImportiereExcel(sender As Object, e As EventArgs)
        Dim ofd As New OpenFileDialog() With {.Filter = "Excel Dateien|*.xlsx"}
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim count As Integer = 0
            Using wb As New XLWorkbook(ofd.FileName)
                Dim ws = wb.Worksheet(1)
                Using conn = DatenbankManager.HoleVerbindung()
                    ' Transaktion: schlägt der Import fehl, bleibt der alte Artikelstamm erhalten
                    Using trans = conn.BeginTransaction()
                        Try
                            Dim cmdDel = New SQLiteCommand("DELETE FROM artikel", conn, trans)
                            cmdDel.ExecuteNonQuery()

                            For r As Integer = 2 To ws.LastRowUsed().RowNumber()
                                Dim nr = ws.Cell(r, 1).GetString().Trim()
                                Dim bez = ws.Cell(r, 2).GetString().Trim()
                                Dim mwstStr = ws.Cell(r, 3).GetString().Trim()
                                Dim prsStr = ws.Cell(r, 4).GetString().Trim()

                                If String.IsNullOrEmpty(nr) Or String.IsNullOrEmpty(bez) Then Continue For

                                Dim mwst As Decimal = ParseMwSt(mwstStr)
                                If Math.Abs(mwst - mwstSatz1) < 0.01D OrElse Math.Abs(mwst - mwstSatz2) < 0.01D Then
                                    ' Spalte 4 enthält seit dem Netto/Brutto-Umbau den Bruttopreis
                                    ' (so wird er auch exportiert) - Netto wird daraus berechnet.
                                    Dim preisBrutto As Decimal = 0
                                    ParseBetrag(prsStr, preisBrutto)
                                    Dim preisNetto As Decimal = preisBrutto / (1 + mwst / 100D)

                                    Dim cmdIns = New SQLiteCommand("INSERT INTO artikel (artikelnummer, bezeichnung, mwst_satz, einzelpreis_netto, einzelpreis_brutto, einheit) VALUES (@nr, @bez, @mwst, @prs, @prsBrutto, 'C62')", conn, trans)
                                    cmdIns.Parameters.AddWithValue("@nr", nr.PadLeft(3, "0"c))
                                    cmdIns.Parameters.AddWithValue("@bez", bez)
                                    cmdIns.Parameters.AddWithValue("@mwst", mwst)
                                    cmdIns.Parameters.AddWithValue("@prs", preisNetto)
                                    cmdIns.Parameters.AddWithValue("@prsBrutto", preisBrutto)
                                    cmdIns.ExecuteNonQuery()
                                    count += 1
                                End If
                            Next

                            trans.Commit()
                        Catch ex As Exception
                            trans.Rollback()
                            MessageBox.Show("Fehler beim Excel-Import - der bestehende Artikelstamm wurde NICHT verändert:" & vbCrLf & vbCrLf & ex.Message,
                                            "Import fehlgeschlagen", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return
                        End Try
                    End Using
                End Using
            End Using
            LadeArtikelTabelle()
            LadeArtikelListeLinks()
            MessageBox.Show($"{count} Artikel erfolgreich importiert.", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    ' =========================================================================
    ' TAB 4: KUNDENVERWALTUNG LOGIK
    ' =========================================================================
    Private Sub DgvMitglieder_SelectionChanged(sender As Object, e As EventArgs)
        If dgvMitglieder.CurrentRow IsNot Nothing AndAlso Not dgvMitglieder.CurrentRow.IsNewRow Then
            aktuelleMitgliedId = CInt(dgvMitglieder.CurrentRow.Cells("id").Value)

            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmd As New SQLiteCommand("SELECT * FROM mitglieder WHERE id = @id", conn)
                cmd.Parameters.AddWithValue("@id", aktuelleMitgliedId)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        txtM_Nr.Text = reader("mitgliedsnummer").ToString()
                        txtM_Name.Text = reader("name").ToString()
                        txtM_Strasse.Text = reader("strasse").ToString()
                        txtM_PLZ.Text = reader("plz").ToString()
                        txtM_Ort.Text = reader("ort").ToString()
                        txtM_Land.Text = reader("land_code").ToString()
                        txtM_Steuer.Text = reader("steuernummer").ToString()
                        txtM_Betrieb.Text = reader("betriebsnummer").ToString()
                        txtM_Email.Text = reader("email").ToString()
                        cbM_Versand.Text = reader("versandart").ToString()
                        Dim standardPreisart As String = reader("standard_preisart").ToString()
                        cbM_Preisart.Text = If(String.IsNullOrWhiteSpace(standardPreisart), "Netto", standardPreisart)
                    End If
                End Using
            End Using
        End If
    End Sub

    Private Sub BtnNeu_Mitglied_Click(sender As Object, e As EventArgs)
        dgvMitglieder.CurrentCell = Nothing
        aktuelleMitgliedId = 0
        txtM_Nr.Clear() : txtM_Name.Clear() : txtM_Strasse.Clear() : txtM_PLZ.Clear()
        txtM_Ort.Clear() : txtM_Land.Text = "DE" : txtM_Steuer.Clear() : txtM_Betrieb.Clear()
        txtM_Email.Clear() : cbM_Versand.SelectedIndex = -1
        cbM_Preisart.SelectedIndex = 0
        txtM_Nr.Focus()
    End Sub

    Private Sub BtnSpeichern_Mitglied_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtM_Nr.Text) OrElse String.IsNullOrWhiteSpace(txtM_Name.Text) Then
            MessageBox.Show("Kundennummer und Name sind Pflichtfelder!", "Fehlende Daten", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmd As New SQLiteCommand(conn)

                If aktuelleMitgliedId = 0 Then
                    cmd.CommandText = "INSERT INTO mitglieder (mitgliedsnummer, name, strasse, plz, ort, land_code, steuernummer, betriebsnummer, email, versandart, standard_preisart) VALUES (@nr, @nam, @str, @plz, @ort, @lan, @steu, @bet, @eml, @ver, @prsart)"
                Else
                    cmd.CommandText = "UPDATE mitglieder SET mitgliedsnummer=@nr, name=@nam, strasse=@str, plz=@plz, ort=@ort, land_code=@lan, steuernummer=@steu, betriebsnummer=@bet, email=@eml, versandart=@ver, standard_preisart=@prsart WHERE id=@id"
                    cmd.Parameters.AddWithValue("@id", aktuelleMitgliedId)
                End If

                cmd.Parameters.AddWithValue("@nr", txtM_Nr.Text.Trim())
                cmd.Parameters.AddWithValue("@nam", txtM_Name.Text.Trim())
                cmd.Parameters.AddWithValue("@str", txtM_Strasse.Text.Trim())
                cmd.Parameters.AddWithValue("@plz", txtM_PLZ.Text.Trim())
                cmd.Parameters.AddWithValue("@ort", txtM_Ort.Text.Trim())
                cmd.Parameters.AddWithValue("@lan", txtM_Land.Text.Trim())
                cmd.Parameters.AddWithValue("@steu", txtM_Steuer.Text.Trim())
                cmd.Parameters.AddWithValue("@bet", txtM_Betrieb.Text.Trim())
                cmd.Parameters.AddWithValue("@eml", txtM_Email.Text.Trim())
                cmd.Parameters.AddWithValue("@ver", cbM_Versand.Text)
                cmd.Parameters.AddWithValue("@prsart", If(cbM_Preisart.SelectedIndex = 1, "Brutto", "Netto"))
                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Kunde erfolgreich gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LadeMitgliederListe()
            LadeDaten()
        Catch ex As SQLiteException When ex.ErrorCode = SQLiteErrorCode.Constraint
            MessageBox.Show("Diese Kundennummer existiert bereits! Bitte wähle eine andere.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Fehler beim Speichern: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnLoeschen_Mitglied_Click(sender As Object, e As EventArgs)
        If aktuelleMitgliedId = 0 Then Return

        Dim result = MessageBox.Show("Soll dieser Kunde wirklich gelöscht werden?", "Löschen bestätigen", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmd As New SQLiteCommand("DELETE FROM mitglieder WHERE id = @id", conn)
                cmd.Parameters.AddWithValue("@id", aktuelleMitgliedId)
                cmd.ExecuteNonQuery()
            End Using
            MessageBox.Show("Kunde gelöscht.", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
            BtnNeu_Mitglied_Click(Nothing, Nothing)
            LadeMitgliederListe()
            LadeDaten()
        End If
    End Sub

    Private Sub ExportiereMitgliederExcel(sender As Object, e As EventArgs)
        Dim sfd As New SaveFileDialog() With {.Filter = "Excel Dateien|*.xlsx", .FileName = "MHRechnung_Kundenstamm.xlsx"}
        If sfd.ShowDialog() = DialogResult.OK Then
            Using wb As New XLWorkbook()
                Dim ws = wb.Worksheets.Add("Kunden")
                Dim headers = {"Kunden-Nr.", "Name", "Straße", "PLZ", "Ort", "Land", "Betriebsnummer", "Email", "Steuernummer"}
                For c As Integer = 0 To headers.Length - 1
                    ws.Cell(1, c + 1).Value = headers(c)
                Next
                ws.Range(1, 1, 1, headers.Length).Style.Font.Bold = True

                Dim dt As New DataTable()
                Using conn = DatenbankManager.HoleVerbindung()
                    Dim da As New SQLiteDataAdapter("SELECT mitgliedsnummer, name, strasse, plz, ort, land_code, betriebsnummer, email, steuernummer FROM mitglieder", conn)
                    da.Fill(dt)
                End Using

                Dim höchsteNr As Integer = 0
                For Each row As DataRow In dt.Rows
                    Dim nrStr As String = row("mitgliedsnummer").ToString()
                    If nrStr.StartsWith("M-") Then
                        Dim zahl As Integer = 0
                        If Integer.TryParse(nrStr.Substring(2), zahl) Then
                            If zahl > höchsteNr Then höchsteNr = zahl
                        End If
                    End If
                Next

                Dim maxZeilen As Integer = Math.Max(50, höchsteNr)
                Dim rowIdx As Integer = 2

                For i As Integer = 1 To maxZeilen
                    Dim suchNr As String = "M-" & i.ToString("D3")
                    Dim gefundeneRows = dt.Select($"mitgliedsnummer = '{suchNr}'")

                    If gefundeneRows.Length > 0 Then
                        Dim row = gefundeneRows(0)
                        For c As Integer = 0 To headers.Length - 1
                            ws.Cell(rowIdx, c + 1).Value = "'" & row(c).ToString()
                        Next
                    Else
                        ws.Cell(rowIdx, 1).Value = suchNr
                        ws.Cell(rowIdx, 6).Value = "DE"
                    End If
                    rowIdx += 1
                Next

                For Each row As DataRow In dt.Rows
                    Dim nrStr As String = row("mitgliedsnummer").ToString()
                    If Not nrStr.StartsWith("M-") Then
                        For c As Integer = 0 To headers.Length - 1
                            ws.Cell(rowIdx, c + 1).Value = "'" & row(c).ToString()
                        Next
                        rowIdx += 1
                    End If
                Next

                ws.Columns().AdjustToContents()
                wb.SaveAs(sfd.FileName)
            End Using
            MessageBox.Show("Export erfolgreich! Vorlage wurde erstellt.", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ImportiereMitgliederExcel(sender As Object, e As EventArgs)
        Dim ofd As New OpenFileDialog() With {.Filter = "Excel Dateien|*.xlsx"}
        If ofd.ShowDialog() = DialogResult.OK Then
            Try
                Using wb As New XLWorkbook(ofd.FileName)
                    Dim ws = wb.Worksheet(1)
                    Dim count As Integer = 0

                    Using conn = DatenbankManager.HoleVerbindung()
                        For r As Integer = 2 To ws.LastRowUsed().RowNumber()
                            Dim nr = ws.Cell(r, 1).GetString().Trim()
                            Dim name = ws.Cell(r, 2).GetString().Trim()
                            If String.IsNullOrEmpty(nr) Or String.IsNullOrEmpty(name) Then Continue For

                            Dim strasse = ws.Cell(r, 3).GetString().Trim()
                            Dim plz = ws.Cell(r, 4).GetString().Trim()
                            Dim ort = ws.Cell(r, 5).GetString().Trim()
                            Dim land = ws.Cell(r, 6).GetString().Trim()
                            Dim betrieb = ws.Cell(r, 7).GetString().Trim()
                            Dim email = ws.Cell(r, 8).GetString().Trim()
                            Dim steuer = ws.Cell(r, 9).GetString().Trim()

                            Dim versandart As String = If(Not String.IsNullOrEmpty(email), "E-Mail", "Post")

                            Dim sql = "INSERT INTO mitglieder (mitgliedsnummer, name, strasse, plz, ort, land_code, betriebsnummer, email, steuernummer, versandart) " &
                                      "VALUES (@nr, @nam, @str, @plz, @ort, @lan, @bet, @eml, @steu, @ver) " &
                                      "ON CONFLICT(mitgliedsnummer) DO UPDATE SET " &
                                      "name=@nam, strasse=@str, plz=@plz, ort=@ort, land_code=@lan, betriebsnummer=@bet, email=@eml, steuernummer=@steu, versandart=@ver;"

                            Dim cmd = New SQLiteCommand(sql, conn)
                            cmd.Parameters.AddWithValue("@nr", nr)
                            cmd.Parameters.AddWithValue("@nam", name)
                            cmd.Parameters.AddWithValue("@str", strasse)
                            cmd.Parameters.AddWithValue("@plz", plz)
                            cmd.Parameters.AddWithValue("@ort", ort)
                            cmd.Parameters.AddWithValue("@lan", If(String.IsNullOrEmpty(land), "DE", land))
                            cmd.Parameters.AddWithValue("@bet", betrieb)
                            cmd.Parameters.AddWithValue("@eml", email)
                            cmd.Parameters.AddWithValue("@steu", steuer)
                            cmd.Parameters.AddWithValue("@ver", versandart)
                            cmd.ExecuteNonQuery()
                            count += 1
                        Next
                    End Using

                    LadeMitgliederListe()
                    LadeDaten()
                    MessageBox.Show($"{count} Kunden erfolgreich importiert/aktualisiert.", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            Catch ex As Exception
                MessageBox.Show("Fehler beim Import: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    ' =========================================================================
    ' DATEN-TÜV
    ' =========================================================================
    Private Sub BtnMitgliederPruefen_Click(sender As Object, e As EventArgs)
        Dim fehlerListe As New List(Of String)
        Dim gepruefteMitglieder As Integer = 0
        Dim fehlerhafteIds As New HashSet(Of Integer)()

        Using conn = DatenbankManager.HoleVerbindung()
            Dim cmd As New SQLiteCommand("SELECT id, mitgliedsnummer, name, plz, email, versandart, steuernummer FROM mitglieder", conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    gepruefteMitglieder += 1
                    Dim mId = CInt(reader("id"))
                    Dim nr = reader("mitgliedsnummer").ToString()
                    Dim name = reader("name").ToString()
                    Dim versand = reader("versandart").ToString()
                    Dim idString = $"{nr} ({name})"
                    Dim hatFehler As Boolean = False

                    Dim plz = reader("plz").ToString().Trim()
                    If plz.Length <> 5 OrElse Not IsNumeric(plz) Then
                        fehlerListe.Add($"{idString}: PLZ muss exakt 5-stellig sein (ist '{plz}').")
                        hatFehler = True
                    End If

                    If versand <> "Post" Then
                        Dim email = reader("email").ToString().Trim()
                        Dim regexEmail As New System.Text.RegularExpressions.Regex("^[^@\s]+@[^@\s]+\.[^@\s]+$")
                        If Not regexEmail.IsMatch(email) Then
                            fehlerListe.Add($"{idString}: E-Mail Adresse ist ungültig.")
                            hatFehler = True
                        End If

                        Dim steuer = reader("steuernummer").ToString().Trim()
                        Dim slashCount = steuer.Length - steuer.Replace("/", "").Length
                        If slashCount <> 2 OrElse steuer.Length < 10 Then
                            fehlerListe.Add($"{idString}: Steuernummer ungültig.")
                            hatFehler = True
                        Else
                            Dim nurZahlen = steuer.Replace("/", "")
                            If Not IsNumeric(nurZahlen) Then
                                fehlerListe.Add($"{idString}: Steuernummer darf außer '/' nur Zahlen enthalten.")
                                hatFehler = True
                            End If
                        End If
                    End If

                    If hatFehler Then fehlerhafteIds.Add(mId)
                End While
            End Using
        End Using

        For Each row As DataGridViewRow In dgvMitglieder.Rows
            If Not row.IsNewRow Then
                Dim rowId As Integer = CInt(row.Cells("id").Value)
                If fehlerhafteIds.Contains(rowId) Then
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235)
                Else
                    row.DefaultCellStyle.BackColor = Color.FromArgb(235, 255, 235)
                End If
            End If
        Next
        dgvMitglieder.ClearSelection()

        If fehlerListe.Count = 0 Then
            MessageBox.Show($"TÜV bestanden! {gepruefteMitglieder} Kunden wurden geprüft — alle Daten sind korrekt.", "Prüfung erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Dim anzeige As String = String.Join(vbCrLf & vbCrLf, fehlerListe.Take(15))
            If fehlerListe.Count > 15 Then anzeige &= vbCrLf & vbCrLf & $"... und {fehlerListe.Count - 15} weitere Fehler."
            MessageBox.Show($"Es wurden {fehlerListe.Count} Fehler gefunden:" & vbCrLf & vbCrLf & anzeige, "Fehlerhafte Daten", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Function IstIbanKorrekt(iban As String) As Boolean
        Try
            Dim rearranged As String = iban.Substring(4) & iban.Substring(0, 4)
            Dim numericIban As String = ""
            For Each c As Char In rearranged
                If Char.IsLetter(c) Then
                    numericIban &= (Asc(c) - 55).ToString()
                Else
                    numericIban &= c
                End If
            Next
            Dim remainder As Integer = 0
            For Each c As Char In numericIban
                Dim digit As Integer = Integer.Parse(c.ToString())
                remainder = (remainder * 10 + digit) Mod 97
            Next
            Return remainder = 1
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' =========================================================================
    ' TAB 5: EINSTELLUNGEN LOGIK
    ' =========================================================================
    Private Sub LadeEinstellungen()
        Using conn = DatenbankManager.HoleVerbindung()
            Dim cmd As New SQLiteCommand("SELECT schluessel, wert FROM einstellungen", conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim key As String = reader("schluessel").ToString()
                    Dim val As String = reader("wert").ToString()
                    Select Case key
                        Case "text_zahlung" : txtE_TextZahlung.Text = val
                        Case "text_gutschrift" : txtE_TextGutschrift.Text = val
                        Case "text_email" : txtE_TextEmail.Text = val
                        Case "prog_mengen_nullen" : chkE_MengenNullen.Checked = (val = "True")
                        Case "prog_auto_backup" : chkE_AutoBackup.Checked = (val = "True")
                        Case "prog_ausgangskopie" : chkE_Ausgangskopie.Checked = (val = "True")
                        Case "prog_ausgangskopie_email" : txtE_AusgangskopieEmail.Text = val
                        Case "prog_druck_anzahl" : cbE_DruckAnzahl.Text = val
                        Case "prog_standard_drucker" : cbE_StandardDrucker.Text = val
                        Case "firma_name" : txtE_FirmaName.Text = val
                        Case "firma_strasse" : txtE_FirmaStrasse.Text = val
                        Case "firma_plz" : txtE_FirmaPLZ.Text = val
                        Case "firma_ort" : txtE_FirmaOrt.Text = val
                        Case "firma_email" : txtE_FirmaMail.Text = val
                        Case "firma_tel" : txtE_FirmaTel.Text = val
                        Case "firma_iban" : txtE_IBAN.Text = val
                        Case "firma_bic" : txtE_BIC.Text = val
                        Case "firma_bank" : txtE_Bank.Text = val
                        Case "firma_steuer" : txtE_Steuer.Text = val
                        Case "mwst_satz_1" : txtE_MwSt1.Text = val.Replace(".", ",")
                        Case "mwst_satz_2" : txtE_MwSt2.Text = val.Replace(".", ",")
                        Case "smtp_server" : txtE_SmtpServer.Text = val
                        Case "smtp_port" : txtE_SmtpPort.Text = val
                        Case "smtp_user" : txtE_SmtpUser.Text = val
                        Case "smtp_pass" : txtE_SmtpPass.Text = val
                        Case "speicherpfad" : txtE_Speicherpfad.Text = val
                    End Select
                End While
            End Using
        End Using
    End Sub

    Private Sub SpeichereEinstellungDB(conn As SQLiteConnection, key As String, value As String)
        Dim cmdCheck As New SQLiteCommand("SELECT COUNT(*) FROM einstellungen WHERE schluessel = @k", conn)
        cmdCheck.Parameters.AddWithValue("@k", key)
        Dim count As Integer = CInt(cmdCheck.ExecuteScalar())

        Dim cmd As New SQLiteCommand(conn)
        If count > 0 Then
            cmd.CommandText = "UPDATE einstellungen SET wert = @v WHERE schluessel = @k"
        Else
            cmd.CommandText = "INSERT INTO einstellungen (schluessel, wert) VALUES (@k, @v)"
        End If
        cmd.Parameters.AddWithValue("@k", key)
        cmd.Parameters.AddWithValue("@v", value)
        cmd.ExecuteNonQuery()
    End Sub

    Private Sub BtnSpeicherpfadAendern_Click(sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog With {
            .Description = "Haupt-Speicherpfad für Rechnungen (PDF/XML/Excel) auswählen",
            .ShowNewFolderButton = True
        }
            If Directory.Exists(txtE_Speicherpfad.Text) Then fbd.SelectedPath = txtE_Speicherpfad.Text
            If fbd.ShowDialog() = DialogResult.OK Then
                txtE_Speicherpfad.Text = fbd.SelectedPath
            End If
        End Using
    End Sub

    Private Sub BtnDbPfadAendern_Click(sender As Object, e As EventArgs)
        Dim neuerPfad As String = ZeigeDatenbankWahlDialog(erzwingeWahl:=False)
        If String.IsNullOrWhiteSpace(neuerPfad) Then Return

        DbKonfiguration.SchreibeDatenbankPfad(neuerPfad)
        txtE_DbPfad.Text = neuerPfad
        MessageBox.Show("Datenbank-Pfad geändert." & vbCrLf & vbCrLf &
                        "Bitte starte MHRechnung neu, damit die Änderung wirkt.",
                        "Neustart nötig", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub BtnSpeichern_Einstellungen_Click(sender As Object, e As EventArgs)
        Try
            Using conn = DatenbankManager.HoleVerbindung()
                SpeichereEinstellungDB(conn, "prog_mengen_nullen", chkE_MengenNullen.Checked.ToString())
                SpeichereEinstellungDB(conn, "prog_auto_backup", chkE_AutoBackup.Checked.ToString())
                SpeichereEinstellungDB(conn, "prog_ausgangskopie", chkE_Ausgangskopie.Checked.ToString())
                SpeichereEinstellungDB(conn, "prog_ausgangskopie_email", txtE_AusgangskopieEmail.Text.Trim())
                SpeichereEinstellungDB(conn, "prog_druck_anzahl", cbE_DruckAnzahl.Text)
                SpeichereEinstellungDB(conn, "prog_standard_drucker", cbE_StandardDrucker.Text)
                SpeichereEinstellungDB(conn, "firma_name", txtE_FirmaName.Text.Trim())
                SpeichereEinstellungDB(conn, "firma_strasse", txtE_FirmaStrasse.Text.Trim())
                SpeichereEinstellungDB(conn, "firma_plz", txtE_FirmaPLZ.Text.Trim())
                SpeichereEinstellungDB(conn, "firma_ort", txtE_FirmaOrt.Text.Trim())
                SpeichereEinstellungDB(conn, "firma_email", txtE_FirmaMail.Text.Trim())
                SpeichereEinstellungDB(conn, "firma_tel", txtE_FirmaTel.Text.Trim())
                SpeichereEinstellungDB(conn, "firma_iban", txtE_IBAN.Text.Trim().Replace(" ", ""))
                SpeichereEinstellungDB(conn, "firma_bic", txtE_BIC.Text.Trim())
                SpeichereEinstellungDB(conn, "firma_bank", txtE_Bank.Text.Trim())
                SpeichereEinstellungDB(conn, "firma_steuer", txtE_Steuer.Text.Trim())

                Dim neuerSatz1 As Decimal = ParseMwSt(txtE_MwSt1.Text)
                Dim neuerSatz2 As Decimal = ParseMwSt(txtE_MwSt2.Text)
                If neuerSatz1 > 0 Then SpeichereEinstellungDB(conn, "mwst_satz_1", neuerSatz1.ToString(Globalization.CultureInfo.InvariantCulture))
                If neuerSatz2 > 0 Then SpeichereEinstellungDB(conn, "mwst_satz_2", neuerSatz2.ToString(Globalization.CultureInfo.InvariantCulture))
                SpeichereEinstellungDB(conn, "text_zahlung", txtE_TextZahlung.Text)
                SpeichereEinstellungDB(conn, "text_gutschrift", txtE_TextGutschrift.Text)
                SpeichereEinstellungDB(conn, "text_email", txtE_TextEmail.Text)
                SpeichereEinstellungDB(conn, "smtp_server", txtE_SmtpServer.Text.Trim())
                SpeichereEinstellungDB(conn, "smtp_port", txtE_SmtpPort.Text.Trim())
                SpeichereEinstellungDB(conn, "smtp_user", txtE_SmtpUser.Text.Trim())
                SpeichereEinstellungDB(conn, "smtp_pass", txtE_SmtpPass.Text)
                SpeichereEinstellungDB(conn, "speicherpfad", txtE_Speicherpfad.Text.Trim())
            End Using

            Dim hinweis As String = ""
            Dim ibanTrim As String = txtE_IBAN.Text.Trim().Replace(" ", "").ToUpper()
            If Not String.IsNullOrWhiteSpace(ibanTrim) AndAlso (ibanTrim.Length <> 22 OrElse Not ibanTrim.StartsWith("DE") OrElse Not IstIbanKorrekt(ibanTrim)) Then
                hinweis &= vbCrLf & vbCrLf & "⚠ Deine eigene IBAN sieht ungültig aus - bitte prüfen, sonst können Kunden nicht korrekt überweisen."
            End If
            If ParseMwSt(txtE_MwSt1.Text) <> mwstSatz1 OrElse ParseMwSt(txtE_MwSt2.Text) <> mwstSatz2 Then
                mwstSatz1 = If(ParseMwSt(txtE_MwSt1.Text) > 0, ParseMwSt(txtE_MwSt1.Text), mwstSatz1)
                mwstSatz2 = If(ParseMwSt(txtE_MwSt2.Text) > 0, ParseMwSt(txtE_MwSt2.Text), mwstSatz2)
                hinweis &= vbCrLf & vbCrLf & "ℹ MwSt-Sätze geändert: Bitte starte das Programm neu, damit alle Dropdowns die neuen Sätze anzeigen."
            End If

            MessageBox.Show("Einstellungen erfolgreich gespeichert!" & hinweis, "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Fehler beim Speichern der Einstellungen: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =========================================================================
    ' ARCHIV LOGIK
    ' =========================================================================
    Private Sub LadeArchiv()
        Using conn = DatenbankManager.HoleVerbindung()
            Dim s1 As String = mwstSatz1.ToString(Globalization.CultureInfo.InvariantCulture)
            Dim s2 As String = mwstSatz2.ToString(Globalization.CultureInfo.InvariantCulture)
            Dim f1 As String = (mwstSatz1 / 100D).ToString(Globalization.CultureInfo.InvariantCulture)
            Dim f2 As String = (mwstSatz2 / 100D).ToString(Globalization.CultureInfo.InvariantCulture)
            Dim sql = "SELECT r.id, r.rechnungsnummer AS 'Re-Nr', r.datum AS 'Datum', m.name AS 'Empfänger', " &
                      "(SELECT ROUND(SUM(ROUND(anzahl*einzelpreis,2)) " &
                      " + ROUND(SUM(CASE WHEN mwst_satz=" & s1 & " THEN ROUND(anzahl*einzelpreis,2) ELSE 0 END)*" & f1 & ",2) " &
                      " + ROUND(SUM(CASE WHEN mwst_satz=" & s2 & " THEN ROUND(anzahl*einzelpreis,2) ELSE 0 END)*" & f2 & ",2),2) " &
                      " FROM rechnungspositionen WHERE rechnung_id = r.id) AS 'Brutto', " &
                      "CASE WHEN m.versandart = 'E-Mail' OR m.versandart = 'Beides' THEN '✓' ELSE '-' END AS 'E-Mail' " &
                      "FROM rechnungen r JOIN mitglieder m ON r.mitglied_id = m.id " &
                      "WHERE r.status = 'Verarbeitet & Exportiert' ORDER BY CAST(r.rechnungsnummer AS INTEGER) DESC"

            Dim da As New SQLiteDataAdapter(sql, conn)
            Dim dt As New DataTable()
            da.Fill(dt)

            dgvArchiv.DataSource = dt
            If dgvArchiv.Columns.Contains("id") Then dgvArchiv.Columns("id").Visible = False
            If dgvArchiv.Columns.Contains("Brutto") Then
                dgvArchiv.Columns("Brutto").DefaultCellStyle.Format = "N2"
                dgvArchiv.Columns("Brutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
            If dgvArchiv.Columns.Contains("E-Mail") Then
                dgvArchiv.Columns("E-Mail").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvArchiv.Columns("E-Mail").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                dgvArchiv.Columns("E-Mail").Width = 70
            End If
        End Using

        If dgvArchiv.Rows.Count > 0 Then
            DgvArchiv_SelectionChanged(Nothing, Nothing)
        Else
            pnlArchivDetails.Controls.Clear()
            lblArchivSumme.Text = $"Netto: 0,00 €" & vbCrLf & $"MwSt {FmtMwSt(mwstSatz1)}: 0,00 €" & vbCrLf & $"MwSt {FmtMwSt(mwstSatz2)}: 0,00 €" & vbCrLf & vbCrLf & "RECHNUNGSBETRAG: 0,00 €"
        End If
    End Sub

    Private Sub DgvArchiv_SelectionChanged(sender As Object, e As EventArgs)
        If dgvArchiv.CurrentRow IsNot Nothing AndAlso Not dgvArchiv.CurrentRow.IsNewRow Then
            Dim reID As Integer = 0
            If Integer.TryParse(dgvArchiv.CurrentRow.Cells("id").Value.ToString(), reID) Then
                LadeDetailsAllgemein(reID, pnlArchivDetails, lblArchivSumme)
            End If
        Else
            pnlArchivDetails.Controls.Clear()
        End If
    End Sub

    Private Function HoleArchivPdfPfad() As String
        If dgvArchiv.CurrentRow Is Nothing OrElse dgvArchiv.CurrentRow.IsNewRow Then Return ""
        Dim reNr = dgvArchiv.CurrentRow.Cells("Re-Nr").Value.ToString()
        Dim datumStr = dgvArchiv.CurrentRow.Cells("Datum").Value.ToString()

        Dim jahr As String = DateTime.Now.Year.ToString()
        If datumStr.Length >= 10 Then jahr = datumStr.Substring(6, 4)

        Dim baseDir As String = "C:\MHRechnung"
        Using conn = DatenbankManager.HoleVerbindung()
            Dim cmdPfad As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'speicherpfad'", conn)
            Dim pfadObj = cmdPfad.ExecuteScalar()
            If pfadObj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(pfadObj.ToString()) Then
                baseDir = pfadObj.ToString()
            End If
        End Using

        Return Path.Combine(baseDir, jahr, "erstellt", "e-rechnungen", reNr & "_zugf.pdf")
    End Function

    Private Sub BtnOeffnenArchiv_Click(sender As Object, e As EventArgs)
        Dim pfad = HoleArchivPdfPfad()
        If String.IsNullOrEmpty(pfad) Then Return
        If File.Exists(pfad) Then
            Process.Start(pfad)
        Else
            MessageBox.Show($"Die PDF-Datei wurde nicht gefunden unter:{vbCrLf}{pfad}", "Datei nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub BtnDruckenArchiv_Click(sender As Object, e As EventArgs)
        Dim pfad = HoleArchivPdfPfad()
        If String.IsNullOrEmpty(pfad) Then Return
        If File.Exists(pfad) Then
            Dim pd As New PrintDialog()
            If pd.ShowDialog() = DialogResult.OK Then DruckeDokument(pfad, pd.PrinterSettings.PrinterName)
        Else
            MessageBox.Show($"Die PDF-Datei wurde nicht gefunden unter:{vbCrLf}{pfad}", "Datei nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    ' =========================================================================
    ' UI AUFBAU METHODEN — MODERNISIERT
    ' =========================================================================

    Private Sub BaueKopfzeile()
        Dim pnlHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 58,
            .BackColor = CLR_HEADER_BG
        }

        ' Weizengold-Akzentlinie unten
        AddHandler pnlHeader.Paint, Sub(s As Object, ev As PaintEventArgs)
                                        ev.Graphics.FillRectangle(New SolidBrush(CLR_GOLD_AKZENT), 0, pnlHeader.Height - 3, pnlHeader.Width, 3)
                                    End Sub

        ' Logo-Text links ("MH"-Monogramm)
        Dim lblL As New Label With {.Text = "M", .Font = New Font("Segoe UI", 18, FontStyle.Bold), .ForeColor = CLR_GOLD_AKZENT, .AutoSize = True, .Location = New Point(18, 10)}
        Dim lblEG As New Label With {.Text = "H", .Font = New Font("Segoe UI", 18, FontStyle.Bold), .ForeColor = CLR_WEISS, .AutoSize = True, .Location = New Point(33, 10)}
        Dim lblSub As New Label With {.Text = "Rechnungs-Manager", .Font = New Font("Segoe UI", 9.5F), .ForeColor = Color.FromArgb(196, 209, 224), .AutoSize = True, .Location = New Point(18, 36)}

        ' Rechnungsnummer rechts
        lblAktuelleReNr.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        lblAktuelleReNr.ForeColor = CLR_GOLD_AKZENT
        lblAktuelleReNr.BackColor = Color.Transparent
        lblAktuelleReNr.AutoSize = False
        lblAktuelleReNr.Width = 160
        lblAktuelleReNr.Height = 30
        lblAktuelleReNr.TextAlign = ContentAlignment.MiddleRight
        lblAktuelleReNr.Location = New Point(1220, 14)
        lblAktuelleReNr.Anchor = AnchorStyles.Top Or AnchorStyles.Right

        Dim lblReNrLabel As New Label With {
            .Text = "NÄCHSTE NR.",
            .Font = New Font("Segoe UI", 7.5F),
            .ForeColor = Color.FromArgb(170, 190, 214),
            .AutoSize = False,
            .Width = 160,
            .Height = 16,
            .TextAlign = ContentAlignment.MiddleRight,
            .Location = New Point(1220, 38),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }

        pnlHeader.Controls.AddRange({lblL, lblEG, lblSub, lblAktuelleReNr, lblReNrLabel})
        Me.Controls.Add(pnlHeader)
    End Sub

    Private Sub BaueTabErfassung()
        TabErfassung.BackColor = CLR_HINTERGRUND

        Dim tlp As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 3}
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 340.0!))
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 340.0!))

        ' ── LINKE SPALTE: Artikelstamm ──────────────────────────────────────
        Dim pnlL As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(10, 10, 5, 10), .BackColor = CLR_HINTERGRUND}

        Dim pnlLHeader As New Panel With {.Dock = DockStyle.Top, .Height = 44, .BackColor = CLR_HINTERGRUND}
        Dim lblArtStamm As New Label With {
            .Text = "  ARTIKELSTAMM",
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .ForeColor = CLR_BLAU_DUNKEL,
            .Dock = DockStyle.Left,
            .Width = 200,
            .TextAlign = ContentAlignment.MiddleLeft
        }
        Dim btnAddArrow = MachePrimaerButton("→ Hinzufügen", 120, 32)
        btnAddArrow.Dock = DockStyle.Right
        AddHandler btnAddArrow.Click, AddressOf BtnAdd_Click
        pnlLHeader.Controls.AddRange({lblArtStamm, btnAddArrow})

        lstArtikel.Dock = DockStyle.Fill
        lstArtikel.BorderStyle = BorderStyle.None
        ' War FONT_KLEIN (8,5pt) - das war eigentlich gemeint, als von der zu kleinen Schrift
        ' im Artikelstamm die Rede war (die Artikelauswahl in der Rechnungserfassung, nicht die
        ' Artikelverwaltung).
        lstArtikel.Font = New Font("Segoe UI", 12)
        lstArtikel.BackColor = CLR_WEISS
        lstArtikel.ForeColor = CLR_TEXT_DUNKEL

        ' Eigenes Zeichnen (Owner-Draw) statt nur eines einzelnen Textstrings pro Zeile: Damit
        ' stehen Preis und MwSt-Satz in zwei festen, rechtsbündigen Spalten immer exakt
        ' untereinander - unabhängig davon, wie lang die jeweilige Artikelbezeichnung ist.
        ' Bei DrawMode.OwnerDrawFixed skaliert die Zeilenhöhe nicht mehr automatisch mit der
        ' Schrift mit, deshalb hier passend zur 12pt-Schrift von Hand gesetzt.
        lstArtikel.DrawMode = DrawMode.OwnerDrawFixed
        lstArtikel.ItemHeight = 28
        AddHandler lstArtikel.DrawItem, AddressOf LstArtikel_DrawItem

        Dim pnlListContainer As New Panel With {.Dock = DockStyle.Fill, .BackColor = CLR_WEISS, .Padding = New Padding(1)}
        AddHandler pnlListContainer.Paint, Sub(s As Object, ev As PaintEventArgs)
                                               ControlPaint.DrawBorder(ev.Graphics, DirectCast(s, Panel).ClientRectangle,
                                                   CLR_BORDER, ButtonBorderStyle.Solid)
                                           End Sub
        pnlListContainer.Controls.Add(lstArtikel)

        pnlL.Controls.Add(pnlListContainer)
        pnlL.Controls.Add(pnlLHeader)

        ' ── MITTLERE SPALTE: Rechnungs-Editor ───────────────────────────────
        Dim pnlM As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(5, 10, 5, 10), .Size = New Size(700, 800), .BackColor = CLR_HINTERGRUND}

        ' Toolbar oben
        Dim pnlMTop As New Panel With {.Dock = DockStyle.Top, .Height = 44, .BackColor = CLR_HINTERGRUND, .Padding = New Padding(0, 6, 10, 6)}
        Dim lblEditor As New Label With {
            .Text = "  RECHNUNGS-EDITOR",
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .ForeColor = CLR_BLAU_DUNKEL,
            .Dock = DockStyle.Left,
            .Width = 200,
            .TextAlign = ContentAlignment.MiddleLeft
        }

        btnRechnungErstellen = New Button With {
            .Text = "RECHNUNG ERSTELLEN",
            .Width = 190,
            .BackColor = Color.FromArgb(160, 160, 160),
            .ForeColor = CLR_WEISS,
            .Enabled = False,
            .FlatStyle = FlatStyle.Flat,
            .Font = FONT_TITLE,
            .Cursor = Cursors.Hand,
            .Dock = DockStyle.Right
        }
        btnRechnungErstellen.FlatAppearance.BorderSize = 0
        AddHandler btnRechnungErstellen.Click, AddressOf SpeichereRechnung

        pnlMTop.Controls.AddRange({lblEditor, btnRechnungErstellen})

        ' Spalten-Header
        Dim headerWidth As Integer = pnlM.Width - 20
        Dim pnlTblHeader As New Panel With {
            .Location = New Point(10, 50),
            .Size = New Size(headerWidth, 30),
            .BackColor = CLR_BLAU_DUNKEL,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        Dim lblHMenge As New Label With {.Text = "Menge", .Location = New Point(8, 7), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = CLR_WEISS}
        Dim lblHBez As New Label With {.Text = "Beschreibung", .Location = New Point(62, 7), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = CLR_WEISS}

        ' Zwei eigene Preis-Spalten statt einer umschaltbaren - Positionen spiegeln exakt die
        ' Felder in ErstelleArtikelZeile (rahmenWidth dort ≈ headerWidth hier).
        Dim lblHNetto As New Label With {.Text = "Netto", .Location = New Point(headerWidth - 422, 7), .AutoSize = False, .Width = 62, .TextAlign = ContentAlignment.TopRight, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = CLR_WEISS, .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        Dim lblHBrutto As New Label With {.Text = "Brutto", .Location = New Point(headerWidth - 354, 7), .AutoSize = False, .Width = 62, .TextAlign = ContentAlignment.TopRight, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = CLR_WEISS, .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        Dim lblHMwSt As New Label With {.Text = "MwSt", .Location = New Point(headerWidth - 282, 7), .AutoSize = False, .Width = 60, .TextAlign = ContentAlignment.TopRight, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = CLR_WEISS, .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        Dim lblHGes As New Label With {.Text = "Gesamt (Brutto)", .Location = New Point(headerWidth - 212, 7), .AutoSize = False, .Width = 100, .TextAlign = ContentAlignment.TopRight, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = CLR_WEISS, .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        pnlTblHeader.Controls.AddRange({lblHMenge, lblHBez, lblHNetto, lblHBrutto, lblHMwSt, lblHGes})

        ' Zeilen-Container
        pnlRows.Location = New Point(10, 80)
        pnlRows.Size = New Size(headerWidth, 680)
        pnlRows.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        pnlRows.AutoScroll = True
        pnlRows.BackColor = CLR_HINTERGRUND
        pnlRows.FlowDirection = FlowDirection.TopDown
        pnlRows.WrapContents = False
        pnlRows.AllowDrop = True
        pnlRows.Padding = New Padding(0, 4, 0, 0)
        AddHandler pnlRows.DragEnter, AddressOf PnlRows_DragEnter
        AddHandler pnlRows.DragDrop, AddressOf PnlRows_DragDrop
        AddHandler pnlRows.Resize, AddressOf PnlRows_Resize

        pnlPlusContainer = New Panel With {.Width = headerWidth - 10, .Height = 54, .Name = "PlusContainer", .BackColor = CLR_HINTERGRUND}
        Dim btnPlus = MacheSekundaerButton("＋  Neue Zeile", 160, 40)
        btnPlus.Location = New Point((pnlPlusContainer.Width - 160) \ 2, 7)
        btnPlus.Anchor = AnchorStyles.Top
        AddHandler btnPlus.Click, AddressOf BtnNewLine_Click
        pnlPlusContainer.Controls.Add(btnPlus)
        pnlRows.Controls.Add(pnlPlusContainer)

        pnlM.Controls.AddRange({pnlMTop, pnlTblHeader, pnlRows})

        ' ── RECHTE SPALTE: Empfänger ─────────────────────────────────────────
        Dim pnlR As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(5, 10, 10, 10), .BackColor = CLR_HINTERGRUND}
        Dim lblEmpfHdr As New Label With {
            .Text = "  EMPFÄNGER",
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .ForeColor = CLR_BLAU_DUNKEL,
            .Dock = DockStyle.Top,
            .Height = 44,
            .TextAlign = ContentAlignment.MiddleLeft
        }

        chkMitglieder.Dock = DockStyle.Fill
        chkMitglieder.BorderStyle = BorderStyle.None
        chkMitglieder.Font = FONT_NORMAL
        chkMitglieder.BackColor = CLR_WEISS
        chkMitglieder.ForeColor = CLR_TEXT_DUNKEL
        chkMitglieder.CheckOnClick = True
        AddHandler chkMitglieder.ItemCheck, AddressOf ChkMitglieder_ItemCheck

        Dim pnlChkContainer As New Panel With {.Dock = DockStyle.Fill, .BackColor = CLR_WEISS, .Padding = New Padding(1)}
        AddHandler pnlChkContainer.Paint, Sub(s As Object, ev As PaintEventArgs)
                                              ControlPaint.DrawBorder(ev.Graphics, DirectCast(s, Panel).ClientRectangle, CLR_BORDER, ButtonBorderStyle.Solid)
                                          End Sub
        pnlChkContainer.Controls.Add(chkMitglieder)

        ' Lieferdatum - eigenes Feld (§14 Abs. 4 Nr. 6 UStG: Liefer-/Leistungsdatum
        ' ist Pflichtangabe, sofern es vom Rechnungsdatum abweicht). Gilt für alle
        ' Empfänger, die in diesem Durchgang angehakt sind.
        Dim pnlLieferdatum As New Panel With {.Dock = DockStyle.Top, .Height = 34, .BackColor = CLR_HINTERGRUND, .Padding = New Padding(0, 2, 0, 6)}
        Dim lblLieferdatum As New Label With {.Text = "Lieferdatum:", .Dock = DockStyle.Left, .Width = 90, .Font = FONT_KLEIN, .ForeColor = CLR_TEXT_GRAU, .TextAlign = ContentAlignment.MiddleLeft}
        txtLieferdatum.Dock = DockStyle.Left
        txtLieferdatum.Width = 100
        txtLieferdatum.Font = FONT_NORMAL
        txtLieferdatum.BorderStyle = BorderStyle.FixedSingle
        txtLieferdatum.Text = DateTime.Now.ToString("dd.MM.yyyy")
        pnlLieferdatum.Controls.AddRange({txtLieferdatum, lblLieferdatum})

        ' Rechnungsart hierher verschoben (war vorher oben über der Artikelliste) - das
        ' betrifft wie das Lieferdatum die ganze Rechnung, nicht einzelne Positionen, deshalb
        ' gehört es zu den anderen rechnungsweiten Angaben in dieser Spalte, nicht in die
        ' Werkzeugleiste direkt über den Artikelzeilen. Wird beim Ankreuzen eines Kunden aus
        ' dessen Stammdaten vorbelegt (siehe ChkMitglieder_ItemCheck), bleibt aber pro
        ' Rechnung überschreibbar.
        Dim pnlPreisart As New Panel With {.Dock = DockStyle.Top, .Height = 34, .BackColor = CLR_HINTERGRUND, .Padding = New Padding(0, 2, 0, 6)}
        Dim lblPreisart As New Label With {.Text = "Rechnungsart:", .Dock = DockStyle.Left, .Width = 90, .Font = FONT_KLEIN, .ForeColor = CLR_TEXT_GRAU, .TextAlign = ContentAlignment.MiddleLeft}
        cbPreisart.Dock = DockStyle.Left
        cbPreisart.Width = 200
        cbPreisart.DropDownStyle = ComboBoxStyle.DropDownList
        cbPreisart.FlatStyle = FlatStyle.Flat
        cbPreisart.BackColor = CLR_WEISS
        cbPreisart.Font = FONT_NORMAL
        cbPreisart.Items.AddRange({"Rechnung zeigt: NETTO", "Rechnung zeigt: BRUTTO"})
        cbPreisart.SelectedIndex = 0
        ' Die Voransicht unten (BerechneSummenTab1) hängt jetzt von der Rechnungsart ab
        ' (Brutto-Modus rechnet anders als Netto-Modus) - ohne diesen Handler würde ein
        ' Umschalten des Dropdowns allein die Voransicht nicht aktualisieren.
        AddHandler cbPreisart.SelectedIndexChanged, AddressOf BerechneSummenTab1
        pnlPreisart.Controls.AddRange({cbPreisart, lblPreisart})

        pnlR.Controls.Add(pnlChkContainer)
        pnlR.Controls.Add(pnlPreisart)
        pnlR.Controls.Add(pnlLieferdatum)
        pnlR.Controls.Add(lblEmpfHdr)

        ' ── FUSSZEILE: Summen ────────────────────────────────────────────────
        Dim pnlF As New Panel With {.Dock = DockStyle.Bottom, .Height = 50, .BackColor = CLR_BLAU_DUNKEL}
        lblSummenTab1.Dock = DockStyle.Fill
        lblSummenTab1.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        lblSummenTab1.ForeColor = CLR_WEISS
        lblSummenTab1.TextAlign = ContentAlignment.MiddleLeft
        pnlF.Controls.Add(lblSummenTab1)

        tlp.Controls.Add(pnlL, 0, 0)
        tlp.Controls.Add(pnlM, 1, 0)
        tlp.Controls.Add(pnlR, 2, 0)
        TabErfassung.Controls.AddRange({tlp, pnlF})
    End Sub

    Private Sub PnlRows_Resize(sender As Object, e As EventArgs)
        If pnlRows Is Nothing Then Return
        Dim newWidth As Integer = pnlRows.ClientSize.Width - 10
        If newWidth < 100 Then Return
        pnlRows.SuspendLayout()
        For Each ctrl As Control In pnlRows.Controls
            ctrl.Width = newWidth
        Next
        pnlRows.ResumeLayout()
    End Sub

    Private Function ErstelleArtikelZeile(anzahl As String, text As String, preisNetto As String, preisBrutto As String, mwst As String) As Panel
        Dim w As Integer = pnlRows.ClientSize.Width - 10
        If w < 500 Then w = 680

        Dim pnlZeile As New Panel With {
            .Width = w, .Height = 68,
            .Margin = New Padding(0, 0, 0, 4),
            .Name = "Zeile",
            .BackColor = CLR_WEISS
        }

        AddHandler pnlZeile.Paint, Sub(s As Object, ev As PaintEventArgs)
                                       ev.Graphics.DrawLine(New Pen(CLR_BORDER, 1), 0, DirectCast(s, Panel).Height - 1, DirectCast(s, Panel).Width, DirectCast(s, Panel).Height - 1)
                                       ' Linker Blau-Akzentstreifen
                                       ev.Graphics.FillRectangle(New SolidBrush(CLR_GOLD_AKZENT), 0, 0, 3, DirectCast(s, Panel).Height - 1)
                                   End Sub

        Dim rahmenWidth As Integer = w - 96

        Dim pnlRahmen As New Panel With {
            .Width = rahmenWidth, .Height = 68,
            .Location = New Point(0, 0),
            .BorderStyle = BorderStyle.None,
            .BackColor = CLR_WEISS,
            .Name = "Rahmen",
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        Dim txtAnzahl As New TextBox With {
            .Name = "txtAnzahl", .Text = anzahl,
            .Location = New Point(10, 22), .Width = 44,
            .TextAlign = HorizontalAlignment.Center,
            .Font = FONT_NORMAL,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = CLR_WEISS
        }

        Dim txtText As New TextBox With {
            .Name = "txtText", .Text = text,
            .Location = New Point(62, 10),
            .Width = rahmenWidth - 494,
            .Multiline = True, .Height = 48,
            .ScrollBars = ScrollBars.None,
            .Font = FONT_NORMAL,
            .BorderStyle = BorderStyle.None,
            .BackColor = CLR_WEISS,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        ' Zwei Preisfelder statt einem - welches auch immer zuletzt getippt wird, gilt als
        ' die "echte" Zahl, das jeweils andere wird live nachgerechnet (siehe Sync-Handler
        ' unten). Kein globaler Eingabemodus mehr nötig.
        Dim txtPreisNetto As New TextBox With {
            .Name = "txtPreisNetto", .Text = preisNetto,
            .Location = New Point(rahmenWidth - 422, 22),
            .Width = 62, .TextAlign = HorizontalAlignment.Right,
            .Font = FONT_NORMAL,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = CLR_WEISS,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }

        Dim txtPreisBrutto As New TextBox With {
            .Name = "txtPreisBrutto", .Text = preisBrutto,
            .Location = New Point(rahmenWidth - 354, 22),
            .Width = 62, .TextAlign = HorizontalAlignment.Right,
            .Font = FONT_NORMAL,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = CLR_WEISS,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }

        Dim cbMwSt As New ComboBox With {
            .Name = "cbMwSt",
            .Location = New Point(rahmenWidth - 282, 22),
            .Width = 60, .DropDownStyle = ComboBoxStyle.DropDownList,
            .FlatStyle = FlatStyle.Flat,
            .BackColor = CLR_WEISS,
            .Font = FONT_NORMAL,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        cbMwSt.Items.AddRange({FmtMwSt(mwstSatz1), FmtMwSt(mwstSatz2)})
        cbMwSt.Text = mwst

        Dim lblGesamt As New Label With {
            .Name = "lblGesamt", .Text = "0,00 €",
            .Location = New Point(rahmenWidth - 212, 23),
            .AutoSize = False, .Width = 100,
            .TextAlign = ContentAlignment.TopRight,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = CLR_BLAU_DUNKEL,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }

        pnlRahmen.Controls.AddRange({txtAnzahl, txtText, txtPreisNetto, txtPreisBrutto, cbMwSt, lblGesamt})

        ' Sync-Logik: Tippen in ein Feld rechnet automatisch das andere passend zum gewählten
        ' MwSt-Satz um. "aktualisiertGerade" verhindert eine Endlosschleife (Feld A ändert
        ' Feld B, das würde ohne die Sperre wieder Feld A auslösen usw.).
        Dim aktualisiertGerade As Boolean = False

        AddHandler txtPreisNetto.TextChanged, Sub()
                                                   If aktualisiertGerade Then Return
                                                   Dim netto As Decimal = 0
                                                   If ParseBetrag(txtPreisNetto.Text, netto) Then
                                                       aktualisiertGerade = True
                                                       Dim satz As Decimal = ParseMwSt(cbMwSt.Text)
                                                       txtPreisBrutto.Text = (netto * (1 + satz / 100D)).ToString("N2")
                                                       aktualisiertGerade = False
                                                   End If
                                               End Sub

        AddHandler txtPreisBrutto.TextChanged, Sub()
                                                    If aktualisiertGerade Then Return
                                                    Dim brutto As Decimal = 0
                                                    If ParseBetrag(txtPreisBrutto.Text, brutto) Then
                                                        aktualisiertGerade = True
                                                        Dim satz As Decimal = ParseMwSt(cbMwSt.Text)
                                                        txtPreisNetto.Text = (brutto / (1 + satz / 100D)).ToString("N2")
                                                        aktualisiertGerade = False
                                                    End If
                                                End Sub

        ' Ändert sich der MwSt-Satz, bleibt Netto die feste Bezugsgröße und Brutto wird
        ' neu berechnet (Netto ist ohnehin die intern gespeicherte Basis für die Steuer).
        AddHandler cbMwSt.SelectedIndexChanged, Sub()
                                                     Dim netto As Decimal = 0
                                                     If ParseBetrag(txtPreisNetto.Text, netto) Then
                                                         aktualisiertGerade = True
                                                         Dim satz As Decimal = ParseMwSt(cbMwSt.Text)
                                                         txtPreisBrutto.Text = (netto * (1 + satz / 100D)).ToString("N2")
                                                         aktualisiertGerade = False
                                                     End If
                                                 End Sub

        ' Speichern-Button (Pfeil nach links = in Stamm übernehmen)
        Dim btnSaveArt As New Button With {
            .Name = "btnSaveArt", .Text = "⬅",
            .Location = New Point(w - 90, 17),
            .Width = 36, .Height = 34,
            .Enabled = False,
            .BackColor = Color.FromArgb(220, 220, 220),
            .ForeColor = CLR_TEXT_GRAU,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .Cursor = Cursors.Hand,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        btnSaveArt.FlatAppearance.BorderSize = 0

        Dim btnDel As New Button With {
            .Text = "✕",
            .Location = New Point(w - 48, 17),
            .Width = 36, .Height = 34,
            .ForeColor = CLR_ROT,
            .BackColor = Color.FromArgb(255, 245, 245),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Cursor = Cursors.Hand,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        btnDel.FlatAppearance.BorderSize = 0
        btnDel.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 220, 220)

        AddHandler txtAnzahl.TextChanged, AddressOf BerechneSummenTab1
        AddHandler txtPreisNetto.TextChanged, AddressOf BerechneSummenTab1
        AddHandler txtPreisBrutto.TextChanged, AddressOf BerechneSummenTab1
        AddHandler cbMwSt.SelectedIndexChanged, AddressOf BerechneSummenTab1
        AddHandler txtText.TextChanged, AddressOf TxtText_TextChanged
        AddHandler txtText.TextChanged, AddressOf BerechneSummenTab1
        AddHandler btnDel.Click, AddressOf BtnDel_Click
        AddHandler btnSaveArt.Click, AddressOf BtnSaveArt_Click
        AddHandler txtText.DoubleClick, AddressOf TxtText_DoubleClick

        pnlZeile.Controls.AddRange({pnlRahmen, btnSaveArt, btnDel})
        Return pnlZeile
    End Function

    Private Sub BaueTabKontrolle()
        TabKontrolle.Name = "TabKontrolle"
        TabKontrolle.BackColor = CLR_HINTERGRUND
        TabKontrolle.Controls.Clear()

        ' ── FOOTER ──────────────────────────────────────────────────────────
        Dim pnlF As New Panel With {.Dock = DockStyle.Bottom, .Height = 70, .BackColor = CLR_PANEL_BG}
        AddHandler pnlF.Paint, Sub(s As Object, ev As PaintEventArgs)
                                   ev.Graphics.DrawLine(New Pen(CLR_BORDER, 1), 0, 0, DirectCast(s, Panel).Width, 0)
                               End Sub

        lblSummenTab2.Location = New Point(16, 10)
        lblSummenTab2.AutoSize = True
        lblSummenTab2.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        lblSummenTab2.ForeColor = CLR_BLAU_DUNKEL

        Dim lblH As New Label With {.Text = "Händler-Rechnung:", .Location = New Point(16, 38), .AutoSize = True, .Font = FONT_KLEIN, .ForeColor = CLR_TEXT_GRAU}
        Dim txtH As New TextBox With {.Name = "txtH", .Location = New Point(138, 35), .Width = 90, .Font = FONT_NORMAL, .BorderStyle = BorderStyle.FixedSingle, .BackColor = CLR_WEISS}
        AddHandler txtH.TextChanged, AddressOf PruefeSummen
        lblKontrolleTab2.Location = New Point(238, 38)
        lblKontrolleTab2.AutoSize = True
        lblKontrolleTab2.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)

        TabKontrolle.Controls.Add(pnlF)

        ' Buttons — von rechts nach links
        Dim btnV = MachePrimaerButton("▶  ALLE VERARBEITEN", 210, 44)
        btnV.Location = New Point(pnlF.Width - 228, 12)
        btnV.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnV.Click, AddressOf VerarbeiteStapel

        Dim btnDelRe = MacheGefahrButton("LÖSCHEN", 130, 44)
        btnDelRe.Location = New Point(pnlF.Width - 368, 12)
        btnDelRe.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnDelRe.Click, AddressOf BtnLoeschenTab2_Click

        Dim btnBearbeiten As New Button With {
            .Text = "BEARBEITEN", .Size = New Size(140, 44),
            .BackColor = CLR_ORANGE, .ForeColor = CLR_WEISS,
            .FlatStyle = FlatStyle.Flat, .Font = FONT_TITLE, .Cursor = Cursors.Hand,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        btnBearbeiten.FlatAppearance.BorderSize = 0
        btnBearbeiten.Location = New Point(pnlF.Width - 518, 12)
        AddHandler btnBearbeiten.Click, AddressOf BtnBearbeiten_Click

        Dim btnImportSatellit = MacheSekundaerButton("SATELLIT-IMPORT", 175, 44)
        btnImportSatellit.Location = New Point(pnlF.Width - 703, 12)
        btnImportSatellit.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler btnImportSatellit.Click, AddressOf ImportiereSatellitenXML

        pnlF.Controls.AddRange({lblSummenTab2, lblH, txtH, lblKontrolleTab2, btnDelRe, btnBearbeiten, btnImportSatellit, btnV})

        ' ── HAUPTBEREICH ────────────────────────────────────────────────────
        Dim tlp As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1}
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0!))
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 60.0!))

        ' Linke Seite
        Dim pnlLeft As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(12)}
        StyleDgv(dgvRechnungen)
        dgvRechnungen.Dock = DockStyle.Fill
        AddHandler dgvRechnungen.SelectionChanged, AddressOf DgvRechnungen_SelectionChanged
        pnlLeft.Controls.Add(dgvRechnungen)
        pnlLeft.Controls.Add(MacheAbschnittsLabel("ERFASSTE RECHNUNGEN"))

        ' Rechte Seite
        Dim pnlRight As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 12, 12, 0)}

        Dim pnlDetailsFooter As New Panel With {.Dock = DockStyle.Bottom, .Height = 150, .BackColor = CLR_PANEL_BG}
        AddHandler pnlDetailsFooter.Paint, Sub(s As Object, ev As PaintEventArgs)
                                               ev.Graphics.DrawLine(New Pen(CLR_BORDER, 1), 0, 0, DirectCast(s, Panel).Width, 0)
                                           End Sub
        lblDetailsSumme.Dock = DockStyle.Fill
        lblDetailsSumme.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        lblDetailsSumme.ForeColor = CLR_BLAU_DUNKEL
        lblDetailsSumme.TextAlign = ContentAlignment.MiddleRight
        lblDetailsSumme.Padding = New Padding(0, 10, 30, 10)
        pnlDetailsFooter.Controls.Add(lblDetailsSumme)

        Dim pnlDetailsHeader As New Panel With {.Dock = DockStyle.Top, .Height = 30, .BackColor = CLR_BLAU_DUNKEL}
        Dim hLabels = {("Anzahl", 10, 50), ("Beschreibung", 70, 200), ("E-Preis", -300, 80), ("MwSt", -200, 60), ("Netto", -120, 80)}
        For Each hl In hLabels
            Dim lx As Integer = If(hl.Item2 < 0, 1000 + hl.Item2, hl.Item2)
            Dim anch = If(hl.Item2 < 0, AnchorStyles.Top Or AnchorStyles.Right, AnchorStyles.Top Or AnchorStyles.Left)
            Dim al = If(hl.Item2 < 0, ContentAlignment.TopRight, ContentAlignment.TopLeft)
            pnlDetailsHeader.Controls.Add(New Label With {
                .Text = hl.Item1, .Location = New Point(lx, 7), .Width = hl.Item3,
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .ForeColor = CLR_WEISS, .TextAlign = al, .Anchor = anch
            })
        Next

        pnlDetailsContent.Dock = DockStyle.Fill
        pnlDetailsContent.AutoScroll = True
        pnlDetailsContent.BackColor = CLR_WEISS
        pnlDetailsContent.FlowDirection = FlowDirection.TopDown
        pnlDetailsContent.WrapContents = False
        pnlDetailsContent.Padding = New Padding(0, 0, 0, 20)

        pnlRight.Controls.Add(pnlDetailsContent)
        pnlRight.Controls.Add(pnlDetailsHeader)
        pnlRight.Controls.Add(MacheAbschnittsLabel("RECHNUNGSDETAILS"))
        pnlRight.Controls.Add(pnlDetailsFooter)

        tlp.Controls.Add(pnlLeft, 0, 0)
        tlp.Controls.Add(pnlRight, 1, 0)
        TabKontrolle.Controls.Add(tlp)
        tlp.BringToFront()
    End Sub

    Private Sub BaueTabArchiv()
        TabArchiv.Name = "TabArchiv"
        TabArchiv.BackColor = CLR_HINTERGRUND

        Dim tlp As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1}
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0!))
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 60.0!))

        Dim pnlLeft As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(12)}
        StyleDgv(dgvArchiv)
        dgvArchiv.Dock = DockStyle.Fill
        AddHandler dgvArchiv.SelectionChanged, AddressOf DgvArchiv_SelectionChanged
        pnlLeft.Controls.Add(dgvArchiv)
        pnlLeft.Controls.Add(MacheAbschnittsLabel("ARCHIVIERTE RECHNUNGEN"))

        Dim pnlRight As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 12, 12, 0)}
        Dim pnlFooter As New Panel With {.Dock = DockStyle.Bottom, .Height = 150, .BackColor = CLR_PANEL_BG}
        AddHandler pnlFooter.Paint, Sub(s As Object, ev As PaintEventArgs)
                                        ev.Graphics.DrawLine(New Pen(CLR_BORDER, 1), 0, 0, DirectCast(s, Panel).Width, 0)
                                    End Sub

        Dim pnlButtons As New Panel With {.Dock = DockStyle.Left, .Width = 320, .BackColor = Color.Transparent}

        Dim btnOpen = MacheSekundaerButton("📄  PDF ÖFFNEN", 140, 40)
        btnOpen.Location = New Point(16, 55)
        AddHandler btnOpen.Click, AddressOf BtnOeffnenArchiv_Click

        Dim btnPrint = MachePrimaerButton("🖨  DRUCKEN", 140, 40)
        btnPrint.Location = New Point(168, 55)
        AddHandler btnPrint.Click, AddressOf BtnDruckenArchiv_Click

        pnlButtons.Controls.AddRange({btnOpen, btnPrint})

        lblArchivSumme.Dock = DockStyle.Fill
        lblArchivSumme.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        lblArchivSumme.ForeColor = CLR_BLAU_DUNKEL
        lblArchivSumme.TextAlign = ContentAlignment.MiddleRight
        lblArchivSumme.Padding = New Padding(0, 10, 30, 10)

        pnlFooter.Controls.Add(pnlButtons)
        pnlFooter.Controls.Add(lblArchivSumme)

        Dim pnlDetailsHeader As New Panel With {.Dock = DockStyle.Top, .Height = 30, .BackColor = CLR_BLAU_DUNKEL}
        Dim hLabels2 = {("Anzahl", 10, 50), ("Beschreibung", 70, 200), ("E-Preis", -300, 80), ("MwSt", -200, 60), ("Netto", -120, 80)}
        For Each hl In hLabels2
            Dim lx As Integer = If(hl.Item2 < 0, 1000 + hl.Item2, hl.Item2)
            Dim anch = If(hl.Item2 < 0, AnchorStyles.Top Or AnchorStyles.Right, AnchorStyles.Top Or AnchorStyles.Left)
            Dim al = If(hl.Item2 < 0, ContentAlignment.TopRight, ContentAlignment.TopLeft)
            pnlDetailsHeader.Controls.Add(New Label With {
                .Text = hl.Item1, .Location = New Point(lx, 7), .Width = hl.Item3,
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .ForeColor = CLR_WEISS, .TextAlign = al, .Anchor = anch
            })
        Next

        pnlArchivDetails.Dock = DockStyle.Fill
        pnlArchivDetails.AutoScroll = True
        pnlArchivDetails.BackColor = CLR_WEISS
        pnlArchivDetails.FlowDirection = FlowDirection.TopDown
        pnlArchivDetails.WrapContents = False
        pnlArchivDetails.Padding = New Padding(0, 0, 0, 20)

        pnlRight.Controls.Add(pnlArchivDetails)
        pnlRight.Controls.Add(pnlDetailsHeader)
        pnlRight.Controls.Add(MacheAbschnittsLabel("RECHNUNGSDETAILS"))
        pnlRight.Controls.Add(pnlFooter)

        tlp.Controls.Add(pnlLeft, 0, 0)
        tlp.Controls.Add(pnlRight, 1, 0)
        TabArchiv.Controls.Add(tlp)
    End Sub

    Private Sub BaueTabArtikel()
        TabArtikel.Name = "TabArtikel"
        TabArtikel.BackColor = CLR_HINTERGRUND

        Dim pnlT = MacheToolbar()
        Dim btnE = MacheSekundaerButton("📤  EXPORT", 120, 36)
        btnE.Location = New Point(12, 10)
        Dim btnI = MachePrimaerButton("IMPORT", 120, 36)
        btnI.Location = New Point(142, 10)
        AddHandler btnE.Click, AddressOf ExportiereExcel
        AddHandler btnI.Click, AddressOf ImportiereExcel
        pnlT.Controls.AddRange({btnE, btnI})

        StyleDgv(dgvArtikelVerwaltung)
        dgvArtikelVerwaltung.Dock = DockStyle.Fill
        dgvArtikelVerwaltung.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvArtikelVerwaltung.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvArtikelVerwaltung.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        AddHandler dgvArtikelVerwaltung.CellDoubleClick, AddressOf DgvArtikel_CellDoubleClick
        AddHandler dgvArtikelVerwaltung.MouseDown, AddressOf DgvArtikel_MouseDown
        AddHandler dgvArtikelVerwaltung.MouseMove, AddressOf DgvArtikel_MouseMove
        AddHandler dgvArtikelVerwaltung.MouseUp, AddressOf DgvArtikel_MouseUp
        AddHandler dgvArtikelVerwaltung.Paint, AddressOf DgvArtikel_Paint

        ' Rechtsklick-Menü: Bearbeiten (identisch zum Doppelklick) und Löschen (direkt, mit
        ' Sicherheitsabfrage). Doppelklick bleibt zusätzlich unverändert bestehen.
        Dim miArtikelBearbeiten As New ToolStripMenuItem("Bearbeiten")
        Dim miArtikelLoeschen As New ToolStripMenuItem("Löschen")
        AddHandler miArtikelBearbeiten.Click, Sub()
                                                   If dgvArtikelVerwaltung.SelectedRows.Count > 0 Then
                                                       BearbeiteArtikelZeile(dgvArtikelVerwaltung.SelectedRows(0))
                                                   End If
                                               End Sub
        AddHandler miArtikelLoeschen.Click, Sub()
                                                 If dgvArtikelVerwaltung.SelectedRows.Count > 0 Then
                                                     LoescheArtikelZeile(dgvArtikelVerwaltung.SelectedRows(0))
                                                 End If
                                             End Sub
        cmsArtikel.Items.AddRange({miArtikelBearbeiten, miArtikelLoeschen})
        dgvArtikelVerwaltung.ContextMenuStrip = cmsArtikel

        Dim pnlMain As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(12, 8, 12, 12)}
        pnlMain.Controls.Add(dgvArtikelVerwaltung)

        TabArtikel.Controls.AddRange({pnlMain, pnlT})
    End Sub

    Private Sub BaueTabMitglieder()
        TabMitglieder.Name = "TabMitglieder"
        TabMitglieder.BackColor = CLR_HINTERGRUND

        Dim pnlTop = MacheToolbar()
        Dim btnExportM = MacheSekundaerButton("📤  EXPORT", 110, 36)
        btnExportM.Location = New Point(12, 10)
        Dim btnImportM = MachePrimaerButton("IMPORT", 110, 36)
        btnImportM.Location = New Point(132, 10)
        Dim btnPruefenM = MacheSekundaerButton("🔎  PRÜFEN", 120, 36)
        btnPruefenM.Location = New Point(252, 10)
        Dim btnExportSatellit As New Button With {
            .Text = "📡  FÜR SATELLIT", .Size = New Size(160, 36),
            .Location = New Point(382, 10),
            .BackColor = Color.FromArgb(255, 249, 220),
            .ForeColor = Color.FromArgb(120, 80, 0),
            .FlatStyle = FlatStyle.Flat, .Font = FONT_NORMAL, .Cursor = Cursors.Hand
        }
        btnExportSatellit.FlatAppearance.BorderColor = Color.FromArgb(200, 160, 60)

        AddHandler btnExportM.Click, AddressOf ExportiereMitgliederExcel
        AddHandler btnImportM.Click, AddressOf ImportiereMitgliederExcel
        AddHandler btnPruefenM.Click, AddressOf BtnMitgliederPruefen_Click
        AddHandler btnExportSatellit.Click, AddressOf ExportiereMitgliederFuerSatellit
        pnlTop.Controls.AddRange({btnExportM, btnImportM, btnPruefenM, btnExportSatellit})

        Dim tlp As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1}
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 35.0!))
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 65.0!))

        ' Linke Seite: Liste
        Dim pnlLeft As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(12)}
        StyleDgv(dgvMitglieder)
        dgvMitglieder.Dock = DockStyle.Fill
        dgvMitglieder.ReadOnly = True
        pnlLeft.Controls.Add(dgvMitglieder)
        pnlLeft.Controls.Add(MacheAbschnittsLabel("KUNDENLISTE"))

        ' Rechte Seite: Formular
        Dim pnlRight As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 8, 12, 12), .AutoScroll = True}

        cbM_Versand.Items.AddRange({"Beides", "E-Mail", "Post"})
        cbM_Preisart.Items.AddRange({"Netto", "Brutto"})

        ' Höhe von 160 auf 215 vergrößert für die dritte Zeile (Standard-Rechnungsart) -
        ' alle nachfolgenden Elemente (gbAdresse, pnlBtnBar) entsprechend um 55px verschoben,
        ' derselbe Abstand wie vorher jeweils beibehalten.
        Dim gbStamm As New GroupBox With {.Text = "1. Stammdaten", .Location = New Point(10, 40), .Size = New Size(720, 215)}
        StyleGroupBox(gbStamm)
        ErstelleFeld(gbStamm, "Kunden-Nr.*", txtM_Nr, 20, 28, 150)
        ErstelleFeld(gbStamm, "Firma / Name*", txtM_Name, 190, 28, 310)
        ErstelleFeld(gbStamm, "Versandart", cbM_Versand, 520, 28, 160)
        ErstelleFeld(gbStamm, "E-Mail Adresse", txtM_Email, 20, 90, 310)
        ErstelleFeld(gbStamm, "Steuernummer", txtM_Steuer, 350, 90, 150)
        ErstelleFeld(gbStamm, "Betriebsnummer", txtM_Betrieb, 520, 90, 160)
        ErstelleFeld(gbStamm, "Rechnungsart (Standard)", cbM_Preisart, 20, 152, 160)

        Dim gbAdresse As New GroupBox With {.Text = "2. Rechnungsadresse", .Location = New Point(10, 265), .Size = New Size(720, 100)}
        StyleGroupBox(gbAdresse)
        ErstelleFeld(gbAdresse, "Straße & Hausnummer", txtM_Strasse, 20, 28, 310)
        ErstelleFeld(gbAdresse, "PLZ", txtM_PLZ, 350, 28, 80)
        ErstelleFeld(gbAdresse, "Ort", txtM_Ort, 450, 28, 150)
        ErstelleFeld(gbAdresse, "Land", txtM_Land, 620, 28, 60)
        txtM_Land.Text = "DE"

        Dim pnlBtnBar As New Panel With {.Location = New Point(10, 385), .Size = New Size(720, 50), .BackColor = Color.Transparent}
        Dim btnNeu = MacheSekundaerButton("➕  NEU LEEREN", 145, 40)
        btnNeu.Location = New Point(0, 5)
        Dim btnLöschen = MacheGefahrButton("LÖSCHEN", 145, 40)
        btnLöschen.Location = New Point(420, 5)
        Dim btnSpeichern = MachePrimaerButton("💾  SPEICHERN", 145, 40)
        btnSpeichern.Location = New Point(575, 5)

        AddHandler btnNeu.Click, AddressOf BtnNeu_Mitglied_Click
        AddHandler btnSpeichern.Click, AddressOf BtnSpeichern_Mitglied_Click
        AddHandler btnLöschen.Click, AddressOf BtnLoeschen_Mitglied_Click
        AddHandler dgvMitglieder.SelectionChanged, AddressOf DgvMitglieder_SelectionChanged
        pnlBtnBar.Controls.AddRange({btnNeu, btnLöschen, btnSpeichern})

        pnlRight.Controls.AddRange({gbStamm, gbAdresse, pnlBtnBar})

        tlp.Controls.Add(pnlLeft, 0, 0)
        tlp.Controls.Add(pnlRight, 1, 0)
        TabMitglieder.Controls.AddRange({tlp, pnlTop})
    End Sub

    Private Sub BaueTabEinstellungen()
        TabEinstellungen.Name = "TabEinstellungen"
        TabEinstellungen.BackColor = CLR_HINTERGRUND
        ' Nur EIN Scroll-Container (pnlMain unten) - ein zusätzliches AutoScroll hier auf der
        ' TabPage selbst hat zuvor dazu geführt, dass die Größenberechnung durcheinanderkam
        ' und der unterste Block (Gefahrenzone) aus dem sichtbaren Bereich herausragte.

        Dim pnlMain As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(20), .AutoScroll = True}

        ' 1. Programmeinstellungen
        Dim gbProg As New GroupBox With {.Text = "1. Programmeinstellungen & Workflow", .Location = New Point(20, 20), .Size = New Size(1050, 215)}
        StyleGroupBox(gbProg)

        chkE_MengenNullen.Text = "Nach Rechnungs-Erstellung: Alle Artikel-Mengen automatisch leeren"
        chkE_MengenNullen.Location = New Point(20, 30)
        chkE_MengenNullen.AutoSize = True
        chkE_MengenNullen.Font = FONT_NORMAL
        chkE_MengenNullen.ForeColor = CLR_TEXT_DUNKEL

        Dim lblKopien As New Label With {.Text = "Druck-Exemplare:", .Location = New Point(20, 65), .AutoSize = True, .Font = FONT_KLEIN, .ForeColor = CLR_TEXT_GRAU}
        cbE_DruckAnzahl.Location = New Point(150, 62)
        cbE_DruckAnzahl.Size = New Size(60, 26)
        cbE_DruckAnzahl.Items.AddRange({"1", "2", "3", "4", "5"})
        cbE_DruckAnzahl.SelectedIndex = 0
        cbE_DruckAnzahl.Font = FONT_NORMAL
        cbE_DruckAnzahl.FlatStyle = FlatStyle.Flat
        cbE_DruckAnzahl.BackColor = CLR_WEISS

        Dim lblDrucker As New Label With {.Text = "Standard-Drucker:", .Location = New Point(240, 65), .AutoSize = True, .Font = FONT_KLEIN, .ForeColor = CLR_TEXT_GRAU}
        cbE_StandardDrucker.Location = New Point(360, 62)
        cbE_StandardDrucker.Size = New Size(320, 26)
        cbE_StandardDrucker.Font = FONT_NORMAL
        cbE_StandardDrucker.FlatStyle = FlatStyle.Flat
        cbE_StandardDrucker.BackColor = CLR_WEISS
        For Each printer As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            cbE_StandardDrucker.Items.Add(printer)
        Next

        chkE_AutoBackup.Text = "Automatisches Backup per E-Mail beim Beenden (wenn neue Rechnungen vorhanden)"
        chkE_AutoBackup.Location = New Point(20, 100)
        chkE_AutoBackup.AutoSize = True
        chkE_AutoBackup.Font = FONT_NORMAL
        chkE_AutoBackup.ForeColor = CLR_TEXT_DUNKEL

        Dim btnBackupManu = MacheSekundaerButton("💾  JETZT PER E-MAIL SICHERN", 280, 32)
        btnBackupManu.Location = New Point(20, 135)
        AddHandler btnBackupManu.Click, AddressOf BtnBackupManu_Click

        chkE_Ausgangskopie.Text = "Kopie jeder versendeten Rechnung per E-Mail archivieren an:"
        chkE_Ausgangskopie.Location = New Point(20, 178)
        chkE_Ausgangskopie.AutoSize = True
        chkE_Ausgangskopie.Font = FONT_NORMAL
        chkE_Ausgangskopie.ForeColor = CLR_TEXT_DUNKEL

        ' Bewusst großzügiger Abstand zum Haken (X=650), da die Breite des Haken-Textes
        ' zur Laufzeit vom tatsächlichen Font abhängt und hier nicht exakt vermessen werden
        ' kann - lieber unnötig viel Luft lassen als ein Überlappungsrisiko eingehen.
        txtE_AusgangskopieEmail.Location = New Point(650, 175)
        txtE_AusgangskopieEmail.Size = New Size(300, 26)
        txtE_AusgangskopieEmail.Font = FONT_NORMAL
        txtE_AusgangskopieEmail.BorderStyle = BorderStyle.FixedSingle
        txtE_AusgangskopieEmail.BackColor = CLR_WEISS

        Dim tipAusgangskopie As New ToolTip()
        tipAusgangskopie.SetToolTip(txtE_AusgangskopieEmail, "Leer lassen, um an die eigene Rechnungs-E-Mail-Adresse (siehe Firmenprofil) zu senden.")

        gbProg.Controls.AddRange({chkE_MengenNullen, lblKopien, cbE_DruckAnzahl, lblDrucker, cbE_StandardDrucker, chkE_AutoBackup, btnBackupManu, chkE_Ausgangskopie, txtE_AusgangskopieEmail})

        ' 2. Firmenprofil
        Dim gbFirma As New GroupBox With {.Text = "2. Firmenprofil & Kontakt", .Location = New Point(20, 255), .Size = New Size(1050, 105)}
        StyleGroupBox(gbFirma)
        ErstelleFeld(gbFirma, "Firmenname (GbR)", txtE_FirmaName, 20, 28, 250)
        ErstelleFeld(gbFirma, "Straße & Hausnummer", txtE_FirmaStrasse, 290, 28, 200)
        ErstelleFeld(gbFirma, "PLZ", txtE_FirmaPLZ, 510, 28, 60)
        ErstelleFeld(gbFirma, "Ort", txtE_FirmaOrt, 590, 28, 120)
        ErstelleFeld(gbFirma, "Telefon", txtE_FirmaTel, 730, 28, 120)
        ErstelleFeld(gbFirma, "E-Mail", txtE_FirmaMail, 870, 28, 160)

        ' 3. Bank, Steuernummer & MwSt-Sätze
        Dim gbBank As New GroupBox With {.Text = "3. Bankverbindung, Steuernummer & MwSt-Sätze (§24 UStG)", .Location = New Point(20, 380), .Size = New Size(1050, 105)}
        StyleGroupBox(gbBank)
        ErstelleFeld(gbBank, "IBAN", txtE_IBAN, 20, 28, 220)
        ErstelleFeld(gbBank, "BIC", txtE_BIC, 260, 28, 120)
        ErstelleFeld(gbBank, "Bankname", txtE_Bank, 400, 28, 180)
        ErstelleFeld(gbBank, "MwSt-Satz 1 (%)", txtE_MwSt1, 600, 28, 90)
        ErstelleFeld(gbBank, "MwSt-Satz 2 (%)", txtE_MwSt2, 700, 28, 90)
        ErstelleFeld(gbBank, "Steuernummer", txtE_Steuer, 810, 28, 210)

        ' 4. Texte
        Dim gbTexte As New GroupBox With {.Text = "4. Rechnungstexte & E-Mail Vorlage", .Location = New Point(20, 505), .Size = New Size(1050, 260)}
        StyleGroupBox(gbTexte)
        Dim lblInfo As New Label With {
            .Text = "  Platzhalter: [RE-nummer]",
            .Location = New Point(20, 26), .AutoSize = True,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Italic),
            .ForeColor = CLR_BLAU_DUNKEL
        }
        gbTexte.Controls.Add(lblInfo)
        ErstelleMultiFeld(gbTexte, "Zahlungsbedingungen (Normale Rechnung)", txtE_TextZahlung, 20, 55, 320, 160)
        ErstelleMultiFeld(gbTexte, "Gutschrifts-Text (Bei Summe < 0 €)", txtE_TextGutschrift, 360, 55, 320, 160)
        ErstelleMultiFeld(gbTexte, "Standard E-Mail Text", txtE_TextEmail, 700, 55, 330, 160)

        ' 5. SMTP
        Dim gbSmtp As New GroupBox With {.Text = "5. E-Mail Postausgangsserver (SMTP)", .Location = New Point(20, 785), .Size = New Size(1050, 105)}
        StyleGroupBox(gbSmtp)
        ErstelleFeld(gbSmtp, "SMTP-Server", txtE_SmtpServer, 20, 28, 260)
        ErstelleFeld(gbSmtp, "Port (587 / 465)", txtE_SmtpPort, 300, 28, 140)
        ErstelleFeld(gbSmtp, "Benutzername (E-Mail)", txtE_SmtpUser, 460, 28, 230)
        ErstelleFeld(gbSmtp, "Passwort", txtE_SmtpPass, 710, 28, 310)

        ' 6. Speicherort
        ' Keine eigene Fußzeilen-Konfiguration mehr nötig: als Einzelunternehmer ohne
        ' Geschäftsführer und mit Bürositz = Firmensitz zieht die Rechnung ihre schlanke
        ' Kontakt-Fußzeile automatisch aus dem Firmenprofil oben (Abschnitt 2).
        ' Höhe 170 für zwei Zeilen (wie ursprünglich) - 65/105 waren zu knapp bemessen.
        Dim gbSystem As New GroupBox With {.Text = "6. Speicherort", .Location = New Point(20, 910), .Size = New Size(1050, 170)}
        StyleGroupBox(gbSystem)
        ErstelleFeld(gbSystem, "Haupt-Speicherpfad", txtE_Speicherpfad, 20, 28, 780)
        Dim btnSpeicherpfadAendern = MacheSekundaerButton("Ändern…", 130, 28)
        btnSpeicherpfadAendern.Location = New Point(820, 46)
        AddHandler btnSpeicherpfadAendern.Click, AddressOf BtnSpeicherpfadAendern_Click
        gbSystem.Controls.Add(btnSpeicherpfadAendern)

        ' Datenbank-Datei: NICHT in den Einstellungen selbst gespeichert (Henne-Ei-Problem -
        ' man müsste die DB erst öffnen, um den Pfad zu kennen), sondern in db-pfad.json
        ' neben der .exe. Hier nur Anzeige + Möglichkeit, sie zu wechseln.
        ErstelleFeld(gbSystem, "Datenbank-Datei", txtE_DbPfad, 20, 90, 780)
        txtE_DbPfad.BackColor = Color.FromArgb(240, 240, 240)
        txtE_DbPfad.Text = DatenbankManager.DatenbankPfad
        Dim btnDbAendern = MacheSekundaerButton("Ändern…", 130, 28)
        btnDbAendern.Location = New Point(820, 107)
        AddHandler btnDbAendern.Click, AddressOf BtnDbPfadAendern_Click
        gbSystem.Controls.Add(btnDbAendern)

        ' Speichern-Button
        Dim btnSpeichern = MachePrimaerButton("💾  EINSTELLUNGEN SPEICHERN", 270, 44)
        btnSpeichern.Location = New Point(20, 1100)
        AddHandler btnSpeichern.Click, AddressOf BtnSpeichern_Einstellungen_Click

        ' Gefahrenzone
        Dim pnlGefahr As New Panel With {
            .Location = New Point(20, 1165),
            .Size = New Size(1050, 210),
            .BackColor = Color.FromArgb(255, 248, 248)
        }
        AddHandler pnlGefahr.Paint, Sub(s As Object, ev As PaintEventArgs)
                                        ev.Graphics.DrawRectangle(New Pen(Color.FromArgb(200, 150, 150), 1), 0, 0, DirectCast(s, Panel).Width - 1, DirectCast(s, Panel).Height - 1)
                                    End Sub

        Dim lblGefahrTitel As New Label With {
            .Text = "  ⚠  SYSTEM & ZÄHLER — GEFAHRENZONE",
            .Location = New Point(0, 10), .Width = 1050,
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .ForeColor = CLR_ROT,
            .BackColor = Color.FromArgb(255, 235, 235),
            .Height = 30, .TextAlign = ContentAlignment.MiddleLeft
        }
        pnlGefahr.Controls.Add(lblGefahrTitel)

        pnlGefahr.Controls.Add(New Label With {.Text = "Neue Start-Rechnungsnummer:", .Location = New Point(16, 54), .AutoSize = True, .Font = FONT_NORMAL, .ForeColor = CLR_TEXT_DUNKEL})
        txtStartReNr.Location = New Point(215, 50)
        txtStartReNr.Size = New Size(110, 28)
        txtStartReNr.BorderStyle = BorderStyle.FixedSingle
        txtStartReNr.BackColor = CLR_WEISS

        Dim lblHinweis As New Label With {
            .Text = $"Format: JJJJXXX (7-stellig, muss mit {DateTime.Now.Year} beginnen, z.B. {DateTime.Now.Year}001)",
            .Location = New Point(215, 78),
            .AutoSize = True, .Font = New Font("Segoe UI", 8.0F, FontStyle.Italic),
            .ForeColor = CLR_TEXT_GRAU
        }

        Dim pnlTrennstrich2 As New Panel With {.Location = New Point(10, 112), .Size = New Size(1030, 1), .BackColor = Color.FromArgb(200, 150, 150)}

        chkSicherLoeschen.Text = "Ich bin mir absolut sicher: Alle Rechnungen unwiderruflich löschen!"
        chkSicherLoeschen.Location = New Point(16, 128)
        chkSicherLoeschen.AutoSize = True
        chkSicherLoeschen.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        chkSicherLoeschen.ForeColor = CLR_ROT

        btnLoeschen = MacheGefahrButton("DATENBANK LÖSCHEN (Eingaben unvollständig)", 420, 38)
        btnLoeschen.Location = New Point(16, 160)
        btnLoeschen.BackColor = Color.FromArgb(180, 180, 180)
        btnLoeschen.Enabled = False

        pnlGefahr.Controls.AddRange({txtStartReNr, lblHinweis, pnlTrennstrich2, chkSicherLoeschen, btnLoeschen})

        pnlMain.Controls.AddRange({gbProg, gbFirma, gbBank, gbTexte, gbSmtp, gbSystem, btnSpeichern, pnlGefahr})

        ' Explizite Scroll-Größe: pnlGefahr (unterster Block) reicht bis Y=1340, mit ihrem
        ' eigenen Rand ("Padding" von pnlMain) macht das rund 1370px Gesamthöhe. Ohne diese
        ' Angabe berechnet WinForms die AutoScroll-Größe bei absolut positionierten Controls
        ' nicht zuverlässig, wodurch die Gefahrenzone unten aus dem sichtbaren Tab herausragt,
        ' statt dass sich ein Scrollbalken zeigt.
        pnlMain.AutoScrollMinSize = New Size(1100, 1405)

        TabEinstellungen.Controls.Add(pnlMain)
    End Sub

    ' =========================================================================
    ' EVENT HANDLER TAB 1 (HILFSFUNKTIONEN)
    ' =========================================================================
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs)
        If lstArtikel.SelectedItem IsNot Nothing Then
            Dim row = DirectCast(lstArtikel.SelectedItem, DataRowView)
            Dim artNetto As Decimal = CDec(row("einzelpreis_netto"))
            Dim artMwst As Decimal = CDec(row("mwst_satz"))
            Dim neueZeile As Panel = ErstelleArtikelZeile("", row("bezeichnung").ToString(), artNetto.ToString("N2"), BruttoText(row("einzelpreis_brutto"), artNetto, artMwst), FmtMwSt(artMwst))
            pnlRows.Controls.Add(neueZeile)
            pnlRows.Controls.SetChildIndex(pnlPlusContainer, pnlRows.Controls.Count - 1)
            pnlRows.ScrollControlIntoView(pnlPlusContainer)
            Dim txtAnz As TextBox = DirectCast(neueZeile.Controls.Find("txtAnzahl", True)(0), TextBox)
            txtAnz.Focus()
            BerechneSummenTab1(Nothing, Nothing)
        End If
    End Sub

    Private Sub BtnNewLine_Click(sender As Object, e As EventArgs)
        Dim leereZeile = ErstelleArtikelZeile("0", "", "0,00", "0,00", FmtMwSt(mwstSatz1))
        pnlRows.Controls.Add(leereZeile)
        pnlRows.Controls.SetChildIndex(pnlPlusContainer, pnlRows.Controls.Count - 1)
        pnlRows.ScrollControlIntoView(pnlPlusContainer)
    End Sub

    Private Sub BtnDel_Click(sender As Object, e As EventArgs)
        pnlRows.Controls.Remove(DirectCast(DirectCast(sender, Button).Parent, Panel))
        BerechneSummenTab1(Nothing, Nothing)
    End Sub

    Private Sub ChkMitglieder_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        Me.BeginInvoke(New Action(AddressOf PruefeEingaben))

        ' Vorbelegung der Rechnungsart aus dem Kundenstamm - nur beim Ankreuzen (nicht beim
        ' Abwählen), und nur als Komfort-Vorschlag: der Nutzer kann cbPreisart danach
        ' jederzeit selbst übersteuern, das wird beim Speichern nicht mehr angetastet.
        If e.NewValue = CheckState.Checked Then
            Dim drv = TryCast(chkMitglieder.Items(e.Index), DataRowView)
            If drv IsNot Nothing Then
                Dim standardPreisart As String = drv("standard_preisart").ToString()
                cbPreisart.SelectedIndex = If(standardPreisart = "Brutto", 1, 0)
            End If
        End If
    End Sub

    Private Sub BtnSaveArt_Click(sender As Object, e As EventArgs)
        Dim btn As Button = DirectCast(sender, Button)
        Dim pnl As Panel = DirectCast(btn.Parent, Panel)

        Dim txtText As TextBox = DirectCast(pnl.Controls.Find("txtText", True)(0), TextBox)
        Dim txtPreisNetto As TextBox = DirectCast(pnl.Controls.Find("txtPreisNetto", True)(0), TextBox)
        Dim txtPreisBrutto As TextBox = DirectCast(pnl.Controls.Find("txtPreisBrutto", True)(0), TextBox)
        Dim cbMwSt As ComboBox = DirectCast(pnl.Controls.Find("cbMwSt", True)(0), ComboBox)

        Dim bezeichnung As String = txtText.Text.Trim()

        If String.IsNullOrWhiteSpace(bezeichnung) Then
            MessageBox.Show("Bitte gib eine Artikelbezeichnung ein.", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Beide Felder sind dank der Sync-Logik in ErstelleArtikelZeile bereits konsistent -
        ' hier einfach unverändert übernehmen, keine Rückrechnung nötig.
        Dim nettoPreis As Decimal = 0
        ParseBetrag(txtPreisNetto.Text, nettoPreis)
        Dim bruttoPreis As Decimal = 0
        ParseBetrag(txtPreisBrutto.Text, bruttoPreis)
        Dim mwst As Decimal = ParseMwSt(cbMwSt.Text)
        If mwst <= 0 Then mwst = mwstSatz1

        Try
            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmdMax As New SQLiteCommand("SELECT MAX(CAST(artikelnummer AS INTEGER)) FROM artikel", conn)
                Dim maxResult = cmdMax.ExecuteScalar()
                Dim naechsteNr As Integer = 1
                If maxResult IsNot Nothing AndAlso Not DBNull.Value.Equals(maxResult) Then
                    naechsteNr = Convert.ToInt32(maxResult) + 1
                End If
                Dim neueArtNr As String = naechsteNr.ToString("D3")

                Dim cmdIns As New SQLiteCommand("INSERT INTO artikel (artikelnummer, bezeichnung, einzelpreis_netto, einzelpreis_brutto, mwst_satz, einheit) VALUES (@artnr, @bez, @prs, @prsBrutto, @mwst, 'C62')", conn)
                cmdIns.Parameters.AddWithValue("@artnr", neueArtNr)
                cmdIns.Parameters.AddWithValue("@bez", bezeichnung)
                cmdIns.Parameters.AddWithValue("@prs", nettoPreis)
                cmdIns.Parameters.AddWithValue("@prsBrutto", bruttoPreis)
                cmdIns.Parameters.AddWithValue("@mwst", mwst)
                cmdIns.ExecuteNonQuery()
            End Using

            btn.Enabled = False
            btn.BackColor = Color.FromArgb(220, 220, 220)
            LadeArtikelListeLinks()
            LadeArtikelTabelle()
        Catch ex As Exception
            MessageBox.Show("Fehler beim Speichern des Artikels: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =========================================================================
    ' EVENT HANDLER GEFAHRENZONE
    ' =========================================================================
    Private Sub PruefeLoeschBedingungen()
        Dim jahr As String = DateTime.Now.Year.ToString()
        Dim eingabe As String = txtStartReNr.Text.Trim()
        Dim isNummerGueltig As Boolean = (eingabe.Length = 7 AndAlso eingabe.StartsWith(jahr) AndAlso IsNumeric(eingabe))
        Dim isHakenGesetzt As Boolean = chkSicherLoeschen.Checked

        If isNummerGueltig AndAlso isHakenGesetzt Then
            btnLoeschen.Enabled = True
            btnLoeschen.BackColor = CLR_ROT
            btnLoeschen.ForeColor = CLR_WEISS
            btnLoeschen.Text = $"DATENBANK LÖSCHEN & ZÄHLER AUF {eingabe} SETZEN"
        Else
            btnLoeschen.Enabled = False
            btnLoeschen.BackColor = Color.FromArgb(180, 180, 180)
            btnLoeschen.ForeColor = CLR_WEISS
            btnLoeschen.Text = "DATENBANK LÖSCHEN (Eingaben unvollständig)"
        End If
    End Sub

    Private Sub TxtStartReNr_TextChanged(sender As Object, e As EventArgs) Handles txtStartReNr.TextChanged
        PruefeLoeschBedingungen()
    End Sub

    Private Sub ChkSicherLoeschen_CheckedChanged(sender As Object, e As EventArgs) Handles chkSicherLoeschen.CheckedChanged
        PruefeLoeschBedingungen()
    End Sub

    Private Sub BtnLoeschen_Click(sender As Object, e As EventArgs) Handles btnLoeschen.Click
        Dim result = MessageBox.Show("LETZTE WARNUNG!" & vbCrLf & vbCrLf &
                                     "Alle bisherigen Rechnungen werden gelöscht. " &
                                     $"Die nächste Rechnung startet mit der Nummer {txtStartReNr.Text.Trim()}!",
                                     "Wirklich löschen?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)

        If result = DialogResult.Yes Then
            Dim neueStartNr As String = txtStartReNr.Text.Trim()
            Try
                Using conn = DatenbankManager.HoleVerbindung()
                    Dim cmd1 As New SQLiteCommand("DELETE FROM rechnungspositionen", conn) : cmd1.ExecuteNonQuery()
                    Dim cmd2 As New SQLiteCommand("DELETE FROM rechnungen", conn) : cmd2.ExecuteNonQuery()
                    Dim cmd3 As New SQLiteCommand("UPDATE sqlite_sequence SET seq = 0 WHERE name = 'rechnungen'", conn) : cmd3.ExecuteNonQuery()
                    Dim cmd4 As New SQLiteCommand("UPDATE sqlite_sequence SET seq = 0 WHERE name = 'rechnungspositionen'", conn) : cmd4.ExecuteNonQuery()
                    SpeichereEinstellung("start_renr", neueStartNr)
                    SpeichereEinstellung("laufende_rechnungsnummer", neueStartNr)
                End Using

                lblAktuelleReNr.Text = "re-" & neueStartNr
                chkSicherLoeschen.Checked = False
                LadeStapelverarbeitung()
                MessageBox.Show($"Neustart ab {neueStartNr} erfolgreich!", "Neustart erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Fehler beim Löschen/Speichern: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub SpeichereEinstellung(schluessel As String, wert As String)
        Try
            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmd As New SQLiteCommand("INSERT OR REPLACE INTO einstellungen (schluessel, wert) VALUES (@key, @val)", conn)
                cmd.Parameters.AddWithValue("@key", schluessel)
                cmd.Parameters.AddWithValue("@val", wert)
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            MessageBox.Show("Fehler beim Speichern der Einstellung: " & ex.Message, "Datenbank-Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnBearbeiten_Click(sender As Object, e As EventArgs)
        If dgvRechnungen.CurrentRow Is Nothing OrElse dgvRechnungen.CurrentRow.IsNewRow Then
            MessageBox.Show("Bitte wähle zuerst eine Rechnung aus der Liste aus.", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim antwort = MessageBox.Show("Möchtest du diese Rechnung wirklich bearbeiten?" & vbCrLf & vbCrLf &
                                      "Sie wird aus dem aktuellen Stapel genommen und zurück in die Erfassung (Tab 1) geladen. " &
                                      "Nach der Bearbeitung kannst du sie einfach neu erstellen.",
                                      "Rechnung bearbeiten", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If antwort = DialogResult.Yes Then
            Dim reID As Integer = CInt(dgvRechnungen.CurrentRow.Cells("id").Value)
            Dim reNrStr As String = dgvRechnungen.CurrentRow.Cells("Re-Nr").Value.ToString()
            Dim geloeschteNr As Integer = CInt(reNrStr)

            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmdMid As New SQLiteCommand("SELECT mitglied_id, preisart FROM rechnungen WHERE id = @id", conn)
                cmdMid.Parameters.AddWithValue("@id", reID)
                Dim mitgliedId As Integer = 0
                Using rMid = cmdMid.ExecuteReader()
                    If rMid.Read() Then
                        If Not DBNull.Value.Equals(rMid("mitglied_id")) Then mitgliedId = CInt(rMid("mitglied_id"))
                        cbPreisart.SelectedIndex = If(rMid("preisart").ToString() = "Brutto", 1, 0)
                    End If
                End Using

                For i As Integer = pnlRows.Controls.Count - 1 To 0 Step -1
                    If pnlRows.Controls(i).Name = "Zeile" Then pnlRows.Controls.RemoveAt(i)
                Next

                For i As Integer = 0 To chkMitglieder.Items.Count - 1
                    Dim drv = DirectCast(chkMitglieder.Items(i), DataRowView)
                    chkMitglieder.SetItemChecked(i, CInt(drv("id")) = mitgliedId)
                Next

                Dim cmdPos As New SQLiteCommand("SELECT anzahl, artikel_bezeichnung, einzelpreis, einzelpreis_brutto, mwst_satz FROM rechnungspositionen WHERE rechnung_id = @id", conn)
                cmdPos.Parameters.AddWithValue("@id", reID)
                Using reader = cmdPos.ExecuteReader()
                    While reader.Read()
                        Dim posNetto As Decimal = CDec(reader("einzelpreis"))
                        Dim posMwst As Decimal = CDec(reader("mwst_satz"))
                        Dim neueZeile = ErstelleArtikelZeile(reader("anzahl").ToString(), reader("artikel_bezeichnung").ToString(), posNetto.ToString("N2"), BruttoText(reader("einzelpreis_brutto"), posNetto, posMwst), FmtMwSt(posMwst))
                        pnlRows.Controls.Add(neueZeile)
                        pnlRows.Controls.SetChildIndex(pnlPlusContainer, pnlRows.Controls.Count - 1)
                    End While
                End Using

                Dim cmdDelPos As New SQLiteCommand("DELETE FROM rechnungspositionen WHERE rechnung_id = @id", conn)
                cmdDelPos.Parameters.AddWithValue("@id", reID) : cmdDelPos.ExecuteNonQuery()
                Dim cmdDelRe As New SQLiteCommand("DELETE FROM rechnungen WHERE id = @id", conn)
                cmdDelRe.Parameters.AddWithValue("@id", reID) : cmdDelRe.ExecuteNonQuery()
                Dim cmdUpdate As New SQLiteCommand("UPDATE rechnungen SET rechnungsnummer = CAST((CAST(rechnungsnummer AS INTEGER) - 1) AS TEXT) WHERE status = 'Erfasst' AND CAST(rechnungsnummer AS INTEGER) > @geloeschteNr", conn)
                cmdUpdate.Parameters.AddWithValue("@geloeschteNr", geloeschteNr) : cmdUpdate.ExecuteNonQuery()
                Dim cmdZaehler As New SQLiteCommand("UPDATE einstellungen SET wert = CAST((CAST(wert AS INTEGER) - 1) AS TEXT) WHERE schluessel = 'laufende_rechnungsnummer'", conn)
                cmdZaehler.ExecuteNonQuery()
            End Using

            LadeStapelverarbeitung()
            LadeDaten()
            BerechneSummenTab1(Nothing, Nothing)
            MainTabs.SelectedTab = TabErfassung
        End If
    End Sub

    Private Sub ImportiereSatellitenXML(sender As Object, e As EventArgs)
        Dim ofd As New OpenFileDialog() With {.Filter = "MHRechnung XML-Daten (*.xml)|*.xml"}
        If ofd.ShowDialog() = DialogResult.OK Then
            Try
                Dim ds As New DataSet()
                ds.ReadXml(ofd.FileName, XmlReadMode.ReadSchema)

                If Not ds.Tables.Contains("Rechnungen") OrElse Not ds.Tables.Contains("Positionen") Then
                    MessageBox.Show("Die XML-Datei hat nicht das erwartete Format.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim dtRe = ds.Tables("Rechnungen")
                Dim dtPos = ds.Tables("Positionen")
                Dim heutigesDatum As String = DateTime.Now.ToString("dd.MM.yyyy")
                Dim importCount As Integer = 0
                Dim fehlerCount As Integer = 0

                Using conn = DatenbankManager.HoleVerbindung()
                    Dim cmdGetNr As New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'laufende_rechnungsnummer'", conn)
                    Dim nrObj = cmdGetNr.ExecuteScalar()
                    If nrObj Is Nothing OrElse DBNull.Value.Equals(nrObj) Then
                        MessageBox.Show("Die laufende Rechnungsnummer fehlt in den Einstellungen. Import abgebrochen.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                    Dim aktuelleReNr As Integer = CInt(nrObj)

                    For Each rowRe As DataRow In dtRe.Rows
                        Dim interneId As Integer = Convert.ToInt32(rowRe("id"))
                        Dim mitgliedsNr As String = rowRe("Mitglieds-Nr").ToString()

                        Dim cmdMid As New SQLiteCommand("SELECT id FROM mitglieder WHERE mitgliedsnummer = @mnr", conn)
                        cmdMid.Parameters.AddWithValue("@mnr", mitgliedsNr)
                        Dim midObj = cmdMid.ExecuteScalar()

                        If midObj IsNot Nothing AndAlso Not DBNull.Value.Equals(midObj) Then
                            Dim mitgliedId As Integer = CInt(midObj)
                            Dim cmdInsertRe As New SQLiteCommand("INSERT INTO rechnungen (rechnungsnummer, datum, lieferdatum, mitglied_id, status) VALUES (@nr, @dat, @dat, @mid, 'Erfasst'); SELECT last_insert_rowid();", conn)
                            cmdInsertRe.Parameters.AddWithValue("@nr", aktuelleReNr.ToString())
                            cmdInsertRe.Parameters.AddWithValue("@dat", heutigesDatum)
                            cmdInsertRe.Parameters.AddWithValue("@mid", mitgliedId)
                            Dim neueReId As Integer = CInt(cmdInsertRe.ExecuteScalar())

                            Dim reihenPos = dtPos.Select($"rechnung_id = {interneId}")
                            For Each rowPos As DataRow In reihenPos
                                Dim cmdInsPos As New SQLiteCommand("INSERT INTO rechnungspositionen (rechnung_id, artikel_bezeichnung, anzahl, einzelpreis, mwst_satz) VALUES (@rid, @bez, @anz, @prs, @mwst)", conn)
                                cmdInsPos.Parameters.AddWithValue("@rid", neueReId)
                                cmdInsPos.Parameters.AddWithValue("@bez", rowPos("artikel_bezeichnung").ToString())
                                cmdInsPos.Parameters.AddWithValue("@anz", CDec(rowPos("anzahl")))
                                cmdInsPos.Parameters.AddWithValue("@prs", CDec(rowPos("einzelpreis")))
                                cmdInsPos.Parameters.AddWithValue("@mwst", CDec(rowPos("mwst_satz")))
                                cmdInsPos.ExecuteNonQuery()
                            Next

                            aktuelleReNr += 1
                            importCount += 1
                        Else
                            fehlerCount += 1
                        End If
                    Next

                    Dim cmdUpdateNr As New SQLiteCommand("UPDATE einstellungen SET wert = @val WHERE schluessel = 'laufende_rechnungsnummer'", conn)
                    cmdUpdateNr.Parameters.AddWithValue("@val", aktuelleReNr.ToString())
                    cmdUpdateNr.ExecuteNonQuery()
                    lblAktuelleReNr.Text = "re-" & aktuelleReNr.ToString()
                End Using

                LadeStapelverarbeitung()
                Dim msg As String = $"{importCount} Rechnungen wurden erfolgreich importiert!"
                If fehlerCount > 0 Then msg &= vbCrLf & vbCrLf & $"ACHTUNG: {fehlerCount} Rechnungen konnten nicht importiert werden (Kundennummer nicht gefunden)."
                MessageBox.Show(msg, "Import abgeschlossen", MessageBoxButtons.OK, If(fehlerCount > 0, MessageBoxIcon.Warning, MessageBoxIcon.Information))
            Catch ex As Exception
                MessageBox.Show("Fehler beim Importieren der XML-Datei: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub ExportiereMitgliederFuerSatellit(sender As Object, e As EventArgs)
        Dim sfd As New SaveFileDialog() With {.Filter = "Textdatei|*.txt", .FileName = "MHRechnung_Kunden.txt"}
        If sfd.ShowDialog() = DialogResult.OK Then
            Try
                Dim zeilen As New List(Of String)
                Using conn = DatenbankManager.HoleVerbindung()
                    Dim cmd As New SQLiteCommand("SELECT mitgliedsnummer, name FROM mitglieder ORDER BY name", conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            zeilen.Add($"{reader("mitgliedsnummer")};{reader("name")}")
                        End While
                    End Using
                End Using
                File.WriteAllLines(sfd.FileName, zeilen, System.Text.Encoding.UTF8)
                MessageBox.Show("Export für Erfassungsprogramm erfolgreich erstellt!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Fehler beim Export: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    ' =========================================================================
    ' BACKUP-LOGIK
    ' =========================================================================
    Private Sub BtnBackupManu_Click(sender As Object, e As EventArgs)
        SendeDatenbankBackup(True)
    End Sub

    Private Sub SendeDatenbankBackup(manuell As Boolean)
        Try
            Dim server As String = "", port As String = "", user As String = "", pass As String = ""
            Dim aktRechnungen As Integer = 0
            Dim letzteSicherung As Integer = 0

            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmdSmtp = New SQLiteCommand("SELECT schluessel, wert FROM einstellungen WHERE schluessel IN ('smtp_server', 'smtp_port', 'smtp_user', 'smtp_pass')", conn)
                Using reader = cmdSmtp.ExecuteReader()
                    While reader.Read()
                        Select Case reader("schluessel").ToString()
                            Case "smtp_server" : server = reader("wert").ToString()
                            Case "smtp_port" : port = reader("wert").ToString()
                            Case "smtp_user" : user = reader("wert").ToString()
                            Case "smtp_pass" : pass = reader("wert").ToString()
                        End Select
                    End While
                End Using

                If String.IsNullOrWhiteSpace(server) OrElse String.IsNullOrWhiteSpace(user) Then
                    If manuell Then MessageBox.Show("Bitte richte zuerst die SMTP-Serverdaten in den Einstellungen ein.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If Not manuell Then
                    Dim cmdAuto = New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'prog_auto_backup'", conn)
                    Dim resAuto = cmdAuto.ExecuteScalar()
                    If resAuto Is Nothing OrElse resAuto.ToString() <> "True" Then Return

                    Dim cmdCount = New SQLiteCommand("SELECT COUNT(*) FROM rechnungen", conn)
                    aktRechnungen = Convert.ToInt32(cmdCount.ExecuteScalar())
                    Dim cmdLast = New SQLiteCommand("SELECT wert FROM einstellungen WHERE schluessel = 'letztes_backup_rechnungen_anzahl'", conn)
                    Dim resLast = cmdLast.ExecuteScalar()
                    If resLast IsNot Nothing AndAlso IsNumeric(resLast) Then letzteSicherung = Convert.ToInt32(resLast)
                    If aktRechnungen <= letzteSicherung Then Return
                Else
                    Dim cmdCount = New SQLiteCommand("SELECT COUNT(*) FROM rechnungen", conn)
                    aktRechnungen = Convert.ToInt32(cmdCount.ExecuteScalar())
                End If
            End Using

            If manuell Then Me.Cursor = Cursors.WaitCursor

            ' Datenbank kann seit der frei wählbaren Speicherort-Funktion überall liegen -
            ' nicht mehr fest neben der .exe annehmen, sondern den tatsächlichen Pfad nehmen.
            Dim dbPfad As String = DatenbankManager.DatenbankPfad
            Dim tempPfad As String = Path.Combine(Path.GetTempPath(), $"Sicherung_MHRechnung_Daten_{DateTime.Now:yyyyMMdd_HHmmss}.sqlite")
            File.Copy(dbPfad, tempPfad, True)

            Using mail As New MailMessage()
                mail.From = New MailAddress(user)
                mail.To.Add(user)
                mail.Subject = "Sicherung Datenbank MHRechnung"
                mail.Body = $"Sicherung vom {DateTime.Now:dd.MM.yyyy HH:mm} Uhr. Aktuelle Rechnungen: {aktRechnungen}"
                mail.Attachments.Add(New Attachment(tempPfad))

                Using smtp As New SmtpClient(server, Convert.ToInt32(port))
                    smtp.Credentials = New NetworkCredential(user, pass)
                    smtp.EnableSsl = True
                    smtp.Send(mail)
                End Using
            End Using

            Using conn = DatenbankManager.HoleVerbindung()
                Dim cmdSave = New SQLiteCommand("INSERT OR REPLACE INTO einstellungen (schluessel, wert) VALUES ('letztes_backup_rechnungen_anzahl', @anz)", conn)
                cmdSave.Parameters.AddWithValue("@anz", aktRechnungen.ToString())
                cmdSave.ExecuteNonQuery()
            End Using

            Try : File.Delete(tempPfad) : Catch : End Try

            If manuell Then
                Me.Cursor = Cursors.Default
                MessageBox.Show("Datenbank wurde erfolgreich per E-Mail gesichert!", "Backup erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            If manuell Then
                Me.Cursor = Cursors.Default
                MessageBox.Show("Fehler beim Senden des Backups: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Sub

    ' =========================================================================
    ' DRAG & DROP E-RECHNUNG IMPORT
    ' =========================================================================
    Private Sub PnlRows_DragEnter(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If files(0).ToLower().EndsWith(".xml") OrElse files(0).ToLower().EndsWith(".pdf") Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.None
            End If
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub PnlRows_DragDrop(sender As Object, e As DragEventArgs)
        Try
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If files.Length = 0 Then Return

            Dim rechnung As ERechnungsImporter.ImportedInvoice = ERechnungsImporter.LeseDatei(files(0))

            If rechnung IsNot Nothing AndAlso rechnung.Positionen.Count > 0 Then
                Dim txtH As TextBox = DirectCast(TabKontrolle.Controls.Find("txtH", True).FirstOrDefault(), TextBox)
                If txtH IsNot Nothing Then txtH.Text = rechnung.GesamtBrutto.ToString("N2")

                For i As Integer = pnlRows.Controls.Count - 1 To 0 Step -1
                    If pnlRows.Controls(i).Name = "Zeile" Then pnlRows.Controls.RemoveAt(i)
                Next

                For Each pos In rechnung.Positionen
                    Dim neueZeile As Panel = ErstelleArtikelZeile("", pos.Bezeichnung, pos.EinzelpreisNetto.ToString("N2"), BruttoText(Nothing, pos.EinzelpreisNetto, pos.MwStSatz), FmtMwSt(pos.MwStSatz))
                    pnlRows.Controls.Add(neueZeile)
                Next

                pnlRows.Controls.SetChildIndex(pnlPlusContainer, pnlRows.Controls.Count - 1)
                pnlRows.ScrollControlIntoView(pnlPlusContainer)
                BerechneSummenTab1(Nothing, Nothing)

                MessageBox.Show($"Erfolg! {rechnung.Positionen.Count} Artikel importiert." & vbCrLf &
                                $"Händler-Kontrollbetrag auf {rechnung.GesamtBrutto:N2} € gesetzt.",
                                "Import abgeschlossen", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Fehler beim Drag & Drop: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxtText_DoubleClick(sender As Object, e As EventArgs)
        Dim txtOriginal = DirectCast(sender, TextBox)

        Using editorForm As New Form With {
            .Text = "Beschreibung bearbeiten",
            .Size = New Size(520, 450),
            .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedToolWindow,
            .BackColor = CLR_HINTERGRUND
        }
            Dim pnlBottom As New Panel With {.Dock = DockStyle.Bottom, .Height = 54, .BackColor = CLR_PANEL_BG}
            Dim btnOK = MachePrimaerButton("Übernehmen", 130, 36)
            btnOK.DialogResult = DialogResult.OK
            btnOK.Location = New Point(pnlBottom.Width - 152, 9)
            btnOK.Anchor = AnchorStyles.Top Or AnchorStyles.Right

            Dim btnCancel = MacheSekundaerButton("Abbrechen", 130, 36)
            btnCancel.DialogResult = DialogResult.Cancel
            btnCancel.Location = New Point(pnlBottom.Width - 292, 9)
            btnCancel.Anchor = AnchorStyles.Top Or AnchorStyles.Right

            pnlBottom.Controls.AddRange({btnOK, btnCancel})

            Dim txtGross As New TextBox With {
                .Multiline = True,
                .Text = txtOriginal.Text,
                .Font = New Font("Arial", 12),
                .ScrollBars = ScrollBars.Vertical,
                .Width = 479,
                .Height = editorForm.ClientSize.Height - pnlBottom.Height - 20,
                .Location = New Point((editorForm.ClientSize.Width - 479) \ 2, 10),
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = CLR_WEISS
            }

            editorForm.Controls.Add(txtGross)
            editorForm.Controls.Add(pnlBottom)
            editorForm.CancelButton = btnCancel

            If editorForm.ShowDialog() = DialogResult.OK Then
                txtOriginal.Text = txtGross.Text
            End If
        End Using
    End Sub

End Class