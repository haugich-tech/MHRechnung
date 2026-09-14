Imports System.IO
Imports System.Text

' =========================================================================
' PFAD ZUR DATENBANK-DATEI
' Kann NICHT in der Datenbank selbst gespeichert werden (Henne-Ei-Problem:
' man müsste die Datenbank erst öffnen, um den Pfad zu lesen). Liegt deshalb
' als eigene kleine JSON-Datei neben der .exe.
'
' Bewusst kein vollwertiger JSON-Parser, sondern eine gezielte Extraktion:
' Es gibt genau ein Feld ("datenbankPfad"), dafür lohnt sich keine zusätzliche
' Abhängigkeit (z.B. Newtonsoft.Json) und kein zusätzlicher Assembly-Verweis
' (DataContractJsonSerializer ist in diesem Projekt nicht referenziert).
' =========================================================================
Public Module DbKonfiguration

    Private ReadOnly KonfigDateiPfad As String = Path.Combine(Application.StartupPath, "db-pfad.json")

    ''' <summary>Liest den gespeicherten Datenbank-Pfad. Leerer String, wenn Datei fehlt,
    ''' leer oder nicht lesbar ist.</summary>
    Public Function LiesDatenbankPfad() As String
        Try
            If Not File.Exists(KonfigDateiPfad) Then Return ""
            Dim inhalt As String = File.ReadAllText(KonfigDateiPfad, Encoding.UTF8)

            Dim marker As String = """datenbankPfad"""
            Dim idx As Integer = inhalt.IndexOf(marker)
            If idx < 0 Then Return ""
            Dim colonIdx As Integer = inhalt.IndexOf(":"c, idx)
            If colonIdx < 0 Then Return ""
            Dim startQuote As Integer = inhalt.IndexOf(""""c, colonIdx)
            If startQuote < 0 Then Return ""

            Dim sb As New StringBuilder()
            Dim i As Integer = startQuote + 1
            While i < inhalt.Length AndAlso inhalt(i) <> """"c
                If inhalt(i) = "\"c AndAlso i + 1 < inhalt.Length Then
                    i += 1
                    Select Case inhalt(i)
                        Case "\"c : sb.Append("\")
                        Case """"c : sb.Append(""""c)
                        Case "n"c : sb.Append(vbLf)
                        Case Else : sb.Append(inhalt(i))
                    End Select
                Else
                    sb.Append(inhalt(i))
                End If
                i += 1
            End While
            Return sb.ToString().Trim()
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>Schreibt den Datenbank-Pfad in die Konfigurationsdatei (legt sie an, falls nötig).</summary>
    Public Sub SchreibeDatenbankPfad(pfad As String)
        Dim escaped As String = pfad.Replace("\", "\\").Replace("""", "\""")
        Dim json As String = "{" & vbCrLf & "  ""datenbankPfad"": """ & escaped & """" & vbCrLf & "}" & vbCrLf
        File.WriteAllText(KonfigDateiPfad, json, New UTF8Encoding(False))
    End Sub

End Module
