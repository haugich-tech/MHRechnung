Imports PdfSharp.Fonts
Imports System.IO

' Diese Klasse verrät PDFsharp, wo es die echten Arial-Dateien auf dem PC findet
Public Class LegFontResolver
    Implements IFontResolver

    Public Function ResolveTypeface(familyName As String, isBold As Boolean, isItalic As Boolean) As FontResolverInfo Implements IFontResolver.ResolveTypeface
        ' Wir fangen Anfragen für Fettdruck ab
        If isBold Then
            Return New FontResolverInfo("Arial-Bold")
        Else
            Return New FontResolverInfo("Arial-Regular")
        End If
    End Function

    Public Function GetFont(faceName As String) As Byte() Implements IFontResolver.GetFont
        ' Der Standard-Pfad zu Windows-Schriftarten
        Dim fontPath As String = "C:\Windows\Fonts\arial.ttf"

        ' Wenn Fettdruck verlangt wird, nehmen wir die "arialbd.ttf"
        If faceName = "Arial-Bold" Then
            fontPath = "C:\Windows\Fonts\arialbd.ttf"
        End If

        ' Datei einlesen und an PDFsharp übergeben
        If File.Exists(fontPath) Then
            Return File.ReadAllBytes(fontPath)
        Else
            Throw New FileNotFoundException("Schriftart nicht gefunden: " & fontPath)
        End If
    End Function
End Class