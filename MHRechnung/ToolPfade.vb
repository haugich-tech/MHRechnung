Imports System.IO

Module ToolPfade
    ' Basis-Ordner, in dem deine MHRechnung.exe liegt
    Private ReadOnly Property AppRoot As String
        Get
            Return Application.StartupPath
        End Get
    End Property

    ' Relative Pfade basierend auf deiner Verzeichnisliste
    Public ReadOnly Property JavaExe As String
        Get
            Return Path.Combine(AppRoot, "Tools\Java\Eclipse Adoptium\jdk-17.0.18.8-hotspot\bin\java.exe")
        End Get
    End Property

    Public ReadOnly Property GhostscriptExe As String
        Get
            Return Path.Combine(AppRoot, "Tools\GS\bin\gswin64c.exe")
        End Get
    End Property

    Public ReadOnly Property MustangJar As String
        Get
            Return Path.Combine(AppRoot, "Tools\Mustang\Mustang-CLI-2.22.0.jar")
        End Get
    End Property
    Public ReadOnly Property LogoPfad As String
        Get
            ' Pfad: [Programm-Ordner]\logo\logo.jpg
            Return Path.Combine(AppRoot, "logo\logo.jpg")
        End Get
    End Property

    ' Check-Funktion für den Systemstart
    Public Function SindAlleToolsBereit() As Boolean
        Return File.Exists(JavaExe) AndAlso File.Exists(GhostscriptExe) AndAlso File.Exists(MustangJar)
    End Function
End Module