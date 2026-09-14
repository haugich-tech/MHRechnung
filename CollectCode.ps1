# Konfiguration
$exportFile = "Projekt_Kontext.txt"
$extensions = @("*.vb", "*.config")

Write-Host "Sammle Code-Dateien (UTF-8 Format)..." -ForegroundColor Cyan

# Datei neu erstellen (leeren) mit explizitem UTF8-Encoding
"" | Set-Content -Path $exportFile -Encoding UTF8

# Alle relevanten Dateien finden (ohne bin/obj Ordner)
$files = Get-ChildItem -Recurse -Include $extensions | Where-Object { 
    $_.FullName -notlike "*\bin\*" -and $_.FullName -notlike "*\obj\*" 
}

foreach ($file in $files) {
    $relativeName = $file.FullName
    Write-Host "Verarbeite: $relativeName"
    
    # Trenner schreiben
    $header = "`n--------------------------------------------------`n" +
              "FILE: $relativeName`n" +
              "--------------------------------------------------`n"
    
    Add-Content -Path $exportFile -Value $header -Encoding UTF8
    
    # Inhalt der Datei einlesen und anhängen
    $content = Get-Content -Path $file.FullName
    Add-Content -Path $exportFile -Value $content -Encoding UTF8
}

Write-Host "Fertig! Die Datei wurde im sauberen UTF-8 Format gespeichert." -ForegroundColor Green