<#
.SYNOPSIS
Kombiniert entpackte CHM-Inhalte stabil in eine HTML-Masterdatei und speichert sie als DOCX.
#>

param (
    [Parameter(Mandatory=$true, HelpMessage="Der komplette Pfad zur .chm Datei")]
    [string]$ChmFilePath,

    [Parameter(Mandatory=$true, HelpMessage="Der Ordner, in dem das Ergebnis gespeichert werden soll")]
    [string]$OutputDirectory
)

$ExtractionDir = Join-Path -Path $OutputDirectory -ChildPath "CHM_Entpackt"
$MasterHtmlPath = Join-Path -Path $OutputDirectory -ChildPath "Copilot_Master_Wissen.html"
$DocxFilePath = Join-Path -Path $OutputDirectory -ChildPath "Copilot_Wissensbasis.docx"
$TempChmPath = Join-Path -Path $OutputDirectory -ChildPath "temp_processing_file.chm"

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

try {
    # 1. Nur entpacken, wenn der Ordner leer ist oder nicht existiert (spart Zeit)
    $htmlFiles = @()
    if (Test-Path -Path $ExtractionDir) {
        $htmlFiles = Get-ChildItem -Path $ExtractionDir -Filter "*.htm*" -Recurse | Sort-Object Name
    }

    if ($htmlFiles.Count -eq 0) {
        Write-Host "Schritt 1: Entpacke CHM-Datei..." -ForegroundColor Cyan
        Unblock-File -Path $ChmFilePath
        Copy-Item -Path $ChmFilePath -Destination $TempChmPath -Force
        
        $hhArgs = @("-decompile", $ExtractionDir, $TempChmPath)
        Start-Process -FilePath "hh.exe" -ArgumentList $hhArgs -Wait -NoNewWindow
        
        if (Test-Path -Path $TempChmPath) { Remove-Item -Path $TempChmPath -Force }
        $htmlFiles = Get-ChildItem -Path $ExtractionDir -Filter "*.htm*" -Recurse | Sort-Object Name
    } else {
        Write-Host "Schritt 1: Bereits entpackte Daten gefunden ($($htmlFiles.Count) Seiten). Nutze diese..." -ForegroundColor Green
    }

    if ($htmlFiles.Count -eq 0) {
        throw "Es wurden keine extrahierten HTML-Dateien gefunden."
    }

    # 2. Per PowerShell zu einer riesigen HTML-Datei zusammenfügen (Super schnell & absturzsicher)
    Write-Host "Schritt 2: Verbinde $($htmlFiles.Count) Seiten blitzschnell zu einer Master-Datei..." -ForegroundColor Cyan
    
    # Header für die HTML-Datei schreiben (damit Umlaute richtig dargestellt werden)
    $HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    body { font-family: Arial, sans-serif; line-height: 1.6; }
    .page-break { page-break-after: always; break-after: page; }
</style>
</head>
<body>
"@
    Set-Content -Path $MasterHtmlPath -Value $HtmlHeader -Encoding Utf8

    $counter = 1
    foreach ($file in $htmlFiles) {
        Write-Progress -Activity "Generiere Master-HTML" -Status "Verarbeite Seite $counter von $($htmlFiles.Count)" -PercentComplete (($counter / $htmlFiles.Count) * 100)
        
        # Inhalt einlesen
        $content = Get-Content -Path $file.FullName -Raw
        
        # Nur den Inhalt aus dem <body> extrahieren, falls vorhanden (sorgt für saubereres HTML)
        if ($content -match '(?s)<body[^>]*>(.*)</body>') {
            $content = $Matches[1]
        }
        
        # Inhalt anhängen + Seitenumbruch-Marker für Word
        Add-Content -Path $MasterHtmlPath -Value $content -Encoding Utf8
        Add-Content -Path $MasterHtmlPath -Value "<div class='page-break'></div>" -Encoding Utf8
        $counter++
    }
    
    Add-Content -Path $MasterHtmlPath -Value "</body></html>" -Encoding Utf8

    # 3. Word nur einmal öffnen, um die fertige Datei als DOCX abzuspeichern
    Write-Host "Schritt 3: Konvertiere Master-Datei sauber in Word-Format..." -ForegroundColor Cyan
    Write-Host "(Word speichert die Datei nun, bitte einen Moment Geduld...)" -ForegroundColor Yellow

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # Öffne die soeben erstellte große HTML-Datei
    $doc = $word.Documents.Open($MasterHtmlPath)
    
    # Speichere sie als echtes Word-Dokument (16 = wdFormatDocumentDefault)
    $doc.SaveAs2([string]$DocxFilePath, 16)
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

    # Optionale Aufräumaktion der Master-HTML
    if (Test-Path -Path $MasterHtmlPath) { Remove-Item -Path $MasterHtmlPath -Force }

    Write-Host "`nErfolgreich abgeschlossen!" -ForegroundColor Green
    Write-Host "Das fertige, prall gefüllte Dokument liegt hier:" -ForegroundColor White
    Write-Host $DocxFilePath -ForegroundColor Yellow

} catch {
    Write-Host "`n[FEHLER] Prozess abgebrochen: $_" -ForegroundColor Red
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
