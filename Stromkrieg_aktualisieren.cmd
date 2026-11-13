<# : Batch-Teil: startet den PowerShell-Teil weiter unten (Doppelklick genuegt)
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression ([System.IO.File]::ReadAllText('%~f0'))"
echo.
pause
exit /b
#>
# Stromkrieg-Referat: holt die neuesten Dateien von GitHub in den lokalen OneDrive-Ordner.
# Kein Git noetig. Ist das Repository privat, oeffnet das Skript den Download im Browser
# (dort musst du bei GitHub angemeldet sein) und nimmt die ZIP-Datei dann aus dem Ordner Downloads.
# Hast du eine Datei selbst geaendert, wird deine Version vorher nach _Sicherung kopiert.

$ErrorActionPreference = 'Stop'
$ziel = 'C:\Users\simon\OneDrive - FHNW\7EIT\wusa\Stromkrieg_Referat'
if ($env:STROMKRIEG_ZIEL) { $ziel = $env:STROMKRIEG_ZIEL }
$url = 'https://github.com/simonbruder-hub/Edison_vs_Westinghouse/archive/refs/heads/claude/sharp-tesla-5z6cdb.zip'
$ordner = @('abgabe', 'vorbereitung')
$skript = 'Stromkrieg_aktualisieren.cmd'

function Get-DownloadOrdner {
    try {
        $wert = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop).'{374DE290-123F-4565-9164-39C4925E467B}'
        if ($wert) { return [Environment]::ExpandEnvironmentVariables($wert) }
    } catch { }
    return (Join-Path $env:USERPROFILE 'Downloads')
}

function Lade-Zip([string]$zip) {
    try {
        Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
        $fs = [System.IO.File]::OpenRead($zip)
        $b1 = $fs.ReadByte(); $b2 = $fs.ReadByte(); $fs.Close()
        if ($b1 -eq 0x50 -and $b2 -eq 0x4B) { return }   # echte ZIP-Datei (beginnt mit PK)
        throw 'keine ZIP-Datei'
    } catch {
        Write-Host 'Direkter Download nicht moeglich, vermutlich weil das Repository privat ist.' -ForegroundColor Yellow
    }
    # Privates Repository: Download im Browser, wo du bei GitHub angemeldet bist
    $downloads = Get-DownloadOrdner
    $start = Get-Date
    Write-Host 'Der Browser oeffnet jetzt den Download. Melde dich dort bei GitHub an, falls noetig.'
    Start-Process $url
    Read-Host "Warte, bis der Download fertig ist (Ordner: $downloads). Dann Enter druecken"
    $datei = Get-ChildItem -LiteralPath $downloads -Filter 'Edison_vs_Westinghouse*.zip' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $datei) { throw "Keine Datei Edison_vs_Westinghouse*.zip im Ordner $downloads gefunden." }
    $zeit = @($datei.CreationTime, $datei.LastWriteTime) | Sort-Object | Select-Object -Last 1
    if ($zeit -lt $start.AddMinutes(-10)) {
        Write-Host "Achtung: $($datei.Name) ist aelter als 10 Minuten. Ist es wirklich die neue Version?" -ForegroundColor Yellow
    }
    Move-Item -LiteralPath $datei.FullName -Destination $zip -Force   # ZIP wird danach geloescht, Downloads bleibt sauber
}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('stromkrieg_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
    Write-Host 'Lade die neueste Version von GitHub ...'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $zip = Join-Path $tmp 'repo.zip'
    Lade-Zip $zip
    Expand-Archive -LiteralPath $zip -DestinationPath $tmp
    $quelle = (Get-ChildItem -LiteralPath $tmp -Directory | Select-Object -First 1).FullName
    New-Item -ItemType Directory -Force -Path $ziel | Out-Null

    # Stand der letzten Aktualisierung: so erkennt das Skript, ob du eine Datei selbst geaendert hast
    $standDatei = Join-Path $ziel '_Stand.txt'
    $stand = @{}
    if (Test-Path -LiteralPath $standDatei) {
        foreach ($zeile in Get-Content -LiteralPath $standDatei) {
            if ($zeile -match '^([0-9A-F]{64})  (.+)$') { $stand[$Matches[2]] = $Matches[1] }
        }
    }

    $stempel = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $neu = 0; $aktualisiert = 0; $gleich = 0
    $gesichert = @(); $fehler = @(); $neuerStand = @()
    foreach ($teil in $ordner) {
        $basis = Join-Path $quelle $teil
        if (-not (Test-Path -LiteralPath $basis)) { continue }
        foreach ($datei in Get-ChildItem -LiteralPath $basis -File -Recurse) {
            $rel = $datei.FullName.Substring($quelle.Length).TrimStart('\', '/')
            $dest = Join-Path $ziel $rel
            $hashNeu = (Get-FileHash -LiteralPath $datei.FullName -Algorithm SHA256).Hash
            try {
                if (Test-Path -LiteralPath $dest) {
                    $hashLokal = (Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash
                    if ($hashLokal -eq $hashNeu) {
                        $gleich++
                        $neuerStand += "$hashNeu  $rel"
                        continue
                    }
                    if ($stand[$rel] -ne $hashLokal) {  # lokal geaendert: vorher sichern
                        $sicherung = Join-Path (Join-Path (Join-Path $ziel '_Sicherung') $stempel) $rel
                        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $sicherung) | Out-Null
                        Copy-Item -LiteralPath $dest -Destination $sicherung -Force
                        $gesichert += $rel
                    }
                    $aktualisiert++
                } else {
                    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
                    $neu++
                }
                Copy-Item -LiteralPath $datei.FullName -Destination $dest -Force
                $neuerStand += "$hashNeu  $rel"
            } catch {
                $fehler += $rel
                if ($stand.ContainsKey($rel)) { $neuerStand += "$($stand[$rel])  $rel" }
            }
        }
    }
    Set-Content -LiteralPath $standDatei -Value (@("Zuletzt aktualisiert: $(Get-Date -Format 'dd.MM.yyyy HH:mm')") + $neuerStand)

    # Dieses Skript einmalig in den Ordner legen (nie ueberschreiben, solange es laeuft)
    $skriptZiel = Join-Path $ziel $skript
    $skriptQuelle = Join-Path $quelle $skript
    if (-not (Test-Path -LiteralPath $skriptZiel) -and (Test-Path -LiteralPath $skriptQuelle)) {
        Copy-Item -LiteralPath $skriptQuelle -Destination $skriptZiel
    }

    Write-Host ''
    Write-Host "Fertig: $ziel" -ForegroundColor Green
    Write-Host "  neu: $neu   aktualisiert: $aktualisiert   unveraendert: $gleich"
    if ($gesichert.Count -gt 0) {
        Write-Host "Deine eigenen Aenderungen sind gesichert in _Sicherung\$stempel :" -ForegroundColor Yellow
        $gesichert | ForEach-Object { Write-Host "  $_" }
    }
    if ($fehler.Count -gt 0) {
        Write-Host 'Nicht aktualisiert - ist die Datei gerade geoeffnet? Bitte schliessen und erneut starten:' -ForegroundColor Yellow
        $fehler | ForEach-Object { Write-Host "  $_" }
    }
    if ($env:OS -eq 'Windows_NT') { Invoke-Item -LiteralPath $ziel }
} catch {
    Write-Host "Fehler: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Pruefe die Internetverbindung und starte das Skript erneut.'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
