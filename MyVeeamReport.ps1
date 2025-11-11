<#
MyVeeamReport_Refactored.ps1
Génère un rapport HTML avec donuts (global jobs + repositories + tape pools)
Compatible Veeam 12.3.x - Exécuter dans PowerShell ISE (Admin recommandé)
#>

# -------------- Charger variables --------------
try {
    $varFile = Join-Path $PSScriptRoot "MyVeeamReport_Variable.ps1"
    if (Test-Path $varFile) { . $varFile } else { Write-Host "Warning: variables file not found: $varFile" -ForegroundColor Yellow }
} catch {
    Write-Host "Impossible de charger MyVeeamReport_Variable.ps1 : $($_.Exception.Message)" -ForegroundColor Red
}

# -------------- Charger module Veeam --------------
try {
    if (Get-Module -ListAvailable -Name "Veeam.Backup.PowerShell") {
        Import-Module Veeam.Backup.PowerShell -ErrorAction Stop
        try { Connect-VBRServer -Server "localhost" -ErrorAction SilentlyContinue | Out-Null } catch {}
    } else {
        Write-Host "Module Veeam.Backup.PowerShell non trouvé. Le script utilisera les variables existantes." -ForegroundColor Yellow
    }
} catch { Write-Host "Erreur import module Veeam: $($_.Exception.Message)" -ForegroundColor Yellow }

# -------------- Fonctions utilitaires --------------
Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

function Convert-HexToColor { param([string]$hex)
    $h = $hex.TrimStart('#'); if ($h.Length -ne 6) { return [System.Drawing.Color]::Black }
    return [System.Drawing.Color]::FromArgb([Convert]::ToInt32($h.Substring(0,2),16),
                                           [Convert]::ToInt32($h.Substring(2,2),16),
                                           [Convert]::ToInt32($h.Substring(4,2),16))
}

function Save-DoughnutChartToBase64 {
    param([double]$Percent, [string]$Title, [string]$Color='#00B336', [int]$Width=500, [int]$Height=350)
    $Percent = [math]::Min([math]::Max($Percent,0),100)
    $tmp = [System.IO.Path]::Combine($env:TEMP, "veeam_donut_$([Guid]::NewGuid()).png")
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.Width = $Width; $chart.Height = $Height
    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea "ca"; $area.BackColor="Transparent"; $chart.ChartAreas.Add($area)
    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series "s"
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut
    $series.Points.AddXY("Succès",$Percent) | Out-Null
    $series.Points.AddXY("Échec",100-$Percent) | Out-Null
    $series["PieLabelStyle"]="Disabled"; $chart.Series.Add($series)
    $chart.Legends.Clear()
    try { $chart.Series[0].Points[0].Color = Convert-HexToColor($Color); $chart.Series[0].Points[1].Color=[System.Drawing.Color]::FromArgb(233,236,239) } catch {}
    $t = New-Object System.Windows.Forms.DataVisualization.Charting.Title
    $t.Text = "$Title - $Percent`%"; $t.Font = New-Object System.Drawing.Font("Arial",12)
    $chart.Titles.Add($t)
    $chart.SaveImage($tmp,"Png")
    $bytes = [System.IO.File]::ReadAllBytes($tmp); Remove-Item $tmp -ErrorAction SilentlyContinue; $chart.Dispose()
    return [Convert]::ToBase64String($bytes)
}

function Compute-SuccessPercent { param([array]$sessions)
    if (-not $sessions) { return 0 }
    $total = $sessions.Count
    $success = ($sessions | Where-Object { $_.Result -eq 'Success' }).Count
    return if ($total -gt 0) { [math]::Round((([decimal]$success / [decimal]$total) * 100),1) } else { 0 }
}

function Get-JobSuccessPercent { param([string]$JobType)
    $sinceDays = if ($null -ne $JobsWindowDays) { [int]$JobsWindowDays } else { 30 }
    $since = (Get-Date).AddDays(-1*$sinceDays)
    $jobs = if ($JobType -eq "Tape") { Get-VBRTapeJob -ErrorAction SilentlyContinue } else { Get-VBRJob | Where-Object { $_.JobType -eq $JobType -and $_.IsActive -eq $true -and $_.IsScheduleEnabled -eq $true } }
    $sessions=@()
    foreach ($j in $jobs) {
        try {
            $s = if ($JobType -eq "Tape") { Get-VBRTapeJobSession -Job $j -Filter @{ CreationTime=@{From=$since} } -ErrorAction SilentlyContinue } else { Get-VBRJobSession -Job $j -Filter @{ CreationTime=@{From=$since} } -ErrorAction SilentlyContinue }
            if ($s) { $sessions += $s }
        } catch {}
    }
    return Compute-SuccessPercent -sessions $sessions
}

# -------------- KPI jobs --------------
$backupSuccessPct  = Get-JobSuccessPercent -JobType "Backup"
$replicaSuccessPct = Get-JobSuccessPercent -JobType "Replica"
$tapeSuccessPct    = Get-JobSuccessPercent -JobType "Tape"

# -------------- Repositories & Tape Pools --------------
$repoStats=@()
try {
    foreach ($r in Get-VBRBackupRepository -ErrorAction SilentlyContinue) {
        $cap=0; $free=0
        try { $cap=$r.Info.Capacity } catch {}; try { $free=$r.Info.FreeSpace } catch {}
        $usedPct = if ($cap -gt 0) { [math]::Round((([decimal]($cap-$free)/[decimal]$cap)*100),1) } else { 0 }
        $repoStats += [PSCustomObject]@{ Name=$r.Name; UsedPercent=$usedPct; Capacity_TB=[math]::Round($cap/1TB,2); Free_TB=[math]::Round($free/1TB,2) }
    }
} catch { Write-Host "Erreur repos: $($_.Exception.Message)" -ForegroundColor Yellow }

$poolStats=@()
try {
    foreach ($p in Get-VBRTapeMediaPool -ErrorAction SilentlyContinue) {
        $medias=Get-VBRTapeMedium -MediaPool $p -ErrorAction SilentlyContinue
        $total = ($medias|Measure-Object).Count
        $used = ($medias|Where-Object { $_.IsExpired -ne $true }).Count
        $usedPct = if ($total -gt 0) { [math]::Round((([decimal]$used/[decimal]$total)*100),1) } else { 0 }
        $poolStats += [PSCustomObject]@{ Name=$p.Name; UsedPercent=$usedPct; TotalTapes=$total; Expired=($medias|Where-Object { $_.IsExpired -eq $true }).Count }
    }
} catch { Write-Host "Erreur tape pools: $($_.Exception.Message)" -ForegroundColor Yellow }

# -------------- Donuts --------------
$colorSuccess='#00B336'; $colorFail='#E53935'

function Generate-DonutHtml { param($Items, [string]$Type)
    $html=""
    foreach ($i in $Items) {
        $pct=[double]$i.UsedPercent
        $color = if ($pct -ge 80) { $colorFail } else { $colorSuccess }
        $img = Save-DoughnutChartToBase64 -Percent $pct -Title ("$Type: " + ($i.Name -as [string])) -Color $color
        $html += "<div style='display:inline-block;width:220px;margin:8px;text-align:center;padding:8px;background:#fff;border-radius:8px;border:1px solid #eee'>"
        $html += "<div style='font-weight:700;margin-bottom:6px;'>$($i.Name)</div>"
        $html += "<img src='data:image/png;base64,$img' style='width:160px;height:auto' alt='$Type'/>"
        if ($Type -eq "Repo") { $html += "<div style='margin-top:6px'>Utilisé: <b>$pct`%</b> — Libre: <b>$($i.Free_TB) TB</b></div>" }
        if ($Type -eq "Pool") { $html += "<div style='margin-top:6px'>Utilisé: <b>$pct`%</b> — Total bandes: <b>$($i.TotalTapes)</b></div>" }
        $html += "</div>"
    }
    if ($html -eq "") { $html="<div style='padding:8px'>Aucun $Type détecté</div>" }
    return $html
}

$repoDonutsHtml = Generate-DonutHtml -Items $repoStats -Type "Repo"
$poolDonutsHtml = Generate-DonutHtml -Items $poolStats -Type "Pool"

# -------------- Tapes expirant 7 jours --------------
$tapesExpiring=@()
try {
    $allTapes=Get-VBRTapeMedium -ErrorAction SilentlyContinue
    foreach ($m in $allTapes) {
        $exp=$null
        try { $exp=$m.ExpirationDate } catch {}
        if (-not $exp -and $m.MediaSet) { try { $exp=$m.MediaSet.ProtectionPeriodEnd } catch {} }
        if ($exp -and $exp -le (Get-Date).AddDays(7) -and $exp -ge (Get-Date)) {
            $tapesExpiring += [PSCustomObject]@{ Barcode=$m.Barcode; ExpirationDate=$exp.ToString("yyyy-MM-dd"); MediaPool=($m.MediaPool.Name -as [string]); IsWorm=($m.IsWorm -as [bool]) }
        }
    }
    $tapesExpiring = $tapesExpiring | Sort-Object ExpirationDate
} catch { Write-Host "Erreur lecture bandes: $($_.Exception.Message)" -ForegroundColor Yellow }

# -------------- Composants inactifs --------------
$inactiveJobsList=@(); $inactiveReposList=@(); $inactiveProxiesList=@()
try {
    $threeMonthsAgo=(Get-Date).AddMonths(-3)
    foreach ($j in Get-VBRJob -ErrorAction SilentlyContinue) {
        $last = Get-VBRJobSession -Job $j -MaxCount 1 -ErrorAction SilentlyContinue
        $lastRun = if ($last) { $last.CreationTime } else { $null }
        if ($j.IsScheduleEnabled -eq $false -or (-not $lastRun) -or ($lastRun -lt $threeMonthsAgo)) {
            $inactiveJobsList += [PSCustomObject]@{ Name=$j.Name; IsScheduled=$j.IsScheduleEnabled; LastRun=if($lastRun){$lastRun.ToString("yyyy-MM-dd")}else{"Never"} }
        }
    }
    foreach ($r in Get-VBRBackupRepository -ErrorAction SilentlyContinue) {
        $la=$null; try{$la=$r.LastAccess}catch{}
        if (($la -and $la -lt $threeMonthsAgo)-or($r.Info.FreeSpace -ge $r.Info.Capacity)) {
            $inactiveReposList+=[PSCustomObject]@{ Name=$r.Name; LastAccess=if($la){$la.ToString("yyyy-MM-dd")}else{"N/A"}; FreeGB=[math]::Round($r.Info.FreeSpace/1GB,2) }
        }
    }
    foreach ($p in Get-VBRViProxy -ErrorAction SilentlyContinue) {
        $enabled=$null;$state=$null; try{$enabled=$p.IsEnabled}catch{}; try{$state=$p.State}catch{}
        if (($enabled -eq $false)-or($state -and $state -ne "Online")) { $inactiveProxiesList+=[PSCustomObject]@{ Name=$p.Name; Enabled=$enabled; State=$state } }
    }
} catch { Write-Host "Erreur composants inactifs: $($_.Exception.Message)" -ForegroundColor Yellow }

# -------------- Génération HTML --------------
$timestamp=Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$html=@"
<!doctype html>
<html>
<head>
<meta charset='utf-8'>
<title>Veeam KPI Report – Infrastructure Backup Overview</title>
<style>
body{font-family:Arial,Helvetica,sans-serif;background:#f6f7f9;color:#222;padding:18px;}
.header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;}
.card{background:#fff;padding:14px;border-radius:8px;box-shadow:0 1px 6px rgba(0,0,0,0.06);margin-bottom:12px;}
.kpi-row{display:flex;gap:12px;flex-wrap:wrap;align-items:flex-start;}
.kpi-card{width:320px;}
.title{color:#0b3a2f;font-weight:700;margin-bottom:6px;}
.small{color:#666;font-size:13px;}
.table{width:100%;border-collapse:collapse;margin-top:8px;}
.table th,.table td{border:1px solid #eee;padding:6px;}
</style>
</head>
<body>
<div class='header'>
<div><h1 style='margin:0'>Veeam KPI Report</h1><div class='small'>Généré: $timestamp</div></div>
<div class='small'>Serveur: $(hostname)</div>
</div>
<div class='kpi-row'>
<div class='card kpi-card'><div class='title'>Sauvegardes (global)</div><img src='data:image/png;base64,$(Save-DoughnutChartToBase64 -Percent $backupSuccessPct -Title "Backup Success" -Color $colorSuccess)' style='width:100%;height:auto' alt='backup'/><div class='small' style='margin-top:8px'>Taux global $backupSuccessPct`%</div></div>
<div class='card kpi-card'><div class='title'>Réplications (global)</div><img src='data:image/png;base64,$(Save-DoughnutChartToBase64 -Percent $replicaSuccessPct -Title "Replica Success" -Color $colorSuccess)' style='width:100%;height:auto' alt='replica'/><div class='small' style='margin-top:8px'>Taux global $replicaSuccessPct`%</div></div>
<div class='card kpi-card'><div class='title'>Bandes (Tape) (global)</div><img src='data:image/png;base64,$(Save-DoughnutChartToBase64 -Percent $tapeSuccessPct -Title "Tape Success" -Color $colorSuccess)' style='width:100%;height:auto' alt='tape'/><div class='small' style='margin-top:8px'>Taux global $tapeSuccessPct`%</div></div>
<div class='card' style='flex:1'><div class='title'>Repositories (utilisation)</div><div style='display:block;margin-top:8px'>$repoDonutsHtml</div></div>
</div>
<div class='card'><div class='title'>Tape Pools (utilisation)</div><div style='display:block;margin-top:8px'>$poolDonutsHtml</div></div>
<div class='card'><div class='title'>Bandes expirant dans 7 jours</div><div style='margin-top:8px'>
"@

if ($tapesExpiring.Count -gt 0) {
    $html+="<table class='table'><tr><th>Barcode</th><th>Expiration</th><th>Pool</th><th>WORM</th></tr>"
    foreach ($t in $tapesExpiring) { $html+="<tr><td>$($t.Barcode)</td><td>$($t.ExpirationDate)</td><td>$($t.MediaPool)</td><td>$($t.IsWorm)</td></tr>" }
    $html+="</table>"
} else { $html+="<div class='small'>Aucune bande n'expire dans les 7 prochains jours.</div>" }

$html+="</div></div>"
$html+="<div class='card'><div class='title'>Composants inactifs (&gt;3 mois)</div><div style='margin-top:8px'>"
$html+="<div style='display:inline-block;padding:10px;border:1px solid #eee;background:#fff;border-radius:6px;margin-right:8px'>Jobs inactifs: <b>$($inactiveJobsList.Count)</b></div>"
$html+="<div style='display:inline-block;padding:10px;border:1px solid #eee;background:#fff;border-radius:6px;margin-right:8px'>Repos inactifs: <b>$($inactiveReposList.Count)</b></div>"
$html+="<div style='display:inline-block;padding:10px;border:1px solid #eee;background:#fff;border-radius:6px'>Proxies inactifs: <b>$($inactiveProxiesList.Count)</b></div>"
$html+="</div></div>"
$html+="<div class='card small'>Note: Données calculées sur les $sinceDays derniers jours. Jobs inactifs et non schedulés sont exclus du calcul des taux.</div>"
$html+="</body></html>"

# -------------- Enregistrer rapport --------------
$dir = Split-Path -Path $ReportPath -Parent; if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
Set-Content -Path $ReportPath -Value $html -Encoding UTF8
Write-Host "Rapport généré: $ReportPath" -ForegroundColor Green
