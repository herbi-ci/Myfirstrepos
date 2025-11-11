<#
MyVeeamReport.ps1
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

# -------------- Charger module Veeam si disponible --------------
try {
    if (Get-Module -ListAvailable -Name "Veeam.Backup.PowerShell") {
        Import-Module Veeam.Backup.PowerShell -ErrorAction Stop
        # Connect to local VBR server (silently)
        try { Connect-VBRServer -Server "localhost" -ErrorAction SilentlyContinue | Out-Null } catch {}
    } else {
        Write-Host "Module Veeam.Backup.PowerShell non trouvé dans cette session. Le script tentera d'utiliser les variables existantes." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Erreur import module Veeam: $($_.Exception.Message)" -ForegroundColor Yellow
}

# -------------- Fonctions Chart (.NET) --------------
Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

function Convert-HexToColor {
    param([string]$hex)
    $h = $hex.TrimStart('#')
    if ($h.Length -ne 6) { return [System.Drawing.Color]::Black }
    $r = [Convert]::ToInt32($h.Substring(0,2),16)
    $g = [Convert]::ToInt32($h.Substring(2,2),16)
    $b = [Convert]::ToInt32($h.Substring(4,2),16)
    return [System.Drawing.Color]::FromArgb($r,$g,$b)
}

function Save-DoughnutChartToBase64 {
    param(
        [double]$Percent,
        [string]$Title,
        [string]$Color = '#00B336',
        [int]$Width = 500,
        [int]$Height = 350
    )
    if ($Percent -lt 0) { $Percent = 0 }
    if ($Percent -gt 100) { $Percent = 100 }

    $tmp = [System.IO.Path]::Combine($env:TEMP, ('veeam_donut_{0}.png' -f ([Guid]::NewGuid().ToString())))
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.Width = $Width; $chart.Height = $Height
    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea "ca"
    $area.BackColor = "Transparent"
    $chart.ChartAreas.Add($area)

    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series "s"
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut
    $series.Points.AddXY("Succès",$Percent) | Out-Null
    $series.Points.AddXY("Échec",100 - $Percent) | Out-Null
    $series["PieLabelStyle"] = "Disabled"
    $chart.Series.Add($series)

    $chart.Legends.Clear()
    try {
        $chart.Series[0].Points[0].Color = Convert-HexToColor($Color)
        $chart.Series[0].Points[1].Color = [System.Drawing.Color]::FromArgb(233,236,239)
    } catch {}

    $t = New-Object System.Windows.Forms.DataVisualization.Charting.Title
    $t.Text = "$Title - $Percent`%"
    $t.Font = New-Object System.Drawing.Font("Arial",12)
    $chart.Titles.Add($t)

    $chart.SaveImage($tmp, "Png")
    $bytes = [System.IO.File]::ReadAllBytes($tmp)
    Remove-Item $tmp -ErrorAction SilentlyContinue
    $chart.Dispose()
    return [Convert]::ToBase64String($bytes)
}

# -------------- Récupération KPI : jobs (exclure inactifs et non schedulés) --------------
$sinceDays = if ($null -ne $JobsWindowDays) { [int]$JobsWindowDays } else { 30 }
$since = (Get-Date).AddDays(-1 * $sinceDays)

# Helper: compute success percentage from list of sessions
function Compute-SuccessPercent {
    param([array]$sessions)
    if (-not $sessions) { return 0 }
    $total = $sessions.Count
    # Count success / consider string equality
    $success = ($sessions | Where-Object { $_.Result -eq 'Success' }).Count
    # Compute percentage with rounding digit-by-digit
    $pct = 0
    if ($total -gt 0) {
        # digit-by-digit safe arithmetic
        $num = [decimal] $success
        $den = [decimal] $total
        $raw = ($num / $den) * 100
        $pct = [math]::Round($raw,1)
    }
    return $pct
}

# --- Backup jobs ---
try {
    $backupJobs = Get-VBRJob | Where-Object { $_.JobType -eq 'Backup' -and $_.IsActive -eq $true -and $_.IsScheduleEnabled -eq $true }
    $backupSessions = @()
    foreach ($j in $backupJobs) {
        $s = Get-VBRJobSession -Job $j -Filter @{ CreationTime = @{ From = $since } } -ErrorAction SilentlyContinue
        if ($s) { $backupSessions += $s }
    }
    $backupSuccessPct = Compute-SuccessPercent -sessions $backupSessions
} catch {
    Write-Host "Erreur récupération Backup jobs: $($_.Exception.Message)" -ForegroundColor Yellow
    $backupSuccessPct = 0
}

# --- Replica jobs ---
try {
    $replicaJobs = Get-VBRJob | Where-Object { $_.JobType -eq 'Replica' -and $_.IsActive -eq $true -and $_.IsScheduleEnabled -eq $true }
    $replicaSessions = @()
    foreach ($j in $replicaJobs) {
        $s = Get-VBRJobSession -Job $j -Filter @{ CreationTime = @{ From = $since } } -ErrorAction SilentlyContinue
        if ($s) { $replicaSessions += $s }
    }
    $replicaSuccessPct = Compute-SuccessPercent -sessions $replicaSessions
} catch {
    Write-Host "Erreur récupération Replica jobs: $($_.Exception.Message)" -ForegroundColor Yellow
    $replicaSuccessPct = 0
}

# --- Tape jobs ---
try {
    # tape jobs are separate object type
    $tapeJobs = Get-VBRTapeJob | Where-Object { $_.IsActive -eq $true -and $_.IsScheduleEnabled -eq $true } -ErrorAction SilentlyContinue
    $tapeSessions = @()
    foreach ($t in $tapeJobs) {
        $s = Get-VBRTapeJobSession -Job $t -Filter @{ CreationTime = @{ From = $since } } -ErrorAction SilentlyContinue
        if ($s) { $tapeSessions += $s }
    }
    $tapeSuccessPct = Compute-SuccessPercent -sessions $tapeSessions
} catch {
    Write-Host "Erreur récupération Tape jobs: $($_.Exception.Message)" -ForegroundColor Yellow
    $tapeSuccessPct = 0
}

# -------------- Repositories & Tape Pools (pour donuts par repo/pool) --------------
$repoStats = @()
try {
    $repos = Get-VBRBackupRepository -ErrorAction SilentlyContinue
    foreach ($r in $repos) {
        $cap = 0; $free = 0
        try { $cap = $r.Info.Capacity } catch { $cap = 0 }
        try { $free = $r.Info.FreeSpace } catch { $free = 0 }
        $usedPct = 0
        if ($cap -gt 0) {
            # Compute used percent digit-by-digit
            $usedPct = [math]::Round((([decimal]($cap - $free) / [decimal]$cap) * 100),1)
        }
        $repoStats += [PSCustomObject]@{ Name = $r.Name; UsedPercent = $usedPct; Capacity_TB = ([math]::Round($cap/1TB,2)); Free_TB = ([math]::Round($free/1TB,2)) }
    }
} catch { Write-Host "Erreur repos: $($_.Exception.Message)" -ForegroundColor Yellow }

$tapePoolStats = @()
try {
    $pools = Get-VBRTapeMediaPool -ErrorAction SilentlyContinue
    foreach ($p in $pools) {
        $medias = Get-VBRTapeMedium -MediaPool $p -ErrorAction SilentlyContinue
        $total = ($medias | Measure-Object).Count
        $used = ($medias | Where-Object { $_.IsExpired -ne $true }).Count
        $usedPct = if ($total -gt 0) { [math]::Round((([decimal]$used / [decimal]$total) * 100),1) } else { 0 }
        $tapePoolStats += [PSCustomObject]@{ PoolName = $p.Name; UsedPercent = $usedPct; TotalTapes = $total; Expired = ($medias | Where-Object { $_.IsExpired -eq $true }).Count }
    }
} catch { Write-Host "Erreur tape pools: $($_.Exception.Message)" -ForegroundColor Yellow }

# -------------- Donuts (global jobs + per repo/pool) --------------
# Colors: reuse your existing colors — adjust hex if you have exact ones; default Veeam green used here.
$colorSuccess = '#00B336'   # vert réussite (changer si tu as une autre couleur dans ton rapport)
$colorFail    = '#E53935'   # rouge échec

# Global job donuts
$img_backup = Save-DoughnutChartToBase64 -Percent $backupSuccessPct -Title "Backup Success" -Color $colorSuccess
$img_replica = Save-DoughnutChartToBase64 -Percent $replicaSuccessPct -Title "Replica Success" -Color $colorSuccess
$img_tape    = Save-DoughnutChartToBase64 -Percent $tapeSuccessPct -Title "Tape Success" -Color $colorSuccess

# Repo donuts (one per repo)
$repoImgMap = @()
foreach ($r in $repoStats) {
    $pct = [double]$r.UsedPercent
    # we want donut showing free vs used: percentUsed => color red if >=80
    $color = if ($pct -ge 80) { $colorFail } else { $colorSuccess }
    $img = Save-DoughnutChartToBase64 -Percent $pct -Title ("Repo: " + $r.Name) -Color $color
    $repoImgMap += [PSCustomObject]@{ Name = $r.Name; ImgBase64 = $img; UsedPercent = $pct; Capacity_TB = $r.Capacity_TB; Free_TB = $r.Free_TB }
}

# Tape pool donuts
$poolImgMap = @()
foreach ($p in $tapePoolStats) {
    $pct = [double]$p.UsedPercent
    $color = if ($pct -ge 80) { $colorFail } else { $colorSuccess }
    $img = Save-DoughnutChartToBase64 -Percent $pct -Title ("Pool: " + $p.PoolName) -Color $color
    $poolImgMap += [PSCustomObject]@{ Name = $p.PoolName; ImgBase64 = $img; UsedPercent = $pct; TotalTapes = $p.TotalTapes; Expired = $p.Expired }
}

# -------------- Tapes expiring (7 jours) --------------
$tapesExpiring = @()
try {
    $allTapes = Get-VBRTapeMedium -ErrorAction SilentlyContinue
    foreach ($m in $allTapes) {
        $exp = $null
        try { $exp = $m.ExpirationDate } catch { $exp = $null }
        if (-not $exp -and $m.MediaSet) { try { $exp = $m.MediaSet.ProtectionPeriodEnd } catch { $exp = $null } }
        if ($exp -and $exp -le (Get-Date).AddDays(7) -and $exp -ge (Get-Date)) {
            $tapesExpiring += [PSCustomObject]@{ Barcode = $m.Barcode; ExpirationDate = $exp.ToString("yyyy-MM-dd"); MediaPool = ($m.MediaPool.Name -as [string]); IsWorm = ($m.IsWorm -as [bool]) }
        }
    }
} catch { Write-Host "Erreur lecture bandes: $($_.Exception.Message)" -ForegroundColor Yellow }

# -------------- Composants inactifs (>3 mois) --------------
$inactiveJobsList = @()
$inactiveReposList = @()
$inactiveProxiesList = @()
try {
    $threeMonthsAgo = (Get-Date).AddMonths(-3)
    # jobs
    $allJobs = Get-VBRJob -ErrorAction SilentlyContinue
    foreach ($j in $allJobs) {
        $last = Get-VBRJobSession -Job $j -MaxCount 1 -ErrorAction SilentlyContinue
        $lastRun = if ($last) { $last.CreationTime } else { $null }
        if ($j.IsScheduleEnabled -eq $false -or (-not $lastRun) -or ($lastRun -lt $threeMonthsAgo)) {
            $inactiveJobsList += [PSCustomObject]@{ Name = $j.Name; IsScheduled = $j.IsScheduleEnabled; LastRun = if ($lastRun) { $lastRun.ToString("yyyy-MM-dd") } else { "Never" } }
        }
    }
    # repos
    foreach ($r in (Get-VBRBackupRepository -ErrorAction SilentlyContinue)) {
        $la = $null
        try { $la = $r.LastAccess } catch { $la = $null }
        if (($la -and $la -lt $threeMonthsAgo) -or ($r.Info.FreeSpace -ge $r.Info.Capacity)) {
            $inactiveReposList += [PSCustomObject]@{ Name = $r.Name; LastAccess = if ($la) { $la.ToString("yyyy-MM-dd") } else { "N/A" }; FreeGB = [math]::Round($r.Info.FreeSpace/1GB,2) }
        }
    }
    # proxies
    foreach ($p in (Get-VBRViProxy -ErrorAction SilentlyContinue)) {
        $enabled = $null; $state = $null
        try { $enabled = $p.IsEnabled } catch { $enabled = $null }
        try { $state = $p.State } catch { $state = $null }
        if (($enabled -eq $false) -or ($state -and $state -ne "Online")) {
            $inactiveProxiesList += [PSCustomObject]@{ Name = $p.Name; Enabled = $enabled; State = $state }
        }
    }
} catch { Write-Host "Erreur composants inactifs: $($_.Exception.Message)" -ForegroundColor Yellow }

# -------------- Génération HTML final --------------
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Build repo donuts HTML (tile grid)
$repoDonutsHtml = ""
if ($repoImgMap.Count -gt 0) {
    foreach ($ri in $repoImgMap) {
        $repoDonutsHtml += "<div style='display:inline-block;width:220px;margin:8px;text-align:center;padding:8px;background:#fff;border-radius:8px;border:1px solid #eee'>"
        $repoDonutsHtml += "<div style='font-weight:700;margin-bottom:6px;'>$($ri.Name)</div>"
        $repoDonutsHtml += "<img src='data:image/png;base64,$($ri.ImgBase64)' style='width:160px;height:auto' alt='repo'/>"
        $repoDonutsHtml += "<div style='margin-top:6px'>Utilisé: <b>$($ri.UsedPercent)%</b> — Libre: <b>$($ri.Free_TB) TB</b></div>"
        $repoDonutsHtml += "</div>"
    }
} else {
    $repoDonutsHtml = "<div style='padding:8px'>Aucun repository détecté</div>"
}

# Build pool donuts HTML
$poolDonutsHtml = ""
if ($poolImgMap.Count -gt 0) {
    foreach ($pi in $poolImgMap) {
        $poolDonutsHtml += "<div style='display:inline-block;width:220px;margin:8px;text-align:center;padding:8px;background:#fff;border-radius:8px;border:1px solid #eee'>"
        $poolDonutsHtml += "<div style='font-weight:700;margin-bottom:6px;'>$($pi.Name)</div>"
        $poolDonutsHtml += "<img src='data:image/png;base64,$($pi.ImgBase64)' style='width:160px;height:auto' alt='pool'/>"
        $poolDonutsHtml += "<div style='margin-top:6px'>Utilisé: <b>$($pi.UsedPercent)%</b> — Total bandes: <b>$($pi.TotalTapes)</b></div>"
        $poolDonutsHtml += "</div>"
    }
} else {
    $poolDonutsHtml = "<div style='padding:8px'>Aucun pool de bande détecté</div>"
}

# Compose final HTML (Style proche de ton rapport existant)
$html = @"
<!doctype html>
<html>
<head>
<meta charset='utf-8'>
<title>Veeam KPI Report – Infrastructure Backup Overview</title>
<style>
body { font-family: Arial, Helvetica, sans-serif; background:#f6f7f9; color:#222; padding:18px; }
.header { display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; }
.card { background:#fff; padding:14px; border-radius:8px; box-shadow:0 1px 6px rgba(0,0,0,0.06); margin-bottom:12px; }
.kpi-row { display:flex; gap:12px; flex-wrap:wrap; align-items:flex-start; }
.kpi-card { width:320px; }
.title { color:#0b3a2f; font-weight:700; margin-bottom:6px; }
.small { color:#666; font-size:13px; }
.table { width:100%; border-collapse:collapse; margin-top:8px; }
.table th, .table td { border:1px solid #eee; padding:6px; }
</style>
</head>
<body>
  <div class='header'>
    <div>
      <h1 style='margin:0'>Veeam KPI Report</h1>
      <div class='small'>Généré: $timestamp</div>
    </div>
    <div class='small'>Serveur: $(hostname)</div>
  </div>

  <div class='kpi-row'>
    <div class='card kpi-card'>
      <div class='title'>Sauvegardes (global)</div>
      <img src='data:image/png;base64,$img_backup' style='width:100%;height:auto' alt='backup'/>
      <div class='small' style='margin-top:8px'>Taux de réussite global sur $sinceDays jours: <b style='color:$colorSuccess'>$backupSuccessPct`%</b></div>
    </div>

    <div class='card kpi-card'>
      <div class='title'>Réplications (global)</div>
      <img src='data:image/png;base64,$img_replica' style='width:100%;height:auto' alt='replica'/>
      <div class='small' style='margin-top:8px'>Taux de réussite global sur $sinceDays jours: <b style='color:$colorSuccess'>$replicaSuccessPct`%</b></div>
    </div>

    <div class='card kpi-card'>
      <div class='title'>Bandes (Tape) (global)</div>
      <img src='data:image/png;base64,$img_tape' style='width:100%;height:auto' alt='tape'/>
      <div class='small' style='margin-top:8px'>Taux de réussite global sur $sinceDays jours: <b style='color:$colorSuccess'>$tapeSuccessPct`%</b></div>
    </div>

    <div class='card' style='flex:1'>
      <div class='title'>Repositories (utilisation)</div>
      <div style='display:block;margin-top:8px'>$repoDonutsHtml</div>
    </div>
  </div>

  <div class='card'>
    <div class='title'>Tape Pools (utilisation)</div>
    <div style='display:block;margin-top:8px'>$poolDonutsHtml</div>
  </div>

  <div class='card'>
    <div class='title'>Bandes expirant dans 7 jours</div>
    <div style='margin-top:8px'>
"@

# Tapes expiring table
if ($tapesExpiring.Count -gt 0) {
    $html += "<table class='table'><tr><th>Barcode</th><th>Expiration</th><th>Pool</th><th>WORM</th></tr>"
    foreach ($t in $tapesExpiring) {
        $html += "<tr><td>$($t.Barcode)</td><td>$($t.ExpirationDate)</td><td>$($t.MediaPool)</td><td>$($t.IsWorm)</td></tr>"
    }
    $html += "</table>"
} else {
    $html += "<div class='small'>Aucune bande n'expire dans les 7 prochains jours.</div>"
}

$html += "</div></div>"

# Inactive components
$html += "<div class='card'><div class='title'>Composants inactifs (&gt;3 mois)</div>"
$html += "<div style='margin-top:8px'>"
$html += "<div style='display:inline-block;padding:10px;border:1px solid #eee;background:#fff;border-radius:6px;margin-right:8px'>Jobs inactifs: <b>$($inactiveJobsList.Count)</b></div>"
$html += "<div style='display:inline-block;padding:10px;border:1px solid #eee;background:#fff;border-radius:6px;margin-right:8px'>Repos inactifs: <b>$($inactiveReposList.Count)</b></div>"
$html += "<div style='display:inline-block;padding:10px;border:1px solid #eee;background:#fff;border-radius:6px'>Proxies inactifs: <b>$($inactiveProxiesList.Count)</b></div>"
$html += "</div></div>"

$html += "<div class='card small'>Note: Données calculées sur les $sinceDays derniers jours. Jobs inactifs et non schedulés sont exclus du calcul des taux.</div>"
$html += "</body></html>"

# -------------- Enregistrer le rapport --------------
$dir = Split-Path -Path $ReportPath -Parent
if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
Set-Content -Path $ReportPath -Value $html -Encoding UTF8
Write-Host "Rapport généré: $ReportPath" -ForegroundColor Green

# -------------- Fin --------------
