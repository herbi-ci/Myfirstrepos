#VEEAM INFRASTRUCTURE KPI REPORT
<#

    .SYNOPSIS
    Show dashboard summury to have an overview on veeam infrastructure health .

    .DESCRIPTION
    Just change the needed patameters and run the script.

    .NOTES
    Authors: Jean-hermann BILE
    Contributors: Cedrick KOUAKOU
    Made by: SGABS IaaS Team
    Last Updated: 13 Febuary 2026
    Script version: 1.0
    Veeam version: 12.3.2.3617
 
    
#> 

# =================================================
# PARAMÈTRES
# =================================================
## Nombres de jour à analyser
$nbjour =  
## Paramèrtres à modifier selon la filiale
$filiale = 
$environnement = 
$smtpServer = 
$smtpPort   = 
$mailFrom   = 
$recipient = 
## Paramètres qui ne change pas
$output = "C:\Temp\Veeam_KPI_Dashboard.html"
$since  = (Get-Date).AddDays(-$nbjour)
$rptTitle = "$filiale $environnement Veeam Infrastructure KPI"
$rptPeriod = "Période analysée : $($since.ToString('dd-MM-yyyy')) → $(Get-Date -Format 'dd-MM-yyyy')"
$reportDate = Get-Date -Format "dd/MM/yyyy HH:mm"
$mailTo     = $recipient -split "," | ForEach-Object { ($_.Trim()) }
$mailSubject = "$filiale $environnement Veeam KPI"

# ===============================================
# 2. FONCTIONS UTILITAIRES (OUTILS)
# ===============================================
# Ces fonctions sont utilisées plus loin pour mettre en forme les données.

# =================================================
# FONCTION REPOS
# =================================================
Function Get-VBRRepoInfo {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, ValueFromPipeline=$true)]
        [PSObject[]]$Repository
    )
    Begin {
        $outputAry = @()

        Function Build-Object {
            param(
                $name, $repohost, $path, $free, $total, $rtype, $rBackupsize
            )

            # Convert to TB
            $freeTB   = [Math]::Round([Decimal]$free / 1TB, 2)
            $totalTB  = [Math]::Round([Decimal]$total / 1TB, 2)
            $backupTB = [Math]::Round([Decimal]$rBackupsize / 1TB, 2)

            $repoObj = [PSCustomObject]@{
                Target          = $name
                RepoHost        = $repohost
                StorePath       = $path
                StorageFreeTB   = $freeTB
                StorageTotalTB  = $totalTB
                FreePercentage  = if ($total -gt 0) { [Math]::Round(($free / $total) * 100, 2) } else { 0 }
                StorageBackupTB = $backupTB
                rType           = $rtype
            }
            return $repoObj
        }
    }
    Process {
        foreach ($r in $Repository) {
            # Refresh Repository Size Info
            [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)

            $rType = switch ($r.Type) {
                "WinLocal"      { "Windows Local" }
                "LinuxLocal"    { "Linux Local" }
                "LinuxHardened" { "Hardened" }
                "CifsShare"     { "CIFS Share" }
                "DataDomain"    { "Data Domain" }
                "ExaGrid"       { "ExaGrid" }
                "HPStoreOnce"   { "HP StoreOnce" }
                "Nfs"           { "NFS Direct" }
                default         { "Unknown" }
            }

            # Calculate total backup size in this repository
            $rBackupsize = 0
            $backups = Get-VBRBackup | Where-Object { $_.RepositoryId -eq $r.Id }
            foreach ($b in $backups) {
                foreach ($storage in $b.GetAllStorages()) {
                    $rBackupsize += $storage.Stats.BackupSize
                }
            }

            $outputObj = Build-Object `
                $r.Name `
                $($r.GetHost()).Name.ToLower() `
                $r.Path `
                $r.GetContainer().CachedFreeSpace.InBytes `
                $r.GetContainer().CachedTotalSpace.InBytes `
                $rType `
                $rBackupsize

            $outputAry += $outputObj
        }
    }
    End {
        $outputAry
    }
}

# =================================================
# FONCTION DONUT JOBS
# =================================================
function Get-Donut {
param($percent,$label)
$color = if ($percent -ge 80) { "#00B336" } else { "#E31E24" }
$offset = 440 - (440 * $percent / 100)

@"
<svg width="160" height="160">
 <circle cx="80" cy="80" r="70" stroke="#eee" stroke-width="15" fill="none"/>
 <circle cx="80" cy="80" r="70" stroke="$color" stroke-width="15"
  fill="none" stroke-dasharray="440" stroke-dashoffset="$offset"
  transform="rotate(-90 80 80)"/>
 <text x="50%" y="50%" text-anchor="middle" dy=".3em"
  font-size="22" font-weight="bold">$label%</text>
</svg>
"@
}

# =================================================
# FONCTION DONUT REPOS (Used %)
# =================================================
function Get-Bar {
param([double]$usedPercent)
$color = if ($usedPercent -ge 70) { "#E31E24" } else { "#00B336" }
$width  = [math]::Round(($usedPercent / 100) * 160)
@"
<svg width="160" height="20">
 <rect width="160" height="20" fill='#eee' rx='5' ry='5'/>
 <rect width="$width" height="20" fill="$color" rx='5' ry='5'/>
 <text x='50%' y='14' text-anchor='middle' alignment-baseline='middle'
       font-size='12' font-weight='bold' fill='#000'>$usedPercent`%</text>
</svg>
"@
}

# ==============================================================================
# RÉCUPÉRATION ET CALCUL DES STATISTIQUES
# ==============================================================================

# =================================================
# BACKUP
# =================================================
##### On calcul les stats de chaque Job
$backupStats = Get-VBRBackupSession |
Where-Object { $_.CreationTime -ge $since -and $_.JobType -eq "Backup"} |
Group-Object JobName |
ForEach-Object {
    $success = ($_.Group | Where-Object Result -eq "Success").Count
    $failed  = ($_.Group | Where-Object Result -eq "Failed").Count
    $warning = ($_.Group | Where-Object Result -eq "Warning").Count
    $total   = $_.Count
    $lastRun = ($_.Group | Sort-Object CreationTime -Descending | Select-Object -First 1).CreationTime
    [PSCustomObject]@{
        Type     = "Backup"
        JobName  = $_.Name
        Success  = $success
        Warning  = $warning
        Failed   = $failed
        Total    = $total
        Percent  = if ($total -gt 0) { [math]::Round((($success + $warning) / $total) * 100,1) } else { 0 }
        LastRun  = $lastRun.ToString("dd/MM/yyyy HH:mm") 
    }
}

##### On calcul le porcetange global pour générer le graphe

$bkSuccess = ($backupStats | Measure-Object Success -Sum).Sum
$bkFailed  = ($backupStats | Measure-Object Failed  -Sum).Sum
$bkWarning = ($backupStats | Measure-Object Warning -Sum).Sum
$bkTotal   = $bkSuccess + $bkFailed + $bkWarning
$bkPercent = if ($bkTotal -gt 0) { [math]::Round((($bkSuccess + $bkWarning) / $bkTotal) * 100,1) } else { 0 }

# =================================================
# REPLICATION
# =================================================

##### On calcul les stats de chaque Job
$replicaStats = Get-VBRBackupSession |
Where-Object { $_.CreationTime -ge $since -and $_.JobType -eq "Replica"} |
Group-Object JobName |
ForEach-Object {
    $success = ($_.Group | Where-Object Result -eq "Success").Count
    $failed  = ($_.Group | Where-Object Result -eq "Failed").Count
    $warning = ($_.Group | Where-Object Result -eq "Warning").Count
    $total   = $_.Count
    $lastRun = ($_.Group | Sort-Object CreationTime -Descending | Select-Object -First 1).CreationTime
    [PSCustomObject]@{
        Type     = "Replica"
        JobName  = $_.Name
        Success  = $success
        Warning  = $warning
        Failed   = $failed
        Total    = $total
        Percent  = if ($total -gt 0) { [math]::Round((($success + $warning) / $total) * 100,1) } else { 0 }
        LastRun  = $lastRun.ToString("dd/MM/yyyy HH:mm")
    }
}

##### On calcul le porcetange global pour générer le graphe

$rpSuccess = ($replicaStats | Measure-Object Success -Sum).Sum
$rpFailed  = ($replicaStats | Measure-Object Failed  -Sum).Sum
$rpWarning = ($replicaStats | Measure-Object Warning -Sum).Sum
$rpTotal   = $rpSuccess + $rpFailed + $rpWarning
$rpPercent = if ($rpTotal -gt 0) { [math]::Round((($rpSuccess + $rpWarning) / $rpTotal) * 100,1) } else { 0 }

# =================================================
# TAPE BACKUP
# =================================================
##### On calcul les stats de chaque Job
$tapeSessions = Get-VBRTapeJob | ForEach-Object { Get-VBRTapeBackupSession -Job $_ } |
Where-Object { $_.CreationTime -ge $since }

$tapeStats = $tapeSessions |
Group-Object Name |
ForEach-Object {
    $success = ($_.Group | Where-Object Result -eq "Success").Count
    $failed  = ($_.Group | Where-Object Result -eq "Failed").Count
    $warning = ($_.Group | Where-Object Result -eq "Warning").Count
    $total   = $_.Count
    $lastRun = ($_.Group | Sort-Object CreationTime -Descending | Select-Object -First 1).CreationTime
    [PSCustomObject]@{
        Type     = "Tape"
        JobName  = $_.Name
        Success  = $success
        Warning  = $warning
        Failed   = $failed
        Total    = $total
        Percent  = if ($total -gt 0) { [math]::Round((($success + $warning) / $total) * 100,1) } else { 0 }
        LastRun  = $lastRun.ToString("dd/MM/yyyy HH:mm")
    }
}

##### On calcul le porcetange global pour générer le graphe

$tpSuccess = ($tapeStats | Measure-Object Success -Sum).Sum
$tpFailed  = ($tapeStats | Measure-Object Failed  -Sum).Sum
$tpWarning = ($tapeStats | Measure-Object Warning -Sum).Sum
$tpTotal   = $tpSuccess + $tpFailed + $tpWarning
$tpPercent = if ($tpTotal -gt 0) { [math]::Round((($tpSuccess + $tpWarning) / $tpTotal) * 100,1) } else { 0 }


# =================================================
# UTILISATION DETAILLEE DES TAPE MEDIA POOL
# =================================================
## On calcul le taux d'utilisation de chaque pool

# Get all tape media pools
$mediaPools = Get-VBRTapeMediaPool

# Prepare results
$results = @()

foreach ($pool in $mediaPools) {
    # Get only ONLINE tapes in this pool (Location = Slot or Drive)
    $onlineTapes = Get-VBRTapeMedium | Where-Object { $_.MediaPoolId -eq $pool.Id -and ($_.Location -match 'Slot' -or $_.Location -match 'Drive' ) }

    if (-not $onlineTapes -or $onlineTapes.Count -eq 0) {
        $results += [PSCustomObject]@{
            PoolName        = $pool.Name
            TotalCapacityTB = 0
            UsedCapacityTB  = 0
            FreeCapacityTB  = 0
            OnlineCount     = 0
        }
        continue
    }

    # Calculate usage in TB (only online tapes)
    $totalCapacity = ($onlineTapes | Measure-Object -Property Capacity -Sum).Sum / 1TB
    $freeCapacity  = ($onlineTapes | Measure-Object -Property Free -Sum).Sum / 1TB
    $usedCapacity  = $totalCapacity - $freeCapacity

    $results += [PSCustomObject]@{
        PoolName        = $pool.Name
        TotalCapacityTB = [math]::Round($totalCapacity, 2)
        UsedCapacityTB  = [math]::Round($usedCapacity, 2)
        FreeCapacityTB  = [math]::Round($freeCapacity, 2)
        OnlineCount     = $onlineTapes.Count
        FreePercentage  = if ($totalCapacity -gt 0) { [Math]::Round(($freeCapacity / $totalCapacity) * 100, 2) } else { 0 }
        UsedPercentage  = if ($totalCapacity -gt 0) { [Math]::Round(($UsedCapacity / $totalCapacity) * 100, 2) } else { 0 }
    }
}


# =================================================
# REPOSITORIES
# =================================================
$repoStats = Get-VBRBackupRepository | Get-VBRRepoInfo

foreach ($r in $repoStats) {
    $r | Add-Member -MemberType NoteProperty -Name UsedPercent -Value ([math]::Round(100 - $r.FreePercentage,2))
}

# =================================================
# LISTE DES TAPES EXPIRE
# =================================================
$expiredTapes = Get-VBRTapeMedium | Where-Object { ($_.Location -match 'Slot' -or $_.Location -match 'Drive' )  -and
    $_.ExpirationDate -and $_.ExpirationDate -lt (Get-Date)
}
$expiredTapesTable = $expiredTapes | ForEach-Object {
    [PSCustomObject]@{
        TapeNumber          = $_.Name
        TapeExpirationDate  = $_.ExpirationDate.ToString('dd/MM/yyyy')
     }
}


# =================================================
# GÉNÉRATION DU RAPPORT HTML
# =================================================

# =================================================
# HTML HEADER
# =================================================
$bodyTop = @"
<body>
<table style="width:100%;border-collapse:collapse;margin-bottom:30px;">
    <tr>
        <td style="width:70%;
                   background-color:#293C52;
                   color:white;
                   font-size:22px;
                   font-weight:bold;
                   padding:15px;
                   text-align:left;">
            $rptTitle
        </td>
        <td style="width:30%;
                   background-color:#293C52;
                   color:white;
                   font-size:12px;
                   padding:15px;
                   text-align:right;
                   vertical-align:top;">
            <strong>$reportDate</strong>
        </td>
    </tr>
    <tr>
        <td colspan="2"
            style="background-color:#293C52;
                   color:white;
                   font-size:12px;
                   padding:6px 15px;
                   text-align:left;">
            $rptPeriod
        </td>
    </tr>
</table>
"@



# =================================================
# AFFICHAGE DE GRAPHES
# =================================================
$html = @"
<html>
<head>
<title>$filiale $environnement VEEAM KPI</title>
<style>
body { font-family:Segoe UI; background:#f4f6f8 }
.container { display:flex; justify-content:center; gap:40px; flex-wrap:wrap }
.card { background:#fff; padding:20px; border-radius:10px;
        box-shadow:0 0 10px rgba(0,0,0,0.1); text-align:center; margin-bottom:20px }
table { width:90%; margin:auto; border-collapse:collapse; margin-top:20px }
th,td { padding:10px; border-bottom:1px solid #ddd }
th { background:#2c3e50; color:white }
.green { color:#00B336; font-weight:bold }
.red { color:#E31E24; font-weight:bold }
.orange { color:#FF7F00; font-weight:bold }
</style>
</head>
<body>

$bodyTop

<div class="container">
 <div class="card">
  <h3>Backup Success</h3>
  $(Get-Donut $bkPercent $bkPercent)
 </div>

 <div class="card">
  <h3>Replication Success</h3>
  $(Get-Donut $rpPercent $rpPercent)
 </div>

 <div class="card">
  <h3>Tape Backup Success</h3>
  $(Get-Donut $tpPercent $tpPercent)
 </div>
</div>

<h2 style="text-align:center;margin-top:40px">Utilisation des Repositories</h2>
<div class="container">
"@

foreach ($r in $repoStats) {
    $html += @"
 <div class='card'>
  <h4>$($r.Name)</h4>
  $(Get-Bar $($r.UsedPercent))
  <p>$($r.Target) </p>
 </div>
"@
}
$html += "</div>"

$html+= @"
<h2 style="text-align:center;margin-top:40px">Utilisation des Tapes Media Pool</h2>
<div class="container">
"@
$results = $results | Where-Object { $_.PoolName -notlike 'Imported' -and $_.PoolName -notlike 'Free' -and $_.PoolName -notlike 'Retired' -and $_.PoolName -notlike 'Unrecognized' }
foreach ($r in $results) {
    $html += @"
 <div class='card'>
  <h4>$($r.Name)</h4>
  $(Get-Bar $($r.UsedPercentage))
  <p>$($r.PoolName) </p>
 </div>
"@
}
$html += "</div>"




# =================================================
# AFFICHAGE DES TABLEAUX
# =================================================

# ================================================
# TABLEAU DÉTAILLÉ DES REPOS ET TAPE MEDIA POOL
# =================================================

# Tableau détaillé Repos
$html += "<h2 style='text-align:center;margin-top:40px'>Détails Repositories</h2>
<table>
<tr><th>Repository</th><th>Host</th><th>Path</th><th>Total TB</th><th>Free TB</th><th>Used %</th><th>BackupSize TB</th><th>Type</th></tr>"
foreach ($r in $repoStats) {
    $html += "<tr><td>$($r.Target)</td><td>$($r.RepoHost)</td><td>$($r.StorePath)</td><td>$($r.StorageTotalTB)</td><td>$($r.StorageFreeTB)</td><td>$($r.UsedPercent)%</td><td>$($r.StorageBackupTB)</td><td>$($r.rType)</td></tr>"
}
$html += "</table>"

# Tableau détaillé Tape Media Pool
$html += "<h2 style='text-align:center;margin-top:40px'>Détails Tape Media Pool</h2>
<table>
<tr><th>Pool Name</th><th>Total TB</th><th>Free TB</th><th>Used TB</th><th>Used %</th></tr>"
foreach ($r in $results) {
    $cls = if ($r.UsedPercentage -ge 70) { "red" } else { "green" }
    $html += "<tr><td>$($r.PoolName)</td><td>$($r.TotalCapacityTB)</td><td>$($r.FreeCapacityTB)</td><td>$($r.UsedCapacityTB)</td><td class='$cls'>$($r.UsedPercentage)%</td></tr>"
}
$html += "</table>"


# =================================================
# TABLEAU DÉTAILLÉ JOBS (Backup, Replica, Tape) 
# =================================================
$html += "<h2 style='text-align:center;margin-top:40px'>Tableau détaillé Jobs Backup</h2>
<table>
<tr><th>Type</th><th>Job</th><th>Success</th><th>Warning</th><th>Failed</th><th>%</th><th>Last Run</th></tr>"
foreach ($j in $backupStats) {
    $cls = if ($j.Warning -ge 1) { "orange" } elseif ($j.Percent -ge 80) { "green" } else { "red" }
    $html += "<tr><td>$($j.Type)</td><td>$($j.JobName)</td><td>$($j.Success)</td><td>$($j.Warning)</td><td>$($j.Failed)</td><td class='$cls'>$($j.Percent)%</td><td>$($j.LastRun)</td></tr>"
}
$html += "</table>"

$html += "<h2 style='text-align:center;margin-top:40px'>Tableau détaillé Jobs Replication</h2>
<table>
<tr><th>Type</th><th>Job</th><th>Success</th><th>Warning</th><th>Failed</th><th>%</th><th>Last Run</th></tr>"
foreach ($j in $replicaStats) {
    $cls = if ($j.Percent -ge 80) { "green" } else { "red" }
    $html += "<tr><td>$($j.Type)</td><td>$($j.JobName)</td><td>$($j.Success)</td><td>$($j.Warning)</td><td>$($j.Failed)</td><td class='$cls'>$($j.Percent)%</td><td>$($j.LastRun)</td></tr>"
}
$html += "</table>"

$html += "<h2 style='text-align:center;margin-top:40px'>Tableau détaillé Jobs Tape Backup</h2>
<table>
<tr><th>Type</th><th>Job</th><th>Success</th><th>Warning</th><th>Failed</th><th>%</th><th>Last Run</th></tr>"
foreach ($j in $tapeStats) {
    $cls = if ($j.Percent -ge 80) { "green" } else { "red" }
    $html += "<tr><td>$($j.Type)</td><td>$($j.JobName)</td><td>$($j.Success)</td><td>$($j.Warning)</td><td>$($j.Failed)</td><td class='$cls'>$($j.Percent)%</td><td>$($j.LastRun)</td></tr>"
}
$html += "</table>"


# =================================================
# TABLEAU DÉTAILLÉ DES BANDES EXPIREES
# =================================================
$html += "<h2 style='text-align:center;margin-top:40px'>Liste des tapes expirées</h2>
<table>
<tr><th>Tape Number</th><th>Expiration Date</th></tr>"
foreach ($j in $expiredTapesTable) {
    $html += "<tr><td>$($j.TapeNumber)</td><td>$($j.TapeExpirationDate)</td></tr>"
}
$html += "</table>"


# ===============================================================
# EXPORT DU RAPPORT EN HTML DANS C:\Temp\Veeam_KPI_Dashboard.html
# ==============================================================
$html += "</body></html>"
$html | Out-File $output -Encoding UTF8


# =================================================
# ENVOIE DU RAPPORT PAR EMAIL
# =================================================

$mailBody = @"
<html>
<head>
<meta charset="UTF-8">
</head>

<body style="font-family:'Segoe UI', Arial, sans-serif; font-size:13px; color:#000; margin:0; padding:0">

<p>Bonjour,</p>

<p style="margin-top:15px">
Veuillez trouver joint <b>le dashboard et les détails de l’état de santé de l'infrastructure Veeam $environnement $filiale</b> pour la période des 7 derniers jours ($rptPeriod).</b>
</p>

<p>
Ci-dessous le tableau recapitulatif:
</p>

<table cellpadding="6" cellspacing="0" border="1"
       style="
       border-collapse:collapse;
       width:100%;
       max-width:100%;
       font-size:13px;
       border-color:#dcdcdc;
       table-layout:auto;
       ">

<tr style="
    background-color:#293C52;
    color:#ffffff;
    text-align:center;
    font-weight:bold;
">
  <th>Type</th>
  <th>Total</th>
  <th>Success</th>
  <th>Warning</th>
  <th>Failed</th>
  <th>Success %</th>
</tr>

<tr style="text-align:center">
  <td><b>Backup</b></td>
  <td>$bkTotal</td>
  <td style="color:#00B336"><b>$bkSuccess</b></td>
  <td style="color:#FF7F00"><b>$bkWarning</b></td>
  <td style="color:#E31E24"><b>$bkFailed</b></td>
  <td><b>$bkPercent %</b></td>
</tr>

<tr style="text-align:center">
  <td><b>Replication</b></td>
  <td>$rpTotal</td>
  <td style="color:#00B336"><b>$rpSuccess</b></td>
  <td style="color:#FF7F00"><b>$rpWarning</b></td>
  <td style="color:#E31E24"><b>$rpFailed</b></td>
  <td><b>$rpPercent %</b></td>
</tr>

<tr style="text-align:center">
  <td><b>Tape Backup</b></td>
  <td>$tpTotal</td>
  <td style="color:#00B336"><b>$tpSuccess</b></td>
  <td style="color:#FF7F00"><b>$tpWarning</b></td>
  <td style="color:#E31E24"><b>$tpFailed</b></td>
  <td><b>$tpPercent %</b></td>
</tr>

</table>



<p>
Cordialement,<br>
<b>SGABS IaaS Teeam</b>
</p>

</body>
</html>
"@


Send-MailMessage `
    -From $mailFrom `
    -To $mailTo `
    -Subject $mailSubject `
    -Body $mailBody `
    -BodyAsHtml `
    -SmtpServer $smtpServer `
    -Port $smtpPort `
    -Encoding UTF8 `
    -Attachments $output `


Write-Host "Dashboard généré : $output" -ForegroundColor Green
