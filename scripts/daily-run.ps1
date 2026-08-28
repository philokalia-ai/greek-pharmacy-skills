<#
  Βραδινή εκτέλεση ημερήσιων αναφορών — με αυτόματη ανάκτηση χαμένων ημερών.

  Γιατί υπάρχει: το scheduled task μπορεί να χάσει μια βραδιά (ο υπολογιστής
  κοιμήθηκε στο κλείσιμο, το Tailscale δεν είχε συνδεθεί, ο server ήταν κλειστός).
  Το «-Today» της daily-report.ps1 κλειδώνει στη σημερινή ημερομηνία, οπότε μια
  καθυστερημένη εκτέλεση δεν ανακτά τη χαμένη ημέρα — απλώς ξαναβγάζει τη σημερινή.

  Αυτό το script κοιτάζει τις τελευταίες N ημέρες, βρίσκει όσες είχαν πωλήσεις
  και δεν έχουν ακόμη αναφορά, και τις στέλνει. Η σημερινή παράγεται πάντα.

  Χρήση:  .\daily-run.ps1            # ό,τι λείπει από τις τελευταίες 7 ημέρες
          .\daily-run.ps1 -Days 14 -NoEmail
#>
[CmdletBinding()]
param(
  [int]$Days = 7,
  [string]$OutDir,
  [switch]$NoEmail
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
. (Join-Path $scriptDir 'lib\report-common.ps1')
if (-not $OutDir) { $OutDir = Join-Path $scriptDir '..\Αναφορές' }
$LogDir = Join-Path $OutDir 'logs'

trap {
  Write-ReportLog -LogDir $LogDir -Level 'ERROR' -Message ("Βραδινή εκτέλεση: " + $_.Exception.Message)
  break
}

$cfg = Get-PharmacyConfig -StartDir $scriptDir
$db  = Open-PharmacyDb -Cfg $cfg -LogDir $LogDir
$cn  = $db.Connection
$today = [datetime](Invoke-Rows -Connection $cn -Sql "SELECT CAST(GETDATE() AS DATE) D")[0].D
$rows = Invoke-Rows -Connection $cn -Sql @"
SELECT CAST([DateTime] AS DATE) D
FROM Transactions
WHERE [DateTime] >= DATEADD(day,-$Days,CAST(GETDATE() AS DATE)) AND False_Tran = 0
GROUP BY CAST([DateTime] AS DATE) ORDER BY D
"@
$cn.Close()

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$outFull = (Resolve-Path $OutDir).Path
$daily   = Join-Path $scriptDir 'daily-report.ps1'

$done = 0; $skipped = 0; $failed = 0
foreach ($r in $rows) {
  $d   = [datetime]$r.D
  $iso = $d.ToString('yyyy-MM-dd')
  $pdf = Join-Path $outFull ("Ημερήσια_$iso.pdf")
  $isToday = ($d -eq $today)

  # Η σημερινή ξαναβγαίνει πάντα (το βράδυ είναι η οριστική). Οι παλιότερες μόνο αν λείπουν.
  if (-not $isToday -and (Test-Path $pdf)) { $skipped++; continue }

  # Hashtable splat: το $args είναι δεσμευμένη μεταβλητή της PowerShell και το
  # splatting πίνακα περνά τα ορίσματα θέσης, όχι ονόματος.
  $opts = @{ Date = $iso; NoOpen = $true; OutDir = $outFull }
  if (-not $NoEmail) { $opts['Email'] = $true }
  try {
    & $daily @opts | Out-Null
    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) { throw "exit $LASTEXITCODE" }
    $tag = if ($isToday) { 'σημερινή' } else { 'ΑΝΑΚΤΗΣΗ χαμένης ημέρας' }
    Write-Output ("OK  $iso  ($tag)")
    $done++
  } catch {
    Write-ReportLog -LogDir $LogDir -Level 'ERROR' -Message ("Βραδινή: απέτυχε η $iso — " + $_.Exception.Message)
    Write-Output ("ΣΦΑΛΜΑ $iso : " + $_.Exception.Message)
    $failed++
  }
}
Write-ReportLog -LogDir $LogDir -Message ("Βραδινή εκτέλεση: $done εστάλησαν, $skipped υπήρχαν ήδη, $failed απέτυχαν (παράθυρο $Days ημερών).")
Write-Output "Σύνολο: $done εστάλησαν · $skipped υπήρχαν ήδη · $failed απέτυχαν"
if ($failed -gt 0) { exit 1 }
