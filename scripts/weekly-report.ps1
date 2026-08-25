<#
  Εβδομαδιαία αναφορά πωλήσεων -> Typst -> PDF   (read-only)
  Τρέχει χειροκίνητα και από scheduled task (κάθε Δευτέρα πρωί).

  Χρήση:
    .\weekly-report.ps1                       # η τελευταία πλήρης εβδομάδα (Δευ-Κυρ)
    .\weekly-report.ps1 -WeekStart 2026-07-27
    .\weekly-report.ps1 -Email                # στέλνει και email (θέλει EUROPHARMACY_SMTP_* στο .env)
#>
[CmdletBinding()]
param(
  [string]$WeekStart,
  [string]$OutDir,
  [switch]$NoOpen,
  [switch]$Email
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

# $PSScriptRoot δεν είναι διαθέσιμο σε param defaults όταν τρέχει με -File (scheduled task)
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $OutDir) { $OutDir = Join-Path $scriptDir '..\Αναφορές' }

# ── κοινή βιβλιοθήκη (.env, σύνδεση, logging, email) ─────────────────────
. (Join-Path $scriptDir 'lib\report-common.ps1')
$cfg = Get-PharmacyConfig -StartDir $scriptDir

$LogDir = Join-Path $OutDir 'logs'
trap {
  Write-ReportLog -LogDir $LogDir -Level 'ERROR' -Message ("Εβδομαδιαία: " + $_.Exception.Message + " @γραμμή " + $_.InvocationInfo.ScriptLineNumber)
  break
}
$db    = Open-PharmacyDb -Cfg $cfg -LogDir $LogDir
$cn    = $db.Connection
$route = $db.Route
$cmd = $cn.CreateCommand()
$cmd.CommandText = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SET DATEFIRST 1;"
$cmd.ExecuteNonQuery() | Out-Null
$cmd.CommandTimeout = 300

function Rows($sql) {
  $cmd.CommandText = $sql
  $rd = $cmd.ExecuteReader(); $out = @()
  while ($rd.Read()) {
    $h = [ordered]@{}
    for ($i=0; $i -lt $rd.FieldCount; $i++) { $h[$rd.GetName($i)] = $rd.GetValue($i) }
    $out += [pscustomobject]$h
  }
  $rd.Close(); return ,$out
}

# ── περίοδος ─────────────────────────────────────────────────────────────
if ($WeekStart) { $ws = [datetime]::ParseExact($WeekStart,'yyyy-MM-dd',$null) }
else {
  $today = (Rows "SELECT CAST(GETDATE() AS DATE) d")[0].d
  $dow = [int]$today.DayOfWeek; if ($dow -eq 0) { $dow = 7 }
  $ws = $today.AddDays(-($dow - 1) - 7)
}
$we  = $ws.AddDays(6)
$f   = $ws.ToString('yyyy-MM-dd');  $t = $ws.AddDays(7).ToString('yyyy-MM-dd')
$pf  = $ws.AddDays(-7).ToString('yyyy-MM-dd')

# κοινό CTE: όλες οι γραμμές πώλησης (ελεύθερες + συνταγές) με αξία λιανικής
$LINES = @"
WITH lines AS (
 SELECT x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] Tx, CAST(x.[DateTime] AS DATE) D, d.[ΒΔ] VD, d.[ΚΩΔ_ΕΙΔΟΥΣ] C,
        d.[ΤΕΜΑΧΙΑ] Q, (d.[ΤΙΜΗ]*d.[ΤΕΜΑΧΙΑ]-ISNULL(d.[ΕΚΠΤΩΣΗ_ΑΞ],0)) V
 FROM FreeSalesDetails d JOIN Transactions x ON x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]=d.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]
 WHERE x.[DateTime]>='$f' AND x.[DateTime]<'$t' AND x.False_Tran=0 AND ISNULL(d.Hidden,0)=0
 UNION ALL
 SELECT x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ], CAST(x.[DateTime] AS DATE), p.[ΒΔ], p.[ΚΩΔ_ΕΙΔΟΥΣ],
        p.[ΤΕΜΑΧΙΑ], (ISNULL(p.[ΛΙΑΝΙΚΗ_ΤΙΜΗ],0)*p.[ΤΕΜΑΧΙΑ])
 FROM PrescDetails p JOIN Transactions x ON x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]=p.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]
 WHERE x.[DateTime]>='$f' AND x.[DateTime]<'$t' AND x.False_Tran=0
)
"@
$NAME = @"
COALESCE((SELECT TOP 1 COMMERCIAL_NAME_ONLY FROM HDIKA WHERE CAST(EOF_CODE AS nvarchar(20))=a.C AND COMMERCIAL_NAME_ONLY IS NOT NULL),
         (SELECT TOP 1 ItemName FROM MHSYFA WHERE CAST(EOFCd AS nvarchar(20))=a.C),
         (SELECT TOP 1 [ΠΕΡΙΓΡΑΦΗ_ΕΙΔΟΥΣ] FROM MedUser WHERE CAST([ΓΕΝ_ΚΩΔΙΚΟΣ] AS nvarchar(20))=a.C),
         '['+a.VD+':'+a.C+']')
"@

# ── δεδομένα ─────────────────────────────────────────────────────────────
$totals = (Rows @"
SELECT
 SUM(CASE WHEN [DateTime]>='$f' AND False_Tran=0 THEN 1 ELSE 0 END) R,
 SUM(CASE WHEN [DateTime]>='$f' AND False_Tran=0 THEN [ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] ELSE 0 END) T,
 SUM(CASE WHEN [DateTime]<'$f'  AND False_Tran=0 THEN 1 ELSE 0 END) PR,
 SUM(CASE WHEN [DateTime]<'$f'  AND False_Tran=0 THEN [ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] ELSE 0 END) PT,
 SUM(CASE WHEN [DateTime]>='$f' AND False_Tran=1 THEN 1 ELSE 0 END) CANC
FROM Transactions WHERE [DateTime]>='$pf' AND [DateTime]<'$t'
"@)[0]

$daily = Rows @"
SELECT CAST([DateTime] AS DATE) D, DATENAME(weekday,[DateTime]) DOW, COUNT(*) R,
 SUM([ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) T,
 SUM(CASE WHEN DATEPART(hour,[DateTime])<14 THEN [ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] ELSE 0 END) AM
FROM Transactions WHERE [DateTime]>='$f' AND [DateTime]<'$t' AND False_Tran=0
GROUP BY CAST([DateTime] AS DATE), DATENAME(weekday,[DateTime]) ORDER BY D
"@

$tills = Rows @"
SELECT CAST([DateTime] AS DATE) D, MachineName M, COUNT(*) R, SUM([ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) T
FROM Transactions WHERE [DateTime]>='$f' AND [DateTime]<'$t' AND False_Tran=0
GROUP BY CAST([DateTime] AS DATE), MachineName ORDER BY D, M
"@

$payDaily = Rows @"
WITH ce AS (
  SELECT [ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] tx, SUM([ΜΕΤΡΗΤΑ]) ca, SUM([ΚΑΡΤΑ]) cd, SUM([ΚΑΤΑΘΕΣΗ]) dp, SUM([ΠΙΣΤΩΣΗ]) cr
  FROM CashExtras WHERE ISNULL(Hidden,0)=0 GROUP BY [ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ])
SELECT CAST(t.[DateTime] AS DATE) D,
 SUM(ISNULL(ce.ca,0)) Cash, SUM(ISNULL(ce.cd,0)) Card, SUM(ISNULL(ce.dp,0)) Dep, SUM(ISNULL(ce.cr,0)) Credit,
 SUM(ISNULL(ce.ca,0)+ISNULL(ce.cd,0)+ISNULL(ce.dp,0)+ISNULL(ce.cr,0)) Tot,
 SUM(t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) Sales
FROM Transactions t LEFT JOIN ce ON ce.tx=t.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]
WHERE t.[DateTime]>='$f' AND t.[DateTime]<'$t' AND t.False_Tran=0
GROUP BY CAST(t.[DateTime] AS DATE) ORDER BY D
"@

$posDaily = Rows @"
SELECT CAST(t.[DateTime] AS DATE) D, c.posId P, SUM(c.[ΚΑΡΤΑ]) T
FROM CashExtras c JOIN Transactions t ON t.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]=c.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]
WHERE t.[DateTime]>='$f' AND t.[DateTime]<'$t' AND t.False_Tran=0
  AND ISNULL(c.Hidden,0)=0 AND c.[ΚΑΡΤΑ]<>0 AND c.posId IS NOT NULL
GROUP BY CAST(t.[DateTime] AS DATE), c.posId ORDER BY 1, 2
"@

# Φάρμακα (ΒΔ=R) vs Παραφάρμακα (ΒΔ=U)
$split = Rows ($LINES + @"
SELECT VD, COUNT(DISTINCT Tx) Rc, SUM(Q) U, SUM(V) Val FROM lines GROUP BY VD
"@)

# Παραφάρμακα ανά κατηγορία
$paraCat = Rows ($LINES + @"
SELECT ISNULL(mc.MKA_DESCRIPTION,'Λοιπά / χωρίς κατηγορία') Cat,
       COUNT(DISTINCT l.Tx) Rc, SUM(l.Q) U, SUM(l.V) Val
FROM lines l
LEFT JOIN MedUser mu ON CAST(mu.[ΓΕΝ_ΚΩΔΙΚΟΣ] AS nvarchar(20)) = l.C
LEFT JOIN MAIN_CATEGORY mc ON mc.MKA_CODE = mu.[ΚΥΡΙΑ_ΚΑΤ]
WHERE l.VD='U'
GROUP BY mc.MKA_DESCRIPTION ORDER BY Val DESC
"@)

$topR = Rows ($LINES + @"
, a AS (SELECT VD,C,COUNT(DISTINCT Tx) R,SUM(Q) U,SUM(V) V FROM lines WHERE VD='R' GROUP BY VD,C)
SELECT TOP 20 a.R,a.U,a.V, $NAME N FROM a ORDER BY a.R DESC
"@)

$topU = Rows ($LINES + @"
, a AS (SELECT VD,C,COUNT(DISTINCT Tx) R,SUM(Q) U,SUM(V) V FROM lines WHERE VD='U' GROUP BY VD,C)
SELECT TOP 20 a.R,a.U,a.V, $NAME N FROM a ORDER BY a.R DESC
"@)

$trend = Rows @"
WITH d AS (SELECT CAST([DateTime] AS DATE) D,
  SUM(CASE WHEN [ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]<10000 THEN [ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] ELSE 0 END) T, COUNT(*) R
 FROM Transactions WHERE [DateTime]>=DATEADD(day,-42,'$f') AND [DateTime]<'$t' AND False_Tran=0
 GROUP BY CAST([DateTime] AS DATE))
SELECT MIN(D) S, MAX(D) E, COUNT(*) Dys, SUM(T) Sales, SUM(R) Rc
FROM d GROUP BY DATEDIFF(day,DATEADD(day,-42,'$f'),D)/7 ORDER BY MIN(D)
"@

$big = Rows @"
SELECT TOP 3 [DateTime] DT, [ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] A, MachineName M FROM Transactions
WHERE [DateTime]>='$f' AND [DateTime]<'$t' AND False_Tran=0 ORDER BY [ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] DESC
"@
$cn.Close()

# ── helpers ──────────────────────────────────────────────────────────────
$gr = [System.Globalization.CultureInfo]::GetCultureInfo('el-GR')
function M($v) { if ($null -eq $v -or $v -is [DBNull]) { '0,00' } else { ([decimal]$v).ToString('N2',$gr) } }
function Esc($s) { ($s -replace '\\','\\' -replace '"','\"' -replace '#','\#' -replace '\$','\$' -replace '@','\@' -replace '_','\_' -replace '\*','\*') }
# Ονόματα ημερών από .NET: το DATENAME εξαρτάται από τη γλώσσα του SQL Server.
$dayNames = @('Κυριακή','Δευτέρα','Τρίτη','Τετάρτη','Πέμπτη','Παρασκευή','Σάββατο')
function DayGr([datetime]$d) { $dayNames[[int]$d.DayOfWeek] }
# Όνομα φαρμακείου για το υποσέλιδο — από το .env, ώστε το script να μένει γενικό
$pharmacyName = if ($cfg['EUROPHARMACY_NAME']) { $cfg['EUROPHARMACY_NAME'] } else { 'Φαρμακείο' }

# Προαιρετικό logo. Σχετικές διαδρομές λύνονται ως προς τον φάκελο του .env.
$logoPath = $null
if ($cfg['EUROPHARMACY_LOGO']) {
  $lp = $cfg['EUROPHARMACY_LOGO']
  if (-not [System.IO.Path]::IsPathRooted($lp)) { $lp = Join-Path $cfg['_envDir'] $lp }
  if (Test-Path $lp) { $logoPath = (Resolve-Path $lp).Path }
  else { Write-Warning "EUROPHARMACY_LOGO δείχνει σε ανύπαρκτο αρχείο: $lp — η αναφορά θα γίνει χωρίς logo." }
}

$T=[decimal]$totals.T; $PT=[decimal]$totals.PT; $R=[int]$totals.R; $PR=[int]$totals.PR
$dT  = if ($PT -ne 0) { ($T-$PT)/$PT*100 } else { 0 }
$dR  = if ($PR -ne 0) { ($R-$PR)/$PR*100 } else { 0 }
$avg = if ($R -gt 0) { $T/$R } else { 0 }
$pavg= if ($PR -gt 0) { $PT/$PR } else { 0 }
$dA  = if ($pavg -ne 0) { ($avg-$pavg)/$pavg*100 } else { 0 }
function Pct($x) { $s = if ($x -ge 0) { '+' } else { '−' }; "$s$(([decimal][math]::Abs($x)).ToString('N1',$gr))%" }
function Delta($x) { if ($x -ge 0) { "#up[$(Pct $x)]" } else { "#dn[$(Pct $x)]" } }

$tillNames = @($tills | Select-Object -Expand M -Unique | Sort-Object)
$days      = @($daily | Select-Object -Expand D)
$posIds    = @($posDaily | Select-Object -Expand P -Unique | Sort-Object)

$valR = [decimal](($split | Where-Object { $_.VD -eq 'R' } | Select-Object -First 1).Val)
$valU = [decimal](($split | Where-Object { $_.VD -eq 'U' } | Select-Object -First 1).Val)
$valAll = $valR + $valU

# ── Typst ────────────────────────────────────────────────────────────────
$sb = New-Object System.Text.StringBuilder
function W($s) { [void]$sb.AppendLine($s) }
# Κάθε ενότητα σε block που ΔΕΝ σπάει: αν δεν χωράει, μετακινείται ολόκληρη
# στην επόμενη σελίδα αντί να κοπεί ο πίνακας στη μέση.
function Sect($title) { W '#block(breakable: false)['; if ($title) { W ('#heading(level: 2)[' + $title + ']') } }
function EndSect { W ']' }

# ── παλέτα ──
# Ταιριάξτε τη με το logo σας μέσω .env:
#   EUROPHARMACY_COLOR_PRIMARY  = σκούρο χρώμα μάρκας (μπάρα κεφαλίδας, επικεφαλίδες)
#   EUROPHARMACY_COLOR_ACCENT   = φωτεινό χρώμα μάρκας (τονισμοί, γραφήματα)
#   EUROPHARMACY_COLOR_SECOND   = συμπληρωματικό (παραφάρμακα)
# Αν το logo έχει σκούρο φόντο ίδιο με το PRIMARY, ενσωματώνεται αθόρυβα στην κεφαλίδα.
$colPrimary = if ($cfg['EUROPHARMACY_COLOR_PRIMARY']) { $cfg['EUROPHARMACY_COLOR_PRIMARY'] } else { '#3f4a5a' }
$colAccent  = if ($cfg['EUROPHARMACY_COLOR_ACCENT'])  { $cfg['EUROPHARMACY_COLOR_ACCENT']  } else { '#7fc9bd' }
$colSecond  = if ($cfg['EUROPHARMACY_COLOR_SECOND'])  { $cfg['EUROPHARMACY_COLOR_SECOND']  } else { '#d9a05b' }
# Παράγωγα: σκουρότερο accent για κείμενο σε λευκό, ανοιχτότερο για δευτερεύουσες μπάρες
function Shade([string]$hex, [double]$f) {   # f<1 σκουραίνει, f>1 ανοίγει προς το λευκό
  $h = $hex.TrimStart('#')
  $r=[Convert]::ToInt32($h.Substring(0,2),16); $g=[Convert]::ToInt32($h.Substring(2,2),16); $b=[Convert]::ToInt32($h.Substring(4,2),16)
  if ($f -le 1) { $r=[int]($r*$f); $g=[int]($g*$f); $b=[int]($b*$f) }
  else { $t=$f-1; $r=[int]($r+(255-$r)*$t); $g=[int]($g+(255-$g)*$t); $b=[int]($b+(255-$b)*$t) }
  '#{0:x2}{1:x2}{2:x2}' -f [math]::Min(255,$r),[math]::Min(255,$g),[math]::Min(255,$b)
}
W ('#let NAVY    = rgb("' + $colPrimary + '")')
W ('#let MINT    = rgb("' + $colAccent + '")')
W ('#let MINT_DK = rgb("' + (Shade $colAccent 0.70) + '")')
W ('#let MINT_LT = rgb("' + (Shade $colAccent 1.55) + '")')
W ('#let SAND    = rgb("' + $colSecond + '")')
W '#let ACCENT  = NAVY'
W '#let ACCENT2 = rgb("#eaf0f4")'   # ανοιχτό slate για κεφαλίδες πινάκων
W '#let INK     = rgb("#25303f")'
W '#let MUTED   = rgb("#78849a")'
W '#let ZEBRA   = rgb("#fafbfc")'
W '#let GOOD    = rgb("#15803d")'
W '#let BAD     = rgb("#b42318")'
W '#let WARN    = rgb("#a9702f")'
W ''
W '#set page(paper: "a4", margin: (x: 1.4cm, top: 1.2cm, bottom: 1.4cm),'
W '  footer: context [#set text(size: 7.5pt, fill: MUTED)'
W '    #line(length: 100%, stroke: 0.3pt + MUTED) #v(-4pt)'
W '    #grid(columns: (1fr, auto), align: (left, right),'
W ('      [' + (Esc $pharmacyName) + ' · read-only αναφορά],')
W '      [#counter(page).display("1 / 1", both: true)])])'
W '#set text(font: ("Calibri", "Arial"), size: 9.5pt, lang: "el", fill: INK)'
W '#show heading.where(level: 2): it => block(above: 12pt, below: 6pt)[#text(size: 11.5pt, weight: "medium", fill: ACCENT)[#it.body]]'
W ''
W '#let up(x) = text(fill: GOOD, weight: "medium", x)'
W '#let dn(x) = text(fill: BAD, weight: "medium", x)'
W '#let hdr(c) = (fill: ACCENT2, y: 0)'
# πίνακας με χρωματιστή κεφαλίδα + zebra
W '#let tbl(cols, align: none, ..args) = table('
W '  columns: cols, align: align,'
W '  fill: (x, y) => if y == 0 { ACCENT2 } else if calc.odd(y) { ZEBRA },'
W '  stroke: (x, y) => (bottom: if y == 0 { 0.9pt + ACCENT } else { 0.25pt + rgb("#e5e7eb") }),'
W '  inset: (x: 5pt, y: 4pt),'
W '  ..args)'
W ''
# ── κεφαλίδα (logo + τίτλος σε brand navy μπάρα) ──
# Το logo έχει ήδη navy φόντο ίδιο με τη μπάρα, οπότε ενσωματώνεται αθόρυβα.
# Το Typst δεν διαβάζει αρχεία εκτός του root του, οπότε αντιγράφουμε το logo
# δίπλα στο .typ και το αναφέρουμε με σκέτο όνομα (κρατά και το .typ αυτοτελές).
$logoTypst = ''
if ($logoPath) {
  if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
  $logoLocal = '_logo' + [System.IO.Path]::GetExtension($logoPath)
  Copy-Item $logoPath (Join-Path (Resolve-Path $OutDir) $logoLocal) -Force
  # Πολλά logo έχουν μεγάλα κενά περιθώρια. Το ZOOM μεγεθύνει την εικόνα μέσα σε κουτί
  # σταθερού μεγέθους και το υπερχείλισμα κόβεται (clip), δηλ. «κροπάρει» τα περιθώρια.
  $zoom = 1.0
  if ($cfg['EUROPHARMACY_LOGO_ZOOM']) { $zoom = [double]::Parse($cfg['EUROPHARMACY_LOGO_ZOOM'], [Globalization.CultureInfo]::InvariantCulture) }
  $boxH = 54.0; $boxW = 74.0
  $imgH = [math]::Round($boxH * $zoom, 1)
  $logoTypst = 'box(width: ' + $boxW + 'pt, height: ' + $boxH + 'pt, clip: true, align(center + horizon, image("' + $logoLocal + '", height: ' + $imgH + 'pt)))'
}
W '#block(fill: NAVY, width: 100%, inset: (x: 0pt, y: 0pt), radius: 4pt, clip: true)['
W '  #grid(columns: (auto, 1fr), align: (left + horizon, right + horizon), inset: (x: 14pt, y: 9pt),'
if ($logoTypst) { W ("    $logoTypst,") } else { W '    box(width: 0pt),' }
W '    align(right)['
W '      #text(fill: white, size: 16pt, weight: "medium")[Εβδομαδιαία Αναφορά Πωλήσεων] \'
W ('      #text(fill: MINT, size: 10.5pt)[' + $ws.ToString('dd/MM/yyyy') + ' – ' + $we.ToString('dd/MM/yyyy') + '] \')
W ('      #text(fill: rgb("#9aa6bd"), size: 7.5pt)[Παράχθηκε ' + (Get-Date).ToString('dd/MM/yyyy HH:mm') + ' · ζωντανά δεδομένα]'  )
W '    ])]'
W '#v(3pt)'

# ── KPI cards ──
$kpi = @(
  @{ l='Πωλήσεις'; v="$(M $T) €"; d=(Delta $dT); s="προηγ. $(M $PT) €" },
  @{ l='Αποδείξεις'; v="$R"; d=(Delta $dR); s="προηγ. $PR" },
  @{ l='Μέσο καλάθι'; v="$(M $avg) €"; d=(Delta $dA); s="προηγ. $(M $pavg) €" }
)
W '#grid(columns: (1fr, 1fr, 1fr), gutter: 8pt,'
foreach ($k in $kpi) {
  W ('  block(fill: ACCENT2, width: 100%, inset: 9pt, radius: 3pt)[')
  W ('    #text(size: 8.5pt, fill: MUTED)[' + $k.l + '] \')
  W ('    #text(size: 15pt, weight: "medium")[' + $k.v + '] #h(4pt) ' + $k.d + ' \')
  W ('    #text(size: 7.5pt, fill: MUTED)[' + $k.s + ']],')
}
W ')'
if ([int]$totals.CANC -gt 0) { W ('#text(size: 8pt, fill: MUTED)[Ακυρωμένες συναλλαγές: ' + $totals.CANC + ' (εξαιρούνται από όλα τα σύνολα)]') }

# ── ημερήσια ──
Sect 'Ημερήσια ανάλυση'
W '#tbl((auto, 1fr, auto, auto, auto, auto), align: (left, left, right, right, right, right),'
W '  [*Ημ/νία*],[*Ημέρα*],[*Αποδ.*],[*Σύνολο*],[*Πρωί*],[*Πρωί %*],'
foreach ($d in $daily) {
  $tt=[decimal]$d.T; $am=[decimal]$d.AM
  $p = if ($tt -ne 0) { [math]::Round($am/$tt*100) } else { 0 }
  W ("  [" + ([datetime]$d.D).ToString('dd/MM') + "],[" + (DayGr ([datetime]$d.D)) + "],[" + $d.R + "],[" + (M $tt) + "],[" + (M $am) + "],[" + $p + "%],")
}
W ("  [*Σύνολο*],[],[*$R*],[*$(M $T)*],[],[],")
W ')'

# ── γράφημα ημέρας: πρωί / απόγευμα ──
$maxDay = ($daily | ForEach-Object { [decimal]$_.T } | Measure-Object -Maximum).Maximum
if ($maxDay -gt 0) {
  $H = 68.0
  W '#v(6pt)'
  W ('#grid(columns: (' + ((@('1fr') * $daily.Count) -join ', ') + '), gutter: 10pt,')
  foreach ($d in $daily) {
    $tt=[decimal]$d.T; $am=[decimal]$d.AM; $pm=$tt-$am
    $hAM = [math]::Round($H * [double]($am/$maxDay),1)
    $hPM = [math]::Round($H * [double]($pm/$maxDay),1)
    W ('  align(center, stack(spacing: 3pt,')
    W ('    box(height: ' + $H + 'pt, align(bottom, stack(dir: btt,')
    W ('      rect(width: 26pt, height: ' + $hAM + 'pt, fill: MINT_DK, stroke: none),')
    W ('      rect(width: 26pt, height: ' + $hPM + 'pt, fill: MINT_LT, stroke: none)))),')
    W ('    text(size: 7.5pt, weight: "medium")[' + ([datetime]$d.D).ToString('dd/MM') + '],')
    W ('    text(size: 7pt, fill: MUTED)[' + (M $tt) + '])),')
  }
  W ')'
  W '#v(2pt)'
  W '#align(center)[#text(size: 7.5pt, fill: MUTED)[#box(width: 8pt, height: 8pt, fill: MINT_DK) πρωί (πριν 14:00) #h(10pt) #box(width: 8pt, height: 8pt, fill: MINT_LT) απόγευμα]]'
}

# ── ταμεία × ημέρα ──
EndSect
Sect 'Ταμεία ανά ημέρα'
if ($tillNames.Count -eq 0) {
  # Χωρίς αυτόν τον έλεγχο, μια περίοδος χωρίς πωλήσεις (π.χ. κλειστά τον Αύγουστο)
  # παρήγαγε «(auto,, auto)» και το Typst απέτυχε με συντακτικό σφάλμα.
  W '#text(fill: MUTED)[Καμία πώληση στην περίοδο.]'
} else {
$nc = $tillNames.Count
W ('#tbl((auto,' + (@('1fr') * $nc -join ',') + ', auto), align: (left,' + (@('right') * $nc -join ',') + ', right),')
W ('  [*Ημ/νία*],' + (($tillNames | ForEach-Object { '[*' + (Esc $_) + '*]' }) -join ',') + ',[*Σύνολο*],')
foreach ($d in $days) {
  $cells = foreach ($m in $tillNames) {
    $r = $tills | Where-Object { $_.D -eq $d -and $_.M -eq $m }
    if ($r) { '[' + (M $r.T) + ']' } else { '[#text(fill: MUTED)[–]]' }
  }
  $rowTot = ($tills | Where-Object { $_.D -eq $d } | Measure-Object -Property T -Sum).Sum
  W ('  [' + ([datetime]$d).ToString('dd/MM') + '],' + ($cells -join ',') + ',[' + (M $rowTot) + '],')
}
$colTots = foreach ($m in $tillNames) { '[*' + (M ($tills | Where-Object { $_.M -eq $m } | Measure-Object -Property T -Sum).Sum) + '*]' }
W ('  [*Σύνολο*],' + ($colTots -join ',') + ',[*' + (M $T) + '*],')
W ')'
}
W '#text(size: 8pt, fill: MUTED)[Ποσά σε € (πληρωτέο πελατών).]'

# ── τρόποι πληρωμής ──
EndSect
Sect 'Τρόποι πληρωμής'
W '#tbl((auto, auto, auto, auto, auto, auto, auto), align: (left, right, right, right, right, right, right),'
W '  [*Ημ/νία*],[*Μετρητά*],[*Κάρτα*],[*Κατάθεση*],[*Πίστωση*],[*Εισπράξεις*],[*Διαφ.*],'
$sC=0;$sK=0;$sD=0;$sP=0;$sT2=0;$sDiff=0
foreach ($p in $payDaily) {
  $sC+=[decimal]$p.Cash; $sK+=[decimal]$p.Card; $sD+=[decimal]$p.Dep; $sP+=[decimal]$p.Credit; $sT2+=[decimal]$p.Tot
  $diff = [decimal]$p.Tot - [decimal]$p.Sales; $sDiff += $diff
  $dCell = if ([math]::Abs($diff) -lt 0.005) { '#text(fill: MUTED)[–]' } else { '#text(fill: WARN)[' + (M $diff) + ']' }
  W ('  [' + ([datetime]$p.D).ToString('dd/MM') + '],[' + (M $p.Cash) + '],[' + (M $p.Card) + '],[' + (M $p.Dep) + '],[' + (M $p.Credit) + '],[' + (M $p.Tot) + '],[' + $dCell + '],')
}
W ("  [*Σύνολο*],[*$(M $sC)*],[*$(M $sK)*],[*$(M $sD)*],[*$(M $sP)*],[*$(M $sT2)*],[*$(M $sDiff)*],")
W ')'
if ($sT2 -ne 0) {
  # στοιβαγμένη μπάρα μείγματος πληρωμών
  $segs = @(
    @{ n='Μετρητά';  v=$sC; c=$colPrimary },
    @{ n='Κάρτα';    v=$sK; c=(Shade $colAccent 0.70) },
    @{ n='Κατάθεση'; v=$sD; c=$colAccent },
    @{ n='Πίστωση';  v=$sP; c=$colSecond }
  ) | Where-Object { $_.v -gt 0 }
  $pcts = @(); $acc = 0
  for ($j=0; $j -lt $segs.Count; $j++) {
    if ($j -eq $segs.Count-1) { $pc = 100 - $acc } else { $pc = [math]::Round([double]($segs[$j].v/$sT2*100),1); $acc += $pc }
    $pcts += $pc
  }
  W '#v(4pt)'
  W '#stack(dir: ltr,'
  for ($j=0; $j -lt $segs.Count; $j++) {
    W ('  box(width: ' + $pcts[$j] + '%, height: 15pt, fill: rgb("' + $segs[$j].c + '")),')
  }
  W ')'
  W '#v(3pt)'
  $leg = for ($j=0; $j -lt $segs.Count; $j++) {
    '#box(width: 8pt, height: 8pt, fill: rgb("' + $segs[$j].c + '")) ' + $segs[$j].n + ' ' + ([decimal]$pcts[$j]).ToString('N1',$gr) + '%'
  }
  W ('#text(size: 7.5pt, fill: MUTED)[' + ($leg -join ' #h(9pt) ') + ']')
  W ('#v(2pt)')
  W ('#text(size: 8pt, fill: MUTED)[«Διαφ.» = εισπράξεις μείον αξία πωλήσεων (π.χ. εξόφληση παλιάς οφειλής). Μικρά ποσά φυσιολογικά.]')
}

EndSect   # κλείσιμο «Τρόποι πληρωμής»

# ── POS ──
if ($posIds.Count -gt 0) {
  Sect 'Κάρτες ανά τερματικό (POS)'
  $pc = $posIds.Count
  W ('#tbl((auto,' + (@('1fr') * $pc -join ',') + ', auto), align: (left,' + (@('right') * $pc -join ',') + ', right),')
  W ('  [*Ημ/νία*],' + (($posIds | ForEach-Object { '[*POS ' + $_ + '*]' }) -join ',') + ',[*Σύνολο*],')
  foreach ($d in ($payDaily | Select-Object -Expand D)) {
    $cells = foreach ($term in $posIds) {
      $r = $posDaily | Where-Object { $_.D -eq $d -and $_.P -eq $term }
      if ($r) { '[' + (M $r.T) + ']' } else { '[#text(fill: MUTED)[–]]' }
    }
    $rt = ($posDaily | Where-Object { $_.D -eq $d } | Measure-Object -Property T -Sum).Sum
    W ('  [' + ([datetime]$d).ToString('dd/MM') + '],' + ($cells -join ',') + ',[' + (M $rt) + '],')
  }
  $pTots = foreach ($term in $posIds) { '[*' + (M ($posDaily | Where-Object { $_.P -eq $term } | Measure-Object -Property T -Sum).Sum) + '*]' }
  W ('  [*Σύνολο*],' + ($pTots -join ',') + ',[*' + (M $sK) + '*],')
  W ')'
  W '#text(size: 8pt, fill: MUTED)[Μη διασυνδεδεμένα τερματικά δεν καταγράφονται με posId και δεν εμφανίζονται ως ξεχωριστή στήλη.]'
  EndSect
}

# ── μεγαλύτερες αποδείξεις ──
if ($big.Count -gt 0) {
  Sect 'Μεγαλύτερες αποδείξεις'
  W '#tbl((auto, auto, 1fr), align: (left, right, left),'
  W '  [*Ημ/νία & ώρα*],[*Ποσό*],[*Ταμείο*],'
  foreach ($b2 in $big) { W ('  [' + ([datetime]$b2.DT).ToString('dd/MM HH:mm') + '],[' + (M $b2.A) + ' €],[' + (Esc ([string]$b2.M)) + '],') }
  W ')'
  $mx = [decimal]$big[0].A
  if ($T -ne 0 -and $mx/$T -gt 0.08) {
    W ('#block(fill: rgb("#fef3f2"), width: 100%, inset: 7pt, radius: 3pt)[#text(size: 8.5pt, fill: BAD)[*Προσοχή:* η μεγαλύτερη απόδειξη είναι ' + [math]::Round($mx/$T*100,1) + '% του εβδομαδιαίου τζίρου. Χωρίς αυτήν οι πωλήσεις είναι ' + (M ($T-$mx)) + ' €.]]')
  }
  EndSect
}

# ── φάρμακα vs παραφάρμακα ──
Sect 'Φάρμακα vs Παραφάρμακα'
$pR = if ($valAll -ne 0) { [math]::Round($valR/$valAll*100) } else { 0 }
$pU = 100 - $pR
W '#grid(columns: (1fr, 1fr), gutter: 8pt,'
W ('  block(fill: ACCENT2, width: 100%, inset: 9pt, radius: 3pt)[')
W ('    #text(size: 8.5pt, fill: MUTED)[Φάρμακα (συνταγογραφούμενα & ΟΤΣ, ΕΟΦ)] \')
W ('    #text(size: 15pt, weight: "medium")[' + (M $valR) + ' €] #h(4pt) #text(fill: ACCENT, weight: "medium")[' + $pR + '%] \')
W ('    #text(size: 7.5pt, fill: MUTED)[' + [int](($split | Where-Object { $_.VD -eq 'R' }).U) + ' τεμάχια]],')
W ('  block(fill: rgb("#fbf1e5"), width: 100%, inset: 9pt, radius: 3pt)[')
W ('    #text(size: 8.5pt, fill: MUTED)[Παραφάρμακα & λοιπά είδη] \')
W ('    #text(size: 15pt, weight: "medium")[' + (M $valU) + ' €] #h(4pt) #text(fill: WARN, weight: "medium")[' + $pU + '%] \')
W ('    #text(size: 7.5pt, fill: MUTED)[' + [int](($split | Where-Object { $_.VD -eq 'U' }).U) + ' τεμάχια]],')
W ')'
W '#text(size: 8pt, fill: MUTED)[Αξία λιανικής των γραμμών πώλησης (ελεύθερες + συνταγές). Διαφέρει ελαφρώς από το «πληρωτέο πελατών», που αφαιρεί τη συμμετοχή των ταμείων.]'

# ── παραφάρμακα ανά κατηγορία ──
EndSect
Sect 'Παραφάρμακα ανά κατηγορία'
W '#tbl((1fr, auto, auto, auto, auto, 4.2cm), align: (left, right, right, right, right, left),'
W '  [*Κατηγορία*],[*Αποδ.*],[*Τεμ.*],[*Αξία*],[*%*],[],'
$maxCat = ($paraCat | ForEach-Object { [decimal]$_.Val } | Measure-Object -Maximum).Maximum
foreach ($c in $paraCat) {
  $cv=[decimal]$c.Val
  $sh = if ($valU -ne 0) { [math]::Round($cv/$valU*100,1) } else { 0 }
  $bw = if ($maxCat -gt 0) { [math]::Round([double]($cv/$maxCat*100),1) } else { 0 }
  $bar = '[#box(width: ' + $bw + '%, height: 7pt, fill: SAND, radius: 1pt)]'
  W ('  [' + (Esc ([string]$c.Cat)) + '],[' + $c.Rc + '],[' + [int]$c.U + '],[' + (M $cv) + '],[' + $sh.ToString('N1',$gr) + '%],' + $bar + ',')
}
W ("  [*Σύνολο*],[],[],[*$(M $valU)*],[*100,0*],[],")
W ')'
EndSect

# ── top ανά κατηγορία ──
$rankNote = '#text(size: 8pt, fill: MUTED)[Κατάταξη κατά πλήθος διαφορετικών αποδείξεων = πραγματικό εύρος ζήτησης. Ο διαχωρισμός ακολουθεί τον χαρακτηρισμό της βάσης (μητρώο ΕΟΦ vs είδη χρήστη)· ένα φάρμακο καταχωρημένο ως «είδος χρήστη» εμφανίζεται στα παραφάρμακα.]'
foreach ($grp in @(
    @{ t = 'Top ' + $topR.Count + ' φάρμακα';      rows = $topR; col = 'ACCENT' },
    @{ t = 'Top ' + $topU.Count + ' παραφάρμακα'; rows = $topU; col = 'SAND'  })) {
  Sect $grp.t
  W '#tbl((auto, 1fr, auto, auto, auto), align: (right, left, right, right, right),'
  W '  [*\#*],[*Προϊόν*],[*Αποδ.*],[*Τεμ.*],[*Αξία*],'
  $i = 0
  foreach ($p in $grp.rows) {
    $i++
    W ("  [$i],[" + (Esc ([string]$p.N)) + "],[" + $p.R + "],[" + [int]$p.U + "],[" + (M $p.V) + "],")
  }
  W ')'
  W $rankNote
  EndSect
}

# ── τάση ──
$maxS = ($trend | ForEach-Object { [decimal]$_.Sales } | Measure-Object -Maximum).Maximum
W '#block(breakable: false)['
W '#heading(level: 2)[Τάση 6 εβδομάδων]'
W '#tbl((auto, auto, auto, auto, 1fr), align: (left, right, right, right, left),'
W '  [*Εβδομάδα*],[*Ημέρες*],[*Πωλήσεις*],[*Μ.Ο. καλαθιού*],[*Σχετικά*],'
foreach ($w in $trend) {
  $wr=[int]$w.Rc; $wt=[decimal]$w.Sales
  $wa = if ($wr -gt 0) { $wt/$wr } else { 0 }
  $lbl = ([datetime]$w.S).ToString('dd/MM') + '–' + ([datetime]$w.E).ToString('dd/MM')
  $cur = ([datetime]$w.S) -ge $ws
  $frac = if ($maxS -gt 0) { [math]::Round($wt/$maxS,3) } else { 0 }
  $col = if ($cur) { 'NAVY' } else { 'MINT_LT' }
  $b = if ($cur) { '*' } else { '' }
  W ("  [$b$lbl$b],[$($w.Dys)],[$b$(M $wt)$b],[$(M $wa)],[#box(width: " + ($frac*100) + "%, height: 7pt, fill: $col, radius: 1pt)],")
}
W ')'
W '#text(size: 8pt, fill: MUTED)[Εξαιρούνται μεμονωμένες συναλλαγές ≥ 10.000 € (χονδρική).]'
W ']'
W '#v(6pt)'
W '#text(size: 7.5pt, fill: MUTED)[«Πωλήσεις» = πληρωτέο πελατών (λιανική + συμμετοχές)· δεν περιλαμβάνεται η αποζημίωση ΕΟΠΥΥ. Οι ακυρωμένες συναλλαγές εξαιρούνται.]'

# ── γράψιμο & compile ────────────────────────────────────────────────────
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$base = Join-Path (Resolve-Path $OutDir) ("Εβδομαδιαία_" + $ws.ToString('yyyy-MM-dd'))
$typ = "$base.typ"; $pdf = "$base.pdf"
[System.IO.File]::WriteAllText($typ, $sb.ToString(), (New-Object Text.UTF8Encoding $false))

$typstExe = @("$env:LOCALAPPDATA\Microsoft\WinGet\Links\typst.exe", 'typst') |
  Where-Object { Get-Command $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
if (-not $typstExe) { throw "Δεν βρέθηκε το typst. Εγκατάσταση: winget install Typst.Typst" }
& $typstExe compile $typ $pdf
if ($LASTEXITCODE -ne 0) { throw "Απέτυχε το compile του typst." }

Write-Output "OK route=$route"
Write-Output "PDF: $pdf"
Write-ReportLog -LogDir $LogDir -Message ("Εβδομαδιαία " + $ws.ToString('yyyy-MM-dd') + " OK — " + (M $T) + " €, " + [int]$totals.R + " αποδ. (route " + $route + ")")

# ── προαιρετικό email ────────────────────────────────────────────────────
if ($Email) {
  foreach ($need in @('EUROPHARMACY_SMTP_HOST','EUROPHARMACY_SMTP_USER','EUROPHARMACY_SMTP_PASS','EUROPHARMACY_SMTP_TO')) {
    if (-not $cfg[$need]) { throw "Λείπει το $need από το .env — δες την ενότητα email στο skill." }
  }
  $port = if ($cfg['EUROPHARMACY_SMTP_PORT']) { [int]$cfg['EUROPHARMACY_SMTP_PORT'] } else { 587 }
  $msg = New-Object Net.Mail.MailMessage
  $msg.From = New-Object Net.Mail.MailAddress($cfg['EUROPHARMACY_SMTP_USER'], 'Αναφορές Φαρμακείου')
  foreach ($to in ($cfg['EUROPHARMACY_SMTP_TO'] -split '[;,]')) { if ($to.Trim()) { $msg.To.Add($to.Trim()) } }
  $msg.Subject = "Εβδομαδιαία αναφορά πωλήσεων " + $ws.ToString('dd/MM') + "–" + $we.ToString('dd/MM/yyyy')
  $msg.BodyEncoding = [Text.Encoding]::UTF8
  $msg.SubjectEncoding = [Text.Encoding]::UTF8
  $msg.IsBodyHtml = $true
  $msg.Body = @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#1f2937">
<p>Καλημέρα,</p>
<p>Η εβδομαδιαία αναφορά για <b>$($ws.ToString('dd/MM'))–$($we.ToString('dd/MM/yyyy'))</b> είναι συνημμένη.</p>
<table style="border-collapse:collapse;font-size:14px">
<tr><td style="padding:2px 12px 2px 0">Πωλήσεις</td><td style="padding:2px 0"><b>$(M $T) €</b> ($(Pct $dT))</td></tr>
<tr><td style="padding:2px 12px 2px 0">Αποδείξεις</td><td style="padding:2px 0">$R ($(Pct $dR))</td></tr>
<tr><td style="padding:2px 12px 2px 0">Μέσο καλάθι</td><td style="padding:2px 0">$(M $avg) € ($(Pct $dA))</td></tr>
</table>
<p style="color:#6b7280;font-size:12px">Αυτόματο μήνυμα από τον υπολογιστή του φαρμακείου.</p></div>
"@
  $msg.Attachments.Add((New-Object Net.Mail.Attachment($pdf)))
  $smtp = New-Object Net.Mail.SmtpClient($cfg['EUROPHARMACY_SMTP_HOST'], $port)
  $smtp.EnableSsl = $true
  $smtp.Credentials = New-Object Net.NetworkCredential($cfg['EUROPHARMACY_SMTP_USER'], $cfg['EUROPHARMACY_SMTP_PASS'])
  $smtp.Send($msg)
  $msg.Dispose()
  Write-Output ("Email στάλθηκε -> " + $cfg['EUROPHARMACY_SMTP_TO'])
}

if (-not $NoOpen) { Invoke-Item $pdf }
