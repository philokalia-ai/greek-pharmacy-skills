<#
  Ημερήσια αναφορά πωλήσεων -> Typst -> PDF   (read-only, μία σελίδα)

  Η σύγκριση γίνεται με την ΙΔΙΑ ΗΜΕΡΑ ΕΒΔΟΜΑΔΑΣ (μ.ο. των τελευταίων 8), όχι με
  «χθες»: η κίνηση φαρμακείου εξαρτάται έντονα από την ημέρα (Δευτέρα ≠ Σάββατο).

  Χρήση:
    .\daily-report.ps1                    # η τελευταία ημέρα με πωλήσεις
    .\daily-report.ps1 -Date 2026-07-31
    .\daily-report.ps1 -Date 2026-07-31 -Email
    .\daily-report.ps1 -Today -Email     # για βραδινό scheduled run (πάντα η σημερινή)
#>
[CmdletBinding()]
param(
  [string]$Date,
  [string]$OutDir,
  [switch]$NoOpen,
  [switch]$Email,
  [switch]$Today,
  [int]$BaselineDays = 8
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
. (Join-Path $scriptDir 'lib\report-common.ps1')
if (-not $OutDir) { $OutDir = Join-Path $scriptDir '..\Αναφορές' }

$LogDir = Join-Path $OutDir 'logs'
# trap: καταγράφει την αιτία πριν βγει το script, ώστε μια αποτυχία του scheduled
# task να μη φαίνεται μόνο ως «LastTaskResult = 1».
trap {
  Write-ReportLog -LogDir $LogDir -Level 'ERROR' -Message ("Ημερήσια: " + $_.Exception.Message + " @γραμμή " + $_.InvocationInfo.ScriptLineNumber)
  break
}

$cfg   = Get-PharmacyConfig -StartDir $scriptDir
$brand = Get-Brand $cfg
$db    = Open-PharmacyDb -Cfg $cfg -LogDir $LogDir
$cn    = $db.Connection
function Q($sql) { Invoke-Rows -Connection $cn -Sql $sql }

# Ταξινόμηση παραστατικών: «λιανική» χωριστά από τιμολόγια ΕΟΠΥΥ/λοιπά.
$JT   = Get-DocJoinSql -Alias 'dc' -TxAlias 't'
$JX   = Get-DocJoinSql -Alias 'dc' -TxAlias 'x'
$CLS  = Get-DocClassSql 'dc'
$RET  = Get-RetailFilterSql 'dc'

# ── ημερομηνία ───────────────────────────────────────────────────────────
if ($Date) { $D = [datetime]::ParseExact($Date,'yyyy-MM-dd',$null) }
elseif ($Today) {
  # Για το βραδινό scheduled run: πάντα η σημερινή ημέρα. Χωρίς αυτό, μια κλειστή
  # ημέρα θα έστελνε ξανά την αναφορά της προηγούμενης ανοιχτής ημέρας.
  $D = [datetime](Q "SELECT CAST(GETDATE() AS DATE) D")[0].D
}
else {
  $r = Q "SELECT TOP 1 CAST([DateTime] AS DATE) D FROM Transactions WHERE False_Tran=0 GROUP BY CAST([DateTime] AS DATE) ORDER BY D DESC"
  if (-not $r) { throw "Δεν υπάρχουν καθόλου πωλήσεις στη βάση." }
  $D = [datetime]$r[0].D
}
$d0 = $D.ToString('yyyy-MM-dd'); $d1 = $D.AddDays(1).ToString('yyyy-MM-dd')

# ── δεδομένα ημέρας ──────────────────────────────────────────────────────
# ISNULL: σε ημέρα χωρίς καμία γραμμή, το SUM επιστρέφει NULL και η μετατροπή σε
# int/decimal σκάει πριν προλάβει ο έλεγχος «κλειστά» παρακάτω.
$tot = (Q @"
SELECT ISNULL(SUM(CASE WHEN t.False_Tran=0 AND $RET THEN 1 ELSE 0 END),0) R,
       ISNULL(SUM(CASE WHEN t.False_Tran=0 AND $RET THEN t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] ELSE 0 END),0) T,
       ISNULL(SUM(CASE WHEN t.False_Tran=1 THEN 1 ELSE 0 END),0) CANC
FROM Transactions t $JT
WHERE t.[DateTime]>='$d0' AND t.[DateTime]<'$d1'
"@)[0]

# Ανάλυση κατά είδος παραστατικού (όλα, όχι μόνο λιανική)
$docs = Q @"
SELECT $CLS Cls, COUNT(*) N, SUM(t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) A
FROM Transactions t $JT
WHERE t.[DateTime]>='$d0' AND t.[DateTime]<'$d1' AND t.False_Tran=0
GROUP BY $CLS
"@
$R = [int]$tot.R; $T = [decimal]$tot.T
if ($R -eq 0) {
  # Κλειστή ημέρα (Κυριακή, αργία, άδεια): δεν είναι σφάλμα. Βγαίνουμε ήσυχα με
  # κωδικό 0, ώστε το scheduled task να μη «κοκκινίζει» κάθε Κυριακή.
  Write-ReportLog -LogDir $LogDir -Message ("Ημερήσια " + $D.ToString('yyyy-MM-dd') + ": καμία πώληση — κλειστά. Παραλείπεται.")
  Write-Output ("Καμία πώληση στις " + $D.ToString('dd/MM/yyyy') + " — το φαρμακείο ήταν κλειστό. Δεν παράγεται αναφορά.")
  $cn.Close()
  exit 0
}
$avg = $T / $R

# baseline: οι τελευταίες Ν ίδιες ημέρες εβδομάδας (χωρίς χονδρικές >=10k)
$baseSql = @"
WITH days AS (
  SELECT TOP $BaselineDays CAST(t.[DateTime] AS DATE) D,
    SUM(t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) T, COUNT(*) R
  FROM Transactions t $JT
  WHERE t.False_Tran=0 AND $RET AND CAST(t.[DateTime] AS DATE) < '$d0'
    AND DATEPART(weekday,t.[DateTime]) = DATEPART(weekday, CAST('$d0' AS date))
  GROUP BY CAST(t.[DateTime] AS DATE) ORDER BY D DESC)
"@
$baseDays = Q ($baseSql + "SELECT D, T, R FROM days ORDER BY D")
$base     = (Q ($baseSql + "SELECT COUNT(*) N, AVG(T) AvgT, AVG(R*1.0) AvgR FROM days"))[0]
$bN = [int]$base.N
$bT = if ($bN -gt 0) { [decimal]$base.AvgT } else { 0 }
$bR = if ($bN -gt 0) { [decimal]$base.AvgR } else { 0 }
$bA = if ($bR -gt 0) { $bT / $bR } else { 0 }
$dT = if ($bT -ne 0) { [double](($T-$bT)/$bT*100) } else { 0 }
$deltaR = if ($bR -ne 0) { [double](($R-$bR)/$bR*100) } else { 0 }
$dA = if ($bA -ne 0) { [double](($avg-$bA)/$bA*100) } else { 0 }

# ωριαίο προφίλ: σήμερα + τυπικό
$hToday = Q @"
SELECT DATEPART(hour,t.[DateTime]) H, COUNT(*) R, SUM(t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) T
FROM Transactions t $JT
WHERE t.[DateTime]>='$d0' AND t.[DateTime]<'$d1' AND t.False_Tran=0 AND $RET
GROUP BY DATEPART(hour,t.[DateTime]) ORDER BY H
"@
$hBase = Q ($baseSql + @"
, h AS (SELECT CAST(t.[DateTime] AS DATE) D, DATEPART(hour,t.[DateTime]) H,
        SUM(t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) T
        FROM Transactions t $JT JOIN days ON days.D = CAST(t.[DateTime] AS DATE)
        WHERE t.False_Tran=0 AND $RET GROUP BY CAST(t.[DateTime] AS DATE), DATEPART(hour,t.[DateTime]))
SELECT H, AVG(T) AvgT FROM h GROUP BY H ORDER BY H
"@)

$tills = Q @"
SELECT t.MachineName M, COUNT(*) R, SUM(t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) T
FROM Transactions t $JT
WHERE t.[DateTime]>='$d0' AND t.[DateTime]<'$d1' AND t.False_Tran=0 AND $RET
GROUP BY t.MachineName ORDER BY T DESC
"@

$pay = (Q @"
WITH ce AS (SELECT [ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] tx, SUM([ΜΕΤΡΗΤΑ]) ca, SUM([ΚΑΡΤΑ]) cd, SUM([ΚΑΤΑΘΕΣΗ]) dp, SUM([ΠΙΣΤΩΣΗ]) cr
            FROM CashExtras WHERE ISNULL(Hidden,0)=0 GROUP BY [ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ])
SELECT SUM(ISNULL(ce.ca,0)) Cash, SUM(ISNULL(ce.cd,0)) Card, SUM(ISNULL(ce.dp,0)) Dep, SUM(ISNULL(ce.cr,0)) Credit,
       SUM(t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ]) Sales
FROM Transactions t LEFT JOIN ce ON ce.tx=t.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] $JT
WHERE t.[DateTime]>='$d0' AND t.[DateTime]<'$d1' AND t.False_Tran=0 AND $RET
"@)[0]

$LINES = @"
WITH lines AS (
 SELECT x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] Tx, d.[ΒΔ] VD, d.[ΚΩΔ_ΕΙΔΟΥΣ] C, d.[ΤΕΜΑΧΙΑ] Q,
        (d.[ΤΙΜΗ]*d.[ΤΕΜΑΧΙΑ]-ISNULL(d.[ΕΚΠΤΩΣΗ_ΑΞ],0)) V
 FROM FreeSalesDetails d JOIN Transactions x ON x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]=d.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] $JX
 WHERE x.[DateTime]>='$d0' AND x.[DateTime]<'$d1' AND x.False_Tran=0 AND ISNULL(d.Hidden,0)=0 AND $RET
 UNION ALL
 SELECT x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ], p.[ΒΔ], p.[ΚΩΔ_ΕΙΔΟΥΣ], p.[ΤΕΜΑΧΙΑ], (ISNULL(p.[ΛΙΑΝΙΚΗ_ΤΙΜΗ],0)*p.[ΤΕΜΑΧΙΑ])
 FROM PrescDetails p JOIN Transactions x ON x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]=p.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] $JX
 WHERE x.[DateTime]>='$d0' AND x.[DateTime]<'$d1' AND x.False_Tran=0 AND $RET)
"@
$split = Q ($LINES + "SELECT VD, SUM(Q) U, SUM(V) Val FROM lines GROUP BY VD")
$top   = Q ($LINES + @"
, a AS (SELECT VD,C,COUNT(DISTINCT Tx) R,SUM(Q) U,SUM(V) V FROM lines GROUP BY VD,C)
SELECT TOP 10 a.VD, a.R, a.U, a.V,
 COALESCE((SELECT TOP 1 COMMERCIAL_NAME_ONLY FROM HDIKA WHERE CAST(EOF_CODE AS nvarchar(20))=a.C AND COMMERCIAL_NAME_ONLY IS NOT NULL),
          (SELECT TOP 1 ItemName FROM MHSYFA WHERE CAST(EOFCd AS nvarchar(20))=a.C),
          (SELECT TOP 1 [ΠΕΡΙΓΡΑΦΗ_ΕΙΔΟΥΣ] FROM MedUser WHERE CAST([ΓΕΝ_ΚΩΔΙΚΟΣ] AS nvarchar(20))=a.C),
          '['+a.VD+':'+a.C+']') N
FROM a ORDER BY a.R DESC, a.V DESC
"@)
$big = Q @"
SELECT TOP 3 t.[DateTime] DT, t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] A, t.MachineName M
FROM Transactions t $JT
WHERE t.[DateTime]>='$d0' AND t.[DateTime]<'$d1' AND t.False_Tran=0 AND $RET
ORDER BY t.[ΠΛΗΡΩΤΕΟ_ΠΩΛΗΣΗΣ] DESC
"@
$rx = (Q @"
SELECT COUNT(DISTINCT p.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]) N FROM PrescDetails p
JOIN Transactions x ON x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]=p.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] $JX
WHERE x.[DateTime]>='$d0' AND x.[DateTime]<'$d1' AND x.False_Tran=0 AND $RET
"@)[0].N
$mix = Q @"
WITH l AS (
 SELECT d.[ΒΔ] VD, (d.[ΤΙΜΗ]*d.[ΤΕΜΑΧΙΑ]-ISNULL(d.[ΕΚΠΤΩΣΗ_ΑΞ],0)) V
 FROM FreeSalesDetails d JOIN Transactions x ON x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]=d.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] $JX
 WHERE x.[DateTime]>='$d0' AND x.[DateTime]<'$d1' AND x.False_Tran=0 AND ISNULL(d.Hidden,0)=0
   AND $RET AND ISNULL(dc.F5,0)=0
 UNION ALL
 SELECT p.[ΒΔ], (ISNULL(p.[ΛΙΑΝΙΚΗ_ΤΙΜΗ],0)*p.[ΤΕΜΑΧΙΑ])
 FROM PrescDetails p JOIN Transactions x ON x.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]=p.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] $JX
 WHERE x.[DateTime]>='$d0' AND x.[DateTime]<'$d1' AND x.False_Tran=0 AND $RET AND ISNULL(dc.F5,0)=0)
SELECT VD, SUM(V) Val FROM l GROUP BY VD
"@
$cn.Close()

# ── σύνθεση τζίρου (για τη μπάρα) ──
# Το «σύνολο» περιλαμβάνει τα τιμολόγια ΕΟΠΥΥ· τα επιμέρους breakdowns παραμένουν λιανική.
function DocAmt($k) {
  $x = $docs | Where-Object { [string]$_.Cls -eq $k } | Select-Object -First 1
  if ($x) { [decimal]$x.A } else { [decimal]0 }
}
$segE   = (DocAmt 'ΕΟΠΥΥ') + (DocAmt 'ΤΙΜΟΛΟΓΙΟ')
$segF5  = DocAmt 'ΔΛΠ'
$alpTot = (DocAmt 'ΑΛΠ') + (DocAmt 'ΠΙΣΤΩΤΙΚΟ') + (DocAmt 'ΧΩΡΙΣ')
$mixR = [decimal](($mix | Where-Object { [string]$_.VD -eq 'R' } | Select-Object -First 1).Val)
$mixU = [decimal](($mix | Where-Object { [string]$_.VD -eq 'U' } | Select-Object -First 1).Val)
$mixT = $mixR + $mixU
# Το ΑΛΠ μέρος μοιράζεται σε φάρμακα/παραφάρμακα με την αναλογία των γραμμών του.
$segF = if ($mixT -gt 0) { $alpTot * $mixR / $mixT } else { $alpTot }
$segP = $alpTot - $segF
$TAll = $segE + $segF + $segP + $segF5

# ── Typst ────────────────────────────────────────────────────────────────
$sb = New-Object System.Text.StringBuilder
function W($s) { [void]$sb.AppendLine($s) }
function Sect($t) { W '#block(breakable: false)['; if ($t) { W ("#heading(level: 2)[$t]") } }
function EndSect { W ']' }
function Delta([double]$x) { if ($x -ge 0) { "#up[$(Format-Pct $x)]" } else { "#dn[$(Format-Pct $x)]" } }
function TT($s) { ConvertTo-TypstText $s }

W (Get-TypstPreamble -Cfg $cfg -Brand $brand -FooterNote 'ημερήσια read-only αναφορά' -Margin '(x: 1.3cm, top: 1.1cm, bottom: 1.3cm)')
W (Get-TypstHeader -Cfg $cfg -OutDir $OutDir `
    -Title 'Ημερήσια Αναφορά' `
    -Subtitle ((Get-GreekDayName $D) + ' ' + $D.Day + ' ' + (Get-GreekMonthName $D.Month) + ' ' + $D.Year) `
    -Note ("Παράχθηκε " + (Get-Date).ToString('dd/MM/yyyy HH:mm') + " · σύγκριση με τις τελευταίες $bN ίδιες ημέρες") `
    -BoxH 46)
W '#v(3pt)'

# KPI
$kpis = @(
  @{ l='Σύνολο ημέρας'; v=((Format-Money $TAll) + ' €'); d='';          s='λιανική + τιμολόγια' },
  @{ l='Λιανική';       v=((Format-Money $T) + ' €');   d=(Delta $dT); s=("τυπική " + (Format-Money $bT) + ' €') },
  @{ l='Αποδείξεις';    v="$R";                         d=(Delta $deltaR); s=("τυπικά " + (Format-Num $bR 0)) },
  @{ l='Μέσο καλάθι';   v=((Format-Money $avg) + ' €'); d=(Delta $dA); s=("τυπικά " + (Format-Money $bA) + ' €') }
)
W '#grid(columns: (1fr, 1fr, 1fr, 1fr), gutter: 7pt,'
foreach ($k in $kpis) {
  W '  block(fill: ACCENT2, width: 100%, inset: 8pt, radius: 3pt)['
  W ('    #text(size: 8pt, fill: MUTED)[' + $k.l + '] \')
  W ('    #text(size: 14pt, weight: "medium")[' + $k.v + '] ' + $(if ($k.d) { '#h(3pt) ' + $k.d }) + ' \')
  W ('    #text(size: 7pt, fill: MUTED)[' + $k.s + ']],')
}
W ')'
if ([int]$tot.CANC -gt 0) { W ('#text(size: 7.5pt, fill: MUTED)[Ακυρωμένες: ' + $tot.CANC + ' (εξαιρούνται)]') }

# ── μπάρα σύνθεσης τζίρου ──
$bar = Get-TypstCompositionBar -Eopyy $segE -Pharma $segF -Para $segP -F5 $segF5
if ($bar) {
  Sect 'Σύνθεση τζίρου'
  W $bar
  W ('#text(size: 7pt, fill: MUTED)[Με συνταγή: ' + $rx + ' αποδείξεις (' + [math]::Round($rx*100.0/$R) + '% της λιανικής). Ο διαχωρισμός φαρμάκων/παραφαρμάκων ακολουθεί την αναλογία των γραμμών των αποδείξεων λιανικής.]')
  EndSect
}

# ── ανάλυση κατά είδος παραστατικού ──
$docLbl = @{ 'ΑΛΠ'='Αποδείξεις λιανικής'; 'ΔΛΠ'='Δελτία λιανικής πώλησης'; 'ΠΙΣΤΩΤΙΚΟ'='Πιστωτικά (επιστροφές)';
          'ΕΟΠΥΥ'='Τιμολόγια Ε.Ο.Π.Υ.Υ.'; 'ΤΙΜΟΛΟΓΙΟ'='Λοιπά τιμολόγια'; 'ΧΩΡΙΣ'='Χωρίς παραστατικό' }
$retailKeys = @('ΑΛΠ','ΔΛΠ','ΠΙΣΤΩΤΙΚΟ','ΧΩΡΙΣ')
$nonRetail = @($docs | Where-Object { $retailKeys -notcontains [string]$_.Cls })
if ($docs.Count -gt 1 -or $nonRetail.Count -gt 0) {
  Sect 'Ανάλυση κατά παραστατικό'
  W '#tbl((1fr, auto, auto, auto), align: (left, right, right, left),'
  W '  [*Παραστατικό*],[*Πλήθος*],[*Ποσό*],[],'
  foreach ($k in @('ΑΛΠ','ΔΛΠ','ΠΙΣΤΩΤΙΚΟ','ΧΩΡΙΣ','ΕΟΠΥΥ','ΤΙΜΟΛΟΓΙΟ')) {
    $docRow = $docs | Where-Object { [string]$_.Cls -eq $k } | Select-Object -First 1
    if (-not $docRow) { continue }
    $isRetail = $retailKeys -contains $k
    $tag = if ($isRetail) { '#text(size: 7pt, fill: MINT_DK)[λιανική]' } else { '#text(size: 7pt, fill: SAND)[εκτός λιανικής]' }
    W ('  [' + $docLbl[$k] + '],[' + $docRow.N + '],[' + (Format-Money $docRow.A) + '],[' + $tag + '],')
  }
  $grand = ($docs | ForEach-Object { [decimal]$_.A } | Measure-Object -Sum).Sum
  W ("  [*Σύνολο όλων*],[],[*" + (Format-Money $grand) + "*],[],")
  W ')'
  $nrSum = if ($nonRetail.Count) { ($nonRetail | ForEach-Object { [decimal]$_.A } | Measure-Object -Sum).Sum } else { 0 }
  W ('#text(size: 7.5pt, fill: MUTED)[Ο δείκτης «Πωλήσεις» πιο πάνω μετράει *μόνο τη λιανική* (' + (Format-Money $T) + ' €). Τα τιμολόγια δεν είναι πωλήσεις ημέρας — τιμολογούνται συγκεντρωτικά και θα αλλοίωναν τη σύγκριση με τις άλλες ημέρες.]')
  if ($nrSum -ne 0) {
    W ('#text(size: 7.5pt, fill: WARN)[Εκτός λιανικής σήμερα: ' + (Format-Money $nrSum) + ' €.]')
  }
  W '#text(size: 7.5pt, fill: MUTED)[Τα «δελτία λιανικής πώλησης» είναι κυρίως επί πιστώσει σε ονομαστικούς πελάτες και δεν φέρουν δικό τους αριθμό — η απόδειξη εκδίδεται με την εξόφληση. Μετρώνται κανονικά στη λιανική.]'
  EndSect
}

# ── ωριαίο προφίλ με «τυπική» γραμμή ──
if ($hToday.Count -gt 0) {
  Sect 'Ωριαίο προφίλ'
  $hrs = @($hToday | Select-Object -Expand H)
  $maxH = [double](($hToday | ForEach-Object { [decimal]$_.T } | Measure-Object -Maximum).Maximum)
  $mb   = [double](($hBase  | ForEach-Object { [decimal]$_.AvgT } | Measure-Object -Maximum).Maximum)
  if ($mb -gt $maxH) { $maxH = $mb }
  $CH = 54.0
  W ('#grid(columns: (' + ((@('1fr') * $hrs.Count) -join ', ') + '), gutter: 4pt,')
  foreach ($h in $hrs) {
    $tv = [double]([decimal](($hToday | Where-Object { $_.H -eq $h }).T))
    $bv = 0.0
    $bh = $hBase | Where-Object { $_.H -eq $h }
    if ($bh) { $bv = [double]([decimal]$bh.AvgT) }
    $hb = if ($maxH -gt 0) { [math]::Round($CH * $tv / $maxH, 1) } else { 0 }
    $hl = if ($maxH -gt 0) { [math]::Round($CH * $bv / $maxH, 1) } else { 0 }
    $col = if ($bv -gt 0 -and $tv -lt $bv) { 'MINT_LT' } else { 'MINT_DK' }
    W '  align(center, stack(spacing: 2pt,'
    W ('    box(width: 100%, height: ' + $CH + 'pt)[')
    W ('      #place(bottom, rect(width: 62%, height: ' + $hb + 'pt, fill: ' + $col + ', stroke: none))')
    if ($bv -gt 0) { W ('      #place(bottom, dy: -' + $hl + 'pt, rect(width: 100%, height: 1.4pt, fill: NAVY, stroke: none))') }
    W '    ],'
    W ('    text(size: 7pt, fill: MUTED)[' + $h + ':00])),')
  }
  W ')'
  W '#v(1pt)'
  W '#align(center)[#text(size: 7.5pt, fill: MUTED)[#box(width: 8pt, height: 8pt, fill: MINT_DK) η ημέρα #h(9pt) #box(width: 10pt, height: 3pt, fill: NAVY) τυπικό επίπεδο ίδιας ημέρας #h(9pt) #box(width: 8pt, height: 8pt, fill: MINT_LT) κάτω από το τυπικό]]'
  EndSect
}

# ── ταμεία + πληρωμές δίπλα-δίπλα ──
$cash=[decimal]$pay.Cash; $card=[decimal]$pay.Card; $dep=[decimal]$pay.Dep; $cred=[decimal]$pay.Credit
$paySum = $cash+$card+$dep+$cred
W '#grid(columns: (1fr, 1fr), gutter: 10pt,'
W '  block[#block(breakable: false)[#heading(level: 2)[Ταμεία]'
W '  #tbl((1fr, auto, auto), align: (left, right, right),'
W '    [*Ταμείο*],[*Αποδ.*],[*Σύνολο*],'
foreach ($m in $tills) { W ('    [' + (TT ([string]$m.M)) + '],[' + $m.R + '],[' + (Format-Money $m.T) + '],') }
W ('    [*Σύνολο*],[*' + $R + '*],[*' + (Format-Money $T) + '*],')
W '  )]],'
W '  block[#block(breakable: false)[#heading(level: 2)[Τρόποι πληρωμής]'
W '  #tbl((1fr, auto, auto), align: (left, right, right),'
W '    [*Τρόπος*],[*Ποσό*],[*%*],'
$segs = @(
  @{n='Μετρητά';  v=$cash; c='NAVY'},
  @{n='Κάρτα';    v=$card; c='MINT_DK'},
  @{n='Κατάθεση'; v=$dep;  c='MINT'},
  @{n='Πίστωση';  v=$cred; c='SAND'}
)
foreach ($s in $segs) {
  $pc = if ($paySum -ne 0) { [math]::Round([double]($s.v/$paySum*100),1) } else { 0 }
  W ('    [#box(width: 7pt, height: 7pt, fill: ' + $s.c + ') ' + $s.n + '],[' + (Format-Money $s.v) + '],[' + (Format-Num $pc 1) + '%],')
}
W ('    [*Εισπράξεις*],[*' + (Format-Money $paySum) + '*],[],')
W '  )]],'
W ')'
$diff = $paySum - [decimal]$pay.Sales
if ([math]::Abs($diff) -ge 0.005) {
  W ('#text(size: 7.5pt, fill: WARN)[Εισπράξεις μείον αξία πωλήσεων: ' + (Format-Money $diff) + ' € — π.χ. εξόφληση παλιάς οφειλής.]')
}

# ── top προϊόντα ──
Sect 'Top προϊόντα της ημέρας'
W '#tbl((auto, 1fr, auto, auto, auto, auto), align: (right, left, center, right, right, right),'
W '  [*\#*],[*Προϊόν*],[*Είδος*],[*Αποδ.*],[*Τεμ.*],[*Αξία*],'
$i=0
foreach ($p in $top) {
  $i++
  $tag = if ([string]$p.VD -eq 'R') { '#text(size: 7pt, fill: MINT_DK)[Φ]' } else { '#text(size: 7pt, fill: SAND)[Π]' }
  W ("  [$i],[" + (TT ([string]$p.N)) + "],[$tag],[" + $p.R + "],[" + [int]$p.U + "],[" + (Format-Money $p.V) + "],")
}
W ')'
$vR = [decimal](($split | Where-Object { $_.VD -eq 'R' } | Select-Object -First 1).Val)
$vU = [decimal](($split | Where-Object { $_.VD -eq 'U' } | Select-Object -First 1).Val)
$vAll = $vR + $vU
if ($vAll -gt 0) {
  W ('#text(size: 7.5pt, fill: MUTED)[Φ = φάρμακα (μητρώο ΕΟΦ) ' + (Format-Money $vR) + ' € · Π = παραφάρμακα ' + (Format-Money $vU) + ' € — δηλαδή ' + [math]::Round($vR/$vAll*100) + '% / ' + [math]::Round($vU/$vAll*100) + '% της αξίας λιανικής.]')
}
EndSect

# ── σύγκριση με τις ίδιες ημέρες ──
if ($baseDays.Count -gt 0) {
  Sect ('Οι τελευταίες ' + (Get-GreekDayPlural $D))
  $allD = @($baseDays | ForEach-Object { [pscustomobject]@{ D=[datetime]$_.D; T=[decimal]$_.T; Cur=$false } })
  $allD += [pscustomobject]@{ D=$D; T=$T; Cur=$true }
  $mx = [double](($allD | ForEach-Object { [double]$_.T } | Measure-Object -Maximum).Maximum)
  $BH = 42.0
  # Οριζόντιες στήλες αντί πίνακα: ίδια πληροφορία, ~1/3 του ύψους.
  W ('#grid(columns: (' + ((@('1fr') * $allD.Count) -join ', ') + '), gutter: 4pt,')
  foreach ($x in $allD) {
    $hh = if ($mx -gt 0) { [math]::Round($BH * [double]($x.T/$mx),1) } else { 0 }
    $c  = if ($x.Cur) { 'NAVY' } else { 'MINT_LT' }
    $w  = if ($x.Cur) { '"medium"' } else { '"regular"' }
    W '  align(center, stack(spacing: 2pt,'
    W ('    box(width: 100%, height: ' + $BH + 'pt)[#place(bottom, rect(width: 58%, height: ' + $hh + 'pt, fill: ' + $c + ', stroke: none))],')
    W ('    text(size: 6.5pt, fill: MUTED)[' + $x.D.ToString('dd/MM') + '],')
    W ('    text(size: 6.5pt, weight: ' + $w + ')[' + (Format-Money $x.T) + '])),')
  }
  W ')'
  W ('#text(size: 7pt, fill: MUTED)[Μέσος όρος ' + (Format-Money $bT) + ' € (εξαιρούνται μεμονωμένες συναλλαγές ≥ 10.000 €). Η σκούρα στήλη είναι η ημέρα της αναφοράς.]')
  EndSect
}

# ── αξιοσημείωτα ──
if ($big.Count -gt 0) {
  $peak = $hToday | Sort-Object { [decimal]$_.T } -Descending | Select-Object -First 1
  W '#v(2pt)'
  W '#block(fill: ZEBRA, width: 100%, inset: 8pt, radius: 3pt)['
  W ('  #text(size: 8.5pt, weight: "medium", fill: ACCENT)[Αξιοσημείωτα] \')
  W ('  #text(size: 8.5pt)[Κορυφαία απόδειξη: *' + (Format-Money $big[0].A) + ' €* (' + ([datetime]$big[0].DT).ToString('HH:mm') + ', ' + (TT ([string]$big[0].M)) + ')' +
      ' #h(12pt) Ώρα αιχμής: *' + $peak.H + ':00* (' + (Format-Money $peak.T) + ' €)' +
      ' #h(12pt) Ακυρώσεις: *' + $tot.CANC + '*]]')
}
W '#v(4pt)'
W ('#text(size: 7pt, fill: MUTED)[«Πωλήσεις» = πληρωτέο πελατών (λιανική + συμμετοχές)· δεν περιλαμβάνεται η αποζημίωση ΕΟΠΥΥ. «Τυπικό» = μέσος όρος των τελευταίων ' + $bN + ' ίδιων ημερών εβδομάδας.]')

# ── compile ──────────────────────────────────────────────────────────────
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$base = Join-Path (Resolve-Path $OutDir) ("Ημερήσια_" + $D.ToString('yyyy-MM-dd'))
$typ = "$base.typ"; $pdf = "$base.pdf"
[System.IO.File]::WriteAllText($typ, $sb.ToString(), (New-Object Text.UTF8Encoding $false))
Invoke-Typst -TypPath $typ -PdfPath $pdf

Write-Output "OK route=$($db.Route)"
Write-Output "PDF: $pdf"

if ($Email) {
  $body = @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#25303f">
<p>Η ημερήσια αναφορά για <b>$((Get-GreekDayName $D)) $($D.ToString('dd/MM/yyyy'))</b>:</p>
<table style="border-collapse:collapse;font-size:14px">
<tr><td style="padding:2px 12px 2px 0">Πωλήσεις</td><td><b>$(Format-Money $T) €</b> ($(Format-Pct $dT) vs τυπική)</td></tr>
<tr><td style="padding:2px 12px 2px 0">Αποδείξεις</td><td>$R ($(Format-Pct $deltaR))</td></tr>
<tr><td style="padding:2px 12px 2px 0">Μέσο καλάθι</td><td>$(Format-Money $avg) € ($(Format-Pct $dA))</td></tr>
</table></div>
"@
  $to = Send-PharmacyMail -Cfg $cfg -Subject ("Ημερήσια αναφορά " + $D.ToString('dd/MM/yyyy')) -BodyHtml $body -Files @($pdf)
  Write-Output "Email -> $to"
  Write-ReportLog -LogDir $LogDir -Message ("Ημερήσια " + $D.ToString('yyyy-MM-dd') + ": email -> " + $to)
}
Write-ReportLog -LogDir $LogDir -Message ("Ημερήσια " + $D.ToString('yyyy-MM-dd') + " OK — " + (Format-Money $T) + " €, " + $R + " αποδ. (route " + $db.Route + ")")
if (-not $NoOpen) { Invoke-Item $pdf }
