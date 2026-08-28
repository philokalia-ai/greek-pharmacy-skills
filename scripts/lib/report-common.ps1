<#
  Κοινός κώδικας για τις αναφορές (εβδομαδιαία / ημερήσια).
  Dot-source: . "$PSScriptRoot\lib\report-common.ps1"
  Read-only: μόνο SELECT, πάντα READ UNCOMMITTED.
#>

# ── .env ─────────────────────────────────────────────────────────────────
function Resolve-EuropharmacyEnv {
  param([string]$StartDir)
  $isOurs = { param($p) (Test-Path $p) -and (Select-String -Path $p -Pattern '^\s*EUROPHARMACY_DB_' -Quiet) }
  if ($env:EUROPHARMACY_ENV -and (& $isOurs $env:EUROPHARMACY_ENV)) { return $env:EUROPHARMACY_ENV }
  $d = if ($StartDir) { $StartDir } else { (Get-Location).Path }
  while ($d) {
    $p = Join-Path $d '.env'
    if (& $isOurs $p) { return $p }
    $par = Split-Path $d -Parent; if ($par -eq $d) { break }; $d = $par
  }
  $h = Join-Path $HOME '.europharmacy\.env'
  if (& $isOurs $h) { return $h }
  throw "Δεν βρέθηκε .env (δες skill europharmacy-setup)."
}

function Get-PharmacyConfig {
  param([string]$StartDir)
  $file = Resolve-EuropharmacyEnv -StartDir $StartDir
  $cfg = @{}
  # -Encoding UTF8: χωρίς αυτό το PS 5.1 διαβάζει ANSI και σπάνε οι ελληνικές τιμές.
  Get-Content $file -Encoding UTF8 | Where-Object { $_ -match '^\s*[^#].*=' } |
    ForEach-Object { $k,$v = $_ -split '=',2; $cfg[$k.Trim()] = $v.Trim() }
  $cfg['_envFile'] = $file
  $cfg['_envDir']  = Split-Path -Parent $file
  return $cfg
}

# ── βάση ─────────────────────────────────────────────────────────────────
function Open-PharmacyDb {
  # Retries: όταν το task ξεκινά λίγο μετά την εκκίνηση του υπολογιστή, το Tailscale
  # μπορεί να μην έχει συνδεθεί ακόμη. Χωρίς επαναλήψεις η αναφορά απλώς αποτυγχάνει.
  param([hashtable]$Cfg, [int]$Timeout = 20, [int]$Retries = 10, [int]$RetryDelaySec = 45, [string]$LogDir)
  for ($attempt = 1; $attempt -le $Retries; $attempt++) {
    foreach ($k in @('EUROPHARMACY_DB_CONNSTR','EUROPHARMACY_DB_CONNSTR_LAN')) {
      if (-not $Cfg[$k]) { continue }
      try {
        $cn = New-Object System.Data.SqlClient.SqlConnection ($Cfg[$k] + ";Connect Timeout=$Timeout")
        $cn.Open()
        $c = $cn.CreateCommand()
        $c.CommandText = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SET DATEFIRST 1;"
        $c.ExecuteNonQuery() | Out-Null
        if ($attempt -gt 1 -and $LogDir) { Write-ReportLog -LogDir $LogDir -Message "Σύνδεση OK ($k) στην προσπάθεια $attempt." }
        return [pscustomobject]@{ Connection = $cn; Route = $k }
      } catch { }
    }
    if ($attempt -lt $Retries) {
      if ($LogDir) { Write-ReportLog -LogDir $LogDir -Level 'WARN' -Message "Αποτυχία σύνδεσης (προσπάθεια $attempt/$Retries) — νέα δοκιμή σε ${RetryDelaySec}s." }
      Start-Sleep -Seconds $RetryDelaySec
    }
  }
  throw "Δεν υπάρχει σύνδεση με τη βάση μετά από $Retries προσπάθειες (server κλειστός ή Tailscale down)."
}

# ── ταξινόμηση παραστατικών ──────────────────────────────────────────────
# Ο πίνακας Transactions περιέχει ΟΛΑ τα παραστατικά μαζί: αποδείξεις λιανικής,
# δελτία λιανικής πώλησης, πιστωτικά, και τα μηνιαία τιμολόγια προς τον ΕΟΠΥΥ.
# Χωρίς διαχωρισμό, η ημέρα που τιμολογείται ο ΕΟΠΥΥ εμφανίζεται ως τζίρος 20x.
#
#   ΑΛΠ        είδος 1, F5=0 — απόδειξη λιανικής με σειρά/αριθμό
#   ΔΛΠ        είδος 1, F5=1 — δελτίο λιανικής πώλησης· χωρίς δικό του αριθμό
#                              (μοιράζεται αρίθμηση), κυρίως επί πιστώσει σε
#                              ονομαστικούς πελάτες — η απόδειξη εκδίδεται μετά
#   ΠΙΣΤΩΤΙΚΟ  είδος 2 — επιστροφές (αρνητικά ποσά)
#   ΕΟΠΥΥ      είδη 5, 9 — τιμολόγια προς τον ΕΟΠΥΥ
#   ΤΙΜΟΛΟΓΙΟ  είδη 4, 6, 7, 10 — ιατρεία, φαρμακαποθήκες, λοιπά
#   ΧΩΡΙΣ      συναλλαγή χωρίς καθόλου παραστατικό (σπάνιο)
#
# «Λιανική» = ΑΛΠ + ΔΛΠ + ΠΙΣΤΩΤΙΚΟ + ΧΩΡΙΣ. Αυτό συγκρίνεται με προηγούμενες
# περιόδους· τα τιμολόγια αναφέρονται ξεχωριστά ώστε να μην κρύβονται.
function Get-DocClassSql { param([string]$Alias = 'ce')
@"
CASE
  WHEN $Alias.[ΕΙΔΟΣ_ΠΑΡΑΣΤΑΤΙΚΟΥ] IS NULL THEN 'ΧΩΡΙΣ'
  WHEN $Alias.[ΕΙΔΟΣ_ΠΑΡΑΣΤΑΤΙΚΟΥ] = 1 AND ISNULL($Alias.F5,0) = 0 THEN 'ΑΛΠ'
  WHEN $Alias.[ΕΙΔΟΣ_ΠΑΡΑΣΤΑΤΙΚΟΥ] = 1 THEN 'ΔΛΠ'
  WHEN $Alias.[ΕΙΔΟΣ_ΠΑΡΑΣΤΑΤΙΚΟΥ] = 2 THEN 'ΠΙΣΤΩΤΙΚΟ'
  WHEN $Alias.[ΕΙΔΟΣ_ΠΑΡΑΣΤΑΤΙΚΟΥ] IN (5,9) THEN 'ΕΟΠΥΥ'
  ELSE 'ΤΙΜΟΛΟΓΙΟ'
END
"@
}
# Το LEFT JOIN που χρειάζεται ο παραπάνω CASE. Προ-αθροίζει ανά συναλλαγή ώστε
# μια συναλλαγή με >1 γραμμές CashExtras να μη διπλομετρηθεί.
function Get-DocJoinSql { param([string]$Alias = 'dc', [string]$TxAlias = 't')
@"
LEFT JOIN (SELECT [ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ] tx, MIN([ΕΙΔΟΣ_ΠΑΡΑΣΤΑΤΙΚΟΥ]) [ΕΙΔΟΣ_ΠΑΡΑΣΤΑΤΙΚΟΥ], MAX(ISNULL(F5,0)) F5
           FROM CashExtras WHERE ISNULL(Hidden,0)=0 GROUP BY [ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]) $Alias
       ON $Alias.tx = $TxAlias.[ΚΩΔ_ΣΥΝΑΛΛΑΓΗΣ]
"@
}
# Φίλτρο «μόνο λιανική» για WHERE.
function Get-RetailFilterSql { param([string]$Alias = 'ce')
  "(ISNULL($Alias.[ΕΙΔΟΣ_ΠΑΡΑΣΤΑΤΙΚΟΥ],1) IN (1,2))"
}

# ── μπάρα σύνθεσης τζίρου ────────────────────────────────────────────────
# [ΕΕΕΕ ΦΦΦΦΦΦ ΠΠΠΠΠ · f5]
#   Ε  = τιμολόγια (ΕΟΠΥΥ και λοιπά)
#   Φ  = φάρμακα λιανικής        Π = παραφάρμακα λιανικής
#   f5 = δελτία λιανικής πώλησης — χωριστά στο τέλος, μετά από κενό
# Τα τέσσερα τμήματα αθροίζουν στο ΣΥΝΟΛΟ, οπότε η μπάρα δείχνει πού πάει
# κάθε ευρώ του τζίρου.
function Get-TypstCompositionBar {
  param([decimal]$Eopyy, [decimal]$Pharma, [decimal]$Para, [decimal]$F5, [double]$Height = 17)
  $gr = [System.Globalization.CultureInfo]::GetCultureInfo('el-GR')
  $total = $Eopyy + $Pharma + $Para + $F5
  if ($total -le 0) { return '' }
  $segs = @(
    @{ k='Ε';  n='Τιμολόγια (ΕΟΠΥΥ κ.λπ.)'; v=$Eopyy;  c='NAVY' },
    @{ k='Φ';  n='Φάρμακα';                 v=$Pharma; c='MINT_DK' },
    @{ k='Π';  n='Παραφάρμακα';             v=$Para;   c='SAND' }
  ) | Where-Object { $_.v -gt 0 }
  $cols = @(); $cells = @()
  foreach ($s in $segs) {
    $cols += ([math]::Round([double]($s.v / $total) * 1000, 2).ToString([Globalization.CultureInfo]::InvariantCulture) + 'fr')
    $cells += ('box(width: 100%, height: ' + $Height + 'pt, fill: ' + $s.c + ')')
  }
  if ($F5 -gt 0) {
    $cols  += '5pt'                       # το «·»: οπτικός διαχωρισμός
    $cells += 'box(width: 100%)'
    $cols  += ([math]::Round([double]($F5 / $total) * 1000, 2).ToString([Globalization.CultureInfo]::InvariantCulture) + 'fr')
    $cells += ('box(width: 100%, height: ' + $Height + 'pt, fill: MINT_LT, stroke: 0.6pt + NAVY)')
  }
  $pc = { param($v) ([decimal]($v / $total * 100)).ToString('N1',$gr) + '%' }
  $legend = @()
  foreach ($s in $segs) {
    $legend += ('#box(width: 8pt, height: 8pt, fill: ' + $s.c + ') *' + $s.k + '* ' + $s.n + ' ' +
                (([decimal]$s.v).ToString('N2',$gr)) + ' € (' + (& $pc $s.v) + ')')
  }
  if ($F5 -gt 0) {
    $legend += ('#box(width: 8pt, height: 8pt, fill: MINT_LT, stroke: 0.6pt + NAVY) *f5* Δελτία λιαν. πώλησης ' +
                (([decimal]$F5).ToString('N2',$gr)) + ' € (' + (& $pc $F5) + ')')
  }
@"
#grid(columns: ($($cols -join ', ')), $($cells -join ', '))
#v(3pt)
#text(size: 7.5pt, fill: MUTED)[$($legend -join ' #h(10pt) ')]
"@
}

# ── logging ──────────────────────────────────────────────────────────────
# Χωρίς αρχείο καταγραφής, μια αποτυχία του scheduled task φαίνεται μόνο ως
# «LastTaskResult = 1» χωρίς καμία ένδειξη για την αιτία.
function Write-ReportLog {
  param([string]$LogDir, [string]$Message, [string]$Level = 'INFO')
  if (-not $LogDir) { return }
  try {
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    $line = '{0} [{1,-5}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path (Join-Path $LogDir ('report-' + (Get-Date).ToString('yyyy-MM') + '.log')) -Value $line -Encoding UTF8
  } catch { }
}

function Invoke-Rows {
  param($Connection, [string]$Sql, [int]$Timeout = 300)
  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql; $cmd.CommandTimeout = $Timeout
  $rd = $cmd.ExecuteReader(); $out = @()
  while ($rd.Read()) {
    $h = [ordered]@{}
    for ($i=0; $i -lt $rd.FieldCount; $i++) { $h[$rd.GetName($i)] = $rd.GetValue($i) }
    $out += [pscustomobject]$h
  }
  $rd.Close(); return ,$out
}

# ── μορφοποίηση ──────────────────────────────────────────────────────────
$script:GR = [System.Globalization.CultureInfo]::GetCultureInfo('el-GR')
function Format-Money { param($v)
  if ($null -eq $v -or $v -is [DBNull]) { '0,00' } else { ([decimal]$v).ToString('N2',$script:GR) } }
function Format-Num { param($v,[int]$dec=1)
  if ($null -eq $v -or $v -is [DBNull]) { '0' } else { ([decimal]$v).ToString("N$dec",$script:GR) } }
function Format-Pct { param([double]$x)
  $s = if ($x -ge 0) { '+' } else { [char]0x2212 }
  "$s$(([decimal][math]::Abs($x)).ToString('N1',$script:GR))%" }

# Escape για Typst content. ΠΡΟΣΟΧΗ: το backslash πρώτο και ΜΙΑ φορά διπλό —
# στο -replace το '\' δεν είναι ειδικό στο replacement, οπότε '\\' = δύο χαρακτήρες.
function ConvertTo-TypstText { param([string]$s)
  if ($null -eq $s) { return '' }
  ($s -replace '\\','\\' -replace '"','\"' -replace '#','\#' -replace '\$','\$' `
      -replace '@','\@' -replace '_','\_' -replace '\*','\*' -replace '<','\<' -replace '>','\>')
}

# Ελληνικά ονόματα ημερών από .NET DateTime — ΟΧΙ από DATENAME, που εξαρτάται
# από τη γλώσσα του SQL Server και θα χαλούσε αν άλλαζε.
function Get-GreekDayName { param([datetime]$d)
  @('Κυριακή','Δευτέρα','Τρίτη','Τετάρτη','Πέμπτη','Παρασκευή','Σάββατο')[[int]$d.DayOfWeek] }
function Get-GreekDayPlural { param([datetime]$d)
  @('Κυριακές','Δευτέρες','Τρίτες','Τετάρτες','Πέμπτες','Παρασκευές','Σάββατα')[[int]$d.DayOfWeek] }
function Get-GreekMonthName { param([int]$m)
  @('','Ιανουαρίου','Φεβρουαρίου','Μαρτίου','Απριλίου','Μαΐου','Ιουνίου','Ιουλίου',
    'Αυγούστου','Σεπτεμβρίου','Οκτωβρίου','Νοεμβρίου','Δεκεμβρίου')[$m] }

# ── παλέτα / Typst preamble ──────────────────────────────────────────────
function Get-Shade { param([string]$hex,[double]$f)
  $h = $hex.TrimStart('#')
  $r=[Convert]::ToInt32($h.Substring(0,2),16); $g=[Convert]::ToInt32($h.Substring(2,2),16); $b=[Convert]::ToInt32($h.Substring(4,2),16)
  if ($f -le 1) { $r=[int]($r*$f); $g=[int]($g*$f); $b=[int]($b*$f) }
  else { $t=$f-1; $r=[int]($r+(255-$r)*$t); $g=[int]($g+(255-$g)*$t); $b=[int]($b+(255-$b)*$t) }
  '#{0:x2}{1:x2}{2:x2}' -f [math]::Min(255,$r),[math]::Min(255,$g),[math]::Min(255,$b) }

function Get-Brand { param([hashtable]$Cfg)
  $p = if ($Cfg['EUROPHARMACY_COLOR_PRIMARY']) { $Cfg['EUROPHARMACY_COLOR_PRIMARY'] } else { '#3f4a5a' }
  $a = if ($Cfg['EUROPHARMACY_COLOR_ACCENT'])  { $Cfg['EUROPHARMACY_COLOR_ACCENT']  } else { '#7fc9bd' }
  $s = if ($Cfg['EUROPHARMACY_COLOR_SECOND'])  { $Cfg['EUROPHARMACY_COLOR_SECOND']  } else { '#d9a05b' }
  [pscustomobject]@{
    Primary=$p; Accent=$a; Second=$s
    AccentDark=(Get-Shade $a 0.70); AccentLight=(Get-Shade $a 1.55)
    Name=$(if ($Cfg['EUROPHARMACY_NAME']) { $Cfg['EUROPHARMACY_NAME'] } else { 'Φαρμακείο' })
  } }

function Get-TypstPreamble {
  param([hashtable]$Cfg, $Brand, [string]$FooterNote = 'read-only αναφορά', [string]$Margin = '(x: 1.4cm, top: 1.2cm, bottom: 1.4cm)')
@"
#let NAVY    = rgb("$($Brand.Primary)")
#let MINT    = rgb("$($Brand.Accent)")
#let MINT_DK = rgb("$($Brand.AccentDark)")
#let MINT_LT = rgb("$($Brand.AccentLight)")
#let SAND    = rgb("$($Brand.Second)")
#let ACCENT  = NAVY
#let ACCENT2 = rgb("#eaf0f4")
#let INK     = rgb("#25303f")
#let MUTED   = rgb("#78849a")
#let ZEBRA   = rgb("#fafbfc")
#let GOOD    = rgb("#15803d")
#let BAD     = rgb("#b42318")
#let WARN    = rgb("#a9702f")

#set page(paper: "a4", margin: $Margin,
  footer: context [#set text(size: 7.5pt, fill: MUTED)
    #line(length: 100%, stroke: 0.3pt + MUTED) #v(-4pt)
    #grid(columns: (1fr, auto), align: (left, right),
      [$(ConvertTo-TypstText $Brand.Name) · $FooterNote],
      [#counter(page).display("1 / 1", both: true)])])
#set text(font: ("Calibri", "Arial"), size: 9.5pt, lang: "el", fill: INK)
#show heading.where(level: 2): it => block(above: 11pt, below: 5pt)[#text(size: 11.5pt, weight: "medium", fill: ACCENT)[#it.body]]

#let up(x) = text(fill: GOOD, weight: "medium", x)
#let dn(x) = text(fill: BAD, weight: "medium", x)
#let tbl(cols, align: none, ..args) = table(
  columns: cols, align: align,
  fill: (x, y) => if y == 0 { ACCENT2 } else if calc.odd(y) { ZEBRA },
  stroke: (x, y) => (bottom: if y == 0 { 0.9pt + ACCENT } else { 0.25pt + rgb("#e5e7eb") }),
  inset: (x: 5pt, y: 4pt),
  ..args)
"@
}

# Κεφαλίδα με logo. Επιστρέφει Typst string· αντιγράφει το logo δίπλα στο .typ
# (το Typst δεν διαβάζει αρχεία εκτός του root του).
function Get-TypstHeader {
  param([hashtable]$Cfg, [string]$OutDir, [string]$Title, [string]$Subtitle, [string]$Note, [double]$BoxH = 54)
  $logoTypst = 'box(width: 0pt)'
  if ($Cfg['EUROPHARMACY_LOGO']) {
    $lp = $Cfg['EUROPHARMACY_LOGO']
    if (-not [System.IO.Path]::IsPathRooted($lp)) { $lp = Join-Path $Cfg['_envDir'] $lp }
    if (Test-Path $lp) {
      if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
      $local = '_logo' + [System.IO.Path]::GetExtension($lp)
      Copy-Item $lp (Join-Path (Resolve-Path $OutDir) $local) -Force
      $zoom = 1.0
      if ($Cfg['EUROPHARMACY_LOGO_ZOOM']) { $zoom = [double]::Parse($Cfg['EUROPHARMACY_LOGO_ZOOM'],[Globalization.CultureInfo]::InvariantCulture) }
      $imgH = [math]::Round($BoxH * $zoom, 1)
      $logoTypst = 'box(width: ' + ($BoxH*1.37) + 'pt, height: ' + $BoxH + 'pt, clip: true, align(center + horizon, image("' + $local + '", height: ' + $imgH + 'pt)))'
    } else { Write-Warning "EUROPHARMACY_LOGO: δεν βρέθηκε $lp — χωρίς logo." }
  }
@"
#block(fill: NAVY, width: 100%, inset: (x: 0pt, y: 0pt), radius: 4pt, clip: true)[
  #grid(columns: (auto, 1fr), align: (left + horizon, right + horizon), inset: (x: 14pt, y: 9pt),
    $logoTypst,
    align(right)[
      #text(fill: white, size: 16pt, weight: "medium")[$Title] \
      #text(fill: MINT, size: 10.5pt)[$Subtitle] \
      #text(fill: rgb("#9aa6bd"), size: 7.5pt)[$Note]
    ])]
"@
}

# ── compile ──────────────────────────────────────────────────────────────
function Invoke-Typst {
  param([string]$TypPath, [string]$PdfPath)
  $exe = @("$env:LOCALAPPDATA\Microsoft\WinGet\Links\typst.exe", 'typst') |
    Where-Object { Get-Command $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
  if (-not $exe) { throw "Δεν βρέθηκε το typst. Εγκατάσταση: winget install Typst.Typst" }
  & $exe compile $TypPath $PdfPath
  if ($LASTEXITCODE -ne 0) { throw "Απέτυχε το compile του typst ($TypPath)." }
}

# ── email ────────────────────────────────────────────────────────────────
function Send-PharmacyMail {
  param([hashtable]$Cfg, [string]$Subject, [string]$BodyHtml, [string[]]$Files)
  foreach ($need in @('EUROPHARMACY_SMTP_HOST','EUROPHARMACY_SMTP_USER','EUROPHARMACY_SMTP_PASS','EUROPHARMACY_SMTP_TO')) {
    if (-not $Cfg[$need]) { throw "Λείπει το $need από το .env." }
  }
  $port = if ($Cfg['EUROPHARMACY_SMTP_PORT']) { [int]$Cfg['EUROPHARMACY_SMTP_PORT'] } else { 587 }
  $msg = New-Object Net.Mail.MailMessage
  $msg.From = New-Object Net.Mail.MailAddress($Cfg['EUROPHARMACY_SMTP_USER'], 'Αναφορές Φαρμακείου')
  foreach ($to in ($Cfg['EUROPHARMACY_SMTP_TO'] -split '[;,]')) { if ($to.Trim()) { $msg.To.Add($to.Trim()) } }
  $msg.Subject = $Subject
  $msg.SubjectEncoding = [Text.Encoding]::UTF8
  $msg.BodyEncoding    = [Text.Encoding]::UTF8
  $msg.IsBodyHtml = $true
  $msg.Body = $BodyHtml
  foreach ($f in $Files) { if ($f -and (Test-Path $f)) { $msg.Attachments.Add((New-Object Net.Mail.Attachment($f))) } }
  $smtp = New-Object Net.Mail.SmtpClient($Cfg['EUROPHARMACY_SMTP_HOST'], $port)
  $smtp.EnableSsl = $true
  $smtp.Credentials = New-Object Net.NetworkCredential($Cfg['EUROPHARMACY_SMTP_USER'], $Cfg['EUROPHARMACY_SMTP_PASS'])
  try { $smtp.Send($msg) } finally { $msg.Dispose(); $smtp.Dispose() }
  return ($Cfg['EUROPHARMACY_SMTP_TO'])
}
