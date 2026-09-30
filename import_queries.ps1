<#
    import_queries.ps1 - load every query from NABAVA_QUERIES.m into NABAVA_PQ.xlsx

    Replaces the manual "paste 10 blocks into Napredni uredjivac" step.
    Re-runnable: existing queries of the same name are removed first, so this
    is safe to run after every edit to the .m file.

    ASCII only, deliberately: PowerShell 5.1 reads a BOM-less file as cp1252,
    where a UTF-8 em dash decodes to a smart quote and terminates a string.
    check_queries.py enforces this.

    Requires: Windows + Excel installed (uses the Excel COM object model).
    The workbook must NOT be open in Excel while this runs.

        powershell -ExecutionPolicy Bypass -File .\import_queries.ps1

    Queries are declared by marker comments in the .m file, one per block:
        query: <name> | load: connection
        query: <name> | load: sheet <Sheet>!$A$4
        query: <name> | load: connection | derive-from: <other> | replace-once: <a> => <b>
#>
param(
    [string]$Workbook = ".\NABAVA_PQ.xlsx",
    [string]$MFile    = ".\NABAVA_QUERIES.m"
)

$ErrorActionPreference = "Stop"

$wbPath = (Resolve-Path $Workbook).Path
$mPath  = (Resolve-Path $MFile).Path
$text   = [System.IO.File]::ReadAllText($mPath, [System.Text.Encoding]::UTF8)

# ---- parse the .m file into ordered query definitions -------------------
# Do NOT name a variable $matches here: it is a PowerShell automatic variable
# and the `switch -Regex` below silently overwrites it.
$markerRx = [regex]'(?m)^/\*@\s*query:\s*(?<spec>[^*]+?)\s*\*/\s*$'
$markers  = @($markerRx.Matches($text))
if ($markers.Count -eq 0) { throw "No query markers found in $mPath" }

$defs = @()
for ($i = 0; $i -lt $markers.Count; $i++) {
    $m     = $markers[$i]
    $start = $m.Index + $m.Length
    $end   = if ($i + 1 -lt $markers.Count) { $markers[$i + 1].Index } else { $text.Length }
    $body  = $text.Substring($start, $end - $start)

    # drop the banner comment that belongs to the NEXT query
    $body = [regex]::Replace($body, '(?s)(\s*/\*((?!\*/).)*\*/\s*)+$', '')
    $body = $body.Trim()

    $parts = $m.Groups['spec'].Value -split '\|'
    $def = [pscustomobject]@{
        Name = $parts[0].Trim(); Load = 'connection'
        Sheet = $null; Cell = $null; DeriveFrom = $null; Find = $null; Replace = $null
        Body = $body
    }
    foreach ($p in $parts[1..($parts.Count - 1)]) {
        $p = $p.Trim()
        switch -Regex ($p) {
            '^load:\s*connection$' { $def.Load = 'connection' }
            '^load:\s*sheet\s+(.+?)!(\$?\w+\$?\d+)$' {
                $g = $Matches; $def.Load = 'sheet'; $def.Sheet = $g[1]; $def.Cell = $g[2]
            }
            '^derive-from:\s*(.+)$' { $def.DeriveFrom = $Matches[1].Trim() }
            '^replace-once:\s*(.+?)\s*=>\s*(.+)$' {
                $g = $Matches; $def.Find = $g[1].Trim(); $def.Replace = $g[2].Trim()
            }
            default { Write-Warning "$($def.Name): unrecognised directive '$p'" }
        }
    }
    $defs += $def
}

# ---- resolve derived queries -------------------------------------------
foreach ($d in $defs) {
    if ($d.DeriveFrom) {
        $src = $defs | Where-Object { $_.Name -eq $d.DeriveFrom }
        if (-not $src)    { throw "$($d.Name): derive-from '$($d.DeriveFrom)' not found" }
        if (-not $d.Find) { throw "$($d.Name): derive-from without replace-once" }
        $hits = ([regex]::Matches($src.Body, [regex]::Escape($d.Find))).Count
        if ($hits -ne 1) { throw "$($d.Name): anchor '$($d.Find)' matched $hits times in $($src.Name), expected exactly 1" }
        $d.Body = $src.Body.Replace($d.Find, $d.Replace)
    }
    if (-not $d.Body) { throw "$($d.Name): empty query body" }
}

Write-Host ("Parsed {0} queries: {1}" -f $defs.Count, (($defs | ForEach-Object { $_.Name }) -join ', '))

# ---- NABAVA sheet layout -------------------------------------------------
# The colour rules, title and frozen panes are (re)applied here on every run,
# so they always sit on the columns the query actually loads to, the same as
# in NABAVA_model_ver03.xlsx. Rule types that take no worksheet formula are
# used on purpose: a CF formula passed through COM is read in the UI language
# (ISNUMBER vs ISBROJ, "," vs ";") and relative to the active cell.
function Get-Bgr([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }

function Add-Cf($fc, $fill, $font) {
    $fc.Interior.Color = $fill
    if ($font -ne $null) { $fc.Font.Color = $font }
}

function Set-Look($rng, [string]$font, [double]$size, [bool]$bold, $color, $fill, [int]$hAlign) {
    # font name/size left to the workbook (Calibri) unless one is given
    if ($font -ne "") { $rng.Font.Name = $font }
    if ($size -gt 0)  { $rng.Font.Size = $size }
    $rng.Font.Bold = $bold
    if ($color -ne $null) { $rng.Font.Color = $color }
    if ($fill -ne $null)  { $rng.Interior.Color = $fill }
    if ($hAlign -ne 0)    { $rng.HorizontalAlignment = $hAlign }
}

function Set-NabavaLayout($excel, $wb, $ws, $lo) {
    $zh = [char]0x017E   # z with caron
    $dot = [char]0x00B7  # middle dot
    $grey   = Get-Bgr 89 89 89
    $white  = Get-Bgr 255 255 255
    $navy   = Get-Bgr 31 78 121
    $blue   = Get-Bgr 74 125 171
    $line   = Get-Bgr 191 191 191
    $pale   = Get-Bgr 234 241 248
    $center = -4108      # xlCenter
    $left   = -4131      # xlLeft

    $ws.Range("A1").Value2 = "NABAVA - planiranje narud" + $zh + "bi"
    Set-Look $ws.Range("A1") "" 14 $true $null $null 0
    $ws.Range("A2").Formula = '="Prag "&Prag&" mjeseci ' + $dot + ' roba na brodu se NE pribraja zalihi ' + $dot + ' do tri dolaska po artiklu"'
    Set-Look $ws.Range("A2") "" 9 $false $grey $null 0
    $ws.Rows.Item(1).RowHeight = 17.35
    $ws.Rows.Item(4).RowHeight = 35.05

    # column widths from ver03; R is the spacer
    $widths = @{ A=14; B=32; C=9; D=28; E=11; F=11; G=11; H=10; I=11; J=9; K=10;
                 L=9; M=10; N=9; O=10; P=11; Q=28; R=3 }
    foreach ($k in $widths.Keys) { $ws.Columns.Item($k).ColumnWidth = $widths[$k] }
    $ws.Range("S:AD").ColumnWidth = 7

    if ($lo -ne $null) {
        # no table style: ver03 is plain cells with grey borders, no banding.
        # Fonts stay Calibri on purpose; only colours, weight and borders follow ver03.
        $lo.TableStyle = ""
        try { $lo.QueryTable.PreserveFormatting = $true } catch { }

        $hdr = $ws.Range("A4:Q4")
        Set-Look $hdr "" 0 $true $white $navy $center
        $hdr.VerticalAlignment = $center
        $hdr.WrapText = $true
        $mh = $ws.Range("S4:AD4")
        Set-Look $mh "" 0 $true $white $blue $center
        $mh.VerticalAlignment = $center
        foreach ($r in @($hdr, $mh)) {
            $r.Borders.LineStyle = 1; $r.Borders.Weight = 2; $r.Borders.Color = $line
        }
        # spacer header: invisible
        Set-Look $ws.Range("R4") "" 0 $false $white $white 0

        $n = $lo.ListRows.Count
        if ($n -gt 0) {
            $last = 4 + $n
            $body = $ws.Range("A5:Q$last")
            Set-Look $body "" 0 $false $null $null $center
            $body.Borders.LineStyle = 1; $body.Borders.Weight = 2; $body.Borders.Color = $line
            Set-Look $ws.Range("A5:A$last") "" 0 $false $null $null $left
            $ws.Range("B5:B$last").HorizontalAlignment = $left
            $ws.Range("D5:D$last").HorizontalAlignment = $left
            $ws.Range("Q5:Q$last").HorizontalAlignment = $left
            $ws.Range("E5:G$last").Interior.Color = $pale
            foreach ($c in @("F", "I", "Q")) { $ws.Range("${c}5:$c$last").Font.Bold = $true }

            $ws.Range("R5:R$last").Borders.LineStyle = -4142   # xlNone

            $mb = $ws.Range("S5:AD$last")
            Set-Look $mb "" 0 $false $grey $null $center
            $mb.Borders.LineStyle = 1; $mb.Borders.Weight = 2; $mb.Borders.Color = $line
        }
    }

    $red    = Get-Bgr 248 203 203; $redF    = Get-Bgr 156 0 6
    $orange = Get-Bgr 252 228 196; $orangeF = Get-Bgr 138 75 0
    $green  = Get-Bgr 212 237 218; $greenF  = Get-Bgr 20 83 45
    $yellow = Get-Bgr 255 240 199
    $m = [Type]::Missing

    # STATUS (Q): xlTextString (9), xlContains (0)
    $q = $ws.Range("Q5:Q5000")
    Add-Cf ($q.FormatConditions.Add(9, $m, $m, $m, "Rupa", 0))          $red    $redF
    Add-Cf ($q.FormatConditions.Add(9, $m, $m, $m, "odmah", 0))         $red    $redF
    Add-Cf ($q.FormatConditions.Add(9, $m, $m, $m, "nije dovoljna", 0)) $orange $orangeF
    Add-Cf ($q.FormatConditions.Add(9, $m, $m, $m, "Sve u redu", 0))    $green  $greenF
    Add-Cf ($q.FormatConditions.Add(9, $m, $m, $m, "na vrijeme", 0))    $green  $greenF

    # Zaliha traje (I): xlCellValue (1). 999 means "no sales", so below the
    # threshold already implies sales > 0, as the ver03 rule AND(I<Prag,H>0) did.
    $i = $ws.Range("I5:I5000")
    Add-Cf ($i.FormatConditions.Add(1, 6, "=Prag"))          $red   $redF     # xlLess
    Add-Cf ($i.FormatConditions.Add(1, 1, "=Prag", "=899"))  $green $greenF   # xlBetween

    # Izlaz (G): xlCellValue (1), xlGreater (5)
    $g = $ws.Range("G5:G5000")
    Add-Cf ($g.FormatConditions.Add(1, 5, "=0")) $yellow $null

    try {
        $ws.Activate()
        $win = $wb.Windows.Item(1)
        $win.FreezePanes = $false
        $win.ScrollRow = 1; $win.ScrollColumn = 1
        $win.SplitColumn = 2; $win.SplitRow = 4
        $win.FreezePanes = $true
        $win.DisplayGridlines = $false
    } catch { Write-Warning "could not freeze panes on $($ws.Name): $_" }
}

# ---- push them into the workbook ---------------------------------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb    = $excel.Workbooks.Open($wbPath)
    $names = @($defs | ForEach-Object { $_.Name })

    # clear previous incarnations: sheet tables first, then queries, then connections
    foreach ($ws in $wb.Worksheets) {
        foreach ($lo in @($ws.ListObjects)) {
            $cmd = $null
            try { $cmd = $lo.QueryTable.CommandText } catch { continue }  # plain table (tblDolasci), skip
            foreach ($n in $names) {
                if ($cmd -match ("\[" + [regex]::Escape($n) + "\]")) {
                    Write-Host ("  - dropping old table for {0} on {1}" -f $n, $ws.Name)
                    $lo.Delete(); break
                }
            }
        }
    }
    foreach ($n in $names) {
        foreach ($q in @($wb.Queries)) {
            if ($q.Name -eq $n) {
                try { $q.Delete() } catch { Write-Warning "could not delete query ${n}: $_" }
            }
        }
        foreach ($c in @($wb.Connections)) {
            if ($c.Name -eq $n -or $c.Name -eq "Query - $n") {
                try { $c.Delete() } catch { }   # usually already removed with the query
            }
        }
    }

    # POSTAVKE!B8 (RollingN) now caps the completed months of the average, as
    # ver03's "Broj zavrsenih mjeseci prodaje"; relabel it in older files
    try {
        $ps = $wb.Worksheets.Item("POSTAVKE")
        $ps.Range("A8").Value2 = "Broj zavrsenih mjeseci prodaje (najvise)"
        $ps.Range("C8").Value2 = "12 = svi zavrseni mjeseci ove godine"
    } catch { Write-Warning "could not relabel POSTAVKE!A8: $_" }

    foreach ($d in $defs) {
        $wb.Queries.Add($d.Name, $d.Body) | Out-Null
        Write-Host ("  + {0}" -f $d.Name)
    }

    foreach ($d in @($defs | Where-Object { $_.Load -eq 'sheet' })) {
        $ws   = $wb.Worksheets.Item($d.Sheet)
        if ($d.Name -eq 'qNabava') {
            # start from a clean sheet: an earlier import may have shifted the
            # title and the colour rules to the right of the table
            [void]$ws.Cells.FormatConditions.Delete()
            [void]$ws.Cells.Clear()
        }
        $dest = $ws.Range($d.Cell)
        $conn = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=' + $d.Name + ';Extended Properties=""'
        $lo   = $ws.ListObjects.Add(0, $conn, $null, 1, $dest)   # xlSrcExternal, xlYes
        $qt   = $lo.QueryTable
        $qt.CommandType       = 2                                 # xlCmdSql
        $qt.CommandText       = "SELECT * FROM [$($d.Name)]"
        $qt.BackgroundQuery   = $false
        $qt.AdjustColumnWidth = $false
        # xlOverwriteCells (0). The default, xlInsertDeleteCells, inserts the
        # table's columns and pushes everything already on the sheet to the right.
        $qt.RefreshStyle      = 0
        # A refresh failure must not cost us the ten queries we just added, so the
        # queries and the table are kept and saved either way.
        try {
            $qt.Refresh($false) | Out-Null
            Write-Host ("  -> {0} loaded to {1}!{2} ({3} rows)" -f $d.Name, $d.Sheet, $d.Cell, $lo.ListRows.Count)
        }
        catch {
            $msg = "$_"
            Write-Warning ("{0} was created but the refresh failed: {1}" -f $d.Name, $msg)
            if ($msg -match 'may not directly access a data source|Formula\.Firewall|rebuild this data combination') {
                Write-Host ""
                Write-Host "This is the Power Query privacy firewall, not a bug in the M." -ForegroundColor Yellow
                Write-Host "Putanja reads the folder from a cell and fnDatoteke then hits the disk," -ForegroundColor Yellow
                Write-Host "which the firewall refuses to combine. To allow it:" -ForegroundColor Yellow
                Write-Host "  Excel -> Data -> Get Data -> Query Options -> Privacy" -ForegroundColor Yellow
                Write-Host "       -> Always ignore Privacy Level settings -> OK" -ForegroundColor Yellow
                Write-Host "Then refresh with Ctrl+Alt+F5, or re-run this script." -ForegroundColor Yellow
                Write-Host ""
            }
            $script:refreshFailed = $true
        }
        # applied after a failed refresh too: the rules sit on fixed columns
        if ($d.Name -eq 'qNabava') {
            Set-NabavaLayout $excel $wb $ws $lo
            Write-Host "  -> NABAVA layout: title, colour rules, frozen panes"
        }
    }

    $wb.Save()
    $wb.Close($true)
    if ($refreshFailed) {
        Write-Host "Saved $wbPath - queries are IN the file, but the data did not refresh (see above)."
    } else {
        Write-Host "Saved $wbPath"
    }
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
