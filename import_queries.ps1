<#
    import_queries.ps1 — load every query from NABAVA_QUERIES.m into NABAVA_PQ.xlsx

    Replaces the manual "paste 10 blocks into Napredni uredjivac" step.
    Re-runnable: existing queries of the same name are removed first, so this
    is safe to run after every edit to the .m file.

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

    foreach ($d in $defs) {
        $wb.Queries.Add($d.Name, $d.Body) | Out-Null
        Write-Host ("  + {0}" -f $d.Name)
    }

    foreach ($d in @($defs | Where-Object { $_.Load -eq 'sheet' })) {
        $ws   = $wb.Worksheets.Item($d.Sheet)
        $dest = $ws.Range($d.Cell)
        $conn = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=' + $d.Name + ';Extended Properties=""'
        $lo   = $ws.ListObjects.Add(0, $conn, $null, 1, $dest)   # xlSrcExternal, xlYes
        $qt   = $lo.QueryTable
        $qt.CommandType       = 2                                 # xlCmdSql
        $qt.CommandText       = "SELECT * FROM [$($d.Name)]"
        $qt.BackgroundQuery   = $false
        $qt.AdjustColumnWidth = $false
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
    }

    $wb.Save()
    $wb.Close($true)
    if ($refreshFailed) {
        Write-Host "Saved $wbPath — queries are IN the file, but the data did not refresh (see above)."
    } else {
        Write-Host "Saved $wbPath"
    }
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
