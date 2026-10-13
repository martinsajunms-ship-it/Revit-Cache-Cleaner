<# :
@echo off
setlocal
title Revit Collaboration Cache Cleaner
set "SCRIPT_PATH=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ([System.IO.File]::ReadAllText($env:SCRIPT_PATH))"
echo.
pause
exit /b
#>

# =====================================================================
#  REVIT COLLABORATION CACHE CLEANER  -  Version 1.0
#  Free to use at your own risk. Not affiliated with Autodesk.
#  1. Scans the Revit collaboration cache and finds the real model
#     name behind every code-named cache file.
#  2. Shows a numbered list of models on screen.
#  3. Clears the cache of the model(s) you choose - from the main
#     folder, CentralCache and LinkedModels.
#  Files are sent to the Windows RECYCLE BIN, so they can be restored.
#  An Excel report can be created at any time (press E in the menu).
# =====================================================================

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$GuidPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

# ---------------------------------------------------------------------
#  Read the real model name stored inside a cached .rvt file
# ---------------------------------------------------------------------
function Get-ModelInfoFromRvt([string]$Path) {
    $keys   = @('Central Model Path:', 'Last Save Path:')
    $found  = @{}
    $fs     = $null
    try {
        $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        $buf     = New-Object byte[] (8MB)
        $overlap = 16384
        $carry   = 0
        while ($true) {
            $read = $fs.Read($buf, $carry, $buf.Length - $carry)
            $atEnd = ($read -le 0)
            $total = $carry + [Math]::Max($read, 0)
            if ($total -le 1) { break }
            foreach ($shift in 0, 1) {
                $text = [System.Text.Encoding]::Unicode.GetString($buf, $shift, $total - $shift)
                foreach ($k in $keys) {
                    if ($found.ContainsKey($k)) { continue }
                    $i = $text.IndexOf($k, [System.StringComparison]::Ordinal)
                    if ($i -lt 0) { continue }
                    # if the text is cut at the end of this chunk, wait for the next chunk
                    if (-not $atEnd -and ($text.Length - $i) -lt 1000) { continue }
                    $part = $text.Substring($i + $k.Length, [Math]::Min(600, $text.Length - $i - $k.Length))
                    $m = [regex]::Match($part, '^\s*([^\r\n\x00]+)')
                    if ($m.Success -and $m.Groups[1].Value.Trim().Length -gt 2) { $found[$k] = $m.Groups[1].Value.Trim() }
                }
            }
            if ($found.ContainsKey('Central Model Path:') -or $atEnd) { break }
            if ($total -gt $overlap) {
                [Array]::Copy($buf, $total - $overlap, $buf, 0, $overlap)
                $carry = $overlap
            } else { $carry = $total }
        }
    } catch {
        return $null
    } finally {
        if ($fs) { $fs.Dispose() }
    }

    foreach ($k in $keys) {
        if (-not $found.ContainsKey($k)) { continue }
        $value = $found[$k]
        $parts = $value -split '[\\/]' | Where-Object { $_ -ne '' }
        if (-not $parts) { continue }
        $name = $parts[-1]
        if ($name -match "^$GuidPattern(\.rvt)?$") { continue }   # only a code, not a real name
        $project = ''
        if ($value -match '://' -and $parts.Count -ge 3) { $project = $parts[-2] }
        return [pscustomobject]@{ Name = $name; Project = $project; FullValue = $value }
    }
    return $null
}

# ---------------------------------------------------------------------
#  Simple Excel (.xlsx) writer - works even if Excel is not installed
# ---------------------------------------------------------------------
function ConvertTo-XmlText([string]$s) {
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [System.Security.SecurityElement]::Escape($s)
}
function Get-ColLetter([int]$n) {
    $s = ''
    while ($n -gt 0) { $r = ($n - 1) % 26; $s = [char](65 + $r) + $s; $n = [int][Math]::Floor(($n - 1) / 26) }
    return $s
}
function Get-SheetXml($sheet) {
    $sb = New-Object System.Text.StringBuilder
    $cols = $sheet.Headers.Count
    $lastCol = Get-ColLetter $cols
    $lastRow = $sheet.Rows.Count + 1
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.Append('<cols>')
    for ($c = 1; $c -le $cols; $c++) {
        [void]$sb.Append("<col min=`"$c`" max=`"$c`" width=`"$($sheet.Widths[$c-1])`" customWidth=`"1`"/>")
    }
    [void]$sb.Append('</cols><sheetData>')
    # header row
    [void]$sb.Append('<row r="1">')
    for ($c = 1; $c -le $cols; $c++) {
        $ref = (Get-ColLetter $c) + '1'
        [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`" s=`"1`"><is><t>$(ConvertTo-XmlText $sheet.Headers[$c-1])</t></is></c>")
    }
    [void]$sb.Append('</row>')
    $base = New-Object DateTime(1899, 12, 30)
    $r = 1
    foreach ($row in $sheet.Rows) {
        $r++
        [void]$sb.Append("<row r=`"$r`">")
        for ($c = 1; $c -le $cols; $c++) {
            $ref = (Get-ColLetter $c) + $r
            $v = $row[$c - 1]
            switch ($sheet.Types[$c - 1]) {
                'd' { $num = [Math]::Floor(($v - $base).TotalDays); [void]$sb.Append("<c r=`"$ref`" s=`"2`"><v>$num</v></c>") }
                't' { $num = $v.TimeOfDay.TotalDays.ToString([Globalization.CultureInfo]::InvariantCulture); [void]$sb.Append("<c r=`"$ref`" s=`"3`"><v>$num</v></c>") }
                'n' { $num = ([double]$v).ToString([Globalization.CultureInfo]::InvariantCulture); [void]$sb.Append("<c r=`"$ref`" s=`"4`"><v>$num</v></c>") }
                'i' { if ("$v" -match '^\d+$') { [void]$sb.Append("<c r=`"$ref`"><v>$v</v></c>") } else { [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t>$(ConvertTo-XmlText ([string]$v))</t></is></c>") } }
                default { [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText ([string]$v))</t></is></c>") }
            }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    [void]$sb.Append("<autoFilter ref=`"A1:$lastCol$lastRow`"/>")
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}
function Save-Xlsx([string]$Path, $Sheets) {
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $add = {
        param($name, $content)
        $e = $zip.CreateEntry($name)
        $w = New-Object System.IO.StreamWriter($e.Open(), $utf8)
        $w.Write($content); $w.Dispose()
    }
    try {
        $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        $wbSheets = ''; $wbRels = ''
        for ($i = 1; $i -le $Sheets.Count; $i++) {
            $ct += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
            $wbSheets += "<sheet name=`"$(ConvertTo-XmlText $Sheets[$i-1].Name)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
            $wbRels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
        }
        $ct += '</Types>'
        $styleId = $Sheets.Count + 1
        $wbRels += "<Relationship Id=`"rId$styleId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>"

        & $add '[Content_Types].xml' $ct
        & $add '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        & $add 'xl/workbook.xml' "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets>$wbSheets</sheets></workbook>"
        & $add 'xl/_rels/workbook.xml.rels' "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$wbRels</Relationships>"
        & $add 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="dd-mmm-yyyy"/><numFmt numFmtId="165" formatCode="h:mm:ss AM/PM"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'
        for ($i = 1; $i -le $Sheets.Count; $i++) {
            & $add "xl/worksheets/sheet$i.xml" (Get-SheetXml $Sheets[$i-1])
        }
    } finally {
        $zip.Dispose()
    }
}


# ---------------------------------------------------------------------
#  Small helpers
# ---------------------------------------------------------------------
$TestMode = [bool]$env:RCR_TEST
$script:Answers = $null
$script:AnsIdx  = 0
if ($env:RCR_ANSWERS) { $script:Answers = @($env:RCR_ANSWERS -split '\|') }

function Ask([string]$Prompt) {
    if ($null -ne $script:Answers) {
        $a = 'Q'
        if ($script:AnsIdx -lt $script:Answers.Count) { $a = $script:Answers[$script:AnsIdx]; $script:AnsIdx++ }
        Write-Host "$Prompt : $a"
        return $a
    }
    return (Read-Host $Prompt)
}

function Format-Size([double]$Bytes) {
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    return ('{0:N0} KB' -f [Math]::Ceiling($Bytes / 1KB))
}

function Limit-Text([string]$s, [int]$n) {
    if ($null -eq $s) { return '' }
    if ($s.Length -le $n) { return $s }
    return $s.Substring(0, $n - 3) + '...'
}

function Test-RevitRunning {
    if ($TestMode) { return [bool]$env:RCR_FAKE_REVIT }
    return [bool](Get-Process -Name 'Revit' -ErrorAction SilentlyContinue)
}

# ---------------------------------------------------------------------
#  Recycle Bin support
# ---------------------------------------------------------------------
$script:ShellReady = $false
if (-not $TestMode) {
    try {
        Add-Type -ErrorAction Stop -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace RcrNative {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct ShFileOp {
        public IntPtr hwnd;
        [MarshalAs(UnmanagedType.U4)] public int wFunc;
        public string pFrom;
        public string pTo;
        public short fFlags;
        [MarshalAs(UnmanagedType.Bool)] public bool fAnyOperationsAborted;
        public IntPtr hNameMappings;
        public string lpszProgressTitle;
    }
    public static class Shell {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern int SHFileOperation(ref ShFileOp op);
    }
}
"@
        $script:ShellReady = $true
    } catch { $script:ShellReady = $false }
    try { Add-Type -AssemblyName Microsoft.VisualBasic } catch { }
}

# Sends one file or folder to the Recycle Bin. Returns $true if it is gone.
function Send-ToRecycleBin([string]$Path, [bool]$IsFolder) {
    if ($TestMode) {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
        return (-not (Test-Path -LiteralPath $Path))
    }
    try {
        if ($script:ShellReady) {
            $op = New-Object RcrNative.ShFileOp
            $op.wFunc  = 3                                   # delete
            $op.pFrom  = $Path + [char]0
            # 0x0040 = to Recycle Bin, 0x0010 = no "are you sure" box,
            # 0x4000 = BUT warn if a file is too big for the Recycle Bin
            $op.fFlags = [int16](0x0040 -bor 0x0010 -bor 0x4000)
            [void][RcrNative.Shell]::SHFileOperation([ref]$op)
        } else {
            $ui = [Microsoft.VisualBasic.FileIO.UIOption]::AllDialogs
            $rb = [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
            if ($IsFolder) { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($Path, $ui, $rb) }
            else           { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($Path, $ui, $rb) }
        }
    } catch { }
    return (-not (Test-Path -LiteralPath $Path))
}

# ---------------------------------------------------------------------
#  Log file (Documents\Revit_Cache_Cleaner_Log.txt)
# ---------------------------------------------------------------------
$logDir = if ($env:RCR_OUT) { $env:RCR_OUT } else { [Environment]::GetFolderPath('MyDocuments') }
if (-not $logDir -or -not (Test-Path -LiteralPath $logDir)) { $logDir = $env:USERPROFILE }
$LogFile = Join-Path $logDir 'Revit_Cache_Cleaner_Log.txt'
function Write-Log([string]$Line) {
    try { Add-Content -LiteralPath $LogFile -Value $Line -Encoding UTF8 } catch { }
}

# =====================================================================
#  START
# =====================================================================
Write-Host ''
Write-Host '=============================================================' -ForegroundColor Cyan
Write-Host '   REVIT COLLABORATION CACHE CLEANER' -ForegroundColor Cyan
Write-Host '=============================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host 'Press ENTER to scan ALL Revit versions on this PC (recommended),'
Write-Host 'or paste a CollaborationCache folder path and press Enter.'
Write-Host 'Example: C:\Users\YourName\AppData\Local\Autodesk\Revit\Autodesk Revit 2026\CollaborationCache' -ForegroundColor DarkGray
Write-Host ''

if ($env:RCR_PATH) { $inputPath = $env:RCR_PATH } else { $inputPath = Read-Host 'Folder path' }
$inputPath = ([string]$inputPath).Trim().Trim('"').Trim("'").Trim()

$roots = @()
if ([string]::IsNullOrWhiteSpace($inputPath)) {
    $revitRoot = ''
    if ($env:LOCALAPPDATA) { $revitRoot = Join-Path $env:LOCALAPPDATA 'Autodesk\Revit' }
    if ($revitRoot -and (Test-Path -LiteralPath $revitRoot)) {
        $roots = @(Get-ChildItem -LiteralPath $revitRoot -Directory -Force -ErrorAction SilentlyContinue |
                   ForEach-Object { Join-Path $_.FullName 'CollaborationCache' } |
                   Where-Object { Test-Path -LiteralPath $_ -PathType Container })
    }
    $relBase = $revitRoot
    if ($roots.Count -eq 0) {
        Write-Host 'No Revit collaboration cache was found on this PC.' -ForegroundColor Yellow
        return
    }
    Write-Host "Scanning all Revit versions in: $revitRoot" -ForegroundColor Yellow
} else {
    if (-not (Test-Path -LiteralPath $inputPath -PathType Container)) {
        Write-Host ''
        Write-Host "Sorry, this folder was not found: $inputPath" -ForegroundColor Red
        Write-Host 'Please check the path and run the tool again.' -ForegroundColor Red
        return
    }
    $roots   = @((Resolve-Path -LiteralPath $inputPath).ProviderPath.TrimEnd('\', '/'))
    $relBase = $roots[0]
}
$sep = [System.IO.Path]::DirectorySeparatorChar

# ---------------------------------------------------------------------
#  STEP 1 - list everything
# ---------------------------------------------------------------------
Write-Host ''
Write-Host 'Scanning the cache, please wait ...' -ForegroundColor Green
$allItems = New-Object System.Collections.Generic.List[object]
foreach ($r in $roots) {
    $allItems.Add((Get-Item -LiteralPath $r -Force))
    foreach ($it in @(Get-ChildItem -LiteralPath $r -Recurse -Force -ErrorAction SilentlyContinue)) { $allItems.Add($it) }
}
$files = @($allItems | Where-Object { -not $_.PSIsContainer })
$dirs  = @($allItems | Where-Object { $_.PSIsContainer })
if ($files.Count -eq 0) {
    Write-Host 'The cache is empty. Nothing to clear.' -ForegroundColor Yellow
    return
}

# children of every folder
$children = @{}
foreach ($it in $allItems) {
    $parent = if ($it.PSIsContainer) { if ($it.Parent) { $it.Parent.FullName } else { '' } } else { $it.DirectoryName }
    $k = $parent.ToLower()
    if (-not $children.ContainsKey($k)) { $children[$k] = New-Object System.Collections.Generic.List[object] }
    $children[$k].Add($it)
}
function Get-Kids([string]$Path) {
    $k = $Path.ToLower()
    if ($children.ContainsKey($k)) { return $children[$k] }
    return @()
}

# find PROJECT folders (they hold <code>.rvt files and/or CentralCache / LinkedModels)
$candidates = @()
foreach ($d in $dirs) {
    if ($d.Name -ieq 'CentralCache' -or $d.Name -ieq 'LinkedModels') { continue }
    foreach ($kid in (Get-Kids $d.FullName)) {
        if (($kid.PSIsContainer -and ($kid.Name -ieq 'CentralCache' -or $kid.Name -ieq 'LinkedModels')) -or
            (-not $kid.PSIsContainer -and $kid.Name -match "^$GuidPattern\.rvt$")) {
            $candidates += $d.FullName
            break
        }
    }
}
$projectFolders = @()
foreach ($c in ($candidates | Sort-Object Length)) {
    $nested = $false
    foreach ($p in $projectFolders) { if ($c.ToLower().StartsWith($p.ToLower() + $sep)) { $nested = $true; break } }
    if (-not $nested) { $projectFolders += $c }
}

# find MODELS inside each project folder
$models    = [ordered]@{}
$itemOwner = @{}      # path of a cache item -> model key
foreach ($p in $projectFolders) {
    $places = @(@{ Path = $p; Kind = 'Main folder' })
    foreach ($sub in 'CentralCache', 'LinkedModels') {
        $sp = Join-Path $p $sub
        if (Test-Path -LiteralPath $sp -PathType Container) { $places += @{ Path = $sp; Kind = $sub } }
    }
    foreach ($place in $places) {
        foreach ($kid in (Get-Kids $place.Path)) {
            $m = [regex]::Match($kid.Name, $GuidPattern)
            if (-not $m.Success) { continue }
            $guid = $m.Value.ToLower()
            $key  = $p.ToLower() + '|' + $guid
            if (-not $models.Contains($key)) {
                $acct = [regex]::Match($p, 'CollaborationCache[\\/]([^\\/]+)').Groups[1].Value
                $ver  = [regex]::Match($p, 'Autodesk Revit (\d{4})').Groups[1].Value
                $models[$key] = [pscustomobject]@{
                    Key = $key; Guid = $guid; ProjectFolder = $p; Account = $acct; Version = $ver
                    Name = ''; Project = ''; Identified = $false
                    Items = New-Object System.Collections.Generic.List[object]
                    Files = New-Object System.Collections.Generic.List[object]
                    Size = 0.0; Latest = [datetime]::MinValue
                    WorkingCopy = $null; CentralCopy = $null; Unsynced = $false; No = 0
                }
            }
            $models[$key].Items.Add([pscustomobject]@{ Item = $kid; Kind = $place.Kind })
            $itemOwner[$kid.FullName.ToLower()] = $key
            if (-not $kid.PSIsContainer -and $kid.Name -ieq "$guid.rvt") {
                if ($place.Kind -eq 'Main folder')  { $models[$key].WorkingCopy = $kid }
                if ($place.Kind -eq 'CentralCache') { $models[$key].CentralCopy = $kid }
            }
        }
    }
}

# give every file to its model
$fileOwner = @{}
foreach ($f in $files) {
    $path = $f.FullName
    $owner = $null
    while ($path -and $path.Length -gt 3) {
        $k = $path.ToLower()
        if ($itemOwner.ContainsKey($k)) { $owner = $itemOwner[$k]; break }
        $path = [System.IO.Path]::GetDirectoryName($path)
    }
    if ($owner) {
        $fileOwner[$f.FullName] = $owner
        $mo = $models[$owner]
        $mo.Files.Add($f)
        $mo.Size += $f.Length
        if ($f.LastWriteTime -gt $mo.Latest) { $mo.Latest = $f.LastWriteTime }
    }
}

# ---------------------------------------------------------------------
#  STEP 2 - read the real model names
# ---------------------------------------------------------------------
$n = 0
foreach ($mo in $models.Values) {
    $n++
    if (-not $TestMode) { Write-Host ("`r   Finding model names ... {0} of {1}   " -f $n, $models.Count) -NoNewline }
    $tries = @()
    if ($mo.WorkingCopy) { $tries += $mo.WorkingCopy }
    if ($mo.CentralCopy) { $tries += $mo.CentralCopy }
    foreach ($x in $mo.Items) { if (-not $x.Item.PSIsContainer -and $x.Item.Extension -ieq '.rvt' -and $tries -notcontains $x.Item) { $tries += $x.Item } }
    foreach ($t in $tries) {
        $info = Get-ModelInfoFromRvt $t.FullName
        if ($info) { $mo.Name = $info.Name; $mo.Project = $info.Project; $mo.Identified = $true; break }
    }
    if (-not $mo.Identified) { $mo.Name = '(not identified) ' + $mo.Guid }

}

Write-Host ''
if ($models.Count -eq 0) {
    Write-Host ''
    Write-Host 'No Revit cloud models were found in this cache.' -ForegroundColor Yellow
    return
}

# project names per project folder (for models without a name)
$projName = @{}
foreach ($mo in $models.Values) { if ($mo.Project -and -not $projName.ContainsKey($mo.ProjectFolder)) { $projName[$mo.ProjectFolder] = $mo.Project } }
foreach ($mo in $models.Values) {
    if (-not $mo.Project) {
        if ($projName.ContainsKey($mo.ProjectFolder)) { $mo.Project = $projName[$mo.ProjectFolder] }
        else { $mo.Project = 'Project folder ' + (Split-Path $mo.ProjectFolder -Leaf) }
    }
}

# clean screen - show only what the user needs
if (-not $TestMode) { try { Clear-Host } catch { } }
Write-Host ''
Write-Host '=============================================================' -ForegroundColor Cyan
Write-Host '   REVIT COLLABORATION CACHE CLEANER' -ForegroundColor Cyan
Write-Host '=============================================================' -ForegroundColor Cyan
Write-Host ("   Scan complete: {0} model(s) found." -f $models.Count) -ForegroundColor Green

# number the models
$list = New-Object System.Collections.Generic.List[object]
$i = 0
foreach ($mo in ($models.Values | Sort-Object Version, Project, @{ Expression = { -not $_.Identified } }, Name)) {
    $i++; $mo.No = $i; $list.Add($mo)
}

# ---------------------------------------------------------------------
#  Excel report
# ---------------------------------------------------------------------
function Export-Report {
    $base = New-Object DateTime(1899, 12, 30)
    $sumRows = New-Object System.Collections.Generic.List[object]
    foreach ($mo in $list) {
        $sumRows.Add(@($mo.No, $mo.Name, $mo.Project, $mo.Version, $mo.Account, $mo.Guid, $mo.Files.Count,
                       [Math]::Round($mo.Size / 1MB, 2), $mo.Latest, $mo.Latest, $mo.ProjectFolder))
    }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($f in ($files | Sort-Object FullName)) {
        if (-not (Test-Path -LiteralPath $f.FullName)) { continue }       # already cleared
        $relDir = $f.DirectoryName
        if ($relBase -and $relDir.ToLower().StartsWith($relBase.ToLower())) { $relDir = $relDir.Substring($relBase.Length).TrimStart('\', '/') }
        if ($relDir -eq '') { $relDir = '(main folder)' }
        $actual = '(not part of a model)'; $project = ''; $no = ''
        if ($fileOwner.ContainsKey($f.FullName)) {
            $mo = $models[$fileOwner[$f.FullName]]
            $actual = $mo.Name; $project = $mo.Project; $no = $mo.No
        }
        $ver  = [regex]::Match($f.FullName, 'Autodesk Revit (\d{4})').Groups[1].Value
        $acct = [regex]::Match($f.FullName, 'CollaborationCache[\\/]([^\\/]+)').Groups[1].Value
        $rows.Add(@($no, $relDir, $f.Name, $actual, $f.LastWriteTime, $f.LastWriteTime,
                    [Math]::Ceiling($f.Length / 1KB), $project, $ver, $acct, $f.FullName))
    }
    $sheets = @(
        @{
            Name    = 'Models Summary'
            Headers = @('No.', 'Actual Model Name', 'Project Name', 'Revit Version', 'Autodesk Account', 'Cache Code (folder / file name)', 'Number of Files', 'Total Size (MB)', 'Last Modified Date', 'Last Modified Time', 'Location of Cache')
            Types   = @('i', 's', 's', 'i', 's', 's', 'i', 'n', 'd', 't', 's')
            Widths  = @(6, 40, 26, 13, 22, 40, 15, 15, 18, 18, 90)
            Rows    = $sumRows
        },
        @{
            Name    = 'All Cache Files'
            Headers = @('Model No.', 'Cache Subfolder Name', 'Cache File Name', 'Actual File Name', 'Date Modified', 'Time Modified', 'File Size (KB)', 'Project Name', 'Revit Version', 'Autodesk Account', 'Full File Path')
            Types   = @('i', 's', 's', 's', 'd', 't', 'i', 's', 'i', 's', 's')
            Widths  = @(11, 60, 45, 40, 16, 16, 15, 26, 13, 22, 110)
            Rows    = $rows
        }
    )
    $outDir = if ($env:RCR_OUT) { $env:RCR_OUT } else { [Environment]::GetFolderPath('Desktop') }
    if (-not $outDir -or -not (Test-Path -LiteralPath $outDir)) { $outDir = $env:USERPROFILE }
    $outFile = Join-Path $outDir ('Revit_Cache_Report_' + (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss') + '.xlsx')
    try {
        Save-Xlsx $outFile $sheets
        Write-Host "Excel report saved to: $outFile" -ForegroundColor Green
        if (-not $TestMode) { try { Invoke-Item -LiteralPath $outFile } catch { } }
    } catch {
        Write-Host "Could not save the Excel file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ---------------------------------------------------------------------
#  Show the list
# ---------------------------------------------------------------------
function Show-List([string]$Filter) {
    $shown = @($list | Where-Object {
        -not $Filter -or $_.Name -like "*$Filter*" -or $_.Project -like "*$Filter*" -or $_.Guid -like "*$Filter*"
    })
    Write-Host ''
    Write-Host '=============================================================================================' -ForegroundColor Cyan
    if ($Filter) { Write-Host "  MODELS MATCHING '$Filter'  ($($shown.Count) of $($list.Count))" -ForegroundColor Cyan }
    else         { Write-Host "  MODELS IN THE CACHE  ($($list.Count))" -ForegroundColor Cyan }
    Write-Host '=============================================================================================' -ForegroundColor Cyan
    if ($shown.Count -eq 0) { return }
    $group = $null
    foreach ($mo in $shown) {
        $g = "Revit $($mo.Version)  |  Project: $($mo.Project)"
        if ($list | Where-Object { $_.Account -ne $mo.Account } | Select-Object -First 1) { $g += "  |  Account: $($mo.Account)" }
        if ($g -ne $group) {
            Write-Host ''
            Write-Host "  $g" -ForegroundColor Yellow
            Write-Host ('  {0,-5}{1,-42}{2,11}   {3,-20}' -f 'No.', 'Model Name', 'Size', 'Last Used') -ForegroundColor DarkGray
            $group = $g
        }
        $line = '  {0,-5}{1,-42}{2,11}   {3,-20}' -f $mo.No, (Limit-Text $mo.Name 40), (Format-Size $mo.Size), $mo.Latest.ToString('dd-MMM-yyyy h:mm tt')
        if (-not $mo.Identified) {
            Write-Host $line -ForegroundColor DarkYellow
        } else {
            Write-Host $line
        }
    }
}

# turns "3", "3,5", "3-6", "2 4" into a list of models
function Get-Selection([string]$Text) {
    $nums = New-Object System.Collections.Generic.List[int]
    foreach ($part in ($Text -split '[,\s]+' | Where-Object { $_ -ne '' })) {
        if ($part -match '^(\d+)-(\d+)$') {
            $a = [int]$Matches[1]; $b = [int]$Matches[2]
            if ($a -gt $b) { $t = $a; $a = $b; $b = $t }
            for ($x = $a; $x -le $b; $x++) { $nums.Add($x) }
        } elseif ($part -match '^\d+$') { $nums.Add([int]$part) }
        else { Write-Host "  '$part' is not a valid number." -ForegroundColor Red; $script:SelMsg = $true; return $null }
    }
    $sel = @()
    foreach ($x in ($nums | Select-Object -Unique)) {
        $mo = $list | Where-Object { $_.No -eq $x } | Select-Object -First 1
        if (-not $mo) { Write-Host "  Number $x is not in the list." -ForegroundColor Red; $script:SelMsg = $true; return $null }
        $sel += $mo
    }
    return $sel
}

# ---------------------------------------------------------------------
#  Clear the cache of the chosen models
# ---------------------------------------------------------------------
function Clear-Models($Selected) {
    Write-Host ''
    Write-Host '=============================================================================================' -ForegroundColor Magenta
    Write-Host '  THESE FILES AND FOLDERS WILL BE DELETED' -ForegroundColor Magenta
    Write-Host '=============================================================================================' -ForegroundColor Magenta
    $total = 0.0
    foreach ($mo in $Selected) {
        Write-Host ''
        Write-Host ("  [{0}] {1}   (Revit {2}, {3})" -f $mo.No, $mo.Name, $mo.Version, $mo.Project) -ForegroundColor Yellow
        foreach ($x in $mo.Items) {
            $it = $x.Item
            if ($it.PSIsContainer) {
                $sz = 0.0; foreach ($f in $mo.Files) { if ($f.FullName.ToLower().StartsWith($it.FullName.ToLower() + $sep)) { $sz += $f.Length } }
                $label = $it.Name + '\  (folder)'
            } else { $sz = $it.Length; $label = $it.Name }
            Write-Host ('      {0,-14}{1,-62}{2,11}' -f $x.Kind, (Limit-Text $label 60), (Format-Size $sz))
        }
        $total += $mo.Size
    }
    Write-Host ''
    Write-Host ("  Total to clear: {0}" -f (Format-Size $total)) -ForegroundColor Cyan
    Write-Host '  Very large files may be too big for the Recycle Bin - if so, Windows will ASK you' -ForegroundColor DarkGray
    Write-Host '  before deleting them permanently (you can say No).' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  WARNING: Sync and Close the Model before proceeding.' -ForegroundColor Red
    Write-Host '  Any work not synced to the cloud will be lost.' -ForegroundColor Red

    # Revit must be closed
    while (Test-RevitRunning) {
        Write-Host ''
        Write-Host '  Revit is OPEN. Please close ALL Revit windows first (save/sync your work).' -ForegroundColor Red
        $a = Ask '  Press Enter when Revit is closed, or type Q to cancel'
        if ($a -match '^\s*[Qq]') { Write-Host '  Cancelled. Nothing was deleted.' -ForegroundColor Yellow; return }
    }

    Write-Host ''
    $a = Ask '  Delete the files? (Y/N)'
    if ([string]$a -notmatch '^\s*(y|yes)\s*$') {
        Write-Host '  Cancelled. Nothing was deleted.' -ForegroundColor Yellow
        return
    }
    if (Test-RevitRunning) { Write-Host '  Revit was opened again. Cancelled - nothing was deleted.' -ForegroundColor Red; return }

    Write-Log ''
    Write-Log ("===== {0}  |  User: {1}  |  PC: {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:USERNAME, $env:COMPUTERNAME)
    foreach ($mo in $Selected) {
        Write-Host ''
        Write-Host ("  Clearing [{0}] {1} ..." -f $mo.No, $mo.Name) -ForegroundColor Green
        Write-Log ("Model: {0}  |  Project: {1}  |  Revit {2}  |  Code: {3}" -f $mo.Name, $mo.Project, $mo.Version, $mo.Guid)
        $ok = 0; $bad = 0
        foreach ($x in $mo.Items) {
            $it = $x.Item
            if (-not (Test-Path -LiteralPath $it.FullName)) { continue }
            $gone = Send-ToRecycleBin $it.FullName $it.PSIsContainer
            Write-Host ("      {0,-62}" -f ("{0}\{1}" -f $x.Kind, $it.Name)) -NoNewline -ForegroundColor Gray
            if ($gone) { $ok++;  Write-Host 'SUCCESS' -ForegroundColor Green; Write-Log ("   Recycled : {0}" -f $it.FullName) }
            else       { $bad++; Write-Host 'FAILED'  -ForegroundColor Red;   Write-Log ("   FAILED   : {0}" -f $it.FullName) }
        }
        Write-Host ''
        if ($bad -eq 0) {
            Write-Host '      Done - this model''s cache is cleared.' -ForegroundColor Green
            [void]$list.Remove($mo)
        } else {
            Write-Host "      $bad item(s) could not be removed. Make sure Revit is closed and try again." -ForegroundColor Red
        }
    }
    Write-Host '      If you need anything back, open the Recycle Bin and choose Restore.' -ForegroundColor DarkGray
    Write-Host "      A record was saved in: $LogFile" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------
#  MENU
# ---------------------------------------------------------------------
$showList = $true
while ($true) {
    if ($list.Count -eq 0) {
        Write-Host ''
        Write-Host 'All listed models have been cleared. Nothing more to do.' -ForegroundColor Green
        break
    }
    if ($showList) {
        Show-List ''
        if (Test-RevitRunning) {
            Write-Host ''
            Write-Host '  Note: Revit is open right now. You will need to close it before clearing.' -ForegroundColor Yellow
        }
        $showList = $false
    }
    Write-Host ''
    Write-Host '  WHAT NEXT?' -ForegroundColor Cyan
    Write-Host '    - Type the No. of the model to clear      e.g.  3    or  3,5    or  3-6'
    Write-Host '    - L = show the list again'
    Write-Host '    - E = create Excel report'
    Write-Host '    - Q = quit'
    $ans = ([string](Ask '  Your choice')).Trim()

    if ($ans -eq '') { continue }
    if ($ans -match '^[Qq]$') { break }
    if ($ans -match '^[Ll]$') { $showList = $true; continue }
    if ($ans -match '^[Ee]$') { Export-Report; continue }
    if ($ans -match '^[\d\s,\-]+$') {
        # @( ) makes sure ONE chosen model is still counted as 1 (older Windows PowerShell)
        $script:SelMsg = $false
        $sel = @(Get-Selection $ans | Where-Object { $null -ne $_ })
        if ($sel.Count -eq 0) {
            if (-not $script:SelMsg) { Write-Host '  Nothing was selected. Please type a model number from the list.' -ForegroundColor Red }
        } else {
            $before = $list.Count
            Clear-Models $sel
            # show the list again only if something was actually cleared
            if ($list.Count -ne $before -and $list.Count -gt 0) {
                Write-Host ''
                [void](Ask '  Press Enter to see the updated list')
                $showList = $true
            }
        }
        continue
    }
    Write-Host '  Please type a model number from the list, or L, E or Q.' -ForegroundColor Red
}
Write-Host ''
Write-Host 'Goodbye!' -ForegroundColor Cyan
