<#
.SYNOPSIS
Generates a local PC health report as HTML or JSON.

.DESCRIPTION
Collects system, CPU, memory, disk, and physical disk health data, then emits an HTML report or JSON payload.

.PARAMETER OutFile
Path to the HTML report output file.

.PARAMETER EmitJson
Switch to emit JSON instead of HTML.

.NOTES
Intended for local use on Windows with CIM and (optionally) Storage module.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutFile = "$env:USERPROFILE\Desktop\LocalPcHealth.html",
    [Parameter()]
    [switch]$EmitJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region helpers
function New-Section {
    <#
    .SYNOPSIS
    Creates a section object for the report.

    .PARAMETER Title
    Section title text.

    .PARAMETER Data
    Section data payload.

    .PARAMETER Severity
    Section severity level: info, warn, or error.

    .OUTPUTS
    PSCustomObject with Title, Data, Severity.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Title,
        # was: [hashtable]$Data
        [Parameter(Mandatory)] [object]$Data,
        [Parameter()] [ValidateSet('info','warn','error')] [string]$Severity = 'info'
    )
    [pscustomobject]@{
        Title    = $Title
        Data     = $Data
        Severity = $Severity
    }
}

function Safe-Invoke {
    <#
    .SYNOPSIS
    Invokes a scriptblock and captures errors as structured output.

    .PARAMETER ScriptBlock
    The scriptblock to execute.

    .OUTPUTS
    Any output from the scriptblock or a structured error object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [scriptblock]$ScriptBlock
    )
    try { & $ScriptBlock } catch {
        [pscustomobject]@{
            Error                   = $true
            Message                 = $_.Exception.Message
            Category                = $_.CategoryInfo.Category
            FullyQualifiedErrorId   = $_.FullyQualifiedErrorId
        }
    }
}

function Get-IsAdmin {
    <#
    .SYNOPSIS
    Checks if the current user has administrator role.

    .OUTPUTS
    System.Boolean
    #>
    [CmdletBinding()] param()
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p  = [Security.Principal.WindowsPrincipal]::new($id)
        [bool]$p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { $false }
}

function Get-Uptime {
    <#
    .SYNOPSIS
    Returns system uptime components.

    .OUTPUTS
    PSCustomObject with Days, Hours, Minutes.
    #>
    [CmdletBinding()] param()
    $ticks = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $uptime = (Get-Date) - $ticks
    [pscustomobject]@{
        Days    = [int]$uptime.TotalDays
        Hours   = [int]$uptime.Hours
        Minutes = [int]$uptime.Minutes
    }
}

function Get-CpuLoad {
    <#
    .SYNOPSIS
    Retrieves CPU model, core count, and load metrics.

    .OUTPUTS
    PSCustomObject with Name, Cores, LoadPercent, PerCore.
    #>
    [CmdletBinding()] param()

    $procs = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue
    $name = ($procs | Select-Object -First 1).Name
    $totalCores = ($procs | Measure-Object -Property NumberOfCores -Sum).Sum

    $avgLoad = $null
    $perCore = @()

    # Try formatted perf data first (fast, no sampling)
    try {
        $perf = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -ErrorAction Stop |
                Where-Object { $_.Name -ne '_Total' }
        if ($perf) {
            $perCore = $perf |
                Sort-Object { [int]($_.Name -replace '[^\d]', '0') } |
                ForEach-Object { [int][math]::Round($_.PercentProcessorTime, 0) }
            if ($perCore.Count -gt 0) {
                $avgLoad = [int][math]::Round((($perCore | Measure-Object -Average).Average), 0)
            }
        }
    } catch { }

    # Fallback: sample counters
    if (-not $perCore -or $perCore.Count -eq 0) {
        try {
            $ctr = Get-Counter '\Processor(*)\% Processor Time' -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
            $samps = $ctr.CounterSamples | Where-Object { $_.InstanceName -ne '_Total' }
            $perCore = $samps |
                Sort-Object { [int]($_.InstanceName -replace '[^\d]', '0') } |
                ForEach-Object { [int][math]::Round($_.CookedValue, 0) }
            if ($perCore.Count -gt 0) {
                $avgLoad = [int][math]::Round((($perCore | Measure-Object -Average).Average), 0)
            }
        } catch { }
    }

    # Final fallback: single WMI value
    if ($null -eq $avgLoad) {
        $load = ($procs | Select-Object -First 1).LoadPercentage
        if ($null -ne $load) { $avgLoad = [int]$load }
    }

    [pscustomobject]@{
        Name        = $name
        Cores       = $totalCores
        LoadPercent = $avgLoad
        PerCore     = $perCore
    }
}

function Get-MemoryUse {
    <#
    .SYNOPSIS
    Reports total, used, free memory and usage percent.

    .OUTPUTS
    PSCustomObject with TotalGB, UsedGB, FreeGB, UsedPct.
    #>
    [CmdletBinding()] param()
    $os    = Get-CimInstance -ClassName Win32_OperatingSystem
    $total = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $free  = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $used  = [math]::Round($total - $free, 2)
    $pct   = if ($total -ne 0) { [math]::Round(($used/$total)*100, 2) } else { 0 }
    [pscustomobject]@{
        TotalGB = $total
        UsedGB  = $used
        FreeGB  = $free
        UsedPct = $pct
    }
}

function Get-Diskspace {
    <#
    .SYNOPSIS
    Returns per-volume disk space metrics.

    .OUTPUTS
    Hashtable keyed by DeviceID with size and usage details.
    #>
    [CmdletBinding()] param()
    $vols = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
    $out = @{}
    foreach ($v in $vols) {
        $total = if ($v.Size) { [math]::Round($v.Size/1GB, 2) } else { 0 }
        $free  = if ($v.FreeSpace) { [math]::Round($v.FreeSpace/1GB, 2) } else { 0 }
        $used  = [math]::Round($total - $free, 2)
        $pct   = if ($total -ne 0) { [math]::Round(($used/$total)*100, 2) } else { 0 }
        $out[$v.DeviceID] = @{
            TotalGB = $total
            UsedGB  = $used
            FreeGB  = $free
            UsedPct = $pct
        }
    }
    $out
}

function Get-PhysicalDiskHealth {
    <#
    .SYNOPSIS
    Retrieves physical disk health and status from Storage module.

    .OUTPUTS
    Hashtable keyed by FriendlyName with health details, or a note/error.
    #>
    [CmdletBinding()] param()
    $result = @{}
    try {
        if (-not (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue)) {
            return @{ Note = 'Get-PhysicalDisk not available (requires Windows Storage module).' }
        }
        $pds = Get-PhysicalDisk -ErrorAction Stop
        foreach ($d in $pds) {
            $result[$d.FriendlyName] = @{
                HealthStatus      = $d.HealthStatus
                OperationalStatus = ($d.OperationalStatus -join ', ')
                MediaType         = $d.MediaType
                SizeGB            = [math]::Round($d.Size/1GB, 2)
            }
        }
        if ($result.Count -eq 0) { return @{ Note = 'No physical disks found.' } }
        $result
    } catch {
        @{ Error = $true; Message = $_.Exception.Message }
    }
}

function Get-Severity {
    <#
    .SYNOPSIS
    Maps a percent value to a severity label.

    .PARAMETER Percent
    Percentage value to evaluate.

    .PARAMETER Warn
    Warning threshold percent.

    .PARAMETER Error
    Error threshold percent.

    .OUTPUTS
    String: info, warn, or error.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [double]$Percent,
        [Parameter()] [double]$Warn  = 80,
        [Parameter()] [double]$Error = 90
    )
    if ($Percent -ge $Error) { 'error' }
    elseif ($Percent -ge $Warn) { 'warn' }
    else { 'info' }
}

function Get-SeverityLabel {
    <#
    .SYNOPSIS
    Converts a severity key to a human label.

    .PARAMETER Severity
    Severity key: info, warn, or error.

    .OUTPUTS
    String label.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('info','warn','error')]
        [string]$Severity
    )
    switch ($Severity) {
        'info'  { 'Normal' }
        'warn'  { 'Warning' }
        'error' { 'Critical' }
    }
}

function As-FlatTable {
    <#
    .SYNOPSIS
    Flattens a dictionary-of-objects into rows.

    .PARAMETER InputObject
    The input dictionary or array.

    .OUTPUTS
    PSCustomObject rows or original input.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)] [object]$InputObject
    )
    process {
        if ($InputObject -is [System.Collections.IDictionary]) {
            foreach ($k in $InputObject.Keys) {
                $val = $InputObject[$k]
                if ($val -is [System.Collections.IDictionary]) {
                    $row = [ordered]@{ Name = [string]$k }
                    foreach ($kk in $val.Keys) { $row[$kk] = $val[$kk] }
                    [pscustomobject]$row
                } elseif ($val -is [pscustomobject]) {
                    $row = [ordered]@{ Name = [string]$k }
                    $val.PSObject.Properties | ForEach-Object { $row[$_.Name] = $_.Value }
                    [pscustomobject]$row
                } else {
                    [pscustomobject]@{ Name = [string]$k; Value = $val }
                }
            }
        } elseif ($InputObject -is [System.Array]) {
            $InputObject
        } else {
            $InputObject
        }
    }
}

function Render-SectionHtml {
    <#
    .SYNOPSIS
    Renders a section object into HTML.

    .PARAMETER Section
    Section object with Title, Data, Severity.

    .OUTPUTS
    HTML fragment string.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Section
    )
    $sevClass = switch ($Section.Severity) {
        'warn'  { 'section-warn' }
        'error' { 'section-error' }
        default { 'section-info' }
    }

    $data = $Section.Data
    $rowsHtml = ''

    if ($data -is [System.Collections.IDictionary]) {
        $flat = As-FlatTable $data | ForEach-Object { [pscustomobject]$_ }
        $rowsHtml = ($flat | Select-Object * | ConvertTo-Html -As Table -Fragment) -replace '<table>', '<table class="kv">'
    } elseif ($data -is [System.Array]) {
        $rowsHtml = ($data | ConvertTo-Html -As Table -Fragment) -replace '<table>', '<table class="kv">'
    } else {
        $display = [ordered]@{}
        # copy props, stringify arrays so they don't show as System.Object[]
        foreach ($p in ($data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)) {
            $v = $data.$p
            if ($v -is [System.Array]) { $v = ($v -join ', ') }
            elseif ($v -is [System.Collections.IDictionary] -or $v -is [pscustomobject]) {
                # compact nested objects for readability
                $v = ($v | ConvertTo-Json -Compress)
            }
            $display[$p] = $v
        }
        $rowsHtml = ([pscustomobject]$display | ConvertTo-Html -As List -Fragment) -replace '<table>', '<table class="kv">'
    }

    @"
<section class="section $sevClass">
  <h2>$($Section.Title)</h2>
  $rowsHtml
</section>
"@
}
#endregion helpers

#region collect
# List used to accumulate section objects
$sections = New-Object System.Collections.Generic.List[object]

# System section
$sysData = Safe-Invoke {
    $u = Get-Uptime
    [ordered]@{
        ComputerName = $env:COMPUTERNAME
        UserName     = $env:USERNAME
        OSVersion    = (Get-CimInstance Win32_OperatingSystem).Caption
        IsAdmin      = Get-IsAdmin
        Uptime       = ('{0}d {1}h {2}m' -f $u.Days, $u.Hours, $u.Minutes)
        TimeStamp    = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
}
$sections.Add( (New-Section -Title 'System' -Data $sysData) )

# CPU section
$cpu = Safe-Invoke { Get-CpuLoad }
$sections.Add( (New-Section -Title 'CPU' -Data @{} ) )
$sections[-1].Data = $cpu | ConvertTo-Json -Compress | ConvertFrom-Json

# Memory section with severity enrichment
$mem = Safe-Invoke { Get-MemoryUse }
$memObj = $mem | ConvertTo-Json -Compress | ConvertFrom-Json

$hasPct = $memObj.PSObject.Properties.Name -contains 'UsedPct'
$usedPctVal = if ($hasPct) { [double]$memObj.UsedPct } else { $null }

$memSev = if ($hasPct) {
    Get-Severity -Percent $usedPctVal -Warn 80 -Error 90
} else { 'warn' } # probe failure → surface as Warning

$memStatus = Get-SeverityLabel -Severity $memSev

# enrich data for readability
if ($hasPct) {
    $memObj | Add-Member -NotePropertyName 'FreePct' -NotePropertyValue ([math]::Round(100 - $usedPctVal, 2)) -Force
}
$memObj | Add-Member -NotePropertyName 'Status' -NotePropertyValue $memStatus -Force

$usedPctText = if ($hasPct) { ('{0:N1}%' -f $usedPctVal) } else { 'n/a' }
$memTitle = "Memory — Overall: $memStatus (Thresholds: Warning ≥80% used; Critical ≥90% used; current used: $usedPctText)"

$sections.Add( (New-Section -Title $memTitle -Data $memObj -Severity $memSev) )

# Volumes section with per-volume severity and overall state
$disks = Safe-Invoke { Get-Diskspace }

$worst = 'info'
$maxUsed = 0.0
$minFree = 100.0

if ($disks -is [System.Collections.IDictionary]) {
    foreach ($k in @($disks.Keys)) {
        $v = $disks[$k]
        if ($v -is [System.Collections.IDictionary] -and $v.ContainsKey('UsedPct')) {
            $u = [double]$v['UsedPct']
            $maxUsed = [math]::Max($maxUsed, $u)
            $minFree = [math]::Min($minFree, 100 - $u)

            $sev = Get-Severity -Percent $u -Warn 85 -Error 90
            $v['Severity'] = $sev
            $v['Status']   = Get-SeverityLabel -Severity $sev
            $v['FreePct']  = [math]::Round(100 - $u, 2)

            switch ($sev) {
                'error' { $worst = 'error' }
                'warn'  { if ($worst -eq 'info') { $worst = 'warn' } }
            }
        } else {
            $worst = 'warn'
        }
    }
} else {
    $worst = 'warn'
}

$overallLabel = Get-SeverityLabel -Severity $worst
$volTitle = "Volumes — Overall: $overallLabel (Thresholds: Warn ≥85% used, Error ≥90% used; max used: {0:N1}%, min free: {1:N1}%)" -f $maxUsed, $minFree
$sections.Add( (New-Section -Title $volTitle -Data $disks -Severity $worst) )

# Physical disks section
$phys = Safe-Invoke { Get-PhysicalDiskHealth }
$sev  = if (($phys -is [System.Collections.IDictionary]) -and $phys.ContainsKey('Error')) { 'warn' } else { 'info' }
$sections.Add( (New-Section -Title 'Physical Disks' -Data $phys -Severity $sev) )
#endregion collect

#region output
# JSON output path
if ($EmitJson.IsPresent) {
    [pscustomobject]@{
        GeneratedAt = (Get-Date)
        Sections    = $sections
    } | ConvertTo-Json -Depth 6
    return
}

# HTML style block
$css = @"
<style>
body { font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; }
header { margin-bottom: 16px; }
h1 { font-size: 22px; margin: 0 0 8px 0; }
.meta { color: #555; font-size: 12px; }
.section { margin: 18px 0; padding: 12px 14px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
.section-info { background: #f8fafc; }
.section-warn { background: #fff7ed; }
.section-error { background: #fef2f2; }
h2 { font-size: 18px; margin: 0 0 10px 0; }
table.kv { width: 100%; border-collapse: collapse; }
table.kv th, table.kv td { border: 1px solid #e5e7eb; padding: 6px 8px; font-size: 13px; }
table.kv th { background: #f3f4f6; text-align: left; }
footer { margin-top: 24px; color: #6b7280; font-size: 12px; }
code { background: #f3f4f6; padding: 2px 4px; border-radius: 6px; }
</style>
"@

# Build HTML body with header and sections
$body = @()
$body += '<header>'
$body += '<h1>Local PC Health Report</h1>'
$body += "<div class='meta'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>"
$body += '</header>'

foreach ($s in $sections) { $body += (Render-SectionHtml -Section $s) }

$body += "<footer>Generated by LocalPcHealth.ps1</footer>"

# Convert to HTML document
$html = ConvertTo-Html -Title 'Local PC Health' -Head $css -Body ($body -join "`n")

# Ensure directory exists before writing
$dir = Split-Path -Parent -Path $OutFile
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

# Write HTML report and print output path
Set-Content -LiteralPath $OutFile -Value $html -Encoding UTF8
Write-Host "Report written to: $OutFile"
#endregion output