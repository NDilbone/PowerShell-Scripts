# Collection of PowerShell Scripts
- ### [Local PC Health Report](#1-local-pc-health-report)

> [!WARNING]
> These scripts are provided as-is and are always a work in progress. Use at your own risk. Review the code before running, and test in a non-production environment.

---

# 1) Local PC Health Report

Generate a lightweight **HTML** or **JSON** health report for your Windows PC: system info, CPU load (incl. per-core), memory usage, volumes, and physical disk health.

![PowerShell](https://img.shields.io/badge/PowerShell-7%2B%20%7C%205.1-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-informational)
![License](https://img.shields.io/badge/License-MIT-green)

---

## Features
- **HTML report** with simple styling; **JSON output** for automation.
- **System**: computer/user, OS, admin, uptime.
- **CPU**: model, core count, **per-core** load, average load.
- **Memory**: totals, %, **severity** with professional labels.
- **Volumes**: per-drive totals, %, per-drive **severity** and overall status.
- **Physical disks** (if available): health/operational status via `Get-PhysicalDisk`.

> [!NOTE]
> No external dependencies; optional Windows Storage module enhances disk health.

---

## Requirements
- Windows with **PowerShell** 5.1 or **PowerShell 7+**.
- **CIM/WMI** available (default on Windows).
- Optional: **Windows Storage** module for `Get-PhysicalDisk` (physical disk health).
- For CPU counters fallback: Performance Counters must be enabled (default).

---

## Install
```powershell
git clone https://github.com/NDilbone/PowerShell-Scripts.git
cd PowerShell-Scripts
# or save the script as .\LocalPcHealthCheck.ps1
```

---

## Usage

### HTML (default)
```powershell
.\LocalPcHealthCheck.ps1
# or specify output path
.\LocalPcHealthCheck.ps1 -OutFile "$env:USERPROFILE\Desktop\LocalPcHealth.html"
```

### JSON (for CI/automation)
```powershell
.\LocalPcHealthCheck.ps1 -EmitJson | ConvertFrom-Json
# example: check max volume usage
$report = .\LocalPcHealthCheck.ps1 -EmitJson | ConvertFrom-Json
$vols = $report.Sections | Where-Object Title -like 'Volumes*'
$vols.Data
```

---

## Parameters

| Name      | Type   | Default                                          | Description                                 |
|-----------|--------|--------------------------------------------------|---------------------------------------------|
| `OutFile` | String | `$env:USERPROFILE\Desktop\LocalPcHealth.html` | Path for the generated HTML report.         |
| `EmitJson`| Switch | *(not set)*                                      | Emit JSON to stdout instead of HTML file.   |

> [!NOTE]
> JSON mode is pipeline-friendly and does not write an HTML file.

---

## Severity Model

Labels are derived from usage percentages.

### Memory
- **Warning** ≥ **80% used**
- **Critical** ≥ **90% used**
- Otherwise **Normal**

### Volumes (per drive & overall)
- **Warning** ≥ **85% used**
- **Critical** ≥ **90% used**
- Otherwise **Normal**

Each volume row includes:
- `UsedGB`, `FreeGB`, `UsedPct`, `FreePct`, machine-readable `Severity`, human label `Status`.

> [!NOTE]
> The **Volumes** section header shows the overall status plus the **max used %** and **min free %** across all volumes.

---

## Output Details

### HTML
- Sections: **System**, **CPU**, **Memory**, **Volumes**, **Physical Disks**.
- Background shading by section severity:
  - Normal → subtle gray
  - Warning → light amber
  - Critical → light red

### JSON (shape overview)
```json
{
  "GeneratedAt": "2025-10-01T12:34:56Z",
  "Sections": [
    { "Title": "System", "Data": { "ComputerName": "...", "Uptime": "Xd Yh Zm", "TimeStamp": "...", "IsAdmin": true, "OSVersion": "...", "UserName": "..." }, "Severity": "info" },
    { "Title": "CPU", "Data": { "Name": "...", "Cores": 8, "LoadPercent": 12, "PerCore": [10,12,9,17] }, "Severity": "info" },
    { "Title": "Memory — Overall: Normal (...)", "Data": { "TotalGB": 31.8, "UsedPct": 42.3, "FreePct": 57.7, "Status": "Normal" }, "Severity": "info" },
    { "Title": "Volumes — Overall: Normal (...)", "Data": { "C:": { "UsedPct": 23.4, "FreePct": 76.6, "Status": "Normal", "Severity": "info" }, "D:": { "UsedPct": 91.2, "FreePct": 8.8, "Status": "Critical", "Severity": "error" } }, "Severity": "error" },
    { "Title": "Physical Disks", "Data": { "Samsung SSD": { "HealthStatus": "Healthy", "OperationalStatus": "OK", "MediaType": "SSD", "SizeGB": 953.87 } }, "Severity": "info" }
  ]
}
```

---

## Examples

**Save HTML to a custom folder**
```powershell
.\LocalPcHealthCheck.ps1 -OutFile 'C:\Reports\PC\health.html'
```

**CI pipeline: assert disks not Critical**
```powershell
$report = .\LocalPcHealth.ps1 -EmitJson | ConvertFrom-Json
$vol = $report.Sections | Where-Object Title -like 'Volumes*'
$hasCritical = ($vol.Data.PSObject.Properties.Value | Where-Object Severity -eq 'error')
if ($hasCritical) { throw "Critical volume usage detected." }
```

**Log CPU per-core as CSV**
```powershell
$cpu = ( .\LocalPcHealth.ps1 -EmitJson | ConvertFrom-Json ).Sections |
       Where-Object Title -eq 'CPU' | Select-Object -Expand Data
[pscustomobject]@{ Timestamp = Get-Date; Avg = $cpu.LoadPercent; PerCore = ($cpu.PerCore -join ';') } |
Export-Csv .\cpu_log.csv -Append -NoTypeInformation
```

---

## Troubleshooting

- **Missing `LoadPercent` or `PerCore`**: Perf counters may be disabled or inaccessible. The script falls back to WMI/counters; try running in an elevated session.
- **No `Physical Disks` data**: `Get-PhysicalDisk` requires the Windows Storage module and may not be present on all SKUs/VMs.
- **Admin rights**: Not required, but some counters/queries are more reliable elevated.
- **Execution policy**: If blocked, run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` for the current session.

---

## Security & Privacy

- Local-only; no network calls.
- Reads system metadata and disk stats; **does not** collect file contents or personally identifiable data beyond machine/user names used for display.

---

## Contributing

Issues and PRs welcome. Please:
1. Describe the scenario and environment (Windows version, PowerShell version).
2. Include sample output (`-EmitJson`) and repro steps.
3. Add/update Pester tests if changing logic.
