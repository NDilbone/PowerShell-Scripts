# PowerShell-Scripts

## 1) Local PC Health Check
Generates a single HTML report of a Windows machine. No external modules required. Works on Windows 10 and 11.

### Checks
- OS and uptime
- CPU load
- Memory usage
- Disk space by volume
- Physical disk health
- Windows Update pending items
- BitLocker status
- Windows Defender status
- Startup programs
- Services set to Automatic but stopped
- Recent critical and error events (last 24 hours)
- Network info

### Usage
Open PowerShell as Admin
```
.\LocalPcHealthCheck.ps1 -OutFile "C:\\Users\\Public\\health.html"
```

### Exit Codes
- 0 success
- 2 warnings found
- 3 errors found

### Outputs
- HTML report at the path you choose
- Optional JSON sidecar if you add -EmitJson

> [!NOTE]
> Admin rights improve coverage. The script still runs without Admin and marks restricted sections. No internet access required.
