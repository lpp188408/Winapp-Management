$ErrorActionPreference = "Continue"

$root = Split-Path -Parent $PSScriptRoot
$logPath = Join-Path $root "excel-diagnostic-log.txt"

function Write-Line {
    param([string]$Text)
    $Text | Tee-Object -FilePath $logPath -Append
}

if (Test-Path $logPath) {
    Remove-Item $logPath -Force
}

Write-Line "Winapp Management Excel Diagnostic"
Write-Line "Time: $(Get-Date)"
Write-Line "Computer: $env:COMPUTERNAME"
Write-Line "User: $env:USERNAME"
Write-Line ""

try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Line "ERROR: 没有获取到正在运行的 Excel COM 对象。"
    Write-Line "可能原因：Excel 没打开、Excel 权限和本工具不一致、Excel 忙碌，或当前不是 Microsoft Excel。"
    Write-Line "Detail: $($_.Exception.Message)"
    exit 1
}

Write-Line "Excel.Application acquired."
Write-Line "Version: $($excel.Version)"
Write-Line ""

Write-Line "== Workbooks =="
try {
    Write-Line "Count: $($excel.Workbooks.Count)"
    for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
        try {
            $wb = $excel.Workbooks.Item($i)
            Write-Line "[$i] Name: $($wb.Name)"
            Write-Line "[$i] FullName: $($wb.FullName)"
            Write-Line "[$i] Path: $($wb.Path)"
            Write-Line "[$i] Saved: $($wb.Saved)"
            Write-Line ""
        } catch {
            Write-Line "[$i] ERROR: $($_.Exception.Message)"
        }
    }
} catch {
    Write-Line "ERROR reading Workbooks: $($_.Exception.Message)"
}

Write-Line "== Excel Windows =="
try {
    Write-Line "Count: $($excel.Windows.Count)"
    for ($i = 1; $i -le $excel.Windows.Count; $i++) {
        try {
            $window = $excel.Windows.Item($i)
            Write-Line "[$i] Caption: $($window.Caption)"
            Write-Line "[$i] Hwnd: $($window.Hwnd)"
            try {
                $sheet = $window.ActiveSheet
                $wb = $sheet.Parent
                Write-Line "[$i] Active Workbook Name: $($wb.Name)"
                Write-Line "[$i] Active Workbook FullName: $($wb.FullName)"
                Write-Line "[$i] Active Workbook Path: $($wb.Path)"
            } catch {
                Write-Line "[$i] Active Workbook ERROR: $($_.Exception.Message)"
            }
            Write-Line ""
        } catch {
            Write-Line "[$i] ERROR: $($_.Exception.Message)"
        }
    }
} catch {
    Write-Line "ERROR reading Excel Windows: $($_.Exception.Message)"
}

Write-Line "Diagnostic log saved to: $logPath"
