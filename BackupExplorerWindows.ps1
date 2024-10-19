# BackupExplorerWindows.ps1

[CmdletBinding()]
param()

Write-Verbose "Starting backup of Explorer windows."

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
}
public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
"@ -ErrorAction Stop

$shellApp = New-Object -ComObject Shell.Application -ErrorAction Stop

$windows = @($shellApp.Windows())
Write-Verbose "Number of open windows: $($windows.Count)"

$windowData = @()

foreach ($window in $windows) {
    Write-Verbose "Processing window: $($window.Name)"
    # Filter only Explorer windows (exclude browsers)
    if ($window.Name -eq "File Explorer" -or $window.Name -eq "Windows Explorer") {
        try {
            $path = $window.Document.Folder.Self.Path
            $hwnd = $window.HWND
            # Get window rectangle
            $rect = New-Object RECT
            [WinAPI]::GetWindowRect([IntPtr]$hwnd, [ref]$rect) | Out-Null
            $windowInfo = @{
                Path = $path
                Left = $rect.Left
                Top = $rect.Top
                Right = $rect.Right
                Bottom = $rect.Bottom
            }
            $windowData += $windowInfo
            Write-Verbose "Captured window at path: $path"
        } catch {
            Write-Warning "Failed to process window: $($_.Exception.Message)"
        }
    } else {
        Write-Verbose "Skipping non-Explorer window: $($window.Name)"
    }
}

if ($windowData.Count -eq 0) {
    Write-Verbose "No Explorer windows found to backup."
    exit 0
}

# Create an array to hold lines of the restore script
$restoreScriptLines = @()
$restoreScriptLines += @'
# RestoreExplorerWindows.ps1
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    public static readonly IntPtr HWND_TOP = new IntPtr(0);
    public const uint SWP_NOZORDER = 0x0004;
    public const uint SWP_NOACTIVATE = 0x0010;
}
"@

$shellApp = New-Object -ComObject Shell.Application

function GetExplorerWindowByPath($path) {
    $windows = @($shellApp.Windows())
    foreach ($window in $windows) {
        if ($window.Name -eq "File Explorer" -or $window.Name -eq "Windows Explorer") {
            if ($window.Document.Folder.Self.Path -eq $path) {
                return $window
            }
        }
    }
    return $null
}

$windowData = @(
'@

# Add window data to the restore script
foreach ($windowInfo in $windowData) {
    $path = $windowInfo.Path.Replace("'", "''")  # Escape single quotes
    $left = $windowInfo.Left
    $top = $windowInfo.Top
    $right = $windowInfo.Right
    $bottom = $windowInfo.Bottom

    $restoreScriptLines += "    @{ Path='$path'; Left=$left; Top=$top; Right=$right; Bottom=$bottom },"
}

# Remove the last comma from the last item
if ($restoreScriptLines[-1].Trim().EndsWith(',')) {
    $restoreScriptLines[-1] = $restoreScriptLines[-1].TrimEnd(',')
}

$restoreScriptLines += @'
)

foreach ($windowInfo in $windowData) {
    $path = $windowInfo.Path
    $left = $windowInfo.Left
    $top = $windowInfo.Top
    $width = $windowInfo.Right - $windowInfo.Left
    $height = $windowInfo.Bottom - $windowInfo.Top

    # Open Explorer window to $path
    $shellApp.Open($path) | Out-Null

    # Wait for the window to open
    Start-Sleep -Milliseconds 500

    # Get the window
    $window = GetExplorerWindowByPath $path

    if ($window -ne $null) {
        $hwnd = $window.HWND
        # Set window position and size
        [WinAPI]::SetWindowPos([IntPtr]$hwnd, [WinAPI]::HWND_TOP, $left, $top, $width, $height, [WinAPI]::SWP_NOZORDER -bor [WinAPI]::SWP_NOACTIVATE) | Out-Null
    } else {
        Write-Host "Failed to find window for path $path"
    }
}
'@

# Combine all lines into a single string with newlines
$restoreScriptContent = $restoreScriptLines -join "`n"

# Save the restore script to a file
$restoreScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "RestoreExplorerWindows.ps1"
$restoreScriptContent | Set-Content -Path $restoreScriptPath -Encoding UTF8 -ErrorAction Stop

Write-Host "Restore script created: $restoreScriptPath"
