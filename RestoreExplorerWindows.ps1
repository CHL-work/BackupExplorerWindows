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
    @{ Path='C:\Users\DDB\OneDrive\Desktop\backupFolder'; Left=625; Top=534; Right=1880; Bottom=1341 },
    @{ Path='D:\SteamLibrary\steamapps\common\FINAL FANTASY VII REMAKE\End\Content\Paks\~mods'; Left=952; Top=445; Right=2207; Bottom=1252 }
)

foreach ($windowInfo in $windowData) {
    $path = $windowInfo.Path
    $left = $windowInfo.Left
    $top = $windowInfo.Top
    $width = $windowInfo.Right - $windowInfo.Left
    $height = $windowInfo.Bottom - $windowInfo.Top

    $shellApp.Open($path) | Out-Null

    Start-Sleep -Milliseconds 500

    $window = GetExplorerWindowByPath $path

    if ($window -ne $null) {
        $hwnd = $window.HWND
        # Set window position and size
        [WinAPI]::SetWindowPos([IntPtr]$hwnd, [WinAPI]::HWND_TOP, $left, $top, $width, $height, [WinAPI]::SWP_NOZORDER -bor [WinAPI]::SWP_NOACTIVATE) | Out-Null
    } else {
        Write-Host "Failed to find window for path $path"
    }
}
