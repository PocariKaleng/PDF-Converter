param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [Parameter(Mandatory = $true)][string]$PidPath
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class WordProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
$taskWord = $null
$taskDocument = $null
$taskProbe = $null
$taskWordPid = 0
$taskExitCode = 0
$taskPreviousIds = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
try {
    $taskWord = New-Object -ComObject Word.Application
    $taskWord.Visible = $false
    $taskWord.DisplayAlerts = 0
    $taskWord.AutomationSecurity = 3 # Disable macros.
    # Word exposes Hwnd on Window, not Application. Create a blank document
    # first so the worker can identify its process before opening the upload.
    $taskProbe = $taskWord.Documents.Add()
    [uint32]$taskWordPid = 0
    [WordProcess]::GetWindowThreadProcessId([IntPtr]$taskProbe.ActiveWindow.Hwnd, [ref]$taskWordPid) | Out-Null
    if ($taskWordPid -gt 0 -and $taskWordPid -notin $taskPreviousIds) {
        [IO.File]::WriteAllText($PidPath, [string]$taskWordPid)
    } else {
        throw 'Word tidak membuat proses terpisah. Tutup dialog Word lalu coba lagi.'
    }
    $taskProbe.Close(0)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskProbe) | Out-Null
    $taskProbe = $null
    $taskWord.Options.UpdateLinksAtOpen = $false
    $taskWord.Options.SaveNormalPrompt = $false
    # ConfirmConversions, ReadOnly, AddToRecentFiles = false, true, false.
    $taskDocument = $taskWord.Documents.Open($InputPath, $false, $true, $false, '', '', $false, '', '', 0, 0, $false)
    $taskDocument.Repaginate()
    # PDF, print quality, all pages, no markup, heading bookmarks, tags,
    # no rasterized missing fonts. Use ordinary PDF because Word substitutes
    # CFF OpenType fonts in PDF/A mode. Python embeds their original CFF data.
    $taskDocument.ExportAsFixedFormat($OutputPath, 17, $false, 0, 0, 1, 1, 0, $true, $false, 1, $true, $false, $false)
    Write-Output 'PDF exported.'
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    $taskExitCode = 1
} finally {
    if ($null -ne $taskProbe) {
        try { $taskProbe.Close(0) } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskProbe) | Out-Null
    }
    if ($null -ne $taskDocument) {
        try { $taskDocument.Close(0) } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskDocument) | Out-Null
    }
    if ($null -ne $taskWord) {
        # Only quit the instance owned by this worker.
        if ($taskWordPid -gt 0 -and $taskWordPid -notin $taskPreviousIds) {
            try { $taskWord.Quit(0) } catch {}
        }
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskWord) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
exit $taskExitCode
