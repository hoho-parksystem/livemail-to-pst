#requires -Version 5.1
<#
.SYNOPSIS
Copies an EML directory tree to a NEW Unicode PST using Outlook MAPI and Redemption.
.DESCRIPTION
Requires classic Outlook and registered Redemption of matching process bitness.
Does not modify EML files. Existing PST/log paths are refused. See README.md.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$EmlFolderPath,
    [Parameter(Mandatory=$true)][string]$PstOutputPath,
    [ValidateNotNullOrEmpty()][string]$TempFolderName = 'LiveMail Import',
    [string]$LogPath,
    [switch]$ScanOnly
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-Com($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

# Enumerate explicitly: do not silently skip unreadable directories or follow junctions.
function Get-EmlFiles([string]$Directory) {
    foreach ($entry in Get-ChildItem -LiteralPath $Directory -Force -ErrorAction Stop) {
        if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Reparse points are not supported; copy the source tree first: $($entry.FullName)"
        }
        if ($entry.PSIsContainer) { Get-EmlFiles $entry.FullName }
        elseif ($entry.Extension -ieq '.eml') { $entry }
    }
}

$session = $stores = $store = $root = $importRoot = $writer = $lockStream = $null
$folderCache = @{}
$success = 0
$failed = 0
$exitCode = 2
try {
    if ($env:OS -ne 'Windows_NT') { throw 'Windows is required.' }
    $source = Get-Item -LiteralPath $EmlFolderPath -Force
    if (-not $source.PSIsContainer -or $source.PSProvider.Name -ne 'FileSystem') {
        throw 'EmlFolderPath must be a filesystem directory.'
    }
    if (($source.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw 'The source directory cannot be a reparse point.'
    }
    $sourcePrefix = $source.FullName.TrimEnd('\') + '\'
    $pst = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PstOutputPath)
    if ([IO.Path]::GetExtension($pst) -ine '.pst') { throw 'PstOutputPath must end in .pst.' }
    if ($pst.StartsWith($sourcePrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Store output outside the source directory.'
    }
    if (Test-Path -LiteralPath $pst) { throw 'Existing PST refused. Choose a new filename.' }
    if ($TempFolderName.IndexOfAny([char[]]'\/') -ge 0 -or [string]::IsNullOrWhiteSpace($TempFolderName)) {
        throw 'TempFolderName must be a single nonempty folder name.'
    }
    if (-not $LogPath) { $LogPath = $pst + '.import.jsonl' }
    $log = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogPath)
    if ($log -ieq $pst -or $log.StartsWith($sourcePrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'LogPath must be separate from the PST and outside the source directory.'
    }
    if (Test-Path -LiteralPath $log) { throw 'Existing log refused. Choose a new filename.' }
    foreach ($path in @($pst, $log)) {
        if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($path)) -PathType Container)) {
            throw "Output parent directory must already exist: $path"
        }
    }
    $files = @(Get-EmlFiles $source.FullName | Sort-Object FullName)
    if ($files.Count -eq 0) { throw 'No EML files found; no PST created.' }
    Write-Host "Found $($files.Count) EML files. Source is never modified."
    if ($ScanOnly) {
        Write-Host 'Scan completed. No PST/log created; COM dependencies not tested.'
        exit 0
    }
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'Use Windows PowerShell with -STA (see README.md).'
    }
    # CreateNew + exclusive lock prevents two instances from targeting the same PST.
    $lockStream = [IO.File]::Open($pst + '.import.lock', [IO.FileMode]::OpenOrCreate,
        [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    if (Test-Path -LiteralPath $pst) { throw 'PST appeared during preflight; refusing to append.' }
    try { $session = New-Object -ComObject Redemption.RDOSession }
    catch { throw 'Redemption.RDOSession is unavailable. Install/register Redemption matching Outlook and PowerShell bitness. See README.md.' }
    $logStream = [IO.File]::Open($log, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    $writer = [IO.StreamWriter]::new($logStream, [Text.UTF8Encoding]::new($false))
    $writer.AutoFlush = $true
    $writer.WriteLine((@{type='start'; source=$source.FullName; pst=$pst; count=$files.Count;
        redemptionVersion=$session.Version; utc=[DateTime]::UtcNow.ToString('o')} | ConvertTo-Json -Compress))
    # A private temporary profile avoids modifying the user's Outlook profile.
    $session.LogonPstStore($pst, 2, 'LiveMail Archive') # olStoreUnicode = 2
    $stores = $session.Stores
    $store = $stores.DefaultStore
    $root = $store.IPMRootFolder
    $rootFolders = $root.Folders
    try { $importRoot = $rootFolders.Add($TempFolderName) }
    finally { Release-Com $rootFolders }
    $folderCache[''] = $importRoot
    foreach ($file in $files) {
        $mail = $check = $items = $null
        $relative = $file.FullName.Substring($sourcePrefix.Length)
        $record = [ordered]@{type='message'; source=$relative; sha256=''; status='failed'; entryId=''; error=''}
        try {
            $record.sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            if ($file.Length -eq 0) { throw 'Empty EML file.' }
            $parent = [IO.Path]::GetDirectoryName($relative)
            $key = ''
            $folder = $importRoot
            if ($parent) {
                foreach ($part in ($parent -split '\\')) {
                    $key = if ($key) { $key + '\' + $part } else { $part }
                    if (-not $folderCache.ContainsKey($key)) {
                        $children = $folder.Folders
                        try { $folderCache[$key] = $children.Add($part) }
                        finally { Release-Com $children }
                    }
                    $folder = $folderCache[$key]
                }
            }
            $items = $folder.Items
            $mail = $items.Add('IPM.Note')
            $mail.Import($file.FullName, 1031) # olRFC822_Redemption; no shell/Inspector fallback
            $mail.Save()
            $record.entryId = [string]$mail.EntryID
            if (-not $record.entryId) { throw 'Save returned an empty EntryID.' }
            Release-Com $mail
            $mail = $null
            $check = $session.GetMessageFromID($record.entryId, $store.EntryID)
            if ($null -eq $check) { throw 'Saved message cannot be reopened.' }
            $record.status = 'saved'
            $success++
        } catch {
            $failed++
            $record.error = $_.Exception.Message
            Write-Warning "Import failed: $relative -- $($record.error)"
            # Keep any partially saved item for inspection; never conceal it as success.
        } finally {
            Release-Com $check
            Release-Com $mail
            Release-Com $items
        }
        # Log errors abort the run; they must not be mistaken for conversion failures.
        $writer.WriteLine(($record | ConvertTo-Json -Compress))
        Write-Progress -Activity 'EML to PST' -Status "$success saved; $failed failed" -PercentComplete (100 * ($success + $failed) / $files.Count)
    }
    $writer.WriteLine((@{type='summary'; total=$files.Count; saved=$success; failed=$failed;
        utc=[DateTime]::UtcNow.ToString('o')} | ConvertTo-Json -Compress))
    Write-Host "Finished: $success saved, $failed failed. Log: $log"
    $exitCode = if ($failed -eq 0) { 0 } else { 1 }
} catch {
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
} finally {
    Write-Progress -Activity 'EML to PST' -Completed
    if ($writer) { $writer.Dispose() }
    # Each cached folder is released once; importRoot is also in this cache.
    foreach ($value in $folderCache.Values) { Release-Com $value }
    Release-Com $root
    Release-Com $store
    Release-Com $stores
    if ($session) {
        try { if ($session.LoggedOn) { $session.Logoff() } }
        catch { Write-Warning "MAPI logoff failed: $($_.Exception.Message)"; $exitCode = 2 }
        finally { Release-Com $session }
    }
    if ($lockStream) { $lockStream.Dispose() }
}
exit $exitCode
