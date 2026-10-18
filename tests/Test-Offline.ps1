#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'eml_to_outlook.ps1'
$tokens = $null
$parseErrors = $null
[void][Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$sandbox = Join-Path ([IO.Path]::GetTempPath()) ('eml-pst-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($sandbox)
$source = Join-Path $sandbox 'mail [sample]'
[void][IO.Directory]::CreateDirectory($source)
$nested = Join-Path $source 'nested'
[void][IO.Directory]::CreateDirectory($nested)
$output = Join-Path $sandbox 'new.pst'
$engine = Join-Path $PSHOME 'powershell.exe'
if (-not (Test-Path -LiteralPath $engine)) { $engine = (Get-Command powershell.exe).Source }
$checks = 0
function Invoke-Case([string]$Name, [string[]]$Arguments, [int]$ExpectedCode, [string]$ExpectedText) {
    $ErrorActionPreference = 'Continue' # Expected child stderr is captured as test evidence.
    $lines = & $engine -NoProfile -ExecutionPolicy Bypass -STA -File $scriptPath @Arguments 2>&1
    $actual = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    # Console-host formatting can wrap stderr inside words; assert messages only on success.
    if ($actual -ne $ExpectedCode -or ($ExpectedCode -eq 0 -and ($lines | Out-String) -notmatch $ExpectedText)) {
        throw "${Name}: exit $actual (expected $ExpectedCode). Output: $lines"
    }
    $script:checks++
    Write-Host "PASS: $Name"
}
try {
    Invoke-Case 'empty source fails' @('-EmlFolderPath',$source,'-PstOutputPath',$output,'-ScanOnly') 2 'No EML files'
    $eml = Join-Path $source 'one.eml'
    [IO.File]::WriteAllText($eml, "From: sender@example.invalid`r`nTo: recipient@example.invalid`r`nSubject: Sample`r`n`r`nBody")
    [IO.File]::WriteAllText((Join-Path $nested 'two.EML'), 'synthetic sample')
    [IO.File]::WriteAllText((Join-Path $source 'ignore.txt'), 'ignored')
    $before = (Get-FileHash -LiteralPath $eml).Hash
    Invoke-Case 'literal brackets and recursive enumeration' @('-EmlFolderPath',$source,'-PstOutputPath',$output,'-ScanOnly') 0 'Found 2 EML'
    if ((Test-Path -LiteralPath $output) -or (Test-Path -LiteralPath ($output + '.import.jsonl'))) { throw 'Scan wrote output' }
    if ((Get-FileHash -LiteralPath $eml).Hash -ne $before) { throw 'Scan modified source' }
    Invoke-Case 'source-contained PST refused' @('-EmlFolderPath',$source,'-PstOutputPath',(Join-Path $source 'bad.pst'),'-ScanOnly') 2 'outside the source'
    Invoke-Case 'wrong extension refused' @('-EmlFolderPath',$source,'-PstOutputPath',(Join-Path $sandbox 'bad.txt'),'-ScanOnly') 2 'end in .pst'
    Invoke-Case 'source-contained log refused' @('-EmlFolderPath',$source,'-PstOutputPath',$output,'-LogPath',(Join-Path $source 'bad.jsonl'),'-ScanOnly') 2 'outside the source'
    Invoke-Case 'PST and log collision refused' @('-EmlFolderPath',$source,'-PstOutputPath',$output,'-LogPath',$output,'-ScanOnly') 2 'separate from the PST'
    Invoke-Case 'invalid import folder refused' @('-EmlFolderPath',$source,'-PstOutputPath',$output,'-TempFolderName','bad\folder','-ScanOnly') 2 'single nonempty'
    [IO.File]::WriteAllText($output, 'existing PST sentinel')
    Invoke-Case 'existing PST refused' @('-EmlFolderPath',$source,'-PstOutputPath',$output,'-ScanOnly') 2 'Existing PST refused'
    if ([IO.File]::ReadAllText($output) -ne 'existing PST sentinel') { throw 'Existing PST was modified' }
    $other = Join-Path $sandbox 'other.pst'
    [IO.File]::WriteAllText(($other + '.import.jsonl'), 'existing log sentinel')
    Invoke-Case 'existing log refused' @('-EmlFolderPath',$source,'-PstOutputPath',$other,'-ScanOnly') 2 'Existing log refused'
    Invoke-Case 'missing parent refused' @('-EmlFolderPath',$source,'-PstOutputPath',(Join-Path $sandbox 'missing\new.pst'),'-ScanOnly') 2 'must already exist'
    Invoke-Case 'file as source refused' @('-EmlFolderPath',$eml,'-PstOutputPath',$other,'-ScanOnly') 2 'filesystem directory'
    Write-Host "All $checks offline checks passed. Real Outlook/Redemption import is NOT tested here."
} finally {
    # Delete only the unique temporary test tree created above.
    $resolvedSandbox = [IO.Path]::GetFullPath($sandbox)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolvedSandbox.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedSandbox).StartsWith('eml-pst-test-')) {
        Remove-Item -LiteralPath $resolvedSandbox -Recurse -Force
    }
}
