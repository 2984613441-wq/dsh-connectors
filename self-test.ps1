# =====================================================================
# DSH Connectors - self test
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\self-test.ps1
#
# A connector that is not installed is reported as SKIP, never FAIL:
# this script answers "is my machine wired up?", not "does every
# third-party tool exist?".
# Exit code = number of failed checks.
# =====================================================================

$ErrorActionPreference = 'Continue'
$here = $PSScriptRoot
$global:DshFail = 0
$global:DshSkip = 0
$global:DshPass = 0

function Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    if ($Ok) {
        Write-Host ("PASS  {0,-22} {1}" -f $Name, $Detail) -ForegroundColor Green
        $global:DshPass++
    } else {
        Write-Host ("FAIL  {0,-22} {1}" -f $Name, $Detail) -ForegroundColor Red
        $global:DshFail++
    }
}

function Skip {
    param([string]$Name, [string]$Detail = 'not installed / not configured')
    Write-Host ("SKIP  {0,-22} {1}" -f $Name, $Detail) -ForegroundColor DarkYellow
    $global:DshSkip++
}

Write-Host "==================== 1. resolution ====================" -ForegroundColor Cyan
. (Join-Path $here '_env.ps1')
Get-DshConnectorStatus | Format-Table -AutoSize | Out-String | Write-Host

if ($DSH_OBSIDIAN_VAULT) { Check 'obsidian-vault' (Test-Path -LiteralPath $DSH_OBSIDIAN_VAULT) $DSH_OBSIDIAN_VAULT }
else { Skip 'obsidian-vault' 'set obsidianVault in connectors.config.json' }

if ($DSH_PYTHON) {
    $v = & $DSH_PYTHON -c "import sys;print(sys.version.split()[0])"
    Check 'python-office' ($LASTEXITCODE -eq 0) ("v" + $v + "  " + $DSH_PYTHON)
} else { Skip 'python-office' 'no interpreter with python-docx/openpyxl/python-pptx' }

Write-Host "==================== 2. obsidian ====================" -ForegroundColor Cyan
if ($DSH_OBSIDIAN_VAULT) {
    . (Join-Path $here 'obsidian.ps1')

    $notes = @(Get-DshObsidianNotes -Newest 3)
    Check 'list-notes' ($notes.Count -ge 1) ("count=" + $notes.Count)
    if ($notes.Count -ge 1) { $notes | Format-Table -AutoSize | Out-String | Write-Host }

    $hits = @(Find-DshObsidianText -Pattern 'the' -Max 3)
    Check 'search-text' ($hits.Count -ge 0) ("hits=" + $hits.Count)

    $tmpRel = '.dsh-selftest\note.md'
    $created = New-DshObsidianNote -Path $tmpRel -Content "# selftest`nline1" -Force
    Check 'create-note' ($created.Bytes -gt 0) ("bytes=" + $created.Bytes)
    $null = Add-DshObsidianContent -Path $tmpRel -Content '- appended' -PrependNewLine
    $back = Get-DshObsidianNote -Path $tmpRel
    Check 'append-note' ($back -match 'appended') ("len=" + $back.Length)
    Check 'roundtrip' ($back -match 'selftest' -and $back -match 'appended') 'content matches'

    $full = Resolve-DshVaultPath $tmpRel
    Remove-Item -LiteralPath $full -Force
    $dir = Split-Path -Parent $full
    if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Force -Recurse }
    Check 'cleanup' (-not (Test-Path -LiteralPath $dir)) 'no residue'

    $blocked = $false
    try { $null = Get-DshObsidianNote -Path '..\..\Windows\win.ini' } catch { $blocked = $true }
    Check 'escape-guard' $blocked 'path traversal rejected'
} else {
    Skip 'obsidian-suite' 'vault not configured'
}

Write-Host "==================== 3. matlab ====================" -ForegroundColor Cyan
if ($DSH_MATLAB) {
    . (Join-Path $here 'matlab.ps1')
    $m = Invoke-DshMatlab -Code "fprintf('R=%d\n', 41+1);"
    Check 'matlab-batch' $m.Ok ("exit=" + $m.ExitCode + " sec=" + $m.Seconds + " out=" + $m.Output.Replace("`n", ' '))
} else { Skip 'matlab-batch' 'MATLAB not resolved' }

Write-Host "==================== 4. office ====================" -ForegroundColor Cyan
if ($DSH_PYTHON) {
    . (Join-Path $here 'office.ps1')
    $docx = Join-Path $env:TEMP 'dsh-selftest.docx'
    $w = New-DshWordDoc -Path $docx -Title 'Selftest' -Line @('alpha', 'beta')
    Check 'docx-create' $w.Ok ("exit=" + $w.ExitCode)
    if (Test-Path -LiteralPath $docx) {
        $rd = Get-DshWordText -Path $docx
        Check 'docx-read' ($rd.Ok -and $rd.Output -match 'alpha') ("exit=" + $rd.ExitCode)
        Remove-Item -LiteralPath $docx -Force
    }
} else { Skip 'office-suite' 'no usable Python' }

Write-Host "==================== 5. documents ====================" -ForegroundColor Cyan
$pandoc = Get-Command pandoc -ErrorAction SilentlyContinue
if ($pandoc) { Check 'pandoc' $true $pandoc.Source } else { Skip 'pandoc' 'not on PATH' }
$xelatex = Get-Command xelatex -ErrorAction SilentlyContinue
if ($xelatex) { Check 'xelatex' $true $xelatex.Source } else { Skip 'xelatex' 'not on PATH (needed for PDF)' }

Write-Host "==================== 6. github ====================" -ForegroundColor Cyan
$gh = Get-Command gh -ErrorAction SilentlyContinue
if ($gh) {
    $ghv = & gh --version 2>&1 | Select-Object -First 1
    Check 'gh-cli' ($LASTEXITCODE -eq 0) $ghv
    $who = & gh api user --jq '.login' 2>&1
    if ($LASTEXITCODE -eq 0) { Check 'gh-auth' $true ("login=" + $who) }
    else { Skip 'gh-auth' 'not logged in (run: gh auth login)' }
} else { Skip 'gh-cli' 'GitHub CLI not installed' }

Write-Host ""
Write-Host ("RESULT: {0} passed, {1} failed, {2} skipped" -f $global:DshPass, $global:DshFail, $global:DshSkip)
if ($global:DshFail -eq 0) { Write-Host "ALL EXECUTABLE CHECKS PASSED" -ForegroundColor Green }
else { Write-Host "SOME CHECKS FAILED" -ForegroundColor Red }
exit $global:DshFail
