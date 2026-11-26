# =====================================================================
# DSH Connectors - one-line entry point
#
# WHY THE PROCESS-LEVEL BYPASS
#   Corporate/enterprise Windows images sometimes ship an execution
#   policy (or a directory ACL) that refuses unsigned local scripts.
#   Setting the policy for THIS PROCESS ONLY keeps the machine's
#   persistent policy untouched:
#       Get-ExecutionPolicy            -> unchanged afterwards
#
# USAGE
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\use.ps1
#   # or, in a shell that already runs scripts:
#   . .\use.ps1
#
# AFTER LOADING
#   Get-DshConnectorStatus
#   Get-DshObsidianNotes / Get-DshObsidianNote / Find-DshObsidianText
#   New-DshObsidianNote  / Add-DshObsidianContent / Start-DshObsidian
#   Invoke-DshMatlab
#   Get-DshWordText / New-DshWordDoc / Get-DshExcelPreview
#   Get-DshPowerPointOutline / Get-DshPdfViaPandoc
# =====================================================================

try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
} catch {
    Write-Warning ("Could not set process execution policy: " + $_.Exception.Message)
}

. (Join-Path $PSScriptRoot '_env.ps1')
. (Join-Path $PSScriptRoot 'obsidian.ps1')
. (Join-Path $PSScriptRoot 'matlab.ps1')
. (Join-Path $PSScriptRoot 'office.ps1')

Write-Host "DSH Connectors loaded. Try: Get-DshConnectorStatus" -ForegroundColor Green
