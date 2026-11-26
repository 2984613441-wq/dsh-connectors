# =====================================================================
# DSH Connectors - Office documents (Word / Excel / PowerPoint) + PDF
# Uses the Python interpreter resolved by _env.ps1 (it probes for a
# distribution that actually has python-docx / openpyxl / python-pptx).
#   . .\office.ps1
#   Get-DshWordText -Path .\doc.docx
# =====================================================================

if (-not $DSH_PYTHON) { . (Join-Path $PSScriptRoot '_env.ps1') }
$script:DshOfficeHelper = Join-Path $PSScriptRoot 'office_helper.py'

function Invoke-DshOfficeHelper {
    param([Parameter(Mandatory)][string[]]$Arguments)
    if (-not $DSH_PYTHON) {
        throw "No Python with python-docx/openpyxl/python-pptx was found. Set 'python' in connectors.config.json."
    }
    if (-not (Test-Path -LiteralPath $script:DshOfficeHelper)) {
        throw "helper not found: $script:DshOfficeHelper"
    }
    $out  = & $DSH_PYTHON $script:DshOfficeHelper @Arguments 2>&1
    $code = $LASTEXITCODE
    [pscustomobject]@{
        ExitCode = $code
        Ok       = ($code -eq 0)
        Output   = ($out | Out-String).Trim()
    }
}

function Get-DshWordText {
    <# Extracts text and table dimensions from a .docx. #>
    param([Parameter(Mandatory)][string]$Path)
    Invoke-DshOfficeHelper @('docx-read', $Path)
}

function New-DshWordDoc {
    <# Creates a .docx from a title plus one paragraph per -Line. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Title = '',
        [string[]]$Line = @()
    )
    $a = @('docx-write', $Path, '--title', $Title)
    foreach ($l in $Line) { $a += @('--line', $l) }
    Invoke-DshOfficeHelper $a
}

function Get-DshExcelPreview {
    <# Lists sheets and previews the first N rows of a .xlsx. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$Rows = 10
    )
    Invoke-DshOfficeHelper @('xlsx-read', $Path, '--rows', "$Rows")
}

function Get-DshPowerPointOutline {
    <# Lists slides and their text from a .pptx. #>
    param([Parameter(Mandatory)][string]$Path)
    Invoke-DshOfficeHelper @('pptx-read', $Path)
}

function Get-DshPdfViaPandoc {
    <# Converts a document to PDF with pandoc (requires a LaTeX engine for PDF output). #>
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [string]$OutputPath = ''
    )
    if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
        throw "pandoc not found on PATH."
    }
    if (-not $OutputPath) { $OutputPath = [System.IO.Path]::ChangeExtension($InputPath, '.pdf') }
    $out = & pandoc $InputPath -o $OutputPath --pdf-engine=xelatex 2>&1
    [pscustomobject]@{
        ExitCode = $LASTEXITCODE
        Ok       = ($LASTEXITCODE -eq 0)
        Output   = ($out | Out-String).Trim()
        Pdf      = $OutputPath
    }
}
