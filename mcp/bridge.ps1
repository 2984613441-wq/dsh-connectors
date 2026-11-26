# =====================================================================
# dsh-connectors MCP bridge
#
# Invoked once per tool call by server.mjs:
#   powershell -NoProfile -ExecutionPolicy Bypass -File bridge.ps1 -Tool <name>
#
# Contract:
#   stdin  : the tool arguments as a JSON object (never argv - note text
#            contains quotes, newlines and non-ASCII characters)
#   stdout : exactly one JSON object
#   stderr : free-form diagnostics (never parsed)
#
# ASCII-only file: Windows PowerShell 5.1 misreads BOM-less UTF-8 scripts.
# =====================================================================
param(
    [Parameter(Mandatory = $true)]
    [string]$Tool
)

# Console must speak UTF-8 in both directions. PS 5.1 otherwise defaults to the
# OEM code page and turns every CJK character into mojibake before Node sees it.
try {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [Console]::OutputEncoding = $utf8
    [Console]::InputEncoding = $utf8
} catch { }

$ErrorActionPreference = 'Stop'
try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue
} catch { }

$here = $PSScriptRoot
$root = Split-Path -Parent $here
. (Join-Path $root '_env.ps1')
. (Join-Path $root 'obsidian.ps1')
. (Join-Path $root 'matlab.ps1')
. (Join-Path $root 'office.ps1')

$raw = [Console]::In.ReadToEnd()
$payload = [pscustomobject]@{}
if ($raw -and $raw.Trim()) {
    $payload = $raw | ConvertFrom-Json
}

function Arg {
    param([string]$Name, $Default = $null)
    $value = $payload.$Name
    if ($null -eq $value) { return $Default }
    return $value
}

$result = [ordered]@{ ok = $true; tool = $Tool }

try {
    switch ($Tool) {

        'connector_status' {
            $result.data = @(Get-DshConnectorStatus | ForEach-Object {
                [pscustomobject]@{
                    connector = $_.Connector
                    resolved  = $_.Resolved
                    path      = $_.Path
                    exists    = $_.Exists
                }
            })
        }

        'obsidian_list' {
            $newest = [int](Arg 'newest' 20)
            $folder = [string](Arg 'folder' '')
            $result.data = @(Get-DshObsidianNotes -Newest $newest -Folder $folder | ForEach-Object {
                [pscustomobject]@{
                    path     = $_.RelPath
                    bytes    = $_.Bytes
                    modified = $_.Modified.ToString('s')
                }
            })
        }

        'obsidian_read' {
            $notePath = [string](Arg 'path' '')
            if (-not $notePath) { throw 'obsidian_read requires path' }
            $content = Get-DshObsidianNote -Path $notePath
            $result.data = [pscustomobject]@{
                path  = $notePath
                bytes = $content.Length
                text  = $content
            }
        }

        'obsidian_search' {
            $pattern = [string](Arg 'pattern' '')
            if (-not $pattern) { throw 'obsidian_search requires pattern' }
            $max = [int](Arg 'max' 30)
            $result.data = @(Find-DshObsidianText -Pattern $pattern -Max $max | ForEach-Object {
                [pscustomobject]@{
                    path = $_.RelPath
                    line = $_.Line
                    text = $_.Text
                }
            })
        }

        'obsidian_create' {
            $notePath = [string](Arg 'path' '')
            if (-not $notePath) { throw 'obsidian_create requires path' }
            $content  = [string](Arg 'content' '')
            $overwrite = [bool](Arg 'overwrite' $false)
            $created = New-DshObsidianNote -Path $notePath -Content $content -Force:$overwrite
            $result.data = [pscustomobject]@{
                path   = $created.RelPath
                bytes  = $created.Bytes
                action = $created.Action
            }
        }

        'obsidian_append' {
            $notePath = [string](Arg 'path' '')
            $content  = [string](Arg 'content' '')
            if (-not $notePath) { throw 'obsidian_append requires path' }
            if (-not $content) { throw 'obsidian_append requires content' }
            $newline = [bool](Arg 'newline' $false)
            $appended = Add-DshObsidianContent -Path $notePath -Content $content -PrependNewLine:$newline
            $result.data = [pscustomobject]@{
                path   = $appended.RelPath
                bytes  = $appended.Bytes
                action = $appended.Action
            }
        }

        'matlab_run' {
            $code = [string](Arg 'code' '')
            $file = [string](Arg 'file' '')
            if (-not $code -and -not $file) { throw 'matlab_run requires code or file' }
            if ($code -and $file) { throw 'matlab_run takes code or file, not both' }
            if ($file) { $run = Invoke-DshMatlab -File $file }
            else       { $run = Invoke-DshMatlab -Code $code }
            $result.data = [pscustomobject]@{
                exitCode = $run.ExitCode
                seconds  = $run.Seconds
                output   = $run.Output
            }
            if (-not $run.Ok) { $result.ok = $false; $result.error = 'MATLAB exited with code ' + $run.ExitCode }
        }

        'office_read' {
            $filePath = [string](Arg 'path' '')
            if (-not $filePath) { throw 'office_read requires path' }
            if (-not (Test-Path -LiteralPath $filePath)) { throw ('file not found: ' + $filePath) }
            $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
            $rows = [int](Arg 'rows' 10)
            switch ($ext) {
                '.docx' { $read = Get-DshWordText -Path $filePath }
                '.xlsx' { $read = Get-DshExcelPreview -Path $filePath -Rows $rows }
                '.pptx' { $read = Get-DshPowerPointOutline -Path $filePath }
                default { throw ('office_read supports .docx/.xlsx/.pptx, got ' + $ext) }
            }
            $result.data = [pscustomobject]@{
                format   = $ext
                exitCode = $read.ExitCode
                output   = $read.Output
            }
            if (-not $read.Ok) { $result.ok = $false; $result.error = 'helper exited with code ' + $read.ExitCode }
        }

        'office_write_docx' {
            $filePath = [string](Arg 'path' '')
            if (-not $filePath) { throw 'office_write_docx requires path' }
            $title = [string](Arg 'title' '')
            $lines = @()
            if ($null -ne $payload.lines) { $lines = @($payload.lines) }
            $written = New-DshWordDoc -Path $filePath -Title $title -Line $lines
            $result.data = [pscustomobject]@{
                path     = $filePath
                exitCode = $written.ExitCode
                output   = $written.Output
            }
            if (-not $written.Ok) { $result.ok = $false; $result.error = 'helper exited with code ' + $written.ExitCode }
        }

        'pdf_convert' {
            $inputPath  = [string](Arg 'input' '')
            $outputPath = [string](Arg 'output' '')
            if (-not $inputPath) { throw 'pdf_convert requires input' }
            if (-not (Test-Path -LiteralPath $inputPath)) { throw ('file not found: ' + $inputPath) }
            $conv = Get-DshPdfViaPandoc -InputPath $inputPath -OutputPath $outputPath
            $bytes = 0
            if (Test-Path -LiteralPath $conv.Pdf) { $bytes = (Get-Item -LiteralPath $conv.Pdf).Length }
            $result.data = [pscustomobject]@{
                pdf      = $conv.Pdf
                bytes    = $bytes
                exitCode = $conv.ExitCode
                output   = $conv.Output
            }
            if (-not $conv.Ok) { $result.ok = $false; $result.error = 'pandoc exited with code ' + $conv.ExitCode }
        }

        default {
            throw ('unknown tool: ' + $Tool)
        }
    }
} catch {
    $result.ok = $false
    $result.error = $_.Exception.Message
}

$result | ConvertTo-Json -Compress -Depth 8
