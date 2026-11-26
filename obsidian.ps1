# =====================================================================
# DSH Connectors - Obsidian (filesystem level, no plugin required)
#   . .\obsidian.ps1
# Verified pattern: write -> read-back identical -> delete -> no residue.
# Every path is confined to the vault root; traversal outside is rejected.
# =====================================================================

if (-not $DSH_OBSIDIAN_VAULT) { . (Join-Path $PSScriptRoot '_env.ps1') }

$script:DshVaultExclude = '\\\.obsidian\\|\\node_modules\\|\\\.backup|\\\.git\\'

function Get-DshVaultRoot {
    if (-not $DSH_OBSIDIAN_VAULT) {
        throw "Obsidian vault not resolved. Set obsidianVault in connectors.config.json (see connectors.config.example.json)."
    }
    if (-not (Test-Path -LiteralPath $DSH_OBSIDIAN_VAULT)) {
        throw "Obsidian vault not found: $DSH_OBSIDIAN_VAULT"
    }
    return (Get-Item -LiteralPath $DSH_OBSIDIAN_VAULT).FullName
}

function Resolve-DshVaultPath {
    <# Resolves a vault-relative path; rejects anything escaping the vault. #>
    param([Parameter(Mandatory)][string]$RelativePath)
    $root = Get-DshVaultRoot
    $full = [System.IO.Path]::GetFullPath((Join-Path $root $RelativePath))
    if (-not $full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refused: path escapes the vault -> $RelativePath"
    }
    return $full
}

function Get-DshObsidianNotes {
    <# Lists markdown notes, newest first. #>
    param([int]$Newest = 20, [string]$Folder = '')
    $root = Get-DshVaultRoot
    $base = if ($Folder) { Resolve-DshVaultPath $Folder } else { $root }
    Get-ChildItem -LiteralPath $base -Filter *.md -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch $script:DshVaultExclude } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First $Newest |
        ForEach-Object {
            [pscustomobject]@{
                RelPath  = $_.FullName.Substring($root.Length).TrimStart('\')
                Bytes    = $_.Length
                Modified = $_.LastWriteTime
            }
        }
}

function Get-DshObsidianNote {
    <# Reads one note by vault-relative path. #>
    param([Parameter(Mandatory)][string]$Path)
    $full = Resolve-DshVaultPath $Path
    if (-not (Test-Path -LiteralPath $full)) { throw "Note not found: $Path" }
    return [System.IO.File]::ReadAllText($full)
}

function Find-DshObsidianText {
    <# Full-text search: returns note path + line number + the matching line. #>
    param([Parameter(Mandatory)][string]$Pattern, [int]$Max = 30)
    $root = Get-DshVaultRoot
    Get-ChildItem -LiteralPath $root -Filter *.md -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch $script:DshVaultExclude } |
        Select-String -Pattern $Pattern -List -ErrorAction SilentlyContinue |
        Select-Object -First $Max |
        ForEach-Object {
            [pscustomobject]@{
                RelPath = $_.Path.Substring($root.Length).TrimStart('\')
                Line    = $_.LineNumber
                Text    = ([string]$_.Line).Trim()
            }
        }
}

function New-DshObsidianNote {
    <# Creates a note; refuses to overwrite unless -Force. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Content = '',
        [switch]$Force
    )
    $full = Resolve-DshVaultPath $Path
    if ((Test-Path -LiteralPath $full) -and -not $Force) {
        throw "Refused: note already exists (use -Force) -> $Path"
    }
    $dir = Split-Path -Parent $full
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($full, $Content, (New-Object System.Text.UTF8Encoding($false)))
    [pscustomobject]@{ RelPath = $Path; Bytes = (Get-Item -LiteralPath $full).Length; Action = 'created' }
}

function Add-DshObsidianContent {
    <# Appends text to an existing note. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content,
        [switch]$PrependNewLine
    )
    $full = Resolve-DshVaultPath $Path
    if (-not (Test-Path -LiteralPath $full)) { throw "Note not found: $Path" }
    $prefix = if ($PrependNewLine) { "`r`n" } else { '' }
    [System.IO.File]::AppendAllText($full, ($prefix + $Content), (New-Object System.Text.UTF8Encoding($false)))
    [pscustomobject]@{ RelPath = $Path; Bytes = (Get-Item -LiteralPath $full).Length; Action = 'appended' }
}

function Start-DshObsidian {
    <# Launches Obsidian if it is not already running. #>
    $p = Get-Process Obsidian -ErrorAction SilentlyContinue
    if ($p) { return [pscustomobject]@{ AlreadyRunning = $true; Pids = ($p.Id -join ',') } }
    if (-not $DSH_OBSIDIAN_EXE) { throw "Obsidian.exe not resolved. Set obsidianExe in connectors.config.json." }
    Start-Process -FilePath $DSH_OBSIDIAN_EXE
    [pscustomobject]@{ AlreadyRunning = $false; Started = $true }
}
