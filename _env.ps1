# =====================================================================
# DSH Connectors - environment resolution
#
# ASCII-only on purpose: Windows PowerShell 5.1 misreads BOM-less UTF-8.
#
# Every path is resolved in this order:
#   1. connectors.config.json  (copy connectors.config.example.json)
#   2. PATH lookup
#   3. Windows registry Uninstall keys (DisplayIcon of a matching product)
#   4. well-known install folders
# Nothing is guessed: an unresolved tool reports $null, never a fake path.
# =====================================================================

$ErrorActionPreference = 'SilentlyContinue'

$script:DshRoot       = $PSScriptRoot
$script:DshConfigPath = Join-Path $PSScriptRoot 'connectors.config.json'

# ---------------------------------------------------------------- config
function Get-DshConfig {
    if (-not (Test-Path -LiteralPath $script:DshConfigPath)) { return $null }
    try {
        return (Get-Content -LiteralPath $script:DshConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch {
        Write-Warning ("connectors.config.json is not valid JSON: " + $_.Exception.Message)
        return $null
    }
}
$script:DshCfg = Get-DshConfig

function Get-DshCfgValue {
    param([string]$Name)
    if (-not $script:DshCfg) { return $null }
    $v = $script:DshCfg.$Name
    if ($v -and (Test-Path -LiteralPath $v)) { return $v }
    return $null
}

# ------------------------------------------------------------- resolvers
function Find-DshFromRegistry {
    <#
      Finds an executable through the Uninstall keys.
      DisplayIcon often points at an .ico rather than the program, so when the
      icon is not an .exe we look for <ExeName> beside it and under the
      registered InstallLocation.
    #>
    param([string]$ProductPattern, [string]$ExeName)

    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $hit = Get-ItemProperty $roots -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match $ProductPattern } |
        Select-Object -First 1
    if (-not $hit) { return $null }

    # 1) DisplayIcon that is already an executable (but never an installer)
    if ($hit.DisplayIcon) {
        $icon = ($hit.DisplayIcon.Trim('"') -split ',')[0]
        $leaf = [System.IO.Path]::GetFileNameWithoutExtension($icon)
        $isInstaller = $leaf -match '^(setup|install|installer|unins\d*|uninstall|update|setuplauncher)$'
        if (-not $isInstaller -and $icon -match '\.exe$' -and (Test-Path -LiteralPath $icon)) { return $icon }
    }

    # 2) the executable beside the icon (icons live in bin, util, resources...)
    if ($hit.DisplayIcon -and $ExeName) {
        $iconDir = Split-Path -Parent (($hit.DisplayIcon.Trim('"') -split ',')[0])
        $searchDirs = @($iconDir, (Join-Path $iconDir 'bin'), (Join-Path $iconDir 'bin\win64'))
        if ($iconDir) {
            $up1 = Split-Path -Parent $iconDir
            $searchDirs += @($up1, (Join-Path $up1 'bin'), (Join-Path $up1 'bin\win64'))
        }
        foreach ($d in $searchDirs) {
            if (-not $d) { continue }
            $cand = Join-Path $d $ExeName
            if (Test-Path -LiteralPath $cand) { return $cand }
        }
    }

    # 3) the registered install location
    if ($hit.InstallLocation -and $ExeName) {
        $loc = $hit.InstallLocation.Trim('"').TrimEnd('\')
        foreach ($d in @($loc, (Join-Path $loc 'bin'), (Join-Path $loc 'bin\win64'))) {
            if (-not $d) { continue }
            $cand = Join-Path $d $ExeName
            if (Test-Path -LiteralPath $cand) { return $cand }
        }
    }
    return $null
}

function Find-DshByGlob {
    <# Expands a wildcard path and returns the first existing match. #>
    param([string[]]$Patterns = @())
    foreach ($p in $Patterns) {
        $hits = Get-Item -Path $p -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hits) { return $hits.FullName }
    }
    return $null
}

function Resolve-DshTool {
    param(
        [string]$ConfigKey,
        [string]$ExeName,
        [string]$RegistryPattern,
        [string[]]$Globs = @(),
        [string[]]$Fallbacks = @()
    )
    $fromCfg = Get-DshCfgValue $ConfigKey
    if ($fromCfg) { return $fromCfg }

    if ($ExeName) {
        $cmd = Get-Command $ExeName -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    if ($RegistryPattern) {
        $r = Find-DshFromRegistry -ProductPattern $RegistryPattern -ExeName $ExeName
        if ($r) { return $r }
    }
    if ($Globs.Count -gt 0) {
        $g = Find-DshByGlob $Globs
        if ($g) { return $g }
    }
    foreach ($f in $Fallbacks) {
        if ($f -and (Test-Path -LiteralPath $f)) { return $f }
    }
    return $null
}

# -------------------------------------------------------------- Obsidian
function Find-DshObsidianVault {
    <# Finds the Obsidian vault: config, else the vault marked open in obsidian.json. #>
    $fromCfg = Get-DshCfgValue 'obsidianVault'
    if ($fromCfg) { return $fromCfg }

    $reg = Join-Path $env:APPDATA 'obsidian\obsidian.json'
    if (Test-Path -LiteralPath $reg) {
        try {
            $j = Get-Content -LiteralPath $reg -Raw -Encoding UTF8 | ConvertFrom-Json
            $vaults = @($j.vaults.PSObject.Properties | ForEach-Object { $_.Value })
            $pick = $vaults | Where-Object { $_.open } | Select-Object -First 1
            if (-not $pick) { $pick = $vaults | Sort-Object ts -Descending | Select-Object -First 1 }
            if ($pick -and (Test-Path -LiteralPath $pick.path)) { return $pick.path }
        } catch { }
    }
    return $null
}

# ---------------------------------------------------------------- Python
function Find-DshPython {
    <# Returns the first interpreter that can import the Office libraries. #>
    $candidates = New-Object System.Collections.ArrayList

    $cfg = Get-DshCfgValue 'python'
    if ($cfg) { [void]$candidates.Add($cfg) }

    foreach ($n in @('python', 'python3', 'py')) {
        $c = Get-Command $n -ErrorAction SilentlyContinue
        if ($c) { [void]$candidates.Add($c.Source) }
    }

    # Python Launcher inventory (py -0p) - covers interpreters not on PATH.
    if (Get-Command 'py' -ErrorAction SilentlyContinue) {
        foreach ($line in (& py -0p 2>$null)) {
            if ($line -match '([A-Za-z]:\\.*?python\.exe)') {
                [void]$candidates.Add($Matches[1])
            }
        }
    }

    foreach ($c in ($candidates | Select-Object -Unique)) {
        if (-not $c -or -not (Test-Path -LiteralPath $c)) { continue }
        & $c -c "import docx, openpyxl, pptx" 2>$null
        if ($LASTEXITCODE -ne 0) { continue }

        # py.exe is a launcher, not an interpreter: resolve what it selects.
        $leaf = [System.IO.Path]::GetFileNameWithoutExtension($c)
        if ($leaf -ieq 'py') {
            $real = & $c -c "import sys;print(sys.executable)" 2>$null
            if ($real -and (Test-Path -LiteralPath $real)) { return $real }
        }
        return $c
    }
    return $null
}

# ------------------------------------------------------------ assignment
$DSH_OBSIDIAN_VAULT = Find-DshObsidianVault
$DSH_OBSIDIAN_EXE   = Resolve-DshTool -ConfigKey 'obsidianExe' -ExeName 'Obsidian.exe' `
                        -RegistryPattern '^Obsidian$' `
                        -Globs @("$env:LOCALAPPDATA\Obsidian\Obsidian.exe", "$env:LOCALAPPDATA\Programs\obsidian\Obsidian.exe", 'C:\Program Files\Obsidian\Obsidian.exe')

$DSH_PYTHON         = Find-DshPython

$DSH_MATLAB         = Resolve-DshTool -ConfigKey 'matlab' -ExeName 'matlab.exe' `
                        -RegistryPattern '^MATLAB R' `
                        -Globs @('C:\Program Files\MATLAB\R*\bin\matlab.exe')

$DSH_COMSOL_BATCH   = Resolve-DshTool -ConfigKey 'comsolBatch' -ExeName 'comsolbatch.exe' `
                        -RegistryPattern '^COMSOL Multiphysics' `
                        -Globs @('C:\Program Files\COMSOL\COMSOL*\Multiphysics\bin\win64\comsolbatch.exe')

$DSH_COMSOL         = Resolve-DshTool -ConfigKey 'comsol' -ExeName 'comsol.exe' `
                        -Globs @('C:\Program Files\COMSOL\COMSOL*\Multiphysics\bin\win64\comsol.exe')

$DSH_CUBEMX         = Resolve-DshTool -ConfigKey 'cubeMx' -ExeName 'STM32CubeMX.exe' `
                        -RegistryPattern '^STM32CubeMX' `
                        -Globs @('C:\Program Files\STMicroelectronics\STM32Cube\STM32CubeMX\STM32CubeMX.exe')

$DSH_PROG_CLI       = Resolve-DshTool -ConfigKey 'stm32ProgrammerCli' -ExeName 'STM32_Programmer_CLI.exe' `
                        -RegistryPattern '^STM32CubeProgrammer' `
                        -Globs @('C:\Program Files\STMicroelectronics\STM32Cube\STM32CubeProgrammer\bin\STM32_Programmer_CLI.exe')

$DSH_KEIL           = Resolve-DshTool -ConfigKey 'keil' -ExeName 'UV4.exe' `
                        -RegistryPattern '^Keil' `
                        -Globs @('C:\Keil_v5\UV4\UV4.exe', 'C:\Keil\UV4\UV4.exe')

$DSH_VSCODE         = Resolve-DshTool -ConfigKey 'vscode' -ExeName 'code.cmd' `
                        -Globs @("$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe", 'C:\Program Files\Microsoft VS Code\Code.exe')

$DSH_NOTION         = Resolve-DshTool -ConfigKey 'notion' -ExeName 'Notion.exe' `
                        -Globs @("$env:LOCALAPPDATA\Programs\Notion\Notion.exe")

$DSH_DRAWIO         = Resolve-DshTool -ConfigKey 'drawio' -ExeName 'draw.io.exe' `
                        -Globs @('C:\Program Files\draw.io\draw.io.exe')

# ---------------------------------------------------------------- status
function Get-DshConnectorStatus {
    <# One row per connector: resolved path + whether it exists. #>
    $map = [ordered]@{
        'Obsidian vault'   = $DSH_OBSIDIAN_VAULT
        'Obsidian.exe'     = $DSH_OBSIDIAN_EXE
        'Python (Office)'  = $DSH_PYTHON
        'MATLAB'           = $DSH_MATLAB
        'COMSOL batch'     = $DSH_COMSOL_BATCH
        'STM32CubeMX'      = $DSH_CUBEMX
        'STM32 Prog CLI'   = $DSH_PROG_CLI
        'Keil uVision'     = $DSH_KEIL
        'VS Code'          = $DSH_VSCODE
        'Notion'           = $DSH_NOTION
        'draw.io'          = $DSH_DRAWIO
    }
    foreach ($k in $map.Keys) {
        $p = $map[$k]
        [pscustomobject]@{
            Connector = $k
            Resolved  = ($p -ne $null)
            Path      = $p
            Exists    = ($p -ne $null -and (Test-Path -LiteralPath $p))
        }
    }
}
