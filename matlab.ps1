# =====================================================================
# DSH Connectors - MATLAB (headless -batch)
#   . .\matlab.ps1
#   Invoke-DshMatlab -Code "disp(version)"
#   Invoke-DshMatlab -File .\script.m
# =====================================================================

if (-not $DSH_MATLAB) { . (Join-Path $PSScriptRoot '_env.ps1') }

function Invoke-DshMatlab {
    <#
      Runs MATLAB headless and returns ExitCode / Ok / Seconds / Output.
      -Code  : inline MATLAB statement
      -File  : path to a .m script (executed from its own folder)
    #>
    param(
        [string]$Code,
        [string]$File,
        [int]$TimeoutSec = 300
    )
    if (-not $DSH_MATLAB) { throw "MATLAB not resolved. Set matlab in connectors.config.json (or put matlab.exe on PATH)." }
    if (-not (Test-Path -LiteralPath $DSH_MATLAB)) { throw "matlab.exe not found: $DSH_MATLAB" }
    if (-not $Code -and -not $File) { throw "Provide -Code or -File" }

    $argList = @()
    if ($File) {
        $resolved = (Resolve-Path -LiteralPath $File -ErrorAction Stop).Path
        $dir  = Split-Path -Parent $resolved
        $name = [System.IO.Path]::GetFileNameWithoutExtension($resolved)
        $argList += @('-sd', $dir, '-batch', "try, $name, catch e, disp(getReport(e)); exit(1); end; exit(0);")
    } else {
        $argList += @('-batch', $Code)
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $raw  = & $DSH_MATLAB @argList 2>&1
    $code = $LASTEXITCODE
    $sw.Stop()

    [pscustomobject]@{
        ExitCode = $code
        Ok       = ($code -eq 0)
        Seconds  = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        Output   = ($raw | Out-String).Trim()
    }
}
