# dsh-connectors

**Let an AI agent actually operate the desktop software you already installed.**

Plain PowerShell functions + one Python helper. No MCP server, no build step, no
account required. Works with any agent that can run a shell command.

```powershell
. .\use.ps1
Get-DshObsidianNotes -Newest 5
Invoke-DshMatlab -Code "disp(version)"
Get-DshWordText -Path .\report.docx
```

---

## Why not MCP?

MCP is the right long-term answer, but it costs a server process per
application, a config edit, and a restart of the host. For a single operator on a
single machine, a dot-sourceable module gets the same capability in one line and
survives host restarts. If you later want MCP, wrap these functions in a thin
server — the hard part (path resolution, safety guards, verified call shapes) is
already done here.

## What it connects

| Applications | Mechanism | Verified |
|---|---|---|
| **Obsidian** | direct filesystem read/write of the vault | yes |
| **MATLAB** | `matlab -batch` headless execution | yes |
| **Word / Excel / PowerPoint** | python-docx, openpyxl, python-pptx | Word yes, Excel/PPT reader written but untested |
| **PDF export** | pandoc + a LaTeX engine | tools detected, conversion untested |
| **GitHub** | the `gh` CLI | yes (when logged in) |
| **COMSOL** | `comsolbatch.exe` | command responds; no real model run |
| **STM32** | `STM32_Programmer_CLI.exe` | version probe only |
| **Keil uVision** | `UV4.exe` | file version read; no real build |
| **VS Code** | `code` CLI | yes |
| Notion / Feishu / Lark | requires an API token or app credentials | **not implemented** |
| WeChat / QQ / Douyin / Bilibili | no official API | **out of scope** |

"Verified" means the repo author ran it on a real machine and recorded the exit
code and output. Anything unverified is labelled as such on purpose.

## Requirements

- Windows 10/11, Windows PowerShell 5.1 or PowerShell 7
- Python **with** `python-docx`, `openpyxl`, `python-pptx` (only for the Office
  functions)
- Whatever third-party application you intend to drive

## Quick start

```powershell
git clone https://github.com/<you>/dsh-connectors.git
cd dsh-connectors

# See what was detected on this machine before doing anything else:
powershell -NoProfile -ExecutionPolicy Bypass -File .\self-test.ps1

# Load the module:
powershell -NoProfile -ExecutionPolicy Bypass -Command ". .\use.ps1; Get-DshConnectorStatus"
```

Optional configuration — copy the example and fill in only what auto-detection
missed:

```powershell
Copy-Item .\connectors.config.example.json .\connectors.config.json
```

`connectors.config.json` is git-ignored on purpose: it is machine-specific and
will contain private paths.

## How paths are resolved

Each tool is resolved in this order, and a value is only accepted if it
**actually exists**:

1. `connectors.config.json`
2. `PATH`
3. Windows registry `Uninstall` keys — product name match, then `DisplayIcon`
   (rejecting installer names such as `setup.exe`), then `InstallLocation`
4. well-known install folders (glob)

Nothing is guessed. An unresolved tool returns `$null` and is reported as
`SKIP`, never as a fake path.

The Obsidian vault is read from `%APPDATA%\obsidian\obsidian.json` (the currently
open vault), falling back to the most recently used one.

## Safety design

1. **Vault confinement.** `Resolve-DshVaultPath` resolves every path and rejects
   anything that escapes the vault root (`..` traversal).
2. **No silent overwrite.** Creating a note that already exists throws unless
   `-Force` is passed.
3. **No credential reads.** No function reads `data.json`, `.credentials.yaml`,
   token files or cookies.
4. **No residue.** The self test creates its fixture under a dot-prefixed folder
   (ignored by Obsidian), verifies it, and deletes it.
5. **Process-scoped execution policy only.** `use.ps1` sets the policy for the
   current process, never for `CurrentUser` or `LocalMachine`.

## Four Windows traps this repo documents and works around

These cost real debugging time; they are the reason this repo exists in its
current shape.

### 1. The `python` on PATH may be a bundled interpreter

Game engines and SDKs ship their own Python and often win the `PATH` race. On
the development machine the first `python` was a **3.7.7** build from a game
runtime with none of the Office libraries installed. `_env.ps1` therefore probes
candidate interpreters by **actually importing** `docx, openpyxl, pptx` and
returns the first one that succeeds, resolving `py.exe` to the interpreter it
selects.

### 2. A GUI app's plugins can silently require a newer app version

Read `plugins\<id>\manifest.json` and compare `minAppVersion` against the
installed app version. On the development machine four Obsidian plugins —
including the backup plugin — required a newer Obsidian than was installed and
were therefore dead, with no visible error.

```powershell
# quick audit for any Obsidian vault
$obsidian = (Get-Item "$env:LOCALAPPDATA\Obsidian\Obsidian.exe").VersionInfo.ProductVersion
Get-ChildItem "$vault\.obsidian\plugins" -Directory | ForEach-Object {
    $m = Get-Content "$($_.FullName)\manifest.json" -Raw | ConvertFrom-Json
    [pscustomobject]@{ Plugin = $m.id; Needs = $m.minAppVersion; Have = $obsidian }
} | Where-Object { [version]$_.Needs -gt [version]$_.Have }
```

### 3. It is Windows PowerShell 5.1, not PowerShell 7

No `-SkipCertificateCheck`, no three-argument `Join-Path`. More importantly:
a BOM-less UTF-8 `.ps1` containing non-ASCII text is misread and produces
mojibake. **All `.ps1` files here are pure ASCII on purpose** — non-ASCII text
belongs in `.md`.

### 4. An unsigned local script can still be refused

If a script's directory carries an unusual ACL, PowerShell can reject an
unsigned local script with *"is not digitally signed"* even though the effective
policy is `RemoteSigned`. Diagnose it with a three-location A/B test before
touching any policy:

```powershell
Copy-Item .\some.ps1 E:\probe.ps1
. E:\probe.ps1        # works? then the directory, not the policy, is the problem
```

Workaround used here: `Set-ExecutionPolicy -Scope Process Bypass`, which is not
persistent and does not change the machine.

## API

### Obsidian

```powershell
Get-DshObsidianNotes   [-Newest 20] [-Folder 'sub/dir']
Get-DshObsidianNote    -Path 'note.md'
Find-DshObsidianText   -Pattern 'regex' [-Max 30]
New-DshObsidianNote    -Path 'inbox/idea.md' -Content '# idea' [-Force]
Add-DshObsidianContent -Path 'inbox/idea.md' -Content '- more' [-PrependNewLine]
Start-DshObsidian
```

### MATLAB

```powershell
Invoke-DshMatlab -Code "disp(version)"
Invoke-DshMatlab -File .\analysis.m
# -> .ExitCode  .Ok  .Seconds  .Output
```

### Office

```powershell
Get-DshWordText          -Path .\doc.docx
New-DshWordDoc           -Path .\new.docx -Title 'T' -Line @('p1','p2')
Get-DshExcelPreview      -Path .\book.xlsx [-Rows 10]
Get-DshPowerPointOutline -Path .\deck.pptx
Get-DshPdfViaPandoc      -InputPath .\doc.md [-OutputPath .\doc.pdf]
```

### Diagnostics

```powershell
Get-DshConnectorStatus   # one row per connector: resolved / path / exists
```

## Project layout

| File | Purpose |
|---|---|
| `_env.ps1` | path resolution for every tool + `Get-DshConnectorStatus` |
| `obsidian.ps1` | vault read/write/append/search/launch, with traversal guard |
| `matlab.ps1` | headless MATLAB execution |
| `office.ps1` | Word/Excel/PowerPoint + pandoc-to-PDF |
| `office_helper.py` | the Python backend for the Office functions |
| `use.ps1` | single entry point (process-scoped execution policy) |
| `self-test.ps1` | 16 checks; exit code = number of failures |
| `connectors.config.example.json` | optional overrides |

## Limitations

- Windows only. Nothing here is conceptually Windows-specific except the
  registry lookup and the PowerShell idioms, but no other platform is tested.
- Excel and PowerPoint readers are written but were never exercised against a
  real workbook or deck.
- MATLAB and COMSOL need a valid license; failures surface as a non-zero exit
  code and the process output.
- The GitHub integration is literally the `gh` CLI. Install it and run
  `gh auth login`.
- Notion, Feishu/Lark and Tencent Meeting are **not** implemented. They need an
  API token or a self-built app, which is a deliberate user decision, not a bug.

## Contributing

Bug reports with the exact command, the exit code, and the output are far more
useful than feature requests. If you add a connector, add a `self-test.ps1`
check that fails loudly when it is broken, and mark it `SKIP` — not `PASS` —
when the dependency is absent.

## License

MIT. See [LICENSE](LICENSE).
