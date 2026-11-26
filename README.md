# dsh-connectors

[![self-test](https://github.com/2984613441-wq/dsh-connectors/actions/workflows/self-test.yml/badge.svg)](https://github.com/2984613441-wq/dsh-connectors/actions/workflows/self-test.yml)

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

| Applications | Mechanism | Self-test |
|---|---|---|
| **Obsidian** | direct filesystem read/write of the vault | passes (skipped if no vault) |
| **MATLAB** | `matlab -batch` headless execution | passes (skipped if absent) |
| **Word** | python-docx | create + read back |
| **Excel** | openpyxl | write + read back |
| **PowerPoint** | python-pptx | write 2 slides + read outline |
| **PDF export** | pandoc + xelatex | real conversion, asserts > 1 KB |
| **GitHub** | the `gh` CLI | version + auth |
| **COMSOL** | `comsolbatch.exe` | resolved and invocable |
| **STM32** | `STM32_Programmer_CLI.exe` | version probe |
| **Keil uVision** | `UV4.exe` | resolved |
| **VS Code** | `code` CLI | resolved |
| Notion / Feishu / Lark | requires an API token or app credentials | **not implemented** |
| WeChat / QQ / Douyin / Bilibili | no official API | **out of scope** |

Each row above is exercised by `self-test.ps1`. Anything that cannot be checked
on a given machine is reported as `SKIP`, never as a pass.

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
4. **No residue.** Every self-test fixture is created in `%TEMP%` (or a
   dot-prefixed folder inside the vault, which Obsidian ignores) and deleted
   afterwards.
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

### Bonus: the file extension is not cosmetic

`openpyxl` refuses to open a workbook whose file name does not end in
`.xlsx`/`.xlsm`/`.xltx`/`.xltm` — it raises `InvalidFileException` even when the
bytes are a perfectly valid workbook. The self test originally generated
fixtures with `GetRandomFileName()`, which embeds its own dot, producing
`something.tmp-xlsx`; the read then failed while the write had succeeded. Names
now come from a GUID plus the real extension.

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

Raw helper subcommands, if you prefer to call Python directly:

```
docx-read  <file> [--max-chars N]
docx-write <file> --title T [--line TEXT ...]
xlsx-read  <file> [--rows N]
xlsx-write <file> [--sheet NAME] --row "a,b,c" ...
pptx-read  <file>
pptx-write <file> --slide "Title::body 1;body 2" ...
```

### Diagnostics

```powershell
Get-DshConnectorStatus   # one row per connector: resolved / path / exists
```

## Continuous integration

`.github/workflows/self-test.yml` runs on every push and pull request:

- **Windows runner** — installs the Office libraries, then runs `self-test.ps1`,
  and always prints the detection table. Applications a runner does not have
  (Obsidian, MATLAB, COMSOL, LaTeX) are reported as `SKIP`, not as failures.
- **Linux runner** — byte-compiles `office_helper.py` and checks that its CLI
  parses.

The self test also honours `DSH_SELFTEST_STRICT=1`, which turns every `SKIP`
into a failure. Use it on a machine that is supposed to have everything
installed:

```powershell
$env:DSH_SELFTEST_STRICT = '1'
powershell -NoProfile -ExecutionPolicy Bypass -File .\self-test.ps1
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
| `self-test.ps1` | 20 checks on a fully equipped machine; exit code = failures |
| `connectors.config.example.json` | optional overrides |
| `.github/workflows/self-test.yml` | CI: Windows full run + Linux syntax check |

## Limitations

- Windows only. Nothing here is conceptually Windows-specific except the
  registry lookup and the PowerShell idioms, but no other platform is tested.
- MATLAB and COMSOL need a valid license; a failure surfaces as a non-zero exit
  code and the process output.
- PDF export needs a LaTeX engine. MiKTeX may want to download packages on first
  use, which the self test cannot do silently — if `pdf-convert` fails on a
  fresh machine, run `xelatex` once by hand.
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
