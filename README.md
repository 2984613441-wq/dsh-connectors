# dsh-connectors

[![self-test](https://github.com/2984613441-wq/dsh-connectors/actions/workflows/self-test.yml/badge.svg)](https://github.com/2984613441-wq/dsh-connectors/actions/workflows/self-test.yml)

**Let an AI agent actually operate the desktop software you already installed.**

Plain PowerShell functions + one Python helper, plus a zero-dependency MCP
server on top. No build step, no account required.

```powershell
. .\use.ps1
Get-DshObsidianNotes -Newest 5
Invoke-DshMatlab -Code "disp(version)"
Get-DshWordText -Path .\report.docx
```

---

## Two ways to use it

| | Direct (PowerShell) | MCP (any agent host) |
|---|---|---|
| Setup | dot-source `use.ps1` | register `mcp/server.mjs` once |
| Best for | you, at a keyboard | a model that needs the tools natively |
| Dependencies | none beyond the apps | none — the server is hand-written JSON-RPC |
| Survives host restart | trivial for a human | needs the host to respawn the server |

Both paths call the same PowerShell functions, so a bug fixed in one is fixed in
the other.

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
- Node.js 18+ (only for the MCP server)
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

## MCP server

`mcp/server.mjs` is a stdio MCP server with **no npm dependencies**. It speaks
newline-delimited JSON-RPC 2.0 directly, so there is nothing to install and
nothing to keep up to date. Tool arguments travel to PowerShell on the child's
stdin, never through `argv`, because note text contains quotes, newlines and
non-ASCII characters.

### Tools

| Tool | Purpose |
|---|---|
| `connector_status` | which applications are resolvable on this machine |
| `obsidian_list` | recent notes, newest first |
| `obsidian_read` | read one note by vault-relative path |
| `obsidian_search` | regex search, returns path + line + matching text |
| `obsidian_create` | create a note (refuses to clobber unless `overwrite`) |
| `obsidian_append` | append text to an existing note |
| `matlab_run` | headless MATLAB, returns exit code / seconds / output |
| `office_read` | docx / xlsx / pptx text, format chosen by extension |
| `office_write_docx` | create a Word document from a title and paragraphs |
| `pdf_convert` | pandoc → PDF, fails loudly instead of writing a partial file |

### Registering it

Add to your MCP client configuration (this is the shape DSH uses):

```yaml
- insert:
    - id: dsh-connectors-mcp
      name: "@deepseek-ai/dsh-mcp-client"
      config:
        serverName: connectors
        transport: stdio
        command: node
        args:
          - C:\path\to\dsh-connectors\mcp\server.mjs
```

The host must be **restarted** after the registration; MCP clients spawn their
servers at startup.

### Testing it without a host

```powershell
node mcp/smoke-test.mjs     # protocol-level: 7 assertions, no app required
```

The smoke test launches the real server over real stdio and asserts handshake,
catalog, schema shape, unknown-tool handling, bad-argument handling, and that
`connector_status` answers. It deliberately does **not** assert anything that
needs a locally installed application, so it is safe on CI.

Environment overrides: `DSH_CONNECTORS_POWERSHELL` (interpreter),
`DSH_CONNECTORS_TIMEOUT_MS` (per-call timeout, default 600000).

## Safety design

1. **Vault confinement.** `Resolve-DshVaultPath` resolves every path and rejects
   anything that escapes the vault root (`..` traversal). This is enforced below
   the MCP layer, so a tool call cannot bypass it.
2. **No silent overwrite.** Creating a note that already exists throws unless
   `-Force` is passed (`overwrite: true` over MCP).
3. **No credential reads.** No function reads `data.json`, `.credentials.yaml`,
   token files or cookies.
4. **No residue.** Every self-test fixture is created in `%TEMP%` (or a
   dot-prefixed folder inside the vault, which Obsidian ignores) and deleted
   afterwards.
5. **Process-scoped execution policy only.** `use.ps1` and the MCP bridge set
   the policy for the current process, never for `CurrentUser` or `LocalMachine`.
6. **stdout is protocol-only.** The MCP server writes nothing but JSON-RPC to
   stdout; diagnostics go to stderr. A single stray line would break the host.

## Five traps this repo documents and works around

These cost real debugging time; they are the reason this repo looks the way it
does.

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

### 5. A file extension is not cosmetic, and neither is a pipe's encoding

Two bugs of the same family, both found only by testing with real data:

- `openpyxl` refuses to open a workbook whose name does not end in
  `.xlsx`/`.xlsm`/`.xltx`/`.xltm` — it raises `InvalidFileException` even when
  the bytes are valid. The self test originally built fixtures with
  `GetRandomFileName()`, which embeds its own dot, producing
  `something.tmp-xlsx`; the write passed and the read failed.
- Python encodes **stdout of a redirected pipe** with the ANSI code page (GBK on
  a Chinese Windows), not UTF-8, unless told otherwise — Python only defaults to
  UTF-8 for pipes from 3.15. PowerShell then decodes those bytes as UTF-8 and
  every CJK character becomes `???`. The fix is to set `PYTHONIOENCODING=utf-8`
  and `PYTHONUTF8=1` for the child process, which `office.ps1` now does. This
  one survived an entire test suite because every assertion used ASCII.

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

- **Windows runner** — installs the Office libraries, runs `self-test.ps1`, runs
  the MCP protocol smoke test, and always prints the detection table.
  Applications a runner does not have (Obsidian, MATLAB, COMSOL, LaTeX) are
  reported as `SKIP`, not as failures.
- **Linux runner** — byte-compiles `office_helper.py`, checks that its CLI
  parses, and syntax-checks the MCP server.

The self test also honours `DSH_SELFTEST_STRICT=1`, which turns every `SKIP`
into a failure. Use it on a machine that is supposed to have everything
installed.

## Project layout

| File | Purpose |
|---|---|
| `_env.ps1` | path resolution for every tool + `Get-DshConnectorStatus` |
| `obsidian.ps1` | vault read/write/append/search/launch, with traversal guard |
| `matlab.ps1` | headless MATLAB execution |
| `office.ps1` | Word/Excel/PowerPoint + pandoc-to-PDF |
| `office_helper.py` | the Python backend for the Office functions |
| `use.ps1` | single entry point (process-scoped execution policy) |
| `self-test.ps1` | 21 checks on a fully equipped machine; exit code = failures |
| `mcp/server.mjs` | zero-dependency stdio MCP server |
| `mcp/bridge.ps1` | per-call PowerShell bridge (args on stdin, JSON on stdout) |
| `mcp/smoke-test.mjs` | protocol-level assertions, safe on CI |
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
- The MCP server spawns one PowerShell process per tool call. That keeps state
  out of the picture but costs roughly 200–400 ms per call, plus whatever the
  tool itself takes.
- The GitHub integration is literally the `gh` CLI. Install it and run
  `gh auth login`.
- Notion, Feishu/Lark and Tencent Meeting are **not** implemented. They need an
  API token or a self-built app, which is a deliberate user decision, not a bug.

## Contributing

Bug reports with the exact command, the exit code, and the output are far more
useful than feature requests. If you add a connector, add a `self-test.ps1`
check that fails loudly when it is broken, and mark it `SKIP` — not `PASS` —
when the dependency is absent. Test with non-ASCII data at least once: trap 5
above is what happens when nobody does.

## License

MIT. See [LICENSE](LICENSE).
