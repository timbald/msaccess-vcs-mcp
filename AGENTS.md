# AGENTS.md - AI Agent Guide to msaccess-vcs-mcp

## Purpose

This repository contains the **msaccess-vcs-mcp** server — a lightweight MCP (Model Context Protocol) bridge that lets AI agents drive the [MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-addin) for version-controlling Microsoft Access databases.

## Development Environment Setup

**Always use the project virtual environment.** The package is installed in editable mode inside `venv/`. Do not install globally or suggest `pip install` outside the venv.

```powershell
# Activate the virtual environment (Windows PowerShell)
cd C:\path\to\msaccess-vcs-mcp
.\venv\Scripts\Activate.ps1

# If the venv doesn't exist yet, create it first:
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

### Running Tests

Tests must run inside the activated virtual environment:

```powershell
.\venv\Scripts\Activate.ps1

# Run all unit tests
pytest

# Run a specific test file
pytest tests/test_usage_logging.py -v

# Skip integration tests (require Access installed)
pytest -m "not integration"

# Run with coverage
pytest --cov=msaccess_vcs_mcp --cov-report=html
```

## Repository Structure

```
msaccess-vcs-mcp/
├── src/msaccess_vcs_mcp/       # Package source
│   ├── __init__.py             # Version (__version__)
│   ├── main.py                 # CLI entry point, startup sequence
│   ├── tools.py                # FastMCP instance + all @vcs_tool handlers
│   ├── config.py               # .env loading, get_config()
│   ├── usage_logging.py        # Structured JSONL usage logging
│   ├── security.py             # Path validation, write guards
│   ├── validation.py           # Startup validation helpers
│   ├── addin_integration.py    # COM calls to VCS add-in
│   ├── operation_manager.py    # Async operation queues
│   ├── callback_server.py      # HTTP callback server for VBA progress
│   └── access_com/             # Low-level COM/DAO helpers
├── tests/                      # pytest test suite
├── docs/                       # Extended documentation
├── .env.example                # Template for environment variables
├── pyproject.toml              # Package metadata & dependencies
└── AGENTS.md                   # This file
```

## Architecture

All tools are registered with the `@vcs_tool("name")` decorator in `tools.py`, which composes three concerns in order:

1. **Config reload** — `load_config()` re-reads `.env` when it changes
2. **Usage logging** — `with_logging(name)` records tool calls to JSONL
3. **MCP registration** — `mcp.tool()` exposes the function to MCP clients

```
AI Agent ──► MCP Server (Python) ──► VCS Add-in (VBA) ──► Access Database
                 │
           Path validation
           Permission checks
           Usage logging
           Async progress tracking
```

### Access window visibility

Any Access instance holding a database the server works with is left **visible**, whether the server created it or attached to one already running. A hidden instance strands the user: an error dialog, a VBA break, or a trust prompt blocks every later call with nothing on screen to explain why, and nobody can dismiss what they cannot see. `COM automation` normally starts Access hidden, so this is deliberate, not incidental.

Two rules follow from that, both enforced in `access_com/connection.py`:

- Open databases through `open_current_database(app, path)`, never a bare `app.OpenCurrentDatabase(...)`. It lowers `Application.UserControl` across the open — `OpenCurrentDatabase` runs the target's AutoExec, and the add-in's own `AutoRun` opens its installer form when that flag says a person is watching, stranding the instance the server is about to drive — and it shows the window afterwards.
- Show the window *after* the database opens, via `ensure_access_visible(app)`. Making a window visible can itself set `UserControl`, which is why the order is not interchangeable.

The one deliberate exception is `validate_access_installation()` in `config.py`: it opens no database and quits immediately, so a window would only flash on screen with nothing to act on.

## Configuration

All settings come from environment variables (loaded from `.env` / `.env.local` in the project root). See `.env.example` for the full list.

Key variables:
- `ACCESS_VCS_DATABASE` — target database path
- `ACCESS_VCS_DISABLE_WRITES` — set `true` to block write operations
- `ACCESS_VCS_ENABLE_LOGGING` — set `true` to enable usage logging
- `ACCESS_VCS_RUN_VBA_TIMEOUT_SEC` — parent-side timeout for `vcs_run_vba` worker processes (default 45s)
- `ACCESS_VCS_CALL_VBA_TIMEOUT_SEC` — parent-side timeout for `vcs_call_vba` `Application.Run` calls (default 45s)
- `ACCESS_VCS_REBUILD_TIMEOUT_SEC` — how long `vcs_rebuild_addin` waits after launch for a terminal status (default 1200s)
- `ACCESS_VCS_BUSY_WAIT_SEC` — how long a second tool call waits for the Access gate before returning `server_busy` (default 15s)
- `ACCESS_VCS_RECOVERY_PROBE_TIMEOUT_SEC` — timeout for automatic Access/add-in recovery probes after a VBA timeout or COM disconnect (default 10s)

One MCP server process is shared across all Cursor windows. Sync tools run in a single COM apartment thread with one Access operation at a time. A long call in one window causes others to get `error_pattern: server_busy` with `busy_with` naming the in-flight tool — retry rather than waiting for a client `-32001` timeout. After any client timeout, call `vcs_get_recent_calls()` to see what actually executed.

`vcs_run_vba` executes Access COM work in a short-lived child Python process. If a snippet hangs because Access is in break mode, blocked on a modal dialog, or otherwise unresponsive, the MCP server kills only that child process and returns a recoverable timeout. It does **not** kill `MSACCESS.EXE` or close user-owned Access windows; after Access becomes responsive, the next call runs an automatic probe and resumes normal operation.

### Rebuilding the VCS add-in

To rebuild `Version Control.accda` from source after editing add-in files, do **not** use `vcs_rebuild_database` (that rebuilds a user project). Call:

```python
vcs_rebuild_addin(r"C:\path\to\msaccess-vcs-addin\Version Control.accda.src")
```

The tool derives the development copy beside that folder, launches
`RebuildAddIn`, and passes its HTTP callback identity through the disconnected
worker to the builder Access process. Detailed `Log.Add` / `Log.Progress`
messages stream during the build; the status file covers compile, install, and
durable terminal recovery. Do **not** poll it yourself in the normal workflow.
`vcs_call_vba(..., ["RebuildAddIn", source])` remains launch-only.

MCP progress notifications are best-effort in Cursor 3.13 (often only "Running..." until the tool returns). For guaranteed live output, run `msaccess-vcs rebuild-addin <source>` from a terminal and keep that command in the foreground so the stream stays in the primary chat. The CLI exits when the operation reaches terminal status; that process exit is the completion signal. Do not background the CLI just to wait on a notification, and do not add a second timer wait, sleep, or `rebuild-status.json` poll after it has already finished.

`refused` and `launch-failed` are returned immediately and mean nothing was rebuilt: `refused` when another `MSACCESS.EXE` holds a file the rebuild must replace or cannot be asked which files it holds (`otherInstances` names what to close; the add-in never closes another process), and `launch-failed` when the helper script never started, which leaves Access open and is safe to retry.

If a client times out (`-32001`) or the tool returns `rebuild_stalled` / `timeout`, recover by reading `<source>/logs/rebuild-status.json` and matching `phaseStarted` against `rebuild_phase_started`. A `complete` whose `phaseStarted` predates the call is an earlier run's record. After any client timeout, call `vcs_get_recent_calls()` before inferring from the status file. A live rebuild always has `MSACCESS.EXE` or `wscript.exe`; neither, with a non-terminal status, means the run died.

### Running the add-in's own tests

Pass the **development copy** in the add-in's repository as `database_path`:

```python
vcs_run_tests(r"C:\path\to\msaccess-vcs-addin\Version Control.accda", "clsTestInstall")
```

A run needs two projects and they are different files: the installed add-in loads as a library and supplies the runner and `TestAssert`, while the code under test is whatever the current database holds. The runner scans the current VBA project, so the host decides which tests are found — aim a run at a user database, or anything in the repo's `Testing` folder, and you get that database's tests reported as a clean pass.

The installed add-in is refused as a host (see below); it also has no source tree beside it for the tests that read one. The add-in refuses such a run itself in `ExecuteTests` via `modInstall.CurrentDbIsInstalledAddIn`, so the server's refusal is the earlier of two.

`AccessConnection` opens the development copy itself (Access refuses to bind a file moniker to an `.accda`, so `GetObject` fails and the explicit `OpenCurrentDatabase` fallback in `_open_as_current_database` takes over), so no manual pre-open is needed.

### The installed add-in is never a target

No tool accepts the installed add-in as `database_path`, `output_path`, or `template_path`. That file exists to be loaded as a library: opening it as a database, or writing into it, resets a VBA project while it is executing. `_refuse_installed_addin_target` runs inside the `vcs_tool` wrapper, ahead of the gate and of any COM work, so the rule holds for every tool rather than the handful that grew their own guards — export, import, rebuild, `run_vba`, `run_tests`, `call_vba`, and the rest. Refusals carry `error_pattern: installed_addin_refused`.

`vcs_get_version_info()` reports the installed add-in's version without opening it, which is the one thing you might legitimately want from that file.

The comparison ignores the file extension, mirroring the add-in's `modInstall.PathsMatchIgnoringExtension`: a compiled install is a `.accde` built from the same `.accda`, and only one of the two is ever named in `ACCESS_VCS_ADDIN_PATH`. Folder parameters (`source_dir`, `output_dir`) are not checked — an export folder beside the install is not the install.

**Resolving the install path.** With `ACCESS_VCS_ADDIN_PATH` unset, `get_default_addin_path()` reads `HKCU\Software\VB and VBA Program Settings\MSAccessVCS\Install`, which is the only place to read it from — the add-in's own `GetInstalledAddInFileName` is built from exactly these two values. `Install Folder` is present only for a folder the user chose (the installer deletes it when the folder is the default, so absence means `%AppData%\MSAccessVCS`, not "not installed"), and `Compile accde` decides the extension. Do not reconstruct the path from `%AppData%` alone or assume `.accda`. The result is cached for the life of the process; `reset_addin_path_cache()` clears it.

Run these tests **through this server**, never from the add-in's own window. `modTestAssert.TestAssert` routes through `Application.Run` to the *installed* add-in path, while the runner singleton that records assertions lives in whichever project received `RunTests`. Invoking them from a development copy puts those in different projects: assertions are discarded and every test reports `EMPTY`. Treat an all-`EMPTY` result as a broken harness, not a pass.

## Logging

The server writes two parallel JSON Lines streams. Both filenames use the `vcs-mcp-` prefix so they don't collide with other tools that share the same logs directory.

### VCS operation logs (written by the add-in, not the server)

Separate from the two streams below, the add-in writes a per-operation log to `{source_dir}/logs/<Base>_<yyyymmdd_hhnnss_fff>.log`, where `<Base>` is `Export`, `Merge`, `Build`, `TestRun`, or `Other`. Note the base name tracks the *operation*, not the tool: `vcs_import_objects` and `vcs_import_object` both produce `Merge_*.log`.

The add-in also writes a `.gitignore` into the source folder containing `logs/` and `*.log`, which means **Cursor's Glob and Grep silently skip these files** — a search returns no matches rather than an error, so an agent can burn several calls before falling back to a shell listing. To avoid that:

- Every operation tool returns `log_path` for the run it just performed. Use it directly with the Read tool.
- On failure, those tools also return `log_excerpt` (tail of the log), so the error is usually available without any follow-up call.
- If the path is no longer at hand, call `vcs_get_log(database_path, log_type=...)` with the base name matching the operation.

The add-in's sync API returns this as camelCase `logPath`; `_addin_json_result` in `tools.py` normalizes it to `log_path` at the boundary and keeps the original key as an alias. Async completion callbacks already use `log_path`. Keep the normalizer even if the add-in changes: a newer server may run against an older add-in build.

### Diagnostic stream (`vcs-mcp-diagnostic.jsonl`) — always on

Captures server lifecycle events: `server_start`, `startup_env_load`, `lazy_env_load`, `lazy_init_started`, `lazy_init_skipped`, `list_roots_failed`, `list_roots_response`, `lazy_init_loaded`, `lazy_init_no_env_in_roots`, `usage_log_status`. Independent of `ACCESS_VCS_ENABLE_LOGGING` so it answers the "why didn't logging work?" question even when usage logging is silent.

- **Location:** `~/.msaccess-vcs-mcp/logs/vcs-mcp-diagnostic.jsonl`
- **Override:** `ACCESS_VCS_DIAGNOSTIC_LOG_DIR`
- **Opt out:** `ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true`
- **Rotation:** 1 MB per file, 3 backups
- **Discoverable from agents:** `vcs_get_version_info()` returns the active `diagnostic_log_path`.

### Usage stream (`vcs-mcp-usage.jsonl`) — default on

When `ACCESS_VCS_ENABLE_LOGGING=true` (the default), every tool call writes a structured entry. Set the env var to `false` to opt out.

- **Development installs:** logs to `{project_root}/logs/vcs-mcp-usage.jsonl`
- **Package installs:** logs to `~/.msaccess-vcs-mcp/logs/vcs-mcp-usage.jsonl`
- **Override:** `ACCESS_VCS_LOG_DIR`
- **Rotation:** `ACCESS_VCS_LOG_MAX_SIZE_MB` (default 10 MB), `ACCESS_VCS_LOG_BACKUP_COUNT` (default 5)

Each `tool_call` entry includes: `timestamp`, `version`, `event`, `tool`, `parameters`, `success`, `error`, `error_pattern`, `execution_time_ms`.

### Tiered audit posture

Three sensitivity tiers in the usage stream, each independently controlled:

1. **Audit metadata** — always written when `ENABLE_LOGGING=true`. Tool name, timing, success/error, sanitized parameter dict.
2. **Code-execution bodies** — `vcs_execute_sql`, `vcs_call_vba`, and `vcs_run_vba` write a `"code_execution"` event *before* execution begins. By default only `code_length` is recorded. Set `ACCESS_VCS_LOG_CODE_CONTENT=true` to record the full `code` field for forensic replay (off by default to limit business-data exposure). `code_length` lets analysts spot anomalies (e.g. "an agent ran a 4 KB VBA block") without seeing the body.
3. **Credential-shaped parameter keys** — any parameter whose name matches `password`, `secret`, `token`, `api_key`, `apikey`, `connection_string`, or `connectionstring` (case-insensitive) is replaced with `"<redacted>"` regardless of other switches. Defense in depth.

## Adding a New Tool

1. Write the handler function in `tools.py`
2. Decorate with `@vcs_tool("vcs_your_tool_name")` — this handles config reload, usage logging, and MCP registration automatically
3. Add tests in `tests/`

## Key Conventions

- All tool names use the `vcs_` prefix
- Tool handlers return `dict[str, Any]` with at least a `success` key
- Error results include `"error"` key (detected by usage logging)
- Async tools (`async def`) are supported by `@vcs_tool` transparently
- `Context` parameters from FastMCP are filtered out of usage logs automatically
