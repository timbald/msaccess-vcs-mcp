# Fable Observations — msaccess-vcs-mcp (timbald fork)

Review date: 2026-09-18, **revised the same day after the fork was brought up to upstream**. Read: branch and remote state before and after the update; the 16 upstream commits (2026-06-02 → 2026-09-10) and the `DECISIONS.md` entries they added; `README.md` registration section; `AGENTS.md`; `.env.example`; `main.py`, `cli.py`, `__main__.py`, `access_gate.py`, `callback_server.py`, `config.py`, `security.py`; the head of `tools.py` (FastMCP instance, `vcs_tool` decorator, workspace-root lazy init); Tim's SSE commit `9f0da84`; and the pinned MCP SDK (`mcp==1.28.0`, downloaded to the scratchpad) for `FastMCP.run`, `Settings` and the transport-security middleware. Nothing was executed; the server only runs on Windows.

The question this revision answers: **how does the MCP server, which must live on the Windows VM beside Access and the add-in, get driven by Claude Code on the Mac?**

---

## A. What was done to the fork today

| Ref | Before | After |
|---|---|---|
| `main` | `ad60fb5` (2026-05-21), 16 behind upstream, 0 ahead | fast-forwarded to upstream `7bfe191` (2026-09-10) and pushed to `origin/main` |
| `feat/sse-transport` | `9f0da84` on the old base; the open upstream PR | **unchanged** — force-pushing the PR branch is your call (§D) |
| `feat/sse-transport-rebased` | — | new branch: `main` + a clean cherry-pick of `9f0da84` (`3e6b3d0`), pushed to the fork |
| `fable-observations.md` | untracked | untracked, preserved |

Upstream gained: the Access gate (one COM apartment thread, one operation at a time, `server_busy` errors), a durable registry of server-created Access processes with safe closure, full-power-core QoS, `vcs_rebuild_addin` and `vcs_run_tests` streamed through `APIAsync`, a `msaccess-vcs` terminal CLI, pywin32 gen_py self-heal, and the rule that the installed add-in is never a target. **No network transport was added upstream; stdio is still the only mode.** Your branch is still the only route to cross-machine use.

---

## B. Running over the LAN — findings

### B.1 The SSE commit will not start

1. 🔴 `main.py` on the branch calls `mcp.run(transport="sse", host=args.host, port=args.port)`. In the pinned SDK (`mcp[cli]>=1.28,<2`) the signature is `FastMCP.run(transport, mount_path=None)`; there are no `host`/`port` parameters, so `--transport sse` raises `TypeError` before anything listens. The default stdio path is unaffected, which is why the PR looked harmless. Host and port are `FastMCP.settings` fields, set either at construction or by assignment before `run()`:

    ```python
    mcp.settings.host = args.host
    mcp.settings.port = args.port
    mcp.run(transport=args.transport)
    ```

2. 🟡 **Zero-code alternative for host and port.** `FastMCP.Settings` is a pydantic `BaseSettings` with `env_prefix="FASTMCP_"` and `env_file=".env"`. So `FASTMCP_HOST=0.0.0.0` and `FASTMCP_PORT=8000` in the environment (or in the `.env` the server reads at start-up) configure the bind with no code. Only the transport choice needs a change, and that can be one line reading an `ACCESS_VCS_TRANSPORT` variable instead of the argparse block, which keeps the console-script entry point argument-free.
3. 🟡 **Prefer `streamable-http` over `sse`.** The SDK supports both; SSE is the legacy transport and Claude Code registers the modern one with `claude mcp add --transport http <name> http://<host>:<port>/mcp`. The endpoint paths are `/mcp` for streamable-http and `/sse` for SSE.

### B.2 What stays inside the VM (correct as is)

1. The VBA → Python **callback server** (`callback_server.py`) binds `127.0.0.1` on an OS-assigned port and is posted to by the add-in running in the same Windows session. Over the LAN nothing changes: `ACCESS_VCS_CALLBACK_HOST` must stay `127.0.0.1`. Progress reaches the Mac as MCP progress notifications over the HTTP transport, which streamable-http carries.
2. The **Access gate**, the **owned-instance registry** (`~/.msaccess-vcs-mcp/owned-instances.json` on the VM), `tasklist`, `GetObject`/ROT lookups and the QoS calls are all Windows-local. None needs to know a client is remote.
3. Upstream's own design already assumes **one long-lived server process shared by several clients** ("every Cursor window"), which is exactly the shape a LAN service has. The `server_busy` answer is what a second Mac session would see while an export runs.

### B.3 Paths, `.env` discovery and the Mac's workspace roots

1. 🟠 **Every path argument is interpreted on the VM.** `database_path`, `output_dir`, `source_dir` must be Windows paths as the VM sees the Mac folder — `G:\GitHub\<project>\…` on the VM that maps `G:`, or `\\Mac\Home\GitHub\<project>\…` (ontap-db's export folder). Three spellings of the same Mac folder already exist across your `vcs-options.json` files; pick one for the LAN setup and write it in the project `CLAUDE.md`.
2. 🟠 **Workspace-root discovery cannot work across machines, and fails safely.** On the first tool call the server asks the client for `roots/list`; Claude Code on the Mac answers `file:///Users/timfinch/GitHub/<project>`. `_file_uri_to_path` turns that into `/Users/timfinch/GitHub/<project>`, which does not exist on Windows, so the `.env` there is never found and the diagnostic log records `lazy_init_no_env_in_roots`. Harmless, but it means the **project `.env` must be found on the VM side**: start the server with `ACCESS_VCS_PROJECT_DIR=G:\GitHub\<project>` (per project) or rely on explicit `database_path` arguments on every call (all Access-touching tools take one).
3. 🟡 With one shared server for all projects, `ACCESS_VCS_PROJECT_DIR` can point at only one; `.env` hot-reload follows that root; usage logs land in `<that root>\logs\`. Either run one server per active project on different ports, or run one server with no project root and always pass paths explicitly. The second is simpler and matches how the CLI already works.
4. 🟡 Mapped drives (`G:`) exist only in an interactive logon session. A server started by Task Scheduler *at logon* sees them; a Windows Service (session 0) does not, and neither does it have a desktop for Access to show windows on — upstream now insists Access windows stay visible so dialogs can be dismissed (`AGENTS.md` "Access window visibility"). **Run the server as a logon task in the VM's interactive session, not as a service.**

### B.4 Security of a network bind

1. 🟠 No authentication exists on the SSE/streamable-http endpoint, and the SDK's DNS-rebinding protection is **disabled by default when the bind host is not localhost** (`TransportSecurityMiddleware` defaults to off "for backwards compatibility"; FastMCP only enables it automatically for `127.0.0.1`/`localhost`). Bound to `0.0.0.0`, the server accepts any Host header from anyone who can reach the port. On Parallels **Shared Networking** the VM sits on a private NAT (`10.211.55.0/24`) reachable only from this Mac, which is acceptable. On **Bridged** networking the whole LAN could drive Access and run VBA (`vcs_run_vba` is gated by the add-in's `McpAllowRunVBA`, but `vcs_execute_sql` and imports are not).
2. Cheap hardening, no code: bind to the VM's shared-network address rather than `0.0.0.0` (`FASTMCP_HOST=10.211.55.x`), and add a Windows Firewall inbound rule for the port scoped to the Mac's address (`10.211.55.2` is the Mac as seen from the VM, per the ontap connection strings). Optional: enable the SDK's Host-header check through the nested env form `FASTMCP_TRANSPORT_SECURITY__ENABLE_DNS_REBINDING_PROTECTION=true` and `FASTMCP_TRANSPORT_SECURITY__ALLOWED_HOSTS=["10.211.55.x:8000"]`.

### B.5 The Mac side today is wrong

1. 🔴 `vba-3d-renderer/.mcp.json` registers `uvx msaccess-vcs-mcp@latest`. That launches the **PyPI** package **on the Mac**, where `validate_access_installation()` fails at `import win32com` and the server exits; and even on Windows it would be upstream's release, not your fork. Replace it with a URL entry once the VM service exists, e.g. `claude mcp add --transport http msaccess-vcs http://10.211.55.x:8000/mcp` (project scope), and remove the `enabledMcpjsonServers` entry from `.claude/settings.local.json` if it no longer applies.
2. 🟡 To run **your** branch on the VM: `pip install git+https://github.com/timbald/msaccess-vcs-mcp@feat/sse-transport-rebased` into a venv (or `uvx --from git+https://… msaccess-vcs-mcp`), so the transport flag exists. The README's `uvx msaccess-vcs-mcp@latest` will never include it until upstream merges.

### B.6 Suggested shape of the fix (for the PR, or a local branch)

1. In `main.py`, replace the argparse block with `transport = os.environ.get("ACCESS_VCS_TRANSPORT", "stdio")`, accept `stdio | sse | streamable-http`, and call `mcp.run(transport=transport)`; document `FASTMCP_HOST`/`FASTMCP_PORT` in `.env.example` as the bind settings. Keep the CLI flags only if the console script must stay self-describing.
2. `README.md`: a "Cross-machine use" section (Windows host runs the server as a logon task; client connects by URL; paths are the server's; `.env` lives beside the server or `ACCESS_VCS_PROJECT_DIR` is set; callback host stays loopback; network is trusted or firewalled).
3. `DECISIONS.md`: one entry — why HTTP transport, why no auth (trusted private network), why loopback callback is unchanged.
4. A unit test that `ACCESS_VCS_TRANSPORT=streamable-http` selects the transport without calling `run()` with unknown kwargs (the missing test is why the original defect shipped).

---

## C. Carried-over code notes (re-checked on `7bfe191`, all still present)

1. 🟡 `config.py` `_find_project_root`: the package-directory search root is appended twice (once under an `if`, once unconditionally). Harmless.
2. 🟡 `security.py`: system-directory check is a case-sensitive `str.startswith` on `C:\Windows` etc., so `c:\windows\…` passes. Use resolved, case-folded paths.
3. 🟡 Root-level `test_correct_syntax.py`, `test_filename_only.py`, `test_preload_first.py`, `test_read_version_property.py` are the maintainer's scratch scripts with hard-coded `C:\Users\Adam Waller\…` paths; `pytest` ignores them (`testpaths = ["tests"]`).
4. 🟡 `pyproject.toml` version is still `0.1.0` after ~36 feature commits; `__version__` is what `vcs_get_version_info` reports.
5. 🟡 **No project `CLAUDE.md`.** `AGENTS.md` assumes a Windows venv at the repo. A ten-line `CLAUDE.md` here should record the VM/Mac split, the VM's address, the port, how the service is started, and which projects register it.

---

## D. Questions

1. **Update the open upstream PR** by force-pushing `feat/sse-transport` to the rebased-and-fixed commit, or close it and open a fresh PR from a new branch once §B.6 is done? (Force-pushing a PR branch on your own fork is normal, but you asked that I never force-push, so I have not.)
2. Do you want me to implement §B.6 on `feat/sse-transport-rebased` (transport selection, settings-based host/port, README and DECISIONS entries, a test), or leave it as a description for you?
3. Parallels networking on the VM: Shared or Bridged? It decides whether §B.4's hardening is optional or required.
4. One server on the VM for all projects (explicit paths on every call, port 8000), or one per active project?
5. Shall I replace the `uvx` entry in `vba-3d-renderer/.mcp.json` with the URL form once the VM address is known?
