"""Tests for ``--transport sse`` and ``--transport streamable-http``: where they
listen, and which requests they accept.

``FastMCP.run()`` takes no host or port (mcp 1.28 to 1.30): they are settings, and
passing them to ``run()`` raised a TypeError at start-up. The server object is built
for 127.0.0.1, which also turns on the Host header check for localhost only, so a
client on another machine (the reason for these transports here: an MCP client on
one machine reaching Access on a Windows machine across the LAN, physical or a VM)
was refused with 421. ``_configure_http`` sets the host and port, and keeps
the check on with the one address listened on allowed as well. Both transports
read the same settings.
"""

import json

import pytest
from mcp.server.fastmcp import FastMCP
from starlette.testclient import TestClient

from msaccess_vcs_mcp.main import _configure_http, _parse_transport_args

# The Windows machine's address: on a LAN, and on a desktop hypervisor's own network.
# Nothing in the configuration depends on which kind it is.
LAN = "192.168.1.50"
PARALLELS = "10.211.55.3"
ADDRESSES = [LAN, PARALLELS]


def _post(app, host, origin=None):
    """POST to the SSE message endpoint with the given Host (and Origin)."""
    headers = {"Host": host, "Content-Type": "application/json"}
    if origin:
        headers["Origin"] = origin
    return TestClient(app).post("/messages/", headers=headers, content=b"{}")


def _configured(host):
    mcp = FastMCP("test")
    _configure_http(mcp, host, 8765)
    return mcp


@pytest.mark.parametrize(
    "argv, transport",
    [
        ([], "stdio"),
        (["--host", LAN], "streamable-http"),
        (["--port", "8765"], "streamable-http"),
        (["--transport", "http", "--host", LAN], "streamable-http"),
        (["--transport", "sse", "--host", LAN], "sse"),
        (["--transport", "stdio", "--host", LAN], "stdio"),
    ],
)
def test_transport_defaults_to_streamable_http_for_network_options(argv, transport):
    """stdio stays the default for local clients; --host or --port means a network one."""
    assert _parse_transport_args(argv).transport == transport


def test_host_and_port_defaults():
    args = _parse_transport_args([])
    assert (args.host, args.port) == ("127.0.0.1", 8000)


def test_run_takes_no_host_or_port():
    """The call this replaces fails before anything listens."""
    with pytest.raises(TypeError):
        FastMCP("test").run(transport="sse", host=LAN, port=8765)


def test_settings_carry_host_and_port():
    mcp = _configured(LAN)
    assert mcp.settings.host == LAN
    assert mcp.settings.port == 8765


@pytest.mark.parametrize("address", ADDRESSES)
def test_unconfigured_server_refuses_another_machine(address):
    """Why the configuration is needed: the default check allows localhost only."""
    assert _post(FastMCP("test").sse_app(), f"{address}:8765").status_code == 421


@pytest.mark.parametrize("address", ADDRESSES)
@pytest.mark.parametrize("host_header", ["{address}:8765", "127.0.0.1:8765", "localhost:8765"])
def test_configured_server_accepts_its_own_address_and_localhost(address, host_header):
    """Its own address for a direct connection; localhost for an SSH tunnel."""
    # 400 "session_id is required" means the request passed the Host check
    response = _post(_configured(address).sse_app(), host_header.format(address=address))
    assert response.status_code == 400
    assert "session_id" in response.text


def test_configured_server_still_refuses_other_hosts():
    assert _post(_configured(LAN).sse_app(), "evil.example:8765").status_code == 421


def test_a_name_for_the_machine_is_refused():
    """The client must use the IP address it listens on: a name is another Host."""
    assert _post(_configured(LAN).sse_app(), "win11.lan:8765").status_code == 421


def test_configured_server_still_refuses_other_origins():
    response = _post(_configured(LAN).sse_app(), f"{LAN}:8765", origin="http://evil.example")
    assert response.status_code == 403


def test_ipv6_address_is_bracketed():
    app = _configured("fe80::1").sse_app()
    assert _post(app, "[fe80::1]:8765").status_code == 400
    assert _post(app, "evil.example:8765").status_code == 421


# Streamable HTTP: one endpoint, /mcp. The session manager starts in the app's
# lifespan, so each request runs inside ``with TestClient(app)``.

def _initialize(version="2025-06-18"):
    return json.dumps({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": version,
            "capabilities": {},
            "clientInfo": {"name": "test", "version": "0"},
        },
    })


def _post_mcp(mcp, host, origin=None, body=None, headers=None):
    """POST to the Streamable HTTP endpoint with the given Host (and Origin)."""
    sent = {
        "Host": host,
        "Content-Type": "application/json",
        "Accept": "application/json, text/event-stream",
    }
    if origin:
        sent["Origin"] = origin
    sent.update(headers or {})
    with TestClient(mcp.streamable_http_app()) as client:
        return client.post("/mcp", headers=sent, content=body or _initialize())


def _result(response):
    """The JSON-RPC message carried in the response's SSE stream."""
    data = [line[5:] for line in response.text.splitlines() if line.startswith("data:")]
    return json.loads(data[-1])["result"]


def test_streamable_http_path_is_mcp():
    assert _configured(LAN).settings.streamable_http_path == "/mcp"


@pytest.mark.parametrize("address", ADDRESSES)
def test_streamable_http_unconfigured_server_refuses_another_machine(address):
    assert _post_mcp(FastMCP("test"), f"{address}:8765").status_code == 421


@pytest.mark.parametrize("address", ADDRESSES)
@pytest.mark.parametrize("host_header", ["{address}:8765", "127.0.0.1:8765", "localhost:8765"])
def test_streamable_http_accepts_its_own_address_and_localhost(address, host_header):
    response = _post_mcp(_configured(address), host_header.format(address=address))
    assert response.status_code == 200
    assert response.headers["mcp-session-id"]


def test_streamable_http_still_refuses_other_hosts():
    assert _post_mcp(_configured(LAN), "evil.example:8765").status_code == 421


def test_streamable_http_refuses_a_name_for_the_machine():
    assert _post_mcp(_configured(LAN), "win11.lan:8765").status_code == 421


def test_streamable_http_still_refuses_other_origins():
    response = _post_mcp(_configured(LAN), f"{LAN}:8765", origin="http://evil.example")
    assert response.status_code == 403


def test_streamable_http_ipv6_address_is_bracketed():
    assert _post_mcp(_configured("fe80::1"), "[fe80::1]:8765").status_code == 200
    assert _post_mcp(_configured("fe80::1"), "evil.example:8765").status_code == 421


# No protocol version is pinned: the server answers with the version the client
# asks for when the SDK supports it. Streamable HTTP exists from 2025-03-26 on.
@pytest.mark.parametrize("version", ["2025-03-26", "2025-06-18", "2025-11-25"])
def test_streamable_http_negotiates_the_client_version(version):
    mcp = _configured(LAN)
    initialized = json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"})
    with TestClient(mcp.streamable_http_app()) as client:
        headers = {
            "Host": f"{LAN}:8765",
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
        }
        response = client.post("/mcp", headers=headers, content=_initialize(version))
        assert response.status_code == 200
        assert _result(response)["protocolVersion"] == version

        headers["mcp-session-id"] = response.headers["mcp-session-id"]
        headers["mcp-protocol-version"] = version
        assert client.post("/mcp", headers=headers, content=initialized).status_code == 202

        headers["mcp-protocol-version"] = "1999-01-01"
        assert client.post("/mcp", headers=headers, content=initialized).status_code == 400
