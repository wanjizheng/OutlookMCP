"""Add only our own MCP table; preserve other TOML settings byte-for-byte."""

import argparse
import datetime
import json
import re
import shutil
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def register(path):
    path = Path(path)
    raw = path.read_bytes() if path.exists() else b""
    text = raw.decode("utf-8-sig")
    tomllib.loads(text)
    block = "\n[mcp_servers.outlook_local]\n" + "".join(
        f"{k} = {json.dumps(v)}\n"
        for k, v in {
            "command": str(ROOT / ".venv/Scripts/python.exe"),
            "args": ["-m", "outlook_mcp.server"],
            "cwd": str(ROOT),
            "startup_timeout_sec": 30,
            "tool_timeout_sec": 120,
            "enabled": True,
        }.items()
    )
    # TOML bools must be lowercase.
    block = block.replace("enabled = true", "enabled = true")
    if re.search(r"^\[mcp_servers\.outlook_local(?:\]|\.)", text, re.M):
        raise RuntimeError(
            "outlook_local already exists; inspect it instead of overwriting"
        )
    candidate = raw + b"\n" + block.encode()
    parsed = tomllib.loads(candidate.decode("utf-8-sig"))
    if parsed["mcp_servers"]["outlook_local"]["cwd"] != str(ROOT):
        raise RuntimeError("Registration validation failed")
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        backup = (
            ROOT
            / "artifacts"
            / (
                "client-config-"
                + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
                + ".toml"
            )
        )
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, backup)
    path.write_bytes(candidate)
    return {
        "registered": True,
        "config": str(path),
        "server": "outlook_local",
        "client_reload_required": True,
    }


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--config", type=Path, default=Path.home() / ".codex/config.toml")
    a = p.parse_args()
    print(json.dumps(register(a.config)))
