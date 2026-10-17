"""Native Windows tunnel setup; credentials encrypted by Windows current-user DPAPI."""

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
PRIVATE = ROOT / "artifacts/tunnel-client"
BIN = PRIVATE / "tunnel-client.exe"
KEY = PRIVATE / "runtime-key.dpapi"


def store_key():
    import tkinter as tk
    from tkinter import messagebox
    import win32crypt

    root = tk.Tk()
    root.title("OutlookMCP — Store tunnel runtime key locally")
    root.geometry("660x210")
    root.attributes("-topmost", True)
    tk.Label(root, text="Paste the dedicated runtime API key here.\nIt is encrypted for this Windows user; never send it in chat.").pack(pady=12)
    entry = tk.Entry(root, show="•", width=75)
    entry.pack(padx=12)

    def save():
        secret = entry.get().strip()
        if not secret.startswith("sk-") or len(secret) < 20:
            messagebox.showerror("Invalid key", "Enter a runtime API key, not an admin key or tunnel ID.")
            return
        if secret.startswith("sk-admin"):
            messagebox.showerror("Admin key rejected", "Use a restricted runtime key with Tunnels Read + Use.")
            return
        PRIVATE.mkdir(parents=True, exist_ok=True)
        try:
            KEY.write_bytes(win32crypt.CryptProtectData(secret.encode(), "OutlookMCP tunnel runtime", None, None, None, 0))
        except Exception:
            messagebox.showerror("Save failed", "Could not encrypt/save locally. No secret is printed. Check the private artifacts directory permissions.")
            return
        entry.delete(0, tk.END)
        root.destroy()

    tk.Button(root, text="Encrypt and save on this PC", command=save).pack(pady=12)
    root.mainloop()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("action", choices=["store-key", "init", "doctor", "run"])
    parser.add_argument("--tunnel-id")
    parser.add_argument("--enable-send", action="store_true")
    args = parser.parse_args()
    if os.name != "nt":
        raise SystemExit("Native interactive Windows is required")
    if args.action == "store-key":
        store_key()
        print("Runtime credential present:", KEY.exists())
        return
    env = dict(os.environ)
    env["TUNNEL_CLIENT_PROFILE_DIR"] = str(PRIVATE / "profiles")
    env["XDG_CONFIG_HOME"] = str(PRIVATE / "config")
    env["XDG_STATE_HOME"] = str(PRIVATE / "state")
    env["OUTLOOK_MCP_ENABLE_SEND"] = "1" if args.enable_send else "0"
    if args.action == "init":
        if not args.tunnel_id or not args.tunnel_id.startswith("tunnel_"):
            parser.error("init requires --tunnel-id from Platform")
        command = f'"{(ROOT / ".venv/Scripts/python.exe").as_posix()}" -m outlook_mcp.tunnel_server'
        argv = [str(BIN), "init", "--profile", "outlook", "--profile-dir", env["TUNNEL_CLIENT_PROFILE_DIR"],
                "--tunnel-id", args.tunnel_id, "--mcp-command", command, "--health-listen-addr", "127.0.0.1:0"]
    else:
        import win32crypt
        if not KEY.exists():
            raise SystemExit("Store the restricted runtime key locally first: store-key")
        env["CONTROL_PLANE_API_KEY"] = win32crypt.CryptUnprotectData(KEY.read_bytes(), None, None, None, 0)[1].decode()
        argv = [str(BIN), args.action, "--profile", "outlook"]
        if args.action == "doctor":
            argv += ["--explain"]
        else:
            argv += ["--health.url-file", str(PRIVATE / "health.url")]
    # Foreground daemon: keep this terminal running for discovery and calls.
    # No secrets in command line, YAML, stdout, or chat.
    raise SystemExit(subprocess.call(argv, cwd=ROOT, env=env))


if __name__ == "__main__":
    main()
