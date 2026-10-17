#!/bin/bash
# Shared helpers for outlook-* wrapper scripts.
# Source this at the top of every outlook wrapper:
#   SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
#   source "$SCRIPT_DIR/_common.sh" 2>/dev/null || source "$HOME/.openclaw/workspace/skills/outlook_reader/scripts/_common.sh"

PS_EXE="/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
SCHTASKS_EXE="/mnt/c/Windows/System32/schtasks.exe"
SKILL_SCRIPTS_DIR="$HOME/.openclaw/workspace/skills/outlook_reader/scripts"
WINDOWS_BRIDGE_DIR="/mnt/c/Users/wanjizheng/openclaw-bridge"
OUTLOOK_COM_LOCK_FILE="${XDG_RUNTIME_DIR:-/tmp}/openclaw-outlook-com-${UID}.lock"

# Fail before reading "$2" when an option value is missing.
require_option_value() {
    local option_name="$1"
    local remaining="$2"
    if [ "$remaining" -lt 2 ]; then
        printf '{"error":"Missing value for %s"}\n' "$option_name"
        return 2
    fi
}

# Reject contradictory action flags instead of silently letting the last one win.
select_outlook_action() {
    local requested="$1"
    if [ -n "${OUTLOOK_ACTION_EXPLICIT:-}" ] && [ "$OUTLOOK_ACTION_EXPLICIT" != "$requested" ]; then
        printf '{"error":"Conflicting action flags: %s and %s"}\n' "$OUTLOOK_ACTION_EXPLICIT" "$requested"
        return 2
    fi
    OUTLOOK_ACTION_EXPLICIT="$requested"
}

# --- Sync deployed bridge scripts if source has changed ---
sync_if_needed() {
    local src="$1"
    local dst="$2"

    if [ ! -f "$src" ]; then
        return
    fi

    mkdir -p "$(dirname "$dst")"
    if [ ! -f "$dst" ] || ! cmp -s "$src" "$dst"; then
        cp "$src" "$dst"
    fi
}

# Return success when WSL interop launches Windows PowerShell with an
# administrator token. Classic Outlook normally runs with a limited desktop
# token, and Outlook COM rejects automation across that integrity boundary.
windows_powershell_is_elevated() {
    local elevated
    elevated=$(
        "$PS_EXE" -NoProfile -NonInteractive -Command \
            '[Console]::Write(([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))' \
            2>/dev/null | tr -d '\r'
    )
    case "${elevated,,}" in
        true) return 0 ;;
        false) return 1 ;;
        *) return 2 ;;
    esac
}

# Run a PowerShell payload under the Windows user's LIMITED interactive token
# and return the payload's combined stdout/stderr plus its real exit code.
#
# schtasks is used only as a privilege-drop launcher. Every task and temporary
# file is uniquely named and removed on success, failure, timeout, or signal.
run_outlook_ps_limited() {
    local ps_command="$1"
    local timeout_sec="${OUTLOOK_RUN_TIMEOUT_SEC:-900}"
    local lock_wait_sec="${OUTLOOK_LOCK_WAIT_SEC:-900}"
    local unique_id="$$-${BASHPID}-${RANDOM}"
    local task_name="OpenClawOutlook-${unique_id}"
    local payload_wsl="$WINDOWS_BRIDGE_DIR/ocl-outlook-payload-${unique_id}.ps1"
    local driver_wsl="$WINDOWS_BRIDGE_DIR/ocl-outlook-driver-${unique_id}.ps1"
    local result_wsl="$WINDOWS_BRIDGE_DIR/ocl-outlook-result-${unique_id}.txt"
    local status_wsl="$WINDOWS_BRIDGE_DIR/ocl-outlook-status-${unique_id}.txt"
    local payload_win driver_win result_win status_win
    local task_created=0
    local elapsed=0
    local max_ticks=0
    local exit_code=1
    local lock_fd=""

    if ! [[ "$timeout_sec" =~ ^[0-9]+$ ]] || [ "$timeout_sec" -lt 1 ] || [ "$timeout_sec" -gt 3600 ]; then
        printf '{"error":"OUTLOOK_RUN_TIMEOUT_SEC must be an integer between 1 and 3600"}\n'
        return 2
    fi
    max_ticks=$((timeout_sec * 5))
    if ! [[ "$lock_wait_sec" =~ ^[0-9]+$ ]] || [ "$lock_wait_sec" -lt 1 ] || [ "$lock_wait_sec" -gt 3600 ]; then
        printf '{"error":"OUTLOOK_LOCK_WAIT_SEC must be an integer between 1 and 3600"}\n'
        return 2
    fi
    if [ ! -x "$SCHTASKS_EXE" ]; then
        printf '{"error":"schtasks.exe not found; cannot drop the elevated Windows token"}\n'
        return 1
    fi
    if ! command -v wslpath >/dev/null 2>&1; then
        printf '{"error":"wslpath command not found"}\n'
        return 1
    fi

    mkdir -p "$WINDOWS_BRIDGE_DIR"

    # Serialize Outlook COM operations. Concurrent scheduled tasks can make
    # Outlook reject one another or produce duplicate drafts/invitations.
    if command -v flock >/dev/null 2>&1; then
        exec {lock_fd}>"$OUTLOOK_COM_LOCK_FILE"
        if ! flock -w "$lock_wait_sec" "$lock_fd"; then
            printf '{"error":"Timed out waiting for the Outlook COM lock"}\n'
            exec {lock_fd}>&-
            return 1
        fi
    fi

    cleanup_outlook_limited_job() {
        trap - INT TERM
        if [ "$task_created" -eq 1 ]; then
            "$SCHTASKS_EXE" /Delete /F /TN "$task_name" >/dev/null 2>&1 || true
        fi
        rm -f -- "$payload_wsl" "$driver_wsl" "$result_wsl" "$status_wsl"
        if [ -n "$lock_fd" ]; then
            flock -u "$lock_fd" 2>/dev/null || true
            exec {lock_fd}>&-
            lock_fd=""
        fi
    }
    trap 'cleanup_outlook_limited_job; return 130' INT TERM

    payload_win=$(wslpath -w "$payload_wsl")
    driver_win=$(wslpath -w "$driver_wsl")
    result_win=$(wslpath -w "$result_wsl")
    status_win=$(wslpath -w "$status_wsl")

    # Windows PowerShell 5.1 needs a UTF-8 BOM to decode non-ASCII mail,
    # meeting, and contact text in script files correctly.
    umask 077
    printf '\xEF\xBB\xBF%s\n' "$ps_command" > "$payload_wsl"
    {
        printf '\xEF\xBB\xBF'
        printf '%s\n' '[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)'
        printf "\$payloadPath = '%s'\n" "${payload_win//\'/\'\'}"
        printf "\$resultPath = '%s'\n" "${result_win//\'/\'\'}"
        printf "\$statusPath = '%s'\n" "${status_win//\'/\'\'}"
        printf '%s\n' '& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $payloadPath 2>&1 | Out-File -LiteralPath $resultPath -Encoding utf8 -Width 1048576'
        printf '%s\n' '$payloadExitCode = $LASTEXITCODE'
        printf '%s\n' '[System.IO.File]::WriteAllText($statusPath, [string]$payloadExitCode, [System.Text.Encoding]::ASCII)'
        printf '%s\n' 'exit $payloadExitCode'
    } > "$driver_wsl"

    local create_output
    create_output=$(
        "$SCHTASKS_EXE" /Create /F /TN "$task_name" \
            /TR "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File \"$driver_win\"" \
            /SC ONCE /ST 00:00 /RL LIMITED 2>&1
    )
    if [ "$?" -ne 0 ]; then
        printf '{"error":"Failed to create LIMITED Outlook task","detail":"%s"}\n' \
            "$(printf '%s' "$create_output" | tr -d '\r\n' | sed -e 's/\\/\\\\/g' -e 's/"/\\"/g')"
        cleanup_outlook_limited_job
        return 1
    fi
    task_created=1

    local run_output
    run_output=$("$SCHTASKS_EXE" /Run /TN "$task_name" 2>&1)
    if [ "$?" -ne 0 ]; then
        printf '{"error":"Failed to start LIMITED Outlook task","detail":"%s"}\n' \
            "$(printf '%s' "$run_output" | tr -d '\r\n' | sed -e 's/\\/\\\\/g' -e 's/"/\\"/g')"
        cleanup_outlook_limited_job
        return 1
    fi

    while [ ! -f "$status_wsl" ] && [ "$elapsed" -lt "$max_ticks" ]; do
        sleep 0.2
        elapsed=$((elapsed + 1))
    done

    if [ ! -f "$status_wsl" ]; then
        "$SCHTASKS_EXE" /End /TN "$task_name" >/dev/null 2>&1 || true
        printf '{"error":"LIMITED Outlook task timed out after %s seconds"}\n' "$timeout_sec"
        cleanup_outlook_limited_job
        return 124
    fi

    if [ -f "$result_wsl" ]; then
        sed $'1s/^\xEF\xBB\xBF//' "$result_wsl" | tr -d '\r'
    fi
    exit_code=$(tr -cd '0-9' < "$status_wsl")
    if ! [[ "$exit_code" =~ ^[0-9]+$ ]]; then
        exit_code=1
    fi
    cleanup_outlook_limited_job
    return "$exit_code"
}

# Run an Outlook PowerShell command at the same integrity level as Outlook.
# Use direct execution for a normal token and schtasks /RL LIMITED only when
# WSL interop is elevated.
run_outlook_ps_command() {
    local ps_command="$1"
    local integrity_status=0
    windows_powershell_is_elevated
    integrity_status=$?
    if [ "$integrity_status" -eq 0 ] || [ "$integrity_status" -eq 2 ]; then
        run_outlook_ps_limited "$ps_command"
    else
        local lock_wait_sec="${OUTLOOK_LOCK_WAIT_SEC:-900}"
        local direct_lock_fd=""
        if command -v flock >/dev/null 2>&1; then
            exec {direct_lock_fd}>"$OUTLOOK_COM_LOCK_FILE"
            if ! flock -w "$lock_wait_sec" "$direct_lock_fd"; then
                printf '{"error":"Timed out waiting for the Outlook COM lock"}\n'
                exec {direct_lock_fd}>&-
                return 1
            fi
        fi
        "$PS_EXE" -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "$ps_command" 2>&1 | tr -d '\r'
        local direct_status="${PIPESTATUS[0]}"
        if [ -n "$direct_lock_fd" ]; then
            flock -u "$direct_lock_fd" 2>/dev/null || true
            exec {direct_lock_fd}>&-
        fi
        return "$direct_status"
    fi
}

# --- Ensure WSL_INTEROP is valid (auto-fix after reboot) ---
# WSL_INTEROP path (e.g. /run/WSL/1329_interop) changes on every Windows/WSL
# restart.  When the gateway runs as a systemd service it keeps the stale value
# from boot time.  This function picks the freshest socket and updates both the
# running env and ~/.openclaw/service.env so the gateway stays healthy.
ensure_wsl_interop() {
    # Already valid — nothing to do.
    if [ -n "$WSL_INTEROP" ] && [ -e "$WSL_INTEROP" ]; then
        return 0
    fi

    # Find the most-recently-created interop socket.
    local latest
    latest=$(ls -t /run/WSL/*_interop 2>/dev/null | head -1)
    if [ -z "$latest" ]; then
        return 1  # No interop socket at all (not WSL2?)
    fi

    export WSL_INTEROP="$latest"

    # Persist into service.env so the gateway picks it up on next command too.
    local env_file="$HOME/.openclaw/service.env"
    if [ -f "$env_file" ]; then
        if grep -q '^WSL_INTEROP=' "$env_file"; then
            sed -i "s|^WSL_INTEROP=.*|WSL_INTEROP=$latest|" "$env_file"
        else
            echo "WSL_INTEROP=$latest" >> "$env_file"
        fi
    fi
}

# --- Run on source ---
ensure_wsl_interop
