"""Fixed native source operations, deliberately not a generic MCP command tool."""
import json
import os
import subprocess
from . import rag


def invoke(operation, items=None):
    if os.name != 'nt' or operation not in {'inventory', 'read_batch'}:
        raise ValueError('Invalid native sync operation')
    if items is not None and (operation != 'read_batch' or len(items) > 25):
        raise ValueError('Invalid sync batch')
    executable = os.path.join(os.environ['SystemRoot'], 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
    result = subprocess.run([executable, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                             '-File', str(rag.ROOT / 'scripts/sync_source.ps1')],
                            input=json.dumps({'operation': operation, 'items': items}),
                            text=True, encoding='utf-8', capture_output=True, timeout=600,
                            cwd=rag.ROOT, creationflags=subprocess.CREATE_NO_WINDOW)
    try:
        data = json.loads(result.stdout)
    except ValueError:
        raise RuntimeError('Invalid native source response') from None
    if result.returncode or data.get('error'):
        raise RuntimeError('Native source unavailable or incomplete')
    return data
