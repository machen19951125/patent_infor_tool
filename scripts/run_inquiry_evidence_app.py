#!/usr/bin/env python3
"""Start the Streamlit app on an OS-assigned free localhost port."""

from __future__ import annotations

import socket
import subprocess
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
APP_PATH = PROJECT_DIR / "scripts" / "inquiry_evidence_app.py"


def find_free_local_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
        listener.bind(("127.0.0.1", 0))
        return int(listener.getsockname()[1])


def main() -> int:
    port = find_free_local_port()
    print(f"正在启动询单产品证据采集器：http://127.0.0.1:{port}")
    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(APP_PATH),
        "--server.address",
        "127.0.0.1",
        "--server.port",
        str(port),
        "--browser.gatherUsageStats",
        "false",
    ]
    return subprocess.call(command, cwd=PROJECT_DIR)


if __name__ == "__main__":
    raise SystemExit(main())
