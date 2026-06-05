from __future__ import annotations

import os
import socket
import subprocess
from pathlib import Path


HOST = "127.0.0.1"
PORT = 8088
PROJECT = Path(__file__).resolve().parents[1]
PYTHON = Path(r"C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe")
SERVER = PROJECT / "app" / "server.py"
LOG = PROJECT / "server_8088_detached.log"


def already_running() -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex((HOST, PORT)) == 0


def main() -> None:
    if already_running():
        return
    env = os.environ.copy()
    env["PYTHONWARNINGS"] = "ignore::DeprecationWarning"
    creationflags = 0x00000008 | 0x00000200
    with LOG.open("ab", buffering=0) as log:
        subprocess.Popen(
            [str(PYTHON), str(SERVER)],
            cwd=str(PROJECT),
            stdin=subprocess.DEVNULL,
            stdout=log,
            stderr=log,
            env=env,
            creationflags=creationflags,
            close_fds=True,
        )


if __name__ == "__main__":
    main()
