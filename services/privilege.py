"""Admin privilege helpers for Windows."""
from __future__ import annotations

import ctypes
import sys
from typing import Final


SHELLEXECUTE_SUCCESS: Final[int] = 42


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except AttributeError:
        return False


def relaunch_as_admin() -> None:
    params = " ".join(f'"{arg}"' for arg in sys.argv[1:])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
