# -*- coding: utf-8 -*-
"""Find the magnifier button window (a 28x28 WinForms top-level window owned by
WINWORD/EXCEL) and click it with a posted WM_LBUTTONDOWN/WM_LBUTTONUP pair.

Usage:
    python tools/clickmag.py list
    python tools/clickmag.py click
"""

import ctypes
import os
import sys
import time
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)

EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
GetClassNameW = user32.GetClassNameW
GetWindowRect = user32.GetWindowRect
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
PostMessageW = user32.PostMessageW
IsWindowVisible = user32.IsWindowVisible

WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
MK_LBUTTON = 0x0001

TARGETS = ("WINWORD.EXE", "EXCEL.EXE")


def target_pids():
    import subprocess
    out = subprocess.run(
        ["tasklist", "/fo", "csv", "/nh"], capture_output=True, text=True
    ).stdout
    pids = []
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[0].upper() in TARGETS:
            try:
                pids.append(int(parts[1]))
            except ValueError:
                pass
    return pids


def find_buttons():
    pids = target_pids()
    found = []

    def cb(hwnd, lparam):
        pid = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value not in pids:
            return True
        buf = ctypes.create_unicode_buffer(256)
        GetClassNameW(hwnd, buf, 256)
        cls = buf.value
        if "WindowsForms10" not in cls:
            return True
        rect = wintypes.RECT()
        GetWindowRect(hwnd, ctypes.byref(rect))
        w = rect.right - rect.left
        h = rect.bottom - rect.top
        if 10 <= w <= 48 and 10 <= h <= 48:
            found.append(
                {
                    "hwnd": hwnd,
                    "class": cls,
                    "rect": (rect.left, rect.top, w, h),
                    "pid": pid.value,
                    "visible": bool(IsWindowVisible(hwnd)),
                }
            )
        return True

    EnumWindows(EnumWindowsProc(cb), 0)
    return found


def click(hwnd, w, h, rect=None):
    x = w // 2
    y = h // 2
    lp = (y << 16) | (x & 0xFFFF)
    PostMessageW(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lp)
    time.sleep(0.05)
    PostMessageW(hwnd, WM_LBUTTONUP, 0, lp)


def real_click(rect):
    """Move the physical cursor onto the button and click, then restore it.

    WinForms decides "was the left button down?" from the real keyboard/mouse
    state, so a posted message is not enough - it needs genuine input.
    """
    pt = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    old = (pt.x, pt.y)

    cx = rect[0] + rect[2] // 2
    cy = rect[1] + rect[3] // 2
    user32.SetCursorPos(cx, cy)
    time.sleep(0.12)
    user32.mouse_event(0x0002, 0, 0, 0, 0)   # LEFTDOWN
    time.sleep(0.05)
    user32.mouse_event(0x0004, 0, 0, 0, 0)   # LEFTUP
    time.sleep(0.2)
    user32.SetCursorPos(old[0], old[1])
    return old, (cx, cy)


def main():
    action = sys.argv[1] if len(sys.argv) > 1 else "list"
    btns = find_buttons()
    print("target pids: %s" % target_pids())
    print("candidate windows: %d" % len(btns))
    for b in btns:
        print(
            "  hwnd=0x%08X class=%s rect=%s visible=%s pid=%d"
            % (b["hwnd"], b["class"], b["rect"], b["visible"], b["pid"])
        )

    if action in ("click", "realclick"):
        if not btns:
            print("NOTHING TO CLICK")
            return 2
        b = next((x for x in btns if x["visible"]), btns[0])
        if action == "realclick":
            old, pos = real_click(b["rect"])
            print("real click at %s (cursor restored to %s)" % (pos, old))
        else:
            print("clicking center of first candidate ...")
            click(b["hwnd"], b["rect"][2], b["rect"][3])
        time.sleep(1.5)
        print("click posted")
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
