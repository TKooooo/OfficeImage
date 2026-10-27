"""
Acceptance test for the add-in's COM exposure - does NOT need Office.

For every interface pointer a host might call OnConnection on (IDispatch and
IDTExtensibility2), it invokes vtable slot 7 and checks that the managed
OnConnection actually ran (by watching the add-in log).

Office's rule: IDTExtensibility2 is a dual interface, OnConnection is slot 7.
Whatever pointer Office ends up using, slot 7 must be OnConnection - not
ToString, not heap garbage.

    python tools/verify.py
"""

import ctypes, os, sys, time
from ctypes import wintypes, byref, c_void_p

CLSID = "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
IID_IDISPATCH = "{00020400-0000-0000-C000-000000000046}"
IID_EXT2 = "{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}"
CLSCTX_INPROC_SERVER = 0x1

LOG = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ImageZoomAddin", "addin.log")

ole32 = ctypes.WinDLL("ole32", use_last_error=True)
ole32.CoInitialize.argtypes = [c_void_p]
ole32.CoInitialize.restype = ctypes.c_long
ole32.CLSIDFromString.argtypes = [wintypes.LPCWSTR, c_void_p]
ole32.CLSIDFromString.restype = ctypes.c_long
ole32.CoCreateInstance.argtypes = [c_void_p, c_void_p, wintypes.DWORD, c_void_p,
                                   ctypes.POINTER(c_void_p)]
ole32.CoCreateInstance.restype = ctypes.c_long


class GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]


def guid(s):
    g = GUID()
    if ole32.CLSIDFromString(s, byref(g)) != 0:
        raise OSError("bad guid " + s)
    return g


def vtable_of(p):
    return ctypes.cast(p, ctypes.POINTER(ctypes.POINTER(c_void_p))).contents


def log_text():
    try:
        with open(LOG, "r", encoding="utf-8", errors="replace") as f:
            return f.read()
    except OSError:
        return ""


def main():
    failed = []

    print("add-in log: %s" % LOG)
    before = log_text()
    ole32.CoInitialize(None)

    p = c_void_p()
    hr = ole32.CoCreateInstance(byref(guid(CLSID)), None, CLSCTX_INPROC_SERVER,
                                byref(guid("{00000000-0000-0000-C000-000000000046}")),
                                byref(p))
    print("CoCreateInstance     hr=0x%08X" % (hr & 0xFFFFFFFF))
    if hr != 0:
        print("FAIL: cannot even create the object")
        return 1

    QI = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.POINTER(GUID),
                            ctypes.POINTER(c_void_p))(vtable_of(p)[0])

    proto = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p, ctypes.c_long,
                               c_void_p, ctypes.POINTER(c_void_p))

    for label, iid in (("IDispatch", IID_IDISPATCH), ("IDTExtensibility2", IID_EXT2)):
        q = c_void_p()
        hr = QI(p, byref(guid(iid)), byref(q))
        print("")
        print("=== %s ===" % label)
        print("  QI                 hr=0x%08X" % (hr & 0xFFFFFFFF))
        if not q:
            failed.append(label + ": QI failed")
            continue

        slot7 = vtable_of(q)[7]
        print("  vtable slot 7      0x%X" % (slot7 or 0))

        before_call = log_text()
        safearr = c_void_p()
        try:
            fn = proto(slot7)
            print("  calling slot 7 as OnConnection(host, mode, addin, &custom) ...")
            sys.stdout.flush()
            hr_call = fn(q, None, 2, None, byref(safearr))
        except Exception as exc:                      # noqa: BLE001
            print("  EXCEPTION: %s" % exc)
            failed.append(label + ": call raised")
            continue

        print("  returned           hr=0x%08X" % (hr_call & 0xFFFFFFFF))
        time.sleep(0.15)
        after = log_text()
        grew = len(after) > len(before_call)
        entered = "OnConnection: enter" in after[len(before_call):]
        if hr_call == 0 and grew and entered:
            print("  RESULT             PASS - slot 7 really is OnConnection")
        else:
            print("  RESULT             FAIL - slot 7 did not reach OnConnection")
            print("                     (log grew=%s, OnConnection logged=%s)"
                  % (grew, entered))
            failed.append(label + ": slot 7 wrong")

    print("")
    print("=" * 60)
    if failed:
        print("VERDICT: FAIL")
        for f in failed:
            print("  - " + f)
        return 2
    print("VERDICT: PASS - every pointer reaches OnConnection at slot 7.")
    print("         Safe to start Word/Excel now.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
