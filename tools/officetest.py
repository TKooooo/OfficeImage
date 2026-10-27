import ctypes, os, sys, time, subprocess
from ctypes import wintypes, byref, c_void_p, POINTER

ole32 = ctypes.WinDLL("ole32")
oleaut32 = ctypes.WinDLL("oleaut32")
ole32.CoInitialize(None)

DISPATCH_METHOD = 1
DISPATCH_PROPERTYPUT = 4
DISPID_PROPERTYPUT = -3
CLSCTX_LOCAL_SERVER = 0x4

class GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

class VUNION(ctypes.Union):
    _fields_ = [("ll", ctypes.c_longlong), ("dbl", ctypes.c_double),
                ("p", c_void_p), ("pad", ctypes.c_ubyte * 16)]

class VARIANT(ctypes.Structure):
    _fields_ = [("vt", ctypes.c_ushort), ("r", ctypes.c_ushort * 3), ("u", VUNION)]

class DISPPARAMS(ctypes.Structure):
    _fields_ = [("rgvarg", POINTER(VARIANT)),
                ("rgdispidNamedArgs", POINTER(ctypes.c_long)),
                ("cArgs", ctypes.c_uint), ("cNamedArgs", ctypes.c_uint)]

def g(s):
    x = GUID(); ole32.CLSIDFromString(s, byref(x)); return x

NULLIID = GUID()

PROGIDS = {
    "Word":  ("{000209FF-0000-0000-C000-000000000046}", "Word.Application"),
    "Excel": ("{00024500-0000-0000-C000-000000000046}", "Excel.Application"),
}

def vcall(pdisp, name, wflags, args=None):
    vtbl = ctypes.cast(pdisp, POINTER(POINTER(c_void_p))).contents
    GetIDsOfNames = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID),
                                      POINTER(wintypes.LPOLESTR), ctypes.c_uint,
                                      ctypes.c_ulong, POINTER(ctypes.c_long))(vtbl[5])
    Invoke = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_long, POINTER(GUID),
                                ctypes.c_ulong, ctypes.c_ushort, POINTER(DISPPARAMS),
                                POINTER(VARIANT), c_void_p, POINTER(ctypes.c_uint))(vtbl[6])
    dispid = ctypes.c_long()
    names = (wintypes.LPOLESTR * 1)(name)
    hr = GetIDsOfNames(pdisp, byref(NULLIID), names, 1, 0, byref(dispid))
    if hr != 0:
        return ("GetIDsOfNames failed", hr)
    arr = (VARIANT * max(1, len(args or [])))()
    if args:
        for i, v in enumerate(args):
            arr[i].vt = 11  # VT_BOOL
            arr[i].u.ll = -1 if v else 0
    dp = DISPPARAMS()
    dp.rgvarg = arr
    dp.cArgs = len(args or [])
    named = (ctypes.c_long * 1)(DISPID_PROPERTYPUT)
    if wflags & DISPATCH_PROPERTYPUT and not args:
        pass
    if wflags == DISPATCH_PROPERTYPUT:
        dp.rgdispidNamedArgs = named
        dp.cNamedArgs = 1
    hr = Invoke(pdisp, dispid.value, byref(NULLIID), 0, wflags, byref(dp), None, None, None)
    return ("ok", hr)

def run(host):
    clsid, progid = PROGIDS[host]
    p = c_void_p()
    hr = ole32.CoCreateInstance(byref(g(clsid)), None, CLSCTX_LOCAL_SERVER,
                                byref(g("{00020400-0000-0000-C000-000000000046}")), byref(p))
    print("%s: CoCreateInstance 0x%08X" % (host, hr & 0xFFFFFFFF))
    if hr != 0:
        return False
    try:
        print("   set Visible=False:", vcall(p, "Visible", DISPATCH_PROPERTYPUT, [0]))
    except Exception as e:
        print("   visible set error:", e)
    time.sleep(5)
    vb = ctypes.cast(p, POINTER(POINTER(c_void_p))).contents
    Release = ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)(vb[2])
    try:
        print("   Quit:", vcall(p, "Quit", DISPATCH_METHOD, []))
    except Exception as e:
        print("   quit error:", e)
    Release(p)
    return True

host = sys.argv[1] if len(sys.argv) > 1 else "Word"
run(host)

log = os.path.join(os.environ["LOCALAPPDATA"], "ImageZoomAddin", "addin.log")
print()
print("=== log:", log, "===")
if os.path.exists(log):
    with open(log, "r", encoding="utf-8", errors="replace") as f:
        lines = f.readlines()
    print("lines:", len(lines))
    for ln in lines[-40:]:
        print("  " + ln.rstrip())
else:
    print("  *** NO LOG FILE -> OnConnection never ran ***")
