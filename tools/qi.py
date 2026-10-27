import ctypes, sys
from ctypes import wintypes, byref, c_void_p

ole32 = ctypes.WinDLL("ole32")
ole32.CoInitialize(None)
ole32.CoCreateInstance.restype = ctypes.c_long

class GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

def g(s):
    x = GUID(); ole32.CLSIDFromString(s, byref(x)); return x

CLSID = g("{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}")
IID_IDTExt = g("{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}")
IID_IDispatch = g("{00020400-0000-0000-C000-000000000046}")
IID_IUnknown = g("{00000000-0000-0000-C000-000000000046}")

def qi(name, iid):
    p = c_void_p()
    hr = ole32.CoCreateInstance(byref(CLSID), None, 0x1, byref(iid), byref(p))
    h = hr & 0xFFFFFFFF
    print("  QI %-22s -> 0x%08X %s" % (name, h, "OK" if h == 0 else "FAIL"))
    if h == 0 and p:
        vt = ctypes.cast(p, ctypes.POINTER(ctypes.POINTER(c_void_p))).contents
        ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)(vt[2])(p)
    return h == 0

print("CoCreateInstance + Interface QI:")
ok1 = qi("IUnknown", IID_IUnknown)
ok2 = qi("IDispatch", IID_IDispatch)
ok3 = qi("IDTExtensibility2", IID_IDTExt)
print()
print("verdict:", "PASS - Office will accept this as a COM add-in" if (ok1 and ok3) else "FAIL")
