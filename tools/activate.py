import ctypes, sys
from ctypes import wintypes, byref, c_void_p

ole32 = ctypes.WinDLL("ole32", use_last_error=True)
ole32.CoInitialize.argtypes = [c_void_p]
ole32.CoInitialize.restype = ctypes.c_long
ole32.CLSIDFromString.argtypes = [wintypes.LPCWSTR, c_void_p]
ole32.CLSIDFromString.restype = ctypes.c_long
ole32.CoCreateInstance.argtypes = [c_void_p, c_void_p, wintypes.DWORD, c_void_p, ctypes.POINTER(c_void_p)]
ole32.CoCreateInstance.restype = ctypes.c_long

class GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

CLSCTX_INPROC_SERVER = 0x1
IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"

def act(clsid):
    hr = ole32.CoInitialize(None)
    g = GUID()
    hr2 = ole32.CLSIDFromString(clsid, byref(g))
    if hr2 != 0:
        return ("BAD_CLSID", hr2)
    iid = GUID()
    ole32.CLSIDFromString(IID_IUnknown, byref(iid))
    p = c_void_p()
    hr3 = ole32.CoCreateInstance(byref(g), None, CLSCTX_INPROC_SERVER, byref(iid), byref(p))
    if hr3 == 0 and p:
        # release
        vtbl = ctypes.cast(p, ctypes.POINTER(ctypes.POINTER(c_void_p))).contents
        Release = ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)(vtbl[2])
        Release(p)
    return ("OK" if hr3 == 0 else "FAIL", hr3)

clsid = sys.argv[1] if len(sys.argv) > 1 else "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
st, hr = act(clsid)
hr_u = hr & 0xFFFFFFFF
print("%-8s HRESULT=0x%08X %s" % (st, hr_u, ctypes.FormatError(hr_u & 0xFFFF) if hr_u >> 16 == 0x8007 else ""))
