"""
Verify the CCW vtable layout of ImageZoomAddin against what Office actually does.

Office calls IDTExtensibility2 through the vtable of the dual interface:
    slot 0-2  IUnknown   (QueryInterface / AddRef / Release)
    slot 3-6  IDispatch  (GetTypeInfoCount / GetTypeInfo / GetIDsOfNames / Invoke)
    slot 7    OnConnection          <-- Office calls THIS
    slot 8    OnDisconnection
    slot 9    OnAddInsUpdate
    slot 10   OnStartupComplete
    slot 11   OnBeginShutdown

If the managed interface was declared InterfaceIsIDispatch, the CCW only has
slots 0-6, so Office calling slot 7 jumps past the vtable -> access violation
-> the host (Word/Excel) dies and gets blacklisted in Resiliency\\DisabledItems.

Usage:
    python tools/vtabletest.py            # safe: introspection only
    python tools/vtabletest.py --call     # also invoke slot 7 (may crash this process)
"""

import ctypes, os, sys
from ctypes import wintypes, byref, c_void_p

CLSID = "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
IID_IDISPATCH = "{00020400-0000-0000-C000-000000000046}"
IID_EXT2 = "{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}"
CLSCTX_INPROC_SERVER = 0x1

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


kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
kernel32.GetModuleHandleExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR,
                                       ctypes.POINTER(c_void_p)]
kernel32.GetModuleHandleExW.restype = wintypes.BOOL
kernel32.GetModuleFileNameW.argtypes = [c_void_p, wintypes.LPWSTR, wintypes.DWORD]
kernel32.GetModuleFileNameW.restype = wintypes.DWORD

GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS = 0x00000004
GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT = 0x00000002


def module_of(addr):
    """Return the module file that contains `addr`, or None if the address is
    not inside any loaded module (= it is heap data, i.e. NOT real code)."""
    if not addr:
        return None
    hmod = c_void_p()
    flags = GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT
    if not kernel32.GetModuleHandleExW(flags, wintypes.LPCWSTR(addr), byref(hmod)):
        return None
    if not hmod:
        return None
    buf = ctypes.create_unicode_buffer(520)
    n = kernel32.GetModuleFileNameW(hmod, buf, 520)
    return buf.value if n else None


def qi(p, iid_str):
    vt = vtable_of(p)
    QI = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.POINTER(GUID),
                            ctypes.POINTER(c_void_p))(vt[0])
    iid = guid(iid_str)
    out = c_void_p()
    hr = QI(p, byref(iid), byref(out))
    return hr, out


def main():
    do_call = "--call" in sys.argv

    hr_init = ole32.CoInitialize(None)
    print("CoInitialize          hr=0x%08X" % (hr_init & 0xFFFFFFFF))
    print("python bitness        %d" % (ctypes.sizeof(c_void_p) * 8))

    p = c_void_p()
    hr = ole32.CoCreateInstance(byref(guid(CLSID)), None, CLSCTX_INPROC_SERVER,
                                byref(guid("{00000000-0000-0000-C000-000000000046}")),
                                byref(p))
    print("CoCreateInstance      hr=0x%08X" % (hr & 0xFFFFFFFF))
    if hr != 0:
        return 1

    hr_disp, pdisp = qi(p, IID_IDISPATCH)
    hr_ext, pext = qi(p, IID_EXT2)
    print("QI IDispatch          hr=0x%08X ptr=%s" % (hr_disp & 0xFFFFFFFF, pdisp))
    print("QI IDTExtensibility2  hr=0x%08X ptr=%s" % (hr_ext & 0xFFFFFFFF, pext))

    if pext:
        vt = vtable_of(pext)
        print("")
        print("vtable of IDTExtensibility2 at 0x%X" % ctypes.cast(vt, c_void_p).value)
        names = ["QueryInterface", "AddRef", "Release",
                 "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke",
                 "OnConnection", "OnDisconnection", "OnAddInsUpdate",
                 "OnStartupComplete", "OnBeginShutdown"]
        for i, n in enumerate(names):
            addr = vt[i] or 0
            mod = module_of(addr)
            if mod:
                mod = os.path.basename(mod)
            else:
                mod = "<NOT IN ANY MODULE = raw heap data>"
            print("  slot %-2d 0x%016X  %-18s %s" % (i, addr, n, mod))

        print("")
        # 注意：CLR 会把接口方法的 thunk 放在动态分配的代码堆里，
        # GetModuleHandleEx 查不到模块并不代表非法（曾因此误判过一次）。
        # 真正的判定标准只有一个：调 slot 7 看能不能跑进 OnConnection，
        # 见 tools/verify.py。
        inmod = [i for i in range(12) if module_of(vt[i] or 0)]
        print("in-module slots: %s" % inmod)
        print("VERDICT: layout only. Run `python tools/verify.py` for the real")
        print("         pass/fail (it calls slot 7 and checks the add-in log).")

    if do_call:
        print("")
        print("--- invoking slot 7 (OnConnection) with a NULL host pointer ---")
        print("--- if the vtable layout is wrong this process will die here ---")
        sys.stdout.flush()
        vt = vtable_of(pext)
        # HRESULT OnConnection(IDispatch* Application, long ConnectMode,
        #                      IDispatch* AddInInst, SAFEARRAY(VARIANT)** custom)
        proto = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p,
                                   ctypes.c_long, c_void_p,
                                   ctypes.POINTER(c_void_p))
        fn = proto(vt[7])
        safearr = c_void_p()
        # ConnectMode 2 = ext_cm_External
        hr_call = fn(pext, None, 2, None, byref(safearr))
        print("OnConnection returned hr=0x%08X" % (hr_call & 0xFFFFFFFF))
        return 0 if hr_call == 0 else 2

    return 0


if __name__ == "__main__":
    sys.exit(main())
