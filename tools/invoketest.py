"""
Reproduce (outside Office) what a client does when it calls
IDTExtensibility2::OnConnection on our managed COM object.

Two call styles are supported, because Office and other hosts differ:

  python tools/invoketest.py disp     # late binding: GetIDsOfNames + IDispatch::Invoke
  python tools/invoketest.py vtable    # early binding: call vtable slot 7 directly

Run it in a throw-away process: if the object's COM exposure is broken this
process dies with the same clr.dll fault Office gets (exit code 0x80131506).
"""

import ctypes, sys
from ctypes import wintypes, byref, c_void_p

CLSID = "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
IID_IDISPATCH = "{00020400-0000-0000-C000-000000000046}"
IID_EXT2 = "{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}"
CLSCTX_INPROC_SERVER = 0x1

DISPATCH_METHOD = 0x1

ole32 = ctypes.WinDLL("ole32", use_last_error=True)
ole32.CoInitialize.argtypes = [c_void_p]
ole32.CoInitialize.restype = ctypes.c_long
ole32.CLSIDFromString.argtypes = [wintypes.LPCWSTR, c_void_p]
ole32.CLSIDFromString.restype = ctypes.c_long
ole32.CoCreateInstance.argtypes = [c_void_p, c_void_p, wintypes.DWORD, c_void_p,
                                   ctypes.POINTER(c_void_p)]
ole32.CoCreateInstance.restype = ctypes.c_long
oleaut32 = ctypes.WinDLL("oleaut32", use_last_error=True)
oleaut32.SysAllocString.argtypes = [wintypes.LPCWSTR]
oleaut32.SysAllocString.restype = c_void_p


class GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]


class VARIANT(ctypes.Structure):
    _fields_ = [("vt", ctypes.c_ushort),
                ("wReserved1", ctypes.c_ushort),
                ("wReserved2", ctypes.c_ushort),
                ("wReserved3", ctypes.c_ushort),
                ("data", ctypes.c_byte * 16)]


class DISPPARAMS(ctypes.Structure):
    _fields_ = [("rgvarg", ctypes.POINTER(VARIANT)),
                ("rgdispidNamedArgs", ctypes.POINTER(wintypes.LONG)),
                ("cArgs", wintypes.UINT),
                ("cNamedArgs", wintypes.UINT)]


def guid(s):
    g = GUID()
    if ole32.CLSIDFromString(s, byref(g)) != 0:
        raise OSError("bad guid " + s)
    return g


def vtable_of(p):
    return ctypes.cast(p, ctypes.POINTER(ctypes.POINTER(c_void_p))).contents


def mk_vt_dispatch_null():
    v = VARIANT()
    v.vt = 9  # VT_DISPATCH
    return v


def mk_vt_i4(n):
    v = VARIANT()
    v.vt = 3  # VT_I4
    ctypes.cast(byref(v.data), ctypes.POINTER(ctypes.c_int))[0] = n
    return v


def mk_vt_empty():
    v = VARIANT()
    v.vt = 0  # VT_EMPTY
    return v


def main():
    style = sys.argv[1] if len(sys.argv) > 1 else "disp"
    ole32.CoInitialize(None)

    p = c_void_p()
    hr = ole32.CoCreateInstance(byref(guid(CLSID)), None, CLSCTX_INPROC_SERVER,
                                byref(guid("{00000000-0000-0000-C000-000000000046}")),
                                byref(p))
    print("[%s] CoCreateInstance hr=0x%08X" % (style, hr & 0xFFFFFFFF))
    if hr != 0:
        return 1
    sys.stdout.flush()

    # QI for IDispatch
    vt = vtable_of(p)
    QI = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.POINTER(GUID),
                            ctypes.POINTER(c_void_p))(vt[0])
    iid = guid(IID_IDISPATCH)
    pdisp = c_void_p()
    hr = QI(p, byref(iid), byref(pdisp))
    print("[%s] QI IDispatch hr=0x%08X" % (style, hr & 0xFFFFFFFF))
    if pdisp:
        dv = vtable_of(pdisp)
        GetIDsOfNames = ctypes.WINFUNCTYPE(
            ctypes.c_long, c_void_p, ctypes.POINTER(GUID),
            ctypes.POINTER(wintypes.LPWSTR), wintypes.UINT, wintypes.DWORD,
            ctypes.POINTER(wintypes.LONG))(dv[5])
        Invoke = ctypes.WINFUNCTYPE(
            ctypes.c_long, c_void_p, wintypes.LONG, ctypes.POINTER(GUID),
            wintypes.DWORD, ctypes.c_ushort, ctypes.POINTER(DISPPARAMS),
            ctypes.POINTER(VARIANT), c_void_p,
            ctypes.POINTER(wintypes.UINT))(dv[6])

        name = oleaut32.SysAllocString("OnConnection")
        names = (wintypes.LPWSTR * 1)()
        names[0] = ctypes.cast(name, wintypes.LPWSTR)
        dispid = wintypes.LONG(-1)
        hr_ids = GetIDsOfNames(pdisp, byref(iid), names, 1, 0, byref(dispid))
        print("[%s] GetIDsOfNames('OnConnection') hr=0x%08X dispid=%d"
              % (style, hr_ids & 0xFFFFFFFF, dispid.value))

        if style == "disp":
            # 参数顺序（IDispatch 里是倒序）：custom, AddInInst, ConnectMode, Application
            args = (VARIANT * 4)()
            args[0] = mk_vt_empty()          # custom  (SAFEARRAY(VARIANT)* in/out)
            args[1] = mk_vt_dispatch_null()  # AddInInst
            args[2] = mk_vt_i4(2)            # ext_cm_External
            args[3] = mk_vt_dispatch_null()  # Application
            dp = DISPPARAMS()
            dp.rgvarg = ctypes.cast(args, ctypes.POINTER(VARIANT))
            dp.cArgs = 4
            dp.cNamedArgs = 0
            dp.rgdispidNamedArgs = None
            res = VARIANT()
            excep = ctypes.create_string_buffer(512)
            argerr = wintypes.UINT(0)
            use = dispid.value if dispid.value >= 0 else 1
            print("[%s] --- calling Invoke(dispid=%d) now ---" % (style, use))
            sys.stdout.flush()
            hr_call = Invoke(pdisp, use, byref(iid), 0, DISPATCH_METHOD,
                             byref(dp), byref(res), ctypes.cast(excep, c_void_p),
                             byref(argerr))
            print("[%s] Invoke returned hr=0x%08X argErr=%d"
                  % (style, hr_call & 0xFFFFFFFF, argerr.value))
            return 0 if hr_call == 0 else 2

    if style == "vtable":
        iid2 = guid(IID_EXT2)
        pext = c_void_p()
        hr = QI(p, byref(iid2), byref(pext))
        print("[vtable] QI IDTExtensibility2 hr=0x%08X" % (hr & 0xFFFFFFFF))
        ev = vtable_of(pext)
        proto = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p, ctypes.c_long,
                                   c_void_p, ctypes.POINTER(c_void_p))
        fn = proto(ev[7])
        print("[vtable] --- calling slot 7 now ---")
        sys.stdout.flush()
        sa = c_void_p()
        hr_call = fn(pext, None, 2, None, byref(sa))
        print("[vtable] slot 7 returned hr=0x%08X" % (hr_call & 0xFFFFFFFF))
        return 0 if hr_call == 0 else 2

    return 0


if __name__ == "__main__":
    sys.exit(main())
