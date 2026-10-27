"""
Dump the type information the CLR actually published for our COM object.

Answers the only question that matters:
    did the CCW expose IDTExtensibility2's 5 methods at all?

  cbSizeVft = 96  -> a 12-slot dual vtable (7 IDispatch + 5 methods)  [what Office wants]
  cbSizeVft = 56  -> only IDispatch's 7 slots, methods reachable via Invoke
  cFuncs    = 0   -> nothing was published at all
"""

import ctypes, sys
from ctypes import wintypes, byref, c_void_p

CLSID = "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
IID_IDISPATCH = "{00020400-0000-0000-C000-000000000046}"
CLSCTX_INPROC_SERVER = 0x1

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


class TYPEATTR_HEAD(ctypes.Structure):
    """First part of TYPEATTR - up to the fields we care about."""
    _fields_ = [("guid", GUID),
                ("lcid", wintypes.DWORD),
                ("dwReserved", wintypes.DWORD),
                ("memidConstructor", ctypes.c_long),
                ("memidDestructor", ctypes.c_long),
                ("lpstrSchema", c_void_p),
                ("cbSizeInstance", wintypes.DWORD),
                ("typekind", ctypes.c_int),
                ("cFuncs", wintypes.WORD),
                ("cVars", wintypes.WORD),
                ("cImplTypes", wintypes.WORD),
                ("cbSizeVft", wintypes.WORD),
                ("cbAlignment", wintypes.WORD),
                ("wTypeFlags", wintypes.WORD),
                ("wMajorVerNum", wintypes.WORD),
                ("wMinorVerNum", wintypes.WORD)]


class FUNCDESC_HEAD(ctypes.Structure):
    _fields_ = [("memid", ctypes.c_long),
                ("lprgscode", c_void_p),
                ("lprgelemdescParam", c_void_p),
                ("funckind", ctypes.c_int),
                ("invkind", ctypes.c_int),
                ("callconv", ctypes.c_int),
                ("cParams", ctypes.c_short),
                ("cParamsOpt", ctypes.c_short),
                ("oVft", ctypes.c_short),
                ("cScodes", ctypes.c_short)]


def guid(s):
    g = GUID()
    ole32.CLSIDFromString(s, byref(g))
    return g


def vtable_of(p):
    return ctypes.cast(p, ctypes.POINTER(ctypes.POINTER(c_void_p))).contents


def main():
    ole32.CoInitialize(None)
    p = c_void_p()
    hr = ole32.CoCreateInstance(byref(guid(CLSID)), None, CLSCTX_INPROC_SERVER,
                                byref(guid("{00000000-0000-0000-C000-000000000046}")),
                                byref(p))
    print("CoCreateInstance hr=0x%08X" % (hr & 0xFFFFFFFF))
    if hr != 0:
        return 1

    QI = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.POINTER(GUID),
                            ctypes.POINTER(c_void_p))(vtable_of(p)[0])
    iid = guid(IID_IDISPATCH)
    pdisp = c_void_p()
    hr = QI(p, byref(iid), byref(pdisp))
    print("QI IDispatch      hr=0x%08X" % (hr & 0xFFFFFFFF))
    if not pdisp:
        return 1

    dv = vtable_of(pdisp)
    GetTypeInfoCount = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p,
                                          ctypes.POINTER(wintypes.UINT))(dv[3])
    GetTypeInfo = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, wintypes.UINT,
                                     wintypes.DWORD,
                                     ctypes.POINTER(c_void_p))(dv[4])
    n = wintypes.UINT(99)
    hr = GetTypeInfoCount(pdisp, byref(n))
    print("GetTypeInfoCount  hr=0x%08X  count=%d" % (hr & 0xFFFFFFFF, n.value))
    if n.value == 0:
        print("")
        print("=> The CLR published NO type info at all: the dispatch map is empty,")
        print("   which is why GetIDsOfNames('OnConnection') returns DISP_E_UNKNOWNNAME.")
        return 0

    pti = c_void_p()
    hr = GetTypeInfo(pdisp, 0, 0, byref(pti))
    print("GetTypeInfo(0)    hr=0x%08X" % (hr & 0xFFFFFFFF))
    if not pti:
        return 1

    tv = vtable_of(pti)
    GetTypeAttr = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p,
                                     ctypes.POINTER(c_void_p))(tv[3])
    GetFuncDesc = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, wintypes.UINT,
                                     ctypes.POINTER(c_void_p))(tv[5])
    GetNames = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_long,
                                  ctypes.POINTER(c_void_p), wintypes.UINT,
                                  ctypes.POINTER(wintypes.UINT))(tv[7])
    GetRefTypeOfImplType = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, wintypes.UINT,
                                              ctypes.POINTER(wintypes.DWORD))(tv[8])
    GetRefTypeInfo = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, wintypes.DWORD,
                                        ctypes.POINTER(c_void_p))(tv[9])
    GetDoc = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_long,
                                ctypes.POINTER(wintypes.LPWSTR),
                                ctypes.POINTER(wintypes.LPWSTR),
                                ctypes.POINTER(wintypes.DWORD),
                                ctypes.POINTER(wintypes.LPWSTR))(tv[12])
    oleaut32.SysFreeString.argtypes = [c_void_p]

    def name_of(pti_, memid):
        pb = c_void_p()
        cnt = wintypes.UINT(0)
        if GetNames(pti_, memid, byref(pb), 1, byref(cnt)) != 0 or not pb.value:
            return "<none>"
        s = ctypes.wstring_at(pb.value)
        oleaut32.SysFreeString(pb)
        return s

    def doc_of(pti_):
        nm = wintypes.LPWSTR()
        d = wintypes.LPWSTR()
        c = wintypes.DWORD()
        h = wintypes.LPWSTR()
        if GetDoc(pti_, -1, byref(nm), byref(d), byref(c), byref(h)) == 0:
            return nm.value or "<none>"
        return "<none>"

    pta = c_void_p()
    hr = GetTypeAttr(pti, byref(pta))
    print("GetTypeAttr       hr=0x%08X" % (hr & 0xFFFFFFFF))
    ta = ctypes.cast(pta, ctypes.POINTER(TYPEATTR_HEAD)).contents

    print("")
    print("type name         %s" % doc_of(pti))
    print("type guid         {%08X-%04X-%04X-%s}" % (
        ta.guid.Data1, ta.guid.Data2, ta.guid.Data3,
        "".join("%02X" % b for b in ta.guid.Data4)))
    print("typekind          %d  (4 = TKIND_INTERFACE, 5 = TKIND_DISPATCH)"
          % ta.typekind)
    print("cFuncs            %d" % ta.cFuncs)
    print("cImplTypes        %d" % ta.cImplTypes)
    print("cbSizeVft         %d bytes = %d slots" % (ta.cbSizeVft, ta.cbSizeVft // 8))
    print("wTypeFlags        0x%04X" % ta.wTypeFlags)
    print("")

    # what does this interface inherit from?
    for i in range(ta.cImplTypes):
        href = wintypes.DWORD()
        if GetRefTypeOfImplType(pti, i, byref(href)) == 0:
            prti = c_void_p()
            if GetRefTypeInfo(pti, href.value, byref(prti)) == 0 and prti:
                name, cfuncs, vft, kind = "<none>", -1, -1, -1
                pta2 = c_void_p()
                if GetTypeAttr(prti, byref(pta2)) == 0 and pta2:
                    ta2 = ctypes.cast(pta2, ctypes.POINTER(TYPEATTR_HEAD)).contents
                    name = doc_of(prti)
                    cfuncs = ta2.cFuncs
                    vft = ta2.cbSizeVft
                    kind = ta2.typekind
                print("  inherits[%d]      %s  kind=%d cFuncs=%d cbSizeVft=%d"
                      % (i, name, kind, cfuncs, vft))

    print("")
    for i in range(ta.cFuncs):
        pfd = c_void_p()
        if GetFuncDesc(pti, i, byref(pfd)) != 0 or not pfd:
            continue
        fd = ctypes.cast(pfd, ctypes.POINTER(FUNCDESC_HEAD)).contents
        print("  slot %-2d: %-20s memid=%-10d invkind=%d oVft=%3d cParams=%d"
              % (fd.oVft // 8, name_of(pti, fd.memid), fd.memid, fd.invkind,
                 fd.oVft, fd.cParams))
    return 0


if __name__ == "__main__":
    sys.exit(main())
