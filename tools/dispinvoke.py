# -*- coding: utf-8 -*-
"""
用 Office 的真实调用方式（IDispatch::Invoke, DISPID 1）触发 OnConnection。

目的：本地复现「Excel 崩在 clr.dll+0x14eb0f / 0x80131506」。
如果这个脚本崩掉，说明崩点就在参数封送阶段，与 Office 无关，可以本地反复迭代。

用法：
    python tools/dispinvoke.py byref-sa      # custom = VT_VARIANT|VT_BYREF -> VT_SAFEARRAY(VARIANT)
    python tools/dispinvoke.py byref-empty   # custom = VT_VARIANT|VT_BYREF -> VT_EMPTY
    python tools/dispinvoke.py empty         # custom = VT_EMPTY
    python tools/dispinvoke.py sa            # custom = VT_SAFEARRAY (by value)
    python tools/dispinvoke.py none          # 只传 3 个参数（缺 custom）
"""

import ctypes
import os
import sys
import time
from ctypes import wintypes, byref, c_void_p, POINTER

CLSID_ADDIN = "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
IID_IDISPATCH = "{00020400-0000-0000-C000-000000000046}"
IID_NULL = "{00000000-0000-0000-0000-000000000000}"

LOG_PATH = os.path.join(os.environ.get("LOCALAPPDATA", os.environ.get("TEMP", ".")),
                        "ImageZoomAddin", "addin.log")

ole32 = ctypes.WinDLL("ole32")
oleaut32 = ctypes.WinDLL("oleaut32")

ole32.CoInitialize.argtypes = [c_void_p]
ole32.CoInitialize.restype = ctypes.c_long
ole32.CLSIDFromString.argtypes = [wintypes.LPCWSTR, c_void_p]
ole32.CLSIDFromString.restype = ctypes.c_long
ole32.CoCreateInstance.argtypes = [c_void_p, c_void_p, wintypes.DWORD, c_void_p, POINTER(c_void_p)]
ole32.CoCreateInstance.restype = ctypes.c_long
oleaut32.SafeArrayCreate.argtypes = [wintypes.DWORD, wintypes.UINT, c_void_p]
oleaut32.SafeArrayCreate.restype = c_void_p


class GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]


def guid(text):
    g = GUID()
    ole32.CLSIDFromString(text, byref(g))
    return g


class VARIANT(ctypes.Structure):
    _fields_ = [("vt", ctypes.c_ushort),
                ("r1", ctypes.c_ushort),
                ("r2", ctypes.c_ushort),
                ("r3", ctypes.c_ushort),
                ("data", ctypes.c_byte * 16)]

    def set_ptr(self, value):
        ctypes.cast(byref(self.data), POINTER(c_void_p))[0] = value

    def set_i4(self, value):
        ctypes.cast(byref(self.data), POINTER(ctypes.c_long))[0] = value

    def get_ptr(self):
        return ctypes.cast(byref(self.data), POINTER(c_void_p))[0]


class SAFEARRAYBOUND(ctypes.Structure):
    _fields_ = [("cElements", wintypes.ULONG), ("lLbound", ctypes.c_long)]


class DISPPARAMS(ctypes.Structure):
    _fields_ = [("rgvarg", POINTER(VARIANT)),
                ("rgdispidNamedArgs", c_void_p),
                ("cArgs", ctypes.c_uint),
                ("cNamedArgs", ctypes.c_uint)]


class EXCEPINFO(ctypes.Structure):
    _fields_ = [("wCode", wintypes.WORD),
                ("wReserved", wintypes.WORD),
                ("bstrSource", c_void_p),
                ("bstrDescription", c_void_p),
                ("bstrHelpFile", c_void_p),
                ("dwHelpContext", wintypes.DWORD),
                ("pvReserved", c_void_p),
                ("pfnDeferredFillIn", c_void_p),
                ("scode", ctypes.c_long)]


VT_EMPTY = 0
VT_I4 = 3
VT_DISPATCH = 9
VT_VARIANT = 12
VT_BYREF = 0x4000
VT_SAFEARRAY = 0x1F

DISPATCH_METHOD = 1


def say(msg):
    print(msg, flush=True)


def vtable(ptr):
    return ctypes.cast(ptr, POINTER(POINTER(c_void_p))).contents


def log_size():
    try:
        return os.path.getsize(LOG_PATH)
    except OSError:
        return -1


def log_tail(n=6):
    try:
        with open(LOG_PATH, "r", encoding="utf-8", errors="replace") as f:
            lines = [l.strip() for l in f if l.strip()]
        return lines[-n:]
    except OSError:
        return []


def make_empty_safearray():
    bound = SAFEARRAYBOUND(0, 0)          # 0 elements
    return oleaut32.SafeArrayCreate(VT_VARIANT, 1, byref(bound))


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "byref-sa"
    say("mode: %s" % mode)

    hr = ole32.CoInitialize(None)
    say("CoInitialize hr=0x%08X" % (hr & 0xFFFFFFFF))

    ip = c_void_p()
    hr = ole32.CoCreateInstance(byref(guid(CLSID_ADDIN)), None, 1,
                                byref(guid(IID_IDISPATCH)), byref(ip))
    say("CoCreateInstance hr=0x%08X ptr=%s" % (hr & 0xFFFFFFFF, hex(ip.value or 0)))
    if not ip.value:
        return 1

    # IDispatch vtable: 0 QueryInterface 1 AddRef 2 Release 3 GetTypeInfoCount
    #                   4 GetTypeInfo 5 GetIDsOfNames 6 Invoke
    vt = vtable(ip)
    Invoke = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_long, c_void_p,
                                wintypes.DWORD, wintypes.WORD, POINTER(DISPPARAMS),
                                c_void_p, c_void_p, c_void_p)(vt[6])
    GetIDsOfNames = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p, POINTER(wintypes.LPWSTR),
                                       wintypes.UINT, wintypes.DWORD, POINTER(ctypes.c_long))(vt[5])

    target = sys.argv[2] if len(sys.argv) > 2 else "conn"
    method, want_dispid = {
        "conn": ("OnConnection", 1),
        "disc": ("OnDisconnection", 2),
        "update": ("OnAddInsUpdate", 3),
        "startup": ("OnStartupComplete", 4),
    }.get(target, ("OnConnection", 1))
    say("target: %s (DISPID %d)" % (method, want_dispid))

    name = wintypes.LPWSTR(method)
    dispid = ctypes.c_long(0)
    hr = GetIDsOfNames(ip, byref(guid(IID_NULL)), byref(name), 1, 0, byref(dispid))
    say("GetIDsOfNames('%s') hr=0x%08X dispid=%d" % (method, hr & 0xFFFFFFFF, dispid.value))
    if hr != 0:
        say("  -> DISPID 查不到，说明 IDispatch 里根本没有这个方法")
        return 2

    # ---- 构造参数（DISPPARAMS 里是逆序：最后一个参数在 rgvarg[0]）----
    args = []
    sa = make_empty_safearray()
    say("SafeArrayCreate -> %s" % hex(sa or 0))

    custom = VARIANT()
    if mode == "byref-array":
        # 晚期绑定里 SAFEARRAY(VARIANT) 的正确表示：VT_VARIANT|VT_BYREF -> VT_ARRAY|VT_VARIANT
        inner = VARIANT()
        inner.vt = 0x200C                          # VT_ARRAY | VT_VARIANT
        inner.set_ptr(sa)
        custom.vt = VT_VARIANT | VT_BYREF
        custom.set_ptr(ctypes.cast(byref(inner), c_void_p).value)
    elif mode == "array":
        custom.vt = 0x200C                         # VT_ARRAY | VT_VARIANT，按值
        custom.set_ptr(sa)
    elif mode == "byref-sa":
        inner = VARIANT()
        inner.vt = VT_SAFEARRAY
        inner.set_ptr(sa)
        custom.vt = VT_VARIANT | VT_BYREF
        custom.set_ptr(ctypes.cast(byref(inner), c_void_p).value)
    elif mode == "byref-empty":
        inner = VARIANT()
        inner.vt = VT_EMPTY
        custom.vt = VT_VARIANT | VT_BYREF
        custom.set_ptr(ctypes.cast(byref(inner), c_void_p).value)
    elif mode == "empty":
        custom.vt = VT_EMPTY
    elif mode == "sa":
        custom.vt = VT_SAFEARRAY
        custom.set_ptr(sa)
    elif mode == "none":
        custom = None
    else:
        say("unknown mode")
        return 1

    addin = VARIANT()
    addin.vt = VT_DISPATCH
    addin.set_ptr(None)

    connect_mode = VARIANT()
    connect_mode.vt = VT_I4
    connect_mode.set_i4(2)                 # ext_cm_External

    app = VARIANT()
    app.vt = VT_DISPATCH
    app.set_ptr(ip.value)                  # 用对象自己冒充 Application，避免 NULL 解引用干扰

    if target == "conn":
        ordered = [custom, addin, connect_mode, app]
    elif target == "disc":
        remove_mode = VARIANT()
        remove_mode.vt = VT_I4
        remove_mode.set_i4(0)
        ordered = [custom, remove_mode]
    else:
        ordered = [custom]                     # OnAddInsUpdate / OnStartupComplete 只有 custom
    ordered = [v for v in ordered if v is not None]
    arr = (VARIANT * len(ordered))(*ordered)
    args.append(arr)

    dp = DISPPARAMS()
    dp.rgvarg = arr
    dp.rgdispidNamedArgs = None
    dp.cArgs = len(ordered)
    dp.cNamedArgs = 0

    before_size = log_size()
    before_tail = log_tail()

    excep = EXCEPINFO()
    argerr = ctypes.c_uint(0xFFFFFFFF)

    say("call Invoke(dispid=%d, DISPATCH_METHOD, cArgs=%d) ..." % (dispid.value, len(ordered)))
    say("  params: custom(vt=0x%04X), AddInInst(vt=%d), ConnectMode(vt=%d), Application(vt=%d)"
        % (custom.vt if custom is not None else -1, VT_DISPATCH, VT_I4, VT_DISPATCH))

    hr = Invoke(ip, dispid.value, byref(guid(IID_NULL)), 0x400, DISPATCH_METHOD,
                byref(dp), None, byref(excep), byref(argerr))
    say("Invoke returned hr=0x%08X" % (hr & 0xFFFFFFFF))
    say("  puArgErr = %d   (0 = 最后一个参数，往上递增)" % argerr.value)
    say("  EXCEPINFO scode=0x%08X wCode=%d" % (excep.scode & 0xFFFFFFFF, excep.wCode))
    if excep.bstrSource:
        say("  source      : %s" % ctypes.cast(excep.bstrSource, wintypes.LPWSTR).value)
    if excep.bstrDescription:
        say("  description : %s" % ctypes.cast(excep.bstrDescription, wintypes.LPWSTR).value)

    time.sleep(0.3)
    after_size = log_size()
    say("log size %d -> %d" % (before_size, after_size))
    new_lines = [l for l in log_tail(20) if l not in before_tail]
    if new_lines:
        say("--- 新增日志 ---")
        for l in new_lines:
            say("    " + l)
    else:
        say("--- 日志无新增：OnConnection 的方法体没有执行 ---")

    if hr == 0 and new_lines:
        say("RESULT: PASS - Invoke 到达 OnConnection 并正常返回")
        return 0
    say("RESULT: FAIL - hr=0x%08X" % (hr & 0xFFFFFFFF))
    return 3


if __name__ == "__main__":
    sys.exit(main())
