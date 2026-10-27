import winreg, subprocess, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
ACT = os.path.join(HERE, "activate.py")
TESTCLSID = "{A1B2C3D4-1111-2222-3333-444455556666}"
KEY = "Software\\Classes\\CLSID\\" + TESTCLSID
INPROC = KEY + "\\InprocServer32"

def activate(clsid):
    r = subprocess.run([sys.executable, ACT, clsid], capture_output=True, text=True, timeout=180)
    return (r.stdout or r.stderr).strip()

def readv(path, name):
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as k:
            v, _ = winreg.QueryValueEx(k, name); return v
    except OSError:
        return "-"

print("===== RuntimeVersion of already-registered classes =====")
for clsid, note in [
    ("{00B01B2E-B1FE-33A6-AD40-57DE8358DC7D}", "mscorlib 2.0 -> OK"),
    ("{010911E2-F61C-479B-B08C-43E6D1299EFE}", "GroupPolicy 2.0 -> OK"),
    ("{000C1286-F68D-4710-AFE8-624B26FC28A9}", "Autodesk 17.0 -> FAIL"),
]:
    p = "SOFTWARE\\Classes\\CLSID\\" + clsid + "\\InprocServer32"
    print("  %s  %-24s RuntimeVersion=%s  CodeBase=%s" % (
        clsid, note, readv(p, "RuntimeVersion"), readv(p, "CodeBase")))

def mkkey():
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, INPROC, 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "mscoree.dll")
        winreg.SetValueEx(k, "ThreadingModel", 0, winreg.REG_SZ, "Both")

def setv(name, val):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, INPROC, 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, name, 0, winreg.REG_SZ, val)

def delv(name):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, INPROC, 0, winreg.KEY_ALL_ACCESS) as k:
            winreg.DeleteValue(k, name)
    except OSError:
        pass

def trial(label, cls, asm, rtv=None, codebase=None):
    try: winreg.DeleteKey(winreg.HKEY_CURRENT_USER, INPROC); winreg.DeleteKey(winreg.HKEY_CURRENT_USER, KEY)
    except OSError: pass
    mkkey()
    setv("Class", cls); setv("Assembly", asm)
    if rtv: setv("RuntimeVersion", rtv)
    if codebase: setv("CodeBase", codebase)
    print("  %-52s %s" % (label, activate(TESTCLSID)))

MSCORLIB4 = "mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
MSCORLIB2 = "mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
OURS = "ImageZoomAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
OURDLL = "file:///C:/Users/BX/WorkBuddy/2026-10-26-11-31-15/imagezoom-addin/ImageZoomAddin.dll"

print()
print("===== matrix =====")
trial("A mscorlib4 + rtv=v4.0.30319", "System.Object", MSCORLIB4, "v4.0.30319")
trial("B mscorlib4 + no rtv", "System.Object", MSCORLIB4)
trial("C mscorlib4 + rtv=v2.0.50727", "System.Object", MSCORLIB4, "v2.0.50727")
trial("D mscorlib2 + rtv=v2.0.50727 (known good)", "System.Object", MSCORLIB2, "v2.0.50727")
trial("E mscorlib2 + NO rtv", "System.Object", MSCORLIB2)
trial("F ours + no rtv + codebase", "ImageZoomAddin.Connect", OURS, None, OURDLL)
trial("G ours + rtv=v4.0.30319 + codebase", "ImageZoomAddin.Connect", OURS, "v4.0.30319", OURDLL)

try: winreg.DeleteKey(winreg.HKEY_CURRENT_USER, INPROC); winreg.DeleteKey(winreg.HKEY_CURRENT_USER, KEY)
except OSError: pass
print("  cleaned test clsid")
