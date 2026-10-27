import subprocess, sys, os, shutil, winreg

CLSID = "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
BASE = "Software\\Classes\\CLSID\\" + CLSID
INPROC = BASE + "\\InprocServer32"
VERKEY = INPROC + "\\1.0.0.0"
DLL = r"C:\Users\BX\WorkBuddy\2026-10-26-11-31-15\imagezoom-addin\ImageZoomAddin.dll"
SIMPLE = r"C:\Users\BX\AppData\Local\ImageZoomAddin\ImageZoomAddin.dll"
HERE = os.path.dirname(os.path.abspath(__file__))
ACT = os.path.join(HERE, "activate.py")

def setv(path, name, val):
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, name, 0, winreg.REG_SZ, val)

def delv(path, name):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_ALL_ACCESS) as k:
            winreg.DeleteValue(k, name)
    except OSError:
        pass

def activate():
    r = subprocess.run([sys.executable, ACT], capture_output=True, text=True, timeout=120)
    return (r.stdout or r.stderr).strip()

# ---- snapshot ----
snap = {}
with winreg.OpenKey(winreg.HKEY_CURRENT_USER, INPROC) as k:
    i = 0
    while True:
        try:
            n, v, t = winreg.EnumValue(k, i); snap[n] = v; i += 1
        except OSError:
            break
print("snapshot:", snap)
print()

def restore():
    # remove version subkey
    try: winreg.DeleteKey(winreg.HKEY_CURRENT_USER, VERKEY)
    except OSError: pass
    for n in list(snap):
        setv(INPROC, n, snap[n])

def reset_to_snapshot():
    restore()

results = []

def trial(name, fn):
    restore()
    fn()
    r = activate()
    results.append((name, r))
    print("[%s] %s" % (name, r))

# V0 baseline
trial("V0 baseline (as installed)", lambda: None)

# V1 RegAsm style: add version subkey 1.0.0.0 with its own values (keep parent too)
def v1():
    for n, v in snap.items():
        if n in ("Class", "Assembly", "RuntimeVersion", "CodeBase"):
            setv(VERKEY, n, v)
trial("V1 +InprocServer32\\1.0.0.0 subkey", v1)

# V2 partial assembly name (no version)
trial("V2 Assembly=partial name", lambda: setv(INPROC, "Assembly", "ImageZoomAddin"))

# V3 CodeBase -> short simple path copy
def v3():
    os.makedirs(os.path.dirname(SIMPLE), exist_ok=True)
    shutil.copy2(DLL, SIMPLE)
    setv(INPROC, "CodeBase", "file:///" + SIMPLE.replace("\\", "/"))
    setv(VERKEY, "CodeBase", "file:///" + SIMPLE.replace("\\", "/"))
trial("V3 CodeBase->short path copy", v3)

# V4 drop RuntimeVersion
trial("V4 no RuntimeVersion", lambda: (delv(INPROC, "RuntimeVersion"), delv(VERKEY, "RuntimeVersion")))

# V5 version subkey only, parent keeps only default+ThreadingModel
def v5():
    restore()
    for n, v in snap.items():
        if n in ("Class", "Assembly", "RuntimeVersion", "CodeBase"):
            setv(VERKEY, n, v)
    for n in ("Class", "Assembly", "RuntimeVersion", "CodeBase"):
        delv(INPROC, n)
trial("V5 RegAsm exact layout (values only in 1.0.0.0)", v5)

# V6 partial name + version subkey + no codebase
def v6():
    delv(INPROC, "CodeBase"); delv(VERKEY, "CodeBase")
trial("V6 no CodeBase at all (GAC lookup only)", v6)

restore()
print()
print("==== summary ====")
for n, r in results:
    print("%-46s %s" % (n, "PASS" if "OK" in r and "FAIL" not in r else "fail"))
