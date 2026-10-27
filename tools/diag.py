import subprocess, sys, os, winreg, glob

HERE = os.path.dirname(os.path.abspath(__file__))
ACT = os.path.join(HERE, "activate.py")
TESTCLSID = "{A1B2C3D4-1111-2222-3333-444455556666}"
KEY = "Software\\Classes\\CLSID\\" + TESTCLSID
INPROC = KEY + "\\InprocServer32"

def activate(clsid):
    r = subprocess.run([sys.executable, ACT, clsid], capture_output=True, text=True, timeout=120)
    return (r.stdout or r.stderr).strip()

def mkkey():
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, INPROC, 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "mscoree.dll")
        winreg.SetValueEx(k, "ThreadingModel", 0, winreg.REG_SZ, "Both")

def setv(name, val):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, INPROC, 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, name, 0, winreg.REG_SZ, val)

def delkey():
    for p in (INPROC, KEY):
        try: winreg.DeleteKey(winreg.HKEY_CURRENT_USER, p)
        except OSError: pass

print("===== T1: harness sanity, known NATIVE com class (Shell.Application) =====")
print(activate("{13709620-C279-11CE-A49E-444553540000}"))

print()
print("===== T2: mscoree + a guaranteed-present GAC assembly (System) =====")
delkey(); mkkey()
setv("Class", "System.Object")
setv("Assembly", "System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
setv("RuntimeVersion", "v4.0.30319")
print(activate(TESTCLSID))

print()
print("===== T3: same but WITH explicit CodeBase to GAC file =====")
gac = glob.glob(r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System\v4.0_4.0.0.0__b77a5c561934e089\System.dll")
print("gac file:", gac)
if gac:
    setv("CodeBase", "file:///" + gac[0].replace("\\", "/"))
    print(activate(TESTCLSID))

print()
print("===== T4: our DLL but class intentionally WRONG (proves binding vs type lookup) =====")
setv("Class", "ImageZoomAddin.NoSuchType")
setv("Assembly", "ImageZoomAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null")
setv("CodeBase", "file:///C:/Users/BX/WorkBuddy/2026-10-26-11-31-15/imagezoom-addin/ImageZoomAddin.dll")
print(activate(TESTCLSID))

print()
print("===== T5: our DLL, correct class, via TEST clsid =====")
setv("Class", "ImageZoomAddin.Connect")
print(activate(TESTCLSID))

delkey()
print()
print("cleaned up test clsid")
