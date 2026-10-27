import subprocess, sys, os, winreg

HERE = os.path.dirname(os.path.abspath(__file__))
ACT = os.path.join(HERE, "activate.py")

CLSID = "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
DLL = r"C:\Users\BX\WorkBuddy\2026-10-26-11-31-15\imagezoom-addin\ImageZoomAddin.dll"
URL = "file:///" + DLL.replace("\\", "/")
ASM = "ImageZoomAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
CLS = "ImageZoomAddin.Connect"

def activate(clsid):
    r = subprocess.run([sys.executable, ACT, clsid], capture_output=True, text=True, timeout=180)
    return (r.stdout or r.stderr).strip()

def write_under(root, base):
    """base e.g. 'Software\\Classes\\CLSID\\{guid}'"""
    ck = base
    with winreg.CreateKeyEx(root, ck, 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "Image Zoom Preview Add-in")
    ip = base + "\\InprocServer32"
    with winreg.CreateKeyEx(root, ip, 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "mscoree.dll")
        winreg.SetValueEx(k, "ThreadingModel", 0, winreg.REG_SZ, "Both")
        winreg.SetValueEx(k, "Class", 0, winreg.REG_SZ, CLS)
        winreg.SetValueEx(k, "Assembly", 0, winreg.REG_SZ, ASM)
        winreg.SetValueEx(k, "RuntimeVersion", 0, winreg.REG_SZ, "v4.0.30319")
        winreg.SetValueEx(k, "CodeBase", 0, winreg.REG_SZ, URL)
    with winreg.CreateKeyEx(root, base + "\\ProgId", 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "ImageZoomAddin.Connect")
    # .NET implemented category (what RegAsm writes)
    with winreg.CreateKeyEx(root, base + "\\Implemented Categories\\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}",
                            0, winreg.KEY_ALL_ACCESS) as k:
        pass
    # ProgID + assembly registry view
    with winreg.CreateKeyEx(root, "Software\\Classes\\ImageZoomAddin.Connect\\CLSID", 0, winreg.KEY_ALL_ACCESS) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, CLSID)

def try_activate(label):
    print("  %-42s %s" % (label, activate(CLSID)))

print("===== step 1: HKCU only (baseline, already installed) =====")
try_activate("HKCU registration")

print()
print("===== step 2: add HKLM 64-bit (SOFTWARE\\Classes) =====")
write_under(winreg.HKEY_LOCAL_MACHINE, "Software\\Classes\\CLSID\\" + CLSID)
try_activate("HKLM 64-bit")

print()
print("===== step 3: also HKLM Wow6432Node =====")
write_under(winreg.HKEY_LOCAL_MACHINE, "Software\\Classes\\Wow6432Node\\CLSID\\" + CLSID)
try_activate("HKLM 64 + Wow6432Node")

print()
print("===== done; HKLM entries kept for inspection =====")
