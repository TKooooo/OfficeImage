import winreg, subprocess, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
ACT = os.path.join(HERE, "activate.py")

def activate(clsid):
    r = subprocess.run([sys.executable, ACT, clsid], capture_output=True, text=True, timeout=120)
    return (r.stdout or r.stderr).strip()

def scan(root, label, limit=6):
    found = []
    path = "Software\\Classes\\CLSID"
    try:
        h = winreg.OpenKey(root, path)
    except OSError as e:
        print(label, "open failed", e); return found
    i = 0
    while True:
        try:
            sub = winreg.EnumKey(h, i); i += 1
        except OSError:
            break
        try:
            ip = winreg.OpenKey(h, sub + "\\InprocServer32")
            dflt, _ = winreg.QueryValueEx(ip, "")
        except OSError:
            continue
        if dflt and "mscoree" in dflt.lower():
            try:
                cls, _ = winreg.QueryValueEx(ip, "Class")
            except OSError:
                cls = "?"
            try:
                asm, _ = winreg.QueryValueEx(ip, "Assembly")
            except OSError:
                asm = "?"
            found.append((sub, cls, asm))
            if len(found) >= limit:
                break
    print("%s : %d mscoree-registered class(es) found" % (label, len(found)))
    return found

for root, label in ((winreg.HKEY_LOCAL_MACHINE, "HKLM"), (winreg.HKEY_CURRENT_USER, "HKCU")):
    items = scan(root, label)
    for clsid, cls, asm in items:
        print("  %s" % clsid)
        print("     Class    : %s" % cls)
        print("     Assembly : %s" % asm)
        print("     ACTIVATE : %s" % activate(clsid))
    print()
