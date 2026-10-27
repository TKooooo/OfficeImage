import struct, sys

path = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\BX\WorkBuddy\2026-10-26-11-31-15\imagezoom-addin\ImageZoomAddin.dll"
d = open(path, "rb").read()

def u16(o): return struct.unpack_from("<H", d, o)[0]
def u32(o): return struct.unpack_from("<I", d, o)[0]

print("file size        :", len(d))

# ---- DOS / PE ----
if d[0:2] != b"MZ":
    print("NOT MZ"); sys.exit(1)
e_lfanew = u32(0x3C)
if d[e_lfanew:e_lfanew+4] != b"PE\0\0":
    print("NOT PE at", e_lfanew); sys.exit(1)

coff = e_lfanew + 4
machine = u16(coff)
nsec = u16(coff + 2)
sizeopt = u16(coff + 16)
opt = coff + 20
magic = u16(opt)
pe32plus = (magic == 0x20B)
print("machine          : 0x%04X (%s)" % (machine, {0x14c: "x86", 0x8664: "x64", 0x0: "anycpu?"}.get(machine, "?")))
print("optional magic   : 0x%03X (%s)" % (magic, "PE32+" if pe32plus else "PE32"))
print("characteristics  : 0x%04X" % u16(coff + 18))

dirdata = opt + (112 if pe32plus else 96)
ndirs = u32(opt + (108 if pe32plus else 92))
print("num data dirs    :", ndirs)

cli_rva = u32(dirdata + 14 * 8)
cli_size = u32(dirdata + 14 * 8 + 4)
print("CLI header RVA   : 0x%X  size=%d %s" % (cli_rva, cli_size, "(OK)" if cli_rva else "*** MISSING: NOT A MANAGED ASSEMBLY ***"))
if not cli_rva:
    sys.exit(2)

# ---- sections ----
secs = []
so = opt + sizeopt
for i in range(nsec):
    b = so + i * 40
    name = d[b:b+8].rstrip(b"\0").decode("latin1")
    vsize = u32(b + 8); va = u32(b + 12)
    rsize = u32(b + 16); rptr = u32(b + 20)
    secs.append((name, va, vsize, rptr, rsize))
    print("section          : %-8s VA=0x%08X VSize=0x%X RawPtr=0x%X RawSize=0x%X" % (name, va, vsize, rptr, rsize))

def rva2off(rva):
    for name, va, vsize, rptr, rsize in secs:
        if va <= rva < va + max(vsize, rsize):
            return rptr + (rva - va)
    return None

cli = rva2off(cli_rva)
if cli is None:
    print("CLI header not mappable"); sys.exit(3)
md_rva = u32(cli + 8)
md_size = u32(cli + 12)
flags = u32(cli + 16)
print("CLI flags        : 0x%X (ILONLY=%s 32BITREQ=%s)" % (flags, bool(flags & 1), bool(flags & 2)))
print("metadata RVA     : 0x%X size=%d" % (md_rva, md_size))

md = rva2off(md_rva)
sig = u32(md)
if sig != 0x424A5342:
    print("bad metadata signature 0x%08X" % sig); sys.exit(4)
ver_len = u32(md + 12)
ver = d[md + 16: md + 16 + ver_len].rstrip(b"\0").decode("latin1")
print("runtime version  : %r" % ver)
p = md + 16 + ver_len
p += 2  # flags
nstreams = u16(p); p += 2
streams = {}
for i in range(nstreams):
    off = u32(p); size = u32(p + 4); p += 8
    end = d.index(b"\0", p)
    n = d[p:end].decode("latin1")
    p = end + 1
    p = (p + 3) & ~3
    streams[n] = (md + off, size)
    print("stream           : %-10s off=0x%X size=0x%X" % (n, off, size))

strings = streams.get("#Strings")
blob = streams.get("#Blob")
guid = streams.get("#GUID")
if "#~" in streams: tbl_off, tbl_size = streams["#~"]; tbl_name = "#~"
elif "#-" in streams: tbl_off, tbl_size = streams["#-"]; tbl_name = "#-"
else: print("no table stream"); sys.exit(5)

def rstr(idx):
    o = strings[0] + idx
    e = d.index(b"\0", o)
    return d[o:e].decode("utf8", "replace")

heap = d[tbl_off + 6]
str4 = bool(heap & 0x01); guid4 = bool(heap & 0x02); blob4 = bool(heap & 0x04)
valid = struct.unpack_from("<Q", d, tbl_off + 8)[0]
rows = {}
q = tbl_off + 24
for t in range(64):
    if valid >> t & 1:
        rows[t] = u32(q); q += 4
print("table stream     : %s  str4=%s guid4=%s blob4=%s" % (tbl_name, str4, guid4, blob4))
print("tables present   :", ", ".join("0x%02X:%d" % (k, v) for k, v in sorted(rows.items())))

S = 4 if str4 else 2
G = 4 if guid4 else 2
B = 4 if blob4 else 2

CODED = {
    "TypeDefOrRef": (2, [0x02, 0x01, 0x1B]),
    "HasConstant": (2, [0x04, 0x08, 0x17]),
    "HasCustomAttribute": (5, [0x06, 0x04, 0x01, 0x02, 0x08, 0x09, 0x0A, 0x00, 0x0E, 0x17, 0x14, 0x11, 0x1A, 0x1B, 0x20, 0x23, 0x26, 0x27, 0x28, 0x2A, 0x2C, 0x2B]),
    "HasFieldMarshal": (1, [0x04, 0x08]),
    "HasDeclSecurity": (2, [0x02, 0x06, 0x20]),
    "MemberRefParent": (3, [0x02, 0x01, 0x1A, 0x06, 0x1B]),
    "HasSemantics": (1, [0x14, 0x17]),
    "MethodDefOrRef": (1, [0x06, 0x0A]),
    "MemberForwarded": (1, [0x04, 0x06]),
    "Implementation": (2, [0x26, 0x23, 0x27]),
    "CustomAttributeType": (3, [0xFF, 0xFF, 0x06, 0x0A, 0xFF]),
    "ResolutionScope": (2, [0x00, 0x1A, 0x23, 0x01]),
    "TypeOrMethodDef": (1, [0x02, 0x06]),
}

SCHEMA = {
    0x00: ["u2", "str", "guid", "guid", "guid"],
    0x01: ["ResolutionScope", "str", "str"],
    0x02: ["u4", "str", "str", "TypeDefOrRef", "rid:0x04", "rid:0x06"],
    0x03: ["rid:0x04"],
    0x04: ["u2", "str", "blob"],
    0x05: ["rid:0x06"],
    0x06: ["u4", "u2", "u2", "str", "blob", "rid:0x08"],
    0x07: ["rid:0x08"],
    0x08: ["u2", "u2", "str"],
    0x09: ["rid:0x02", "TypeDefOrRef"],
    0x0A: ["MemberRefParent", "str", "blob"],
    0x0B: ["u2", "HasConstant", "blob"],
    0x0C: ["HasCustomAttribute", "CustomAttributeType", "blob"],
    0x0D: ["HasFieldMarshal", "blob"],
    0x0E: ["u2", "HasDeclSecurity", "blob"],
    0x0F: ["u2", "u4", "rid:0x02"],
    0x10: ["u4", "rid:0x04"],
    0x11: ["blob"],
    0x12: ["rid:0x02", "rid:0x14"],
    0x13: ["rid:0x14"],
    0x14: ["u2", "str", "TypeDefOrRef"],
    0x15: ["rid:0x02", "rid:0x17"],
    0x16: ["rid:0x17"],
    0x17: ["u2", "str", "blob"],
    0x18: ["u2", "rid:0x06", "HasSemantics"],
    0x19: ["rid:0x02", "MethodDefOrRef", "MethodDefOrRef"],
    0x1A: ["str"],
    0x1B: ["blob"],
    0x1C: ["u2", "MemberForwarded", "str", "rid:0x1A"],
    0x1D: ["u4", "rid:0x04"],
    0x1E: ["u4", "u4"],
    0x1F: ["u4"],
    0x20: ["u4", "u2", "u2", "u2", "u2", "u4", "blob", "str", "str"],
    0x21: ["u4"],
    0x22: ["u4", "u4", "u4"],
    0x23: ["u2", "u2", "u2", "u2", "u4", "blob", "str", "str", "blob"],
    0x24: ["u4", "rid:0x23"],
    0x25: ["u4", "u4", "u4", "rid:0x23"],
    0x26: ["u4", "str", "blob"],
    0x27: ["u4", "u4", "str", "str", "Implementation"],
    0x28: ["u4", "u4", "str", "Implementation"],
    0x29: ["rid:0x02", "rid:0x02"],
    0x2A: ["u2", "u2", "TypeOrMethodDef", "str"],
    0x2B: ["MethodDefOrRef", "blob"],
    0x2C: ["rid:0x2A", "TypeDefOrRef"],
}

def cidx_size(name):
    bits, tables = CODED[name]
    mx = 0
    for t in tables:
        if t != 0xFF:
            mx = max(mx, rows.get(t, 0))
    return 4 if mx >= (1 << (16 - bits)) else 2

def col_size(c):
    if c == "u2": return 2
    if c == "u4": return 4
    if c == "str": return S
    if c == "guid": return G
    if c == "blob": return B
    if c.startswith("rid:"): return 4 if rows.get(int(c[4:], 16), 0) >= 65536 else 2
    return cidx_size(c)

sizes = {t: sum(col_size(c) for c in cols) for t, cols in SCHEMA.items()}

# offset of each table start
offs = {}
cur = q
for t in sorted(rows.keys()):
    offs[t] = cur
    cur += rows[t] * sizes.get(t, 0)

def read_row(t, i):
    o = offs[t] + i * sizes[t]
    vals = []
    for c in SCHEMA[t]:
        s = col_size(c)
        vals.append(read_val(c, o, s))
        o += s
    return vals

def read_val(c, o, s):
    if c == "str": return rstr(u32(o) or u16(o)) if s == 4 else rstr(u16(o))
    if c == "blob":
        idx = u32(o) if s == 4 else u16(o)
        bo = blob[0] + idx
        b0 = d[bo]
        if b0 & 0x80 == 0: ln, hs = b0, 1
        elif b0 & 0xC0 == 0x80: ln, hs = ((b0 & 0x3F) << 8) | d[bo+1], 2
        else: ln, hs = ((b0 & 0x1F) << 24) | (d[bo+1] << 16) | (d[bo+2] << 8) | d[bo+3], 4
        return d[bo+hs: bo+hs+ln]
    if c == "guid": return u32(o) if s == 4 else u16(o)
    if c == "u2": return u16(o)
    if c == "u4": return u32(o)
    return u32(o) if s == 4 else u16(o)

print()
print("===== ASSEMBLY (manifest identity) =====")
if 0x20 in rows:
    for i in range(rows[0x20]):
        v = read_row(0x20, i)
        ver = "%d.%d.%d.%d" % (v[1], v[2], v[3], v[4])
        print("  Name        : %s" % v[7])
        print("  Version     : %s   <<< mscoree binds on THIS" % ver)
        print("  Culture     : %r" % v[8])
        print("  Flags       : 0x%X" % v[5])
        print("  PublicKey len: %d" % len(v[6]))
else:
    print("  *** NO Assembly TABLE -> NOT A MANIFEST MODULE ***")

print()
print("===== ASSEMBLY REFS (dependencies mscoree/fusion must resolve) =====")
if 0x23 in rows:
    for i in range(rows[0x23]):
        v = read_row(0x23, i)
        print("  %-34s %d.%d.%d.%d  flags=0x%X pkt=%d" % (v[6], v[0], v[1], v[2], v[3], v[4], len(v[5])))
else:
    print("  (none)")

print()
print("===== MODULE =====")
if 0x00 in rows:
    for i in range(rows[0x00]):
        v = read_row(0x00, i)
        print("  Name: %r  MvidIdx=%d" % (v[1], v[2]))

print()
print("===== TYPE DEFS (top-level, looking for Connect) =====")
if 0x02 in rows:
    for i in range(rows[0x02]):
        v = read_row(0x02, i)
        full = (v[2] + "." if v[2] else "") + v[1]
        fl = v[0]
        vis = {0: "NotPublic", 1: "Public", 2: "NestedPublic", 3: "NestedPrivate"}.get(fl & 7, "?")
        if "Connect" in full or "IDTExtensibility" in full or vis == "Public":
            print("  %-45s flags=0x%08X (%s)" % (full, fl, vis))
