# -*- coding: utf-8 -*-
"""检查官方 Extensibility PIA 的元数据，重点看 custom 参数是否带 MarshalAs。"""
import struct
import re
import sys

PIA = r"C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll"

d = open(PIA, "rb").read()
pe = struct.unpack_from("<I", d, 0x3C)[0]
nsec = struct.unpack_from("<H", d, pe + 6)[0]
optsz = struct.unpack_from("<H", d, pe + 20)[0]
opt = pe + 24
magic = struct.unpack_from("<H", d, opt)[0]
ddoff = opt + (96 if magic == 0x10b else 112)
cli_rva = struct.unpack_from("<I", d, ddoff + 14 * 8)[0]
sec = opt + optsz


def r2o(rva):
    for i in range(nsec):
        s = sec + i * 40
        vsz, vaddr, rsz, raddr = struct.unpack_from("<IIII", d, s + 8)
        if vaddr <= rva < vaddr + max(vsz, rsz):
            return raddr + (rva - vaddr)
    return None


cli = r2o(cli_rva)
md = r2o(struct.unpack_from("<I", d, cli + 8)[0])
vlen = struct.unpack_from("<I", d, md + 12)[0]
p = md + 16 + vlen + 2
n = struct.unpack_from("<H", d, p)[0]
p += 2
st = {}
for _ in range(n):
    off, size = struct.unpack_from("<II", d, p)
    p += 8
    e = d.index(b"\x00", p)
    nm = d[p:e].decode()
    p = e + 1
    p = (p + 3) & ~3
    st[nm] = (md + off, size)

so, ss = st["#Strings"]
heap = d[so:so + ss]
bo, bs = st["#Blob"]
blob = d[bo:bo + bs]

print("=== PIA #Strings ===")
for t in [b"SafeArraySubType", b"VarEnum", b"MarshalAsAttribute", b"UnmanagedType",
          b"IDTExtensibility2", b"ext_ConnectMode", b"ext_DisconnectMode"]:
    print("  %-20s %s" % (t.decode(), "FOUND" if t in heap else "missing"))

print()
print("=== #Blob（命名参数名存这里）===")
for t in [b"SafeArraySubType", b"VT_VARIANT", b"VT_ARRAY"]:
    print("  %-20s %s" % (t.decode(), "FOUND" if t in blob else "missing"))

print()
print("=== 全部类型名字符串 ===")
names = sorted(set(x.decode(errors="replace")
                   for x in re.findall(rb"[A-Za-z_][A-Za-z0-9_]{2,}", heap)))
print(", ".join(names))
