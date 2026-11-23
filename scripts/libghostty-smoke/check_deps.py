#!/usr/bin/env python3
"""Fail if a native library imports anything outside a small allowlist.

Usage: python3 -I check_deps.py pe  <file.dll>
       python3 -I check_deps.py elf <file.so>

Stdlib only, no external tools. PE: import + delay-load directories.
ELF64: DT_NEEDED entries of the dynamic section.
"""
import fnmatch
import struct
import sys

PE_ALLOWED = ["kernel32.dll", "ntdll.dll", "advapi32.dll", "ucrtbase.dll", "msvcrt.dll",
              "bcrypt.dll", "ws2_32.dll", "user32.dll", "api-ms-win-*.dll"]
ELF_ALLOWED = ["libc.so.6", "libm.so.6", "libpthread.so.0", "libdl.so.2", "librt.so.1",
               "ld-linux-x86-64.so.2", "ld-linux-aarch64.so.1"]


def cstr(data, off):
    if off < 0 or off >= len(data):
        raise ValueError("string offset out of range")
    end = data.index(b"\0", off)
    return data[off:end].decode("ascii", "replace")


def parse_pe(data):
    if data[:2] != b"MZ":
        raise ValueError("not a PE file (no MZ)")
    pe = struct.unpack_from("<I", data, 0x3C)[0]
    if data[pe:pe + 4] != b"PE\0\0":
        raise ValueError("bad PE signature")
    nsec, = struct.unpack_from("<H", data, pe + 6)
    optsz, = struct.unpack_from("<H", data, pe + 20)
    opt = pe + 24
    magic, = struct.unpack_from("<H", data, opt)
    if magic == 0x20B:
        ddoff = opt + 112
    elif magic == 0x10B:
        ddoff = opt + 96
    else:
        raise ValueError("unknown optional header magic 0x%x" % magic)
    ndd, = struct.unpack_from("<I", data, ddoff - 4)
    dirs = [struct.unpack_from("<II", data, ddoff + 8 * i) for i in range(min(ndd, 16))]
    secs = []
    so = opt + optsz
    for i in range(nsec):
        vsize, va, rawsz, rawptr = struct.unpack_from("<IIII", data, so + 40 * i + 8)
        secs.append((va, max(vsize, rawsz), rawptr))

    def rva(r):
        for va, sz, ptr in secs:
            if va <= r < va + sz:
                return r - va + ptr
        raise ValueError("RVA 0x%x not in any section" % r)

    names = []
    if len(dirs) > 1 and dirs[1][0]:  # import table, 20-byte descriptors
        o = rva(dirs[1][0])
        while True:
            oft, _, _, name, ft = struct.unpack_from("<IIIII", data, o)
            if not (oft or name or ft):
                break
            names.append(cstr(data, rva(name)))
            o += 20
    if len(dirs) > 13 and dirs[13][0]:  # delay-load, 32-byte descriptors
        o = rva(dirs[13][0])
        while True:
            attrs, name = struct.unpack_from("<II", data, o)
            rest = struct.unpack_from("<6I", data, o + 8)
            if not (attrs or name or any(rest)):
                break
            # attrs bit 0 set => RVAs (modern); otherwise VAs (legacy, unsupported)
            names.append(cstr(data, rva(name)) + " (delay-load)")
            o += 32
    return names


def parse_elf(data):
    if data[:4] != b"\x7fELF":
        raise ValueError("not an ELF file")
    if data[4] != 2 or data[5] != 1:
        raise ValueError("only little-endian ELF64 is supported")
    phoff, = struct.unpack_from("<Q", data, 0x20)
    phentsz, phnum = struct.unpack_from("<HH", data, 0x36)
    loads, dyn = [], None
    for i in range(phnum):
        o = phoff + i * phentsz
        ptype, _, poff, pvaddr, _, pfilesz, _, _ = struct.unpack_from("<IIQQQQQQ", data, o)
        if ptype == 1:
            loads.append((pvaddr, pfilesz, poff))
        elif ptype == 2:
            dyn = (poff, pfilesz)
    if dyn is None:
        return []

    def off(vaddr):
        for va, sz, po in loads:
            if va <= vaddr < va + sz:
                return vaddr - va + po
        raise ValueError("vaddr 0x%x not in a PT_LOAD" % vaddr)

    needed, strtab = [], None
    for o in range(dyn[0], dyn[0] + dyn[1], 16):
        tag, val = struct.unpack_from("<qQ", data, o)
        if tag == 0:
            break
        if tag == 1:
            needed.append(val)
        elif tag == 5:
            strtab = val
    if strtab is None:
        return []
    base = off(strtab)
    return [cstr(data, base + n) for n in needed]


def main(argv):
    if len(argv) != 3 or argv[1] not in ("pe", "elf"):
        print(__doc__)
        return 2
    with open(argv[2], "rb") as f:
        data = f.read()
    try:
        imports = parse_pe(data) if argv[1] == "pe" else parse_elf(data)
    except (ValueError, struct.error) as e:
        print("error: cannot parse %s: %s" % (argv[2], e))
        return 1
    allowed = PE_ALLOWED if argv[1] == "pe" else ELF_ALLOWED
    bad = []
    print("%s imports (%d):" % (argv[2], len(imports)))
    for n in imports:
        base = n.split(" ")[0].lower()
        ok = any(fnmatch.fnmatchcase(base, p.lower()) for p in allowed)
        print("  %s %s" % ("ok " if ok else "BAD", n))
        if not ok:
            bad.append(n)
    if bad:
        print("error: disallowed dependencies in %s: %s" % (argv[2], ", ".join(bad)))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
