# -*- coding: utf-8 -*-
"""
MS-OVBA 2.4.1 Compression Container и сборка потоков VBA-проекта.
Формат: [MS-OVBA] https://learn.microsoft.com/openspecs/office_file_formats/ms-ovba
"""
import struct

SIGNATURE = 0x01
CHUNK_MAX = 4096
COMPRESSED_CHUNK_MAX = 4096


def _bit_count(value: int) -> int:
    """ceil(log2(value)), минимум 4 (MS-OVBA 2.4.1.3.19)."""
    n = 0
    while (1 << n) < value:
        n += 1
    return max(4, n)


def _compress_chunk(chunk: bytes) -> bytes:
    """Сжать один блок <=4096 байт (словарь локален для блока)."""
    out = bytearray()
    pos = 0
    n = len(chunk)
    while pos < n:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for token in range(8):
            if pos >= n:
                break
            # битность как у декомпрессора: от ТЕКУЩЕЙ позиции в чанке
            bit = _bit_count(max(1, pos))
            # длина - младшие `bit` бит, смещение - старшие (MS-OVBA 2.4.1.3.19)
            length_mask = 0xFFFF >> bit
            offset_mask = (~length_mask) & 0xFFFF
            max_len = min(length_mask + 3, n - pos)
            max_off = min(pos, offset_mask >> (16 - bit))
            best_off, best_len = 0, 0
            lo = pos - max_off
            for cand in range(lo, pos):
                max_here = min(max_len, n - pos)
                if max_here < 3:
                    continue
                off = pos - cand
                # сравнение с виртуальным выводом (корректно для overlap-копий)
                L = 0
                while L < max_here and chunk[cand + (L % off)] == chunk[pos + L]:
                    L += 1
                if L > best_len:
                    best_len, best_off = L, off
                    if L == max_here:
                        break
            if best_len >= 3:
                flags |= 1 << token
                token16 = ((best_off - 1) << (16 - bit)) | (best_len - 3)
                assert token16 <= 0xFFFF
                out += struct.pack("<H", token16)
                pos += best_len
            else:
                out.append(chunk[pos])
                pos += 1
        out[flag_pos] = flags
    if len(out) > CHUNK_MAX:
        # несжимаемый блок: флаг 0, сырые данные
        header = ((len(chunk) - 1) & 0x0FFF) | (0b011 << 12)
        return struct.pack("<H", header) + chunk
    header = ((len(out) - 1) & 0x0FFF) | (0b011 << 12) | (1 << 15)
    return struct.pack("<H", header) + bytes(out)


def compress_container(data: bytes) -> bytes:
    """Полный контейнер: сигнатура + набор чанков."""
    out = bytearray([SIGNATURE])
    for i in range(0, len(data), CHUNK_MAX):
        out += _compress_chunk(data[i:i + CHUNK_MAX])
    return bytes(out)


# ----------------------------------------------------------------------
# Сборка записей dir-потока [MS-OVBA 2.3.4.2]
# ----------------------------------------------------------------------
def rec(rid: int, payload: bytes) -> bytes:
    return struct.pack("<HI", rid, len(payload)) + payload


# Канонические Libid стандартных библиотек (вид у реальных VBA-проектов)
LIBID_VBA = (r"*\G{000204EF-0000-0000-C000-000000000046}#4.2#9#"
             r"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6.DLL"
             r"#Visual Basic For Applications")
LIBID_STDOLE = (r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
                r"C:\Windows\System32\stdole2.tlb#OLE Automation")


def build_dir(project_name: str, modules: list, cp: int = 1251) -> bytes:
    """modules: список кортежей (имя, тип): 0 procedural, 1 document.

    Раскладка записей выверена по реальному Corel .gms (см. внешние источники)
    и парсеру oletools: внутренние терминаторы 0x0000 не пишутся, их роль
    играют поля Reserved (0x0040/0x003D/0x003C/0x0032/0x0048).
    """
    d = bytearray()
    d += rec(0x0001, struct.pack("<I", 3))                 # PROJECTSYSKIND: 32-bit Windows
    d += rec(0x0002, struct.pack("<I", 0x409))             # PROJECTLCID en-US
    d += rec(0x0014, struct.pack("<I", 0x409))             # PROJECTLCIDINVOKE
    d += rec(0x0003, struct.pack("<H", cp))                # PROJECTCODEPAGE
    d += rec(0x0004, project_name.encode("cp1251"))        # PROJECTNAME

    # Строковые записи с MBCS+Unicode: поле Size покрывает ТОЛЬКО MBCS-часть,
    # Reserved и Unicode-размер следуют отдельными фиксированными полями.
    def mbcs_unicode_record(rid: int, reserved: int, text: str) -> bytes:
        mb = text.encode("cp1251")
        mu = text.encode("utf-16-le")
        return (struct.pack("<HI", rid, len(mb)) + mb
                + struct.pack("<HI", reserved, len(mu)) + mu)

    d += mbcs_unicode_record(0x0005, 0x0040, "")            # PROJECTDOCSTRING
    d += mbcs_unicode_record(0x0006, 0x003D, "")            # PROJECTHELPFILEPATH
    d += rec(0x0007, struct.pack("<I", 0))                 # PROJECTHELPCONTEXT
    d += rec(0x0008, struct.pack("<I", 0))                 # PROJECTLIBFLAGS
    # PROJECTVERSION: Id, Reserved (L, значение 4), VersionMajor (L), VersionMinor (H)
    d += struct.pack("<H", 0x0009) + struct.pack("<IIH", 4, 1, 0)
    d += mbcs_unicode_record(0x000C, 0x003C, "")            # PROJECTCONSTANTS

    def ref_name(name: str) -> bytes:
        return mbcs_unicode_record(0x0016, 0x003E, name)    # REFERENCENAME

    def ref_registered(libid: str) -> bytes:
        lib = libid.encode("ascii")
        # SizeOfLibid, Libid MBCS, Reserved1=0 (I), Reserved2=0 (H)
        payload = struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0)
        return rec(0x000D, payload)

    # ссылки: VBA + stdole (каждая = REFERENCENAME + REFERENCEREGISTERED)
    for nm, lib in (("VBA", LIBID_VBA), ("stdole", LIBID_STDOLE)):
        d += ref_name(nm)
        d += ref_registered(lib)

    # PROJECTMODULES: count (H), затем отдельной записью PROJECTCOOKIE
    d += rec(0x000F, struct.pack("<H", len(modules)))
    d += rec(0x0013, struct.pack("<H", 0xFFFF))

    for mname, mtype in modules:
        mb = mname.encode("ascii")
        mu = mname.encode("utf-16-le")
        d += rec(0x0019, mb)                                   # MODULENAME
        d += rec(0x0047, mu)                                   # MODULENAMEUNICODE
        d += mbcs_unicode_record(0x001A, 0x0032, mname)        # MODULESTREAMNAME
        d += mbcs_unicode_record(0x001C, 0x0048, "")           # MODULEDOCSTRING
        d += rec(0x0031, struct.pack("<I", 0))                 # MODULEOFFSET = 0
        d += rec(0x001E, struct.pack("<I", 0))                 # MODULEHELPCONTEXT
        d += rec(0x002C, struct.pack("<H", 0xFFFF))            # MODULECOOKIE
        d += rec(0x0022 if mtype == 1 else 0x0021, b"")        # MODULETYPE
        d += rec(0x002B, b"")                                  # MODULE terminator
    d += rec(0x0010, b"")                                      # PROJECTMODULES terminator
    return compress_container(bytes(d))


def build_projectwm(modules: list) -> bytes:
    """PROJECTwm: для каждого модуля MBCS-имя\\0 + UTF-16-имя\\0 (MS-OVBA 2.3.4.1)."""
    out = bytearray()
    for mname, _ in modules:
        out += mname.encode("ascii") + b"\x00"
        out += mname.encode("utf-16-le") + b"\x00\x00"
    return bytes(out)


def build_project_ini(project_name: str, modules: list, doc_module: str) -> bytes:
    lines = [
        'ID="{00000000-0000-0000-0000-000000000000}"',
        f"Document={doc_module}/&H00000000",
    ]
    for mname, mtype in modules:
        if mtype == 0:
            lines.append(f"Module={mname}")
    lines += [
        f'Name="{project_name}"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for mname, _ in modules:
        lines.append(f"{mname}=0, 0, 0, 0, C")
    lines.append("")
    return ("\r\n".join(lines)).encode("cp1251")


def build_vba_project_stream() -> bytes:
    # 32-байтный минимальный _VBA_PROJECT (как пишет GMSViewer):
    # Reserved1=0x61CC (LE CC 61), Version=0xFFFF, остальное нули;
    # performance-кэш отсутствует - хост пересоберёт его при открытии.
    return struct.pack("<HH", 0x61CC, 0xFFFF) + b"\x00" * 28
