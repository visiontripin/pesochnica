# -*- coding: utf-8 -*-
"""
Сборка FasadPanel.gms:
  1. читает .bas из gms-source (UTF-8), кодирует в cp1251;
  2. сжимает исходники по MS-OVBA;
  3. формирует PROJECT, PROJECTwm, dir, _VBA_PROJECT, потоки модулей;
  4. пишет манифест для tools/cfb_wrap.js (OLE CFBF через npm cfb).
"""
import json
import os
import sys
import io
import struct

import olefile

sys.path.insert(0, os.path.dirname(__file__))
from ovba import (build_dir, build_project_ini, build_projectwm,
                  build_vba_project_stream, compress_container)

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "gms-source")
BUILD = os.path.join(ROOT, "build_streams")
OUT_GMS = os.path.join(ROOT, "output", "FasadPanel.gms")

TEMPLATE = os.path.join(ROOT, "ShipoYaschik_V2_01.gms")
HOST_PREFIX = "apc/The VBA Project/"
PROJECT_NAME = "FasadPanel"
# (файл, имя модуля/потока, тип: 1 документ, 0 модуль)
MODULES = [
    ("ThisMacroStorage.bas", "ThisMacroStorage", 1),
    ("modMain.bas", "modMain", 0),
    ("modUtils.bas", "modUtils", 0),
    ("modEngine.bas", "modEngine", 0),
    ("modFormRuntime.bas", "modFormRuntime", 0),
]


def main():
    os.makedirs(BUILD, exist_ok=True)
    os.makedirs(os.path.dirname(OUT_GMS), exist_ok=True)
    entries = []

    def dump(fname, data):
        path = os.path.join(BUILD, fname)
        with open(path, "wb") as fh:
            fh.write(data)
        entries.append({"file": fname, "size": len(data)})

    mod_records = []
    for fname, mname, mtype in MODULES:
        raw = open(os.path.join(SRC, fname), encoding="utf-8").read()
        raw = raw.replace("\r\n", "\n").replace("\n", "\r\n")
        source = raw.encode("cp1251")
        dump(mname, compress_container(source))
        mod_records.append((mname, mtype))

    dump("PROJECT", build_project_ini(PROJECT_NAME, mod_records, "ThisMacroStorage"))
    dump("PROJECTwm", build_projectwm(mod_records))
    dump("dir", build_dir(PROJECT_NAME, mod_records))
    dump("_VBA_PROJECT", build_vba_project_stream())

    # Служебная обёртка Corel из рабочего проекта пользователя. VBA-код,
    # форма и кэш шипоящика НЕ копируются. Путь владельца удаляется.
    with open(TEMPLATE, "rb") as fh:
        template = fh.read()
    if template[:4] != b"GMS\x01":
        raise ValueError("Ожидался шаблон Corel GMS: " + TEMPLATE)
    offset = struct.unpack_from("<I", template, 4)[0] + 8
    with olefile.OleFileIO(io.BytesIO(template[offset:])) as ole:
        host_entries = []
        for parts in ole.listdir(streams=True, storages=False):
            path = "/".join(parts)
            if not path.startswith(HOST_PREFIX) or path.startswith(HOST_PREFIX + "_VBA_Project/"):
                continue
            data = ole.openstream(parts).read()
            if path == HOST_PREFIX + "VBA Project Data":
                # Записи формата host: id (uint16), размер (uint32), payload.
                records = []
                seen = set()
                pos = 0
                while pos < len(data):
                    rid, length = struct.unpack_from("<HI", data, pos)
                    value = data[pos + 6:pos + 6 + length]
                    if len(value) != length:
                        raise ValueError("Повреждён VBA Project Data")
                    seen.add(rid)
                    if rid == 0x01b8:  # длина следующего пути в символах
                        value = struct.pack("<I", len("FasadPanel.gms"))
                    elif rid == 0x01c2:  # путь проекта, без личных данных
                        value = "FasadPanel.gms".encode("utf-16le")
                    records.append(struct.pack("<HI", rid, len(value)) + value)
                    pos += 6 + length
                if not {0x01b8, 0x01c2}.issubset(seen):
                    raise ValueError("В шаблоне не найдены поля пути проекта")
                data = b"".join(records)
            fname = "host_%02d.bin" % len(host_entries)
            dump(fname, data)
            host_entries.append({"path": path, "file": fname})

    manifest = {
        "project": PROJECT_NAME,
        "out": OUT_GMS,
        "build_dir": BUILD,
        "root_entries": ["PROJECT", "PROJECTwm"],
        "vba_entries": ["_VBA_PROJECT", "dir"] + [m[0] for m in mod_records],
        "entries": entries,
        "host_entries": host_entries,
    }
    with open(os.path.join(BUILD, "manifest.json"), "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    print("потоки подготовлены:", len(entries))
    for e in entries:
        print(" ", e["file"], e["size"], "байт")


if __name__ == "__main__":
    main()
