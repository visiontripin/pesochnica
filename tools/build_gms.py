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

sys.path.insert(0, os.path.dirname(__file__))
from ovba import (build_dir, build_project_ini, build_projectwm,
                  build_vba_project_stream, compress_container)

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "gms-source")
BUILD = os.path.join(ROOT, "build_streams")
OUT_GMS = os.path.join(ROOT, "output", "FasadPanel.gms")

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

    manifest = {
        "project": PROJECT_NAME,
        "out": OUT_GMS,
        "build_dir": BUILD,
        "root_entries": ["PROJECT", "PROJECTwm"],
        "vba_entries": ["_VBA_PROJECT", "dir"] + [m[0] for m in mod_records],
        "entries": entries,
    }
    with open(os.path.join(BUILD, "manifest.json"), "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    print("потоки подготовлены:", len(entries))
    for e in entries:
        print(" ", e["file"], e["size"], "байт")


if __name__ == "__main__":
    main()
