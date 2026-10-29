# -*- coding: utf-8 -*-
"""
Валидация собранного output/FasadPanel.gms:
  1. 18-байтный GMS-заголовок (выверен по реальному Corel .gms / GMSViewer);
  2. OLE/CFBF структура: Corel host-обёртка и вложенный VBA-проект;
  3. MS-OVBA: байтовый round-trip каждого модуля против gms-source/*.bas;
  4. полный разбор парсером oletools.olevba (без warning'ов, cp1251, 5 модулей).
Выход: код 0 при успехе, 1 при ошибке.
"""
import logging
import os
import struct
import sys
import tempfile

import olefile
from oletools.olevba import VBA_Parser, decompress_stream

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
GMS = os.path.join(ROOT, "output", "FasadPanel.gms")
SRC = os.path.join(ROOT, "gms-source")

GMS_HEADER = bytes([0x47, 0x4D, 0x53, 0x01, 0x0A, 0x00, 0x00, 0x00,
                    0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
                    0x00, 0x00])
MODULES = [
    ("ThisMacroStorage", "ThisMacroStorage.bas", True),
    ("modMain", "modMain.bas", False),
    ("modUtils", "modUtils.bas", False),
    ("modEngine", "modEngine.bas", False),
    ("modFormRuntime", "modFormRuntime.bas", False),
]
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def fail(msg: str) -> None:
    print("FAIL:", msg)
    sys.exit(1)


def main() -> None:
    raw = open(GMS, "rb").read()
    print(f"файл: {GMS} ({len(raw)} байт)")

    # 1. GMS-заголовок
    if raw[:18] != GMS_HEADER:
        fail(f"заголовок GMS не совпал: {raw[:18].hex(' ')}")
    ole_offset = struct.unpack_from("<I", raw, 4)[0] + 8
    if ole_offset != 18 or raw[18:26] != OLE_MAGIC:
        fail("смещение OLE или сигнатура неверны")
    print("[1] GMS-заголовок OK (OLE с байта 18)")

    ole = olefile.OleFileIO(raw[18:])
    entries = {"/".join(e) for e in ole.listdir(streams=True, storages=True)}
    base = "apc/The VBA Project/_VBA_Project/"
    host = "apc/The VBA Project/"
    required = {base + x for x in ("PROJECT", "PROJECTwm", "VBA", "VBA/dir", "VBA/_VBA_PROJECT")}
    required |= {base + f"VBA/{m[0]}" for m in MODULES}
    required |= {host + x for x in ("Host Project Items", "VBA Project Data", "VBA Project Signature", "Host Project Item Names/Host Project Item List Data", "Host Project Item Names/ThisMacroStorage/Project Item Data") }
    missing = required - entries
    if missing:
        fail(f"отсутствуют потоки: {sorted(missing)}")
    data = ole.openstream(host + "VBA Project Data").read()
    if "FasadPanel.gms".encode("utf-16le") not in data or "Bjalk".encode("utf-16le") in data:
        fail("неверный или неочищенный путь проекта в host-данных")
    if any("Sh33tJ5" in e for e in entries):
        fail("присутствует служебный seed-поток cfb-библиотеки")
    print(f"[2] OLE-структура OK ({len(entries)} записей)")

    # 3. байтовый round-trip исходников
    for mname, fname, _ in MODULES:
        comp = ole.openstream(base + f"VBA/{mname}").read()
        got = decompress_stream(comp)
        orig = (open(os.path.join(SRC, fname), encoding="utf-8").read()
                .replace("\r\n", "\n").replace("\n", "\r\n").encode("cp1251"))
        if got != orig:
            fail(f"round-trip {mname}: {len(got)} != {len(orig)}")
    ole.close()
    print("[3] байтовый round-trip 5 модулей OK (cp1251, CRLF)")

    # 4. полный разбор olevba; ловим все warning'и парсера
    records = []
    handler = logging.Handler()
    handler.emit = lambda r: records.append(r)
    logging.getLogger("olevba").addHandler(handler)
    logging.getLogger("olevba").setLevel(logging.WARNING)

    with tempfile.NamedTemporaryFile(suffix=".ole", delete=False) as tmp:
        tmp.write(raw[18:])
        tmp_path = tmp.name
    try:
        vp = VBA_Parser(tmp_path)
        got = {n.split(".")[0]: c for (_, _, n, c) in vp.extract_macros()}
        vp.close()
    finally:
        os.unlink(tmp_path)
    warnings = [r for r in records if r.levelno >= logging.WARNING]
    if warnings:
        for r in warnings:
            print("  warning:", r.getMessage())
        fail("olevba выдал предупреждения о структуре dir/PROJECT")
    if set(got) != {m[0] for m in MODULES}:
        fail(f"набор модулей olevba: {sorted(got)}")
    for mname, fname, is_doc in MODULES:
        orig = open(os.path.join(SRC, fname), encoding="utf-8").read()
        if got[mname].replace("\r\n", "\n").strip("\n") != \
                orig.replace("\r\n", "\n").strip("\n"):
            fail(f"текст {mname} после распаковки отличается")
        if not any("А" <= ch <= "я" or ch in "ёЁ" for ch in got[mname]):
            fail(f"кириллица не читается в {mname} (кодовая страница?)")
    print("[4] olevba: полный разбор OK, 5 модулей, cp1251 без предупреждений")
    print("ИТОГ: структурная валидация пройдена; запуск в CorelDRAW не проверен")


if __name__ == "__main__":
    main()
