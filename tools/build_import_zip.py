#!/usr/bin/env python3
"""Prepare VBA import files for the Windows-1251 VBA editor, not a GMS."""
from pathlib import Path
from shutil import copyfile
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parent.parent
VERSION = (ROOT / "VERSION").read_text(encoding="ascii").strip()
NAMES = ("modMain", "modUtils", "modEngine", "modFormRuntime")
OUT = ROOT / "output" / f"FasadPanel_VBA_import_v{VERSION}.zip"
LATEST = ROOT / "output" / "FasadPanel_VBA_import.zip"


def main():
    OUT.parent.mkdir(exist_ok=True)
    with ZipFile(OUT, "w", ZIP_DEFLATED) as z:
        for name in NAMES:
            source = (ROOT / "gms-source" / (name + ".bas")).read_text(encoding="utf-8")
            data = source.replace("\r\n", "\n").replace("\n", "\r\n").encode("cp1251")
            z.writestr(name + ".bas", data)
        cls = (ROOT / "gms-source" / "clsFacadePreview.cls").read_text(encoding="utf-8")
        z.writestr("clsFacadePreview.cls", cls.replace("\n", "\r\n").encode("cp1251"))
        form = (ROOT / "gms-source" / "frmFasad-code.txt").read_text(encoding="utf-8")
        z.writestr("frmFasad-code.txt", form.replace("\n", "\r\n").encode("utf-8-sig"))
        z.writestr("VERSION.txt", VERSION + "\r\n")
    with ZipFile(OUT) as z:
        assert z.testzip() is None
        for name in NAMES:
            source = (ROOT / "gms-source" / (name + ".bas")).read_text(encoding="utf-8")
            assert z.read(name + ".bas").decode("cp1251").replace("\r\n", "\n") == source
    copyfile(OUT, LATEST)
    print(f"{OUT}: {len(NAMES)} модулей, round-trip OK; latest: {LATEST}")


if __name__ == "__main__":
    main()
