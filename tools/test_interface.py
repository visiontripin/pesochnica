"""Static consistency checks; CorelDRAW/VBA runtime is not available here."""
from pathlib import Path
import re
import unittest
import zipfile

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "gms-source"


class InterfaceChecks(unittest.TestCase):
    def test_form_control_names_exist(self):
        module = (SRC / "modFormRuntime.bas").read_text(encoding="utf-8")
        form = (SRC / "frmFasad-code.txt").read_text(encoding="utf-8")
        names = set(re.findall(r'Add(?:Label|Text|Combo|Check) d, "([^"]+)"', module))
        names.update(re.findall(r'Controls.Add\("Forms\.\w+\.1", "([^"]+)"\)', module))
        names.update(("btnOk", "btnCancel"))
        used = set(re.findall(r'Me.Controls\("([^"]+)"\)', form))
        self.assertFalse(used - names, f"Unknown form controls: {used - names}")
        for callback in ("SetupPreviewEvents", "StopPreviewEvents", "RefreshPreview"):
            self.assertRegex(form, r"(?im)^(?:Public |Private )Sub " + callback + r"\b")
        self.assertIn("As clsFacadePreview", form)

    def test_preview_class_wired(self):
        cls = (SRC / "clsFacadePreview.cls").read_text(encoding="utf-8")
        for name in ("mText_Change", "mCombo_Change", "mCheck_Click"):
            self.assertIn("Private Sub " + name + "()", cls)
        self.assertIn("Public Sub Detach()", cls)
        archive = ROOT / "output" / "FasadPanel_VBA_import.zip"
        with zipfile.ZipFile(archive) as z:
            self.assertIn("clsFacadePreview.cls", z.namelist())
            self.assertIn("frmFasad-code.txt", z.namelist())
            self.assertEqual(z.read("clsFacadePreview.cls").decode("cp1251").replace("\r\n", "\n"), cls)

    def test_entrance_only_on_lowest_floor(self):
        engine = (SRC / "modEngine.bas").read_text(encoding="utf-8")
        self.assertIn("IsEntrance = (f = 0 And si > mNSec - NEntrances)", engine)
        self.assertIn("If IsEntrance(si, f) Then", engine)
        self.assertNotIn("If f = NPlinth And NEntrances > 0 Then", engine)
        self.assertIn("If NEntrances > NSections Then NEntrances = NSections", (SRC / "modMain.bas").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
