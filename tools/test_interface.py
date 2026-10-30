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
        slider_fields = re.findall(r'InstallSlider d, "txt([^"]+)"', module)
        names.update("scr" + field for field in slider_fields)
        names.update("val" + field for field in slider_fields)
        names.update(("btnOk", "btnCancel"))
        used = set(re.findall(r'Me.Controls\("([^"]+)"\)', form))
        self.assertFalse(used - names, f"Unknown form controls: {used - names}")
        for callback in ("SetupPreviewEvents", "StopPreviewEvents", "RefreshPreview"):
            self.assertRegex(form, r"(?im)^(?:Public |Private )Sub " + callback + r"\b")
        self.assertIn("As clsFacadePreview", form)

    def test_version_and_reserved_preview_variable(self):
        version = (ROOT / "VERSION").read_text(encoding="ascii").strip()
        form = (SRC / "frmFasad-code.txt").read_text(encoding="utf-8")
        self.assertIn(f'Генератор фасада - v{version}', form)
        # VBA Scale is a reserved statement, not a safe variable name.
        self.assertNotRegex(form, r"(?im)\bDim\s+scale\s+As\b")
        self.assertNotRegex(form, r"(?im)^\s*scale\s*=")
        archive = ROOT / "output" / f"FasadPanel_VBA_import_v{version}.zip"
        with zipfile.ZipFile(archive) as z:
            self.assertEqual(z.read("VERSION.txt").decode("ascii").strip(), version)

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
        self.assertIn("IsEntrance = (f = 0 And EntranceAt(si, mNSec, NEntrances, EntranceSections))", engine)
        self.assertIn("If IsEntrance(si, f) Then", engine)
        self.assertNotIn("If f = NPlinth And NEntrances > 0 Then", engine)
        self.assertIn("If NEntrances > NSections Then NEntrances = NSections", (SRC / "modMain.bas").read_text(encoding="utf-8"))

    def test_detail_controls_and_template_persistence(self):
        module = (SRC / "modFormRuntime.bas").read_text(encoding="utf-8")
        form = (SRC / "frmFasad-code.txt").read_text(encoding="utf-8")
        cls = (SRC / "clsFacadePreview.cls").read_text(encoding="utf-8")
        for name in ("txtSecTypes", "txtEntList", "chkAC", "chkPipes", "chkCracks", "chkBalconies", "cmbTemplate", "cmbCladding", "btnSaveTemplate"):
            self.assertIn('"' + name + '"', module)
            self.assertIn('"' + name + '"', form)
        self.assertIn("mOwner.SaveCurrentTemplate", cls)
        self.assertIn("SaveSetting \"FasadPanel\"", form)
        self.assertIn("LoadPresetFields spec", form)

    def test_selectors_and_textures_are_wired(self):
        form = (SRC / "frmFasad-code.txt").read_text(encoding="utf-8")
        module = (SRC / "modFormRuntime.bas").read_text(encoding="utf-8")
        cls = (SRC / "clsFacadePreview.cls").read_text(encoding="utf-8")
        engine = (SRC / "modEngine.bas").read_text(encoding="utf-8")
        main = (SRC / "modMain.bas").read_text(encoding="utf-8")
        for name in ("cmbSectionNo", "cmbSectionType", "cmbEntrance", "cmbSeed"):
            self.assertIn('"' + name + '"', module)
            self.assertIn('"' + name + '"', form)
            self.assertIn('"' + name + '"', cls)
        for old in ("txtSecTypes", "txtEntList", "txtSeed"):
            self.assertIn('"' + old + '"', form)
        self.assertIn("mOwner.SectionChoiceChanged", cls)
        self.assertIn("SetupSectionSelectors\n    SyncSliders\n    RefreshPreview", form)
        self.assertIn("SectionTypeAt(i, oldList)", form)
        self.assertIn("EntranceAt(i, count, ec, oldList)", form)
        self.assertIn("TextureType < 0 Or TextureType > 7", main)
        for kind in ("4", "5", "6", "7"):
            self.assertIn("TextureType = " + kind, engine)
        self.assertIn('texture = Me.Controls("cmbTexture").ListIndex', form)
        self.assertIn('PreviewNumber("txtTexA"', form)
        self.assertIn('PreviewNumber("txtTexB"', form)

    def test_sliders_grouping_and_layout(self):
        form = (SRC / "frmFasad-code.txt").read_text(encoding="utf-8")
        runtime = (SRC / "modFormRuntime.bas").read_text(encoding="utf-8")
        engine = (SRC / "modEngine.bas").read_text(encoding="utf-8")
        utils = (SRC / "modUtils.bas").read_text(encoding="utf-8")
        cls = (SRC / "clsFacadePreview.cls").read_text(encoding="utf-8")
        for field in ("Floors", "Plinth", "Sections", "Entrances", "FoundH", "TexA", "TexB", "TexGap"):
            self.assertIn('InstallSlider d, "txt' + field + '"', runtime)
            self.assertIn('"scr' + field + '"', form)
        self.assertIn("mOwner.SliderChanged mScroll.Name", cls)
        self.assertIn('bar.Orientation = 1', runtime)
        self.assertIn('gapSteps + 1', form)
        self.assertIn('Set facadeGroup = gFacadeShapes.Group', engine)
        self.assertEqual(utils.count('gColorShapes(layerIdx).Add sh'), 2)
        self.assertIn('Set colorGroup = gColorShapes(colorIndex).Group', engine)
        self.assertIn('gFacadeShapes.Add colorGroup', engine)
        self.assertLess(engine.index('Set colorGroup = gColorShapes(colorIndex).Group'),
                        engine.index('Set facadeGroup = gFacadeShapes.Group'))
        self.assertIn('If FourSides Then', engine)
        self.assertIn('For side = 0 To sides - 1', engine)
        self.assertIn('If oldSeed > 0 Then Seed = oldSeed', engine)
        self.assertIn('UBound(values) < UBound(fields) - 1', form)

    def test_continuous_texture_and_swatch(self):
        engine = (SRC / "modEngine.bas").read_text(encoding="utf-8")
        utils = (SRC / "modUtils.bas").read_text(encoding="utf-8")
        form = (SRC / "frmFasad-code.txt").read_text(encoding="utf-8")
        texture = engine[engine.index("Private Sub DrawTextures()"):engine.index("Private Sub DrawStoneTexture")]
        self.assertIn('DrawWallTexture', texture)
        self.assertNotIn('DrawTextureSection', texture)
        self.assertIn('x0 = mX0: x1 = mX1: ya = mFY(0): ybnd = mRoofBot', texture)
        self.assertIn('row = CLng(Int(y / rowH + 0.000001))', texture)
        self.assertIn('TextureHorizontalAllowed', texture)
        self.assertIn('TextureVerticalAllowed', texture)
        self.assertIn('Public Sub DrawTextureEdge', utils)
        self.assertIn('SegmentBox(x1, y1, dx, dy', utils)
        self.assertIn('RefreshTextureSwatch', form)
        self.assertIn('СХЕМА (не итоговые векторы)', form)

    def test_obstacle_checks_used_for_details(self):
        engine = (SRC / "modEngine.bas").read_text(encoding="utf-8")
        self.assertIn("If Not RectFree(x, y, x + w, y + h, GapMin + 0.3) Then Exit Sub", engine)
        self.assertIn("ACPosition = RectFree", engine)
        self.assertIn("If IsEntrance(si, f) Then", engine)
        self.assertIn("If ACOn And mSecType(i) = 2 Then", engine)
        self.assertIn("Private Sub DrawStoneTexture", engine)


if __name__ == "__main__":
    unittest.main()
