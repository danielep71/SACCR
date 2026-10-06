"""Offline negative fixtures for the SACCR source gate."""
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

import check_source
import check_vba_public_api as public_api

ROOT = Path(__file__).resolve().parents[1]
CHANGELOG = "# Changelog\n\n## [Unreleased]\n\n[Unreleased]: https://example.invalid\n"
MODULE = 'Attribute VB_Name = "M_Test"\r\nOption Explicit\r\n\r\nPublic Sub Run()\r\nEnd Sub\r\n'


class SourceGateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.git("init", "-q")
        shutil.copy(ROOT / ".gitattributes", self.root / ".gitattributes")
        self.write("CHANGELOG.md", CHANGELOG)

    def git(self, *args):
        return subprocess.run(["git", "-C", str(self.root), *args], check=True,
                              capture_output=True, text=True).stdout

    def write(self, path, text):
        target = self.root / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(text.encode("cp1252"))

    def findings(self):
        self.git("add", "-A")
        return check_source.run_check(self.root)["findings"]

    def test_clean_component_passes_and_is_stored_lf(self):
        self.write("src/modules/M_Test.bas", MODULE)
        self.assertEqual(self.findings(), [])
        self.assertEqual(self.git("ls-files", "--eol", "src/modules/M_Test.bas").split()[0], "i/lf")

    def test_vb_name_must_match_filename(self):
        self.write("src/modules/M_Other.bas", MODULE)
        self.assertIn("src/modules/M_Other.bas: VB_Name must match the filename exactly", self.findings())

    def test_duplicate_component_name(self):
        self.write("src/modules/M_Test.bas", MODULE)
        self.write("tests/M_Test.bas", MODULE)
        self.assertTrue(any("duplicate component name" in f for f in self.findings()))

    def test_option_explicit_in_comment_does_not_count(self):
        self.write("src/modules/M_Test.bas", MODULE.replace("Option Explicit", "' Option Explicit"))
        self.assertIn("src/modules/M_Test.bas: missing Option Explicit", self.findings())

    def test_form_requires_tracked_frx_within_bounds(self):
        form = 'VERSION 5.00\r\nBegin {X} F_Test\r\n   OleObjectBlob = "F_Test.frx":0010\r\nEnd\r\n'
        form += MODULE.replace("M_Test", "F_Test")
        self.write("src/forms/F_Test.frm", form)
        self.assertIn("src/forms/F_Test.frm: missing or unsafe form resource F_Test.frx", self.findings())
        (self.root / "src/forms/F_Test.frx").write_bytes(b"\0" * 8)
        self.assertIn("src/forms/F_Test.frm: resource offset is outside F_Test.frx", self.findings())
        (self.root / "src/forms/F_Test.frx").write_bytes(b"\0" * 32)
        self.assertEqual(self.findings(), [])

    def test_vba_outside_documented_locations_is_rejected(self):
        self.write("src/M_Test.bas", MODULE)
        self.write("misc/Old.frx", "x")
        found = self.findings()
        self.assertIn("src/M_Test.bas: VBA component outside the documented source locations", found)
        self.assertIn("misc/Old.frx: VBA component outside the documented source locations", found)

    def test_core_module_requires_option_private_module(self):
        self.write("src/core/M_Test.bas", MODULE)
        self.assertIn("src/core/M_Test.bas: core modules must declare Option Private Module", self.findings())
        self.write("src/core/M_Test.bas", MODULE.replace("Option Explicit", "Option Explicit\r\nOption Private Module"))
        self.assertEqual(self.findings(), [])

    def test_crlf_in_index_is_rejected(self):
        (self.root / ".gitattributes").write_text("* -text\n")
        self.write("notes.txt", "a\r\nb\r\n")
        found = self.findings()
        self.assertIn("notes.txt: stored with CRLF in Git; renormalize to LF", found)

    def test_lone_cr_in_declared_text_is_rejected(self):
        self.write("notes.md", "a\rb\r")
        (self.root / "blob.dat").write_bytes(b"a\rb\r")
        found = self.findings()
        self.assertIn("notes.md: declared text but Git classifies the blob as non-text (lone CR?)", found)
        self.assertFalse(any(f.startswith("blob.dat") for f in found))

    def test_vba_without_crlf_checkout_attribute_is_rejected(self):
        (self.root / ".gitattributes").write_text("* text=auto eol=lf\n")
        self.write("src/modules/M_Test.bas", MODULE.replace("\r\n", "\n"))
        self.assertIn("src/modules/M_Test.bas: .gitattributes must check VBA source out as CRLF", self.findings())

    def test_changelog_requires_unreleased_first(self):
        self.write("CHANGELOG.md", "## [0.0.1] - 2026-01-01\n\n## [Unreleased]\n")
        found = self.findings()
        self.assertIn("CHANGELOG.md: the first version heading must be '## [Unreleased]'", found)
        self.assertIn("CHANGELOG.md: missing link reference for [0.0.1]", found)

    def test_changelog_release_heading_format(self):
        self.write("CHANGELOG.md", CHANGELOG + "\n## [v0.0.1] 2026-01-01\n\n[v0.0.1]: x\n")
        self.assertTrue(any("release heading must be" in f for f in self.findings()))

    def test_changelog_release_date_must_exist(self):
        for bad in ("2026-02-31", "2026-99-99"):
            self.write("CHANGELOG.md", CHANGELOG + f"\n## [0.0.1] - {bad}\n\n[0.0.1]: x\n")
            self.assertIn(f"CHANGELOG.md: [0.0.1] date {bad} is not a calendar date", self.findings())
        self.write("CHANGELOG.md", CHANGELOG + "\n## [0.0.1] - 2028-02-29\n\n[0.0.1]: x\n")
        self.assertEqual(self.findings(), [])

    def test_missing_changelog(self):
        (self.root / "CHANGELOG.md").unlink()
        self.assertIn("CHANGELOG.md is missing", self.findings())


class PublicApiRoleTests(unittest.TestCase):
    FACADE = ('Attribute VB_Name = "Facade"\nOption Explicit\n'
              'Public Function Echo(ByVal value As Long) As Long\nEnd Function\n')
    MANIFEST = ["Facade\tFunction\tEcho",
                "# SIG\tFacade\tFunction\tEcho\tPublic Function Echo(ByVal value As Long) As Long"]

    def test_facade_declaration_listed_passes(self):
        self.assertEqual(public_api.fixture(self.FACADE, self.MANIFEST)["status"], "pass")

    def test_unlisted_facade_declaration_fails(self):
        self.assertEqual(public_api.fixture(self.FACADE, [])["status"], "fail")

    def test_core_declaration_cannot_be_listed(self):
        manifest = self.MANIFEST + ["Core\tFunction\tInternalOnly"]
        self.assertEqual(public_api.fixture(self.FACADE, manifest)["status"], "fail")

    def test_roles_follow_repository_structure(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            for path in ("src/modules/A.bas", "src/core/B.bas", "src/classes/C.cls",
                         "tests/modules/D.bas", "examples/modules/E.bas", "misc/F.bas"):
                (root / path).parent.mkdir(parents=True, exist_ok=True)
                (root / path).write_text("x")
            subprocess.run(["git", "-C", str(root), "add", "-A"], check=True)
            self.assertEqual(public_api.component_roles(root), {
                "src/modules/A.bas": "public", "src/core/B.bas": "internal",
                "src/classes/C.cls": "internal", "tests/modules/D.bas": "test",
                "examples/modules/E.bas": "example"})


if __name__ == "__main__":
    unittest.main()
