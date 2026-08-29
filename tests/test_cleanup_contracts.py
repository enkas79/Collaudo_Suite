from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class CleanupContractTests(unittest.TestCase):
    def test_obsolete_api_integration_is_removed(self):
        self.assertFalse((ROOT / "collaudo_suite" / "jarvis").exists())
        requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8").lower()
        self.assertNotIn("requests-negotiate-sspi", requirements)
        build = (ROOT / "build_windows.bat").read_text(encoding="utf-8").lower()
        self.assertNotIn("api_config.json", build)
        self.assertNotIn("requests_negotiate_sspi", build)

    def test_map_period_control_is_wired_and_persisted(self):
        checklist_source = (ROOT / "collaudo_suite" / "checklist" / "app.py").read_text(encoding="utf-8")
        self.assertIn('self.map_period_combo.addItem("1 mese", 1)', checklist_source)
        self.assertIn('self.map_period_combo.addItem("3 mesi", 3)', checklist_source)
        self.assertIn('self.map_period_combo.addItem("6 mesi", 6)', checklist_source)
        self.assertIn('self.map_period_combo.addItem("1 anno", 12)', checklist_source)
        self.assertIn('reference_date=self._collaudo_reference_date()', checklist_source)
        self.assertIn('months_back=self._selected_map_period_months()', checklist_source)
        self.assertIn('"map_period_months": self._selected_map_period_months()', checklist_source)

    def test_save_and_save_as_commands_exist(self):
        checklist_source = (ROOT / "collaudo_suite" / "checklist" / "app.py").read_text(encoding="utf-8")
        main_source = (ROOT / "collaudo_suite" / "main.py").read_text(encoding="utf-8")
        self.assertIn("def save_work(self)", checklist_source)
        self.assertIn("def save_work_as(self)", checklist_source)
        self.assertIn("callback=self.checklist.save_work", main_source)
        self.assertIn("callback=self.checklist.save_work_as", main_source)
        self.assertNotIn('QGroupBox("Output")', checklist_source)


if __name__ == "__main__":
    unittest.main()
