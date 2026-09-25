import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from config import Settings


class SettingsRecoveryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "settings.json"
        self.env = patch.dict(os.environ, RECEIPT_LEDGER_DATA_DIR=self.temp.name)
        self.env.start()
        self.addCleanup(self.env.stop)

    def test_non_object_json_falls_back_without_crashing(self):
        for value in (None, [], ["model"], "model", 1, True):
            with self.subTest(value=value):
                self.path.write_text(json.dumps(value), encoding="utf-8")
                self.assertEqual(Settings.load(), Settings())

    def test_invalid_fields_recover_without_discarding_valid_settings(self):
        data = {"server_url": 123, "model": None, "workbook_path": None,
                "image_folder": [], "theme": {}, "keep_images": "false",
                "backup_generations": True, "model_choices": ["valid-model", None, 12, "valid-model"]}
        self.path.write_text(json.dumps(data), encoding="utf-8")
        loaded = Settings.load()
        self.assertEqual(loaded.server_url, Settings().server_url)
        self.assertEqual(loaded.model, "")
        self.assertEqual(loaded.workbook_path, "")
        self.assertEqual(loaded.image_folder, "")
        self.assertEqual(loaded.theme, "system")
        self.assertTrue(loaded.keep_images)
        self.assertEqual(loaded.backup_generations, 20)
        self.assertEqual(loaded.model_choices, ["valid-model"])

    def test_invalid_optional_field_keeps_existing_ledger_location(self):
        ledger = str(Path(self.temp.name) / "existing.xlsx")
        self.path.write_text(json.dumps({"model": "saved-model", "theme": None,
                                        "workbook_path": ledger}), encoding="utf-8")
        loaded = Settings.load()
        self.assertEqual(loaded.model, "saved-model")
        self.assertEqual(loaded.workbook_path, ledger)
        self.assertEqual(loaded.theme, "system")


if __name__ == "__main__":
    unittest.main()
