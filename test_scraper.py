import os
import unittest
from pathlib import Path
import pandas as pd
from paginebianche_scraper import (
    Contatto,
    deduplicate_contatti,
    save_to_excel,
    append_checkpoint_csv,
    validate_and_clean_phone
)

class TestScraperRefactored(unittest.TestCase):

    def test_phone_validation(self):
        self.assertEqual(validate_and_clean_phone("0376 123456"), "0376123456")
        self.assertEqual(validate_and_clean_phone("+39 340 1234567"), "+393401234567")
        self.assertEqual(validate_and_clean_phone("3391234567"), "3391234567")
        self.assertEqual(validate_and_clean_phone("via Roma 15, Milano"), "")

    def test_deduplication(self):
        c1 = Contatto("Mario Rossi", "Via Roma 10", "0376123456", "Suzzara")
        c2 = Contatto("mario rossi", "via roma 10", "0376123456", "Mantova") # duplicate
        c3 = Contatto("Luigi Bianchi", "Corso Italia 5", "0376654321", "Mantova")

        deduped = deduplicate_contatti([c1, c2, c3])
        self.assertEqual(len(deduped), 2)
        self.assertEqual(deduped[0].nome, "Mario Rossi")
        self.assertEqual(deduped[1].nome, "Luigi Bianchi")

    def test_per_page_append_checkpoint_csv(self):
        out_dir = Path("test_out")
        out_dir.mkdir(exist_ok=True)
        chk_file = out_dir / "checkpoint_mario_rossi_suzzara_20260906.csv"

        c1 = Contatto("Mario Rossi", "Via Roma 10", "0376123456", "Suzzara")
        c2 = Contatto("Giuseppe Rossi", "Via Milano 2", "0376654321", "Suzzara")

        append_checkpoint_csv(chk_file, [c1])
        append_checkpoint_csv(chk_file, [c2])

        self.assertTrue(chk_file.exists())

        df = pd.read_csv(chk_file)
        self.assertEqual(len(df), 2)
        self.assertListEqual(list(df.columns), ["Nome", "Indirizzo", "Telefono", "Comune di Ricerca"])

        # Clean up
        if chk_file.exists():
            chk_file.unlink()
        if out_dir.exists():
            out_dir.rmdir()

    def test_excel_export(self):
        out_dir = Path("test_out_excel")
        out_dir.mkdir(exist_ok=True)
        run_ts = "20260906_120000"

        c1 = Contatto("Mario Rossi", "Via Roma 10", "0376123456", "Suzzara")
        file_path = save_to_excel("Mario Rossi", ["Suzzara"], [c1], output_dir=out_dir, run_timestamp=run_ts)

        self.assertTrue(os.path.exists(file_path))

        excel_file = pd.ExcelFile(file_path)
        self.assertIn("Riepilogo", excel_file.sheet_names)
        self.assertIn("Dati", excel_file.sheet_names)

        df_dati = pd.read_excel(file_path, sheet_name="Dati")
        self.assertEqual(len(df_dati), 1)

        # Clean up
        if os.path.exists(file_path):
            os.remove(file_path)
        if out_dir.exists():
            out_dir.rmdir()

if __name__ == "__main__":
    unittest.main()
