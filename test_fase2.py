import os
import unittest
import pandas as pd
from paginebianche_scraper import (
    Contatto,
    deduplicate_contatti,
    save_to_excel,
    save_checkpoint_csv,
    validate_and_clean_phone
)

class TestFase2Refactoring(unittest.TestCase):

    def test_deduplication(self):
        c1 = Contatto("Mario Rossi", "Via Roma 10, Suzzara", "0376123456", "Suzzara")
        c2 = Contatto("mario rossi", "via roma 10, suzzara", "0376123456", "Mantova") # duplicate
        c3 = Contatto("Luigi Bianchi", "Corso Italia 5", "0376654321", "Mantova")

        deduped = deduplicate_contatti([c1, c2, c3])
        self.assertEqual(len(deduped), 2)
        self.assertEqual(deduped[0].nome, "Mario Rossi")
        self.assertEqual(deduped[1].nome, "Luigi Bianchi")

    def test_checkpoint_csv(self):
        comune = "Suzzara"
        c1 = Contatto("Mario Rossi", "Via Roma 10", "0376123456", comune)
        save_checkpoint_csv(comune, [c1])

        expected_csv = "checkpoint_suzzara.csv"
        self.assertTrue(os.path.exists(expected_csv))

        df_csv = pd.read_csv(expected_csv)
        self.assertEqual(len(df_csv), 1)
        self.assertEqual(df_csv.iloc[0]["Nome"], "Mario Rossi")

        # Cleanup
        os.remove(expected_csv)

    def test_excel_export_filename(self):
        c1 = Contatto("Mario Rossi", "Via Roma 10", "0376123456", "Suzzara")
        filename = save_to_excel("Mario Rossi", ["Suzzara"], [c1])
        self.assertTrue(os.path.exists(filename))
        self.assertIn("ricerca_mario_rossi_", filename)

        excel_file = pd.ExcelFile(filename)
        self.assertIn("Riepilogo", excel_file.sheet_names)
        self.assertIn("Dati", excel_file.sheet_names)

        df_dati = pd.read_excel(filename, sheet_name="Dati")
        self.assertEqual(len(df_dati), 1)

        os.remove(filename)

if __name__ == "__main__":
    unittest.main()
