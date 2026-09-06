import os
import unittest
import pandas as pd
from paginebianche_scraper import save_to_excel

class TestExcelExport(unittest.TestCase):
    def test_save_to_excel(self):
        target_name = "Mario Rossi"
        comuni_list = ["Suzzara", "Mantova"]
        mock_data = [
            {
                "Nome": "Rossi Mario",
                "Indirizzo": "Via Roma 10, 46029 Suzzara",
                "Telefono": "0376 123456",
                "Comune di Ricerca": "Suzzara"
            },
            {
                "Nome": "Rossi Mario",
                "Indirizzo": "Corso Vittorio Emanuele 5, 46100 Mantova",
                "Telefono": "0376 654321",
                "Comune di Ricerca": "Mantova"
            },
            {
                "Nome": "Rossi Mario Jr",
                "Indirizzo": "Via Milano 1, 46100 Mantova",
                "Telefono": "0376 999888",
                "Comune di Ricerca": "Mantova"
            }
        ]

        filename = save_to_excel(target_name, comuni_list, mock_data)
        self.assertTrue(os.path.exists(filename))

        # Check sheets using pandas ExcelFile
        excel_file = pd.ExcelFile(filename)
        sheet_names = excel_file.sheet_names
        self.assertIn("Riepilogo", sheet_names)
        self.assertIn("Dati", sheet_names)

        # Read Riepilogo sheet
        df_riepilogo = pd.read_excel(filename, sheet_name="Riepilogo")
        self.assertEqual(len(df_riepilogo), 2)
        self.assertListEqual(list(df_riepilogo.columns), ["Comune di Ricerca", "Numero Totale Persone Trovate"])

        suzzara_row = df_riepilogo[df_riepilogo["Comune di Ricerca"] == "Suzzara"]
        mantova_row = df_riepilogo[df_riepilogo["Comune di Ricerca"] == "Mantova"]
        self.assertEqual(suzzara_row["Numero Totale Persone Trovate"].values[0], 1)
        self.assertEqual(mantova_row["Numero Totale Persone Trovate"].values[0], 2)

        # Read Dati sheet
        df_dati = pd.read_excel(filename, sheet_name="Dati")
        self.assertEqual(len(df_dati), 3)
        self.assertListEqual(list(df_dati.columns), ["Nome", "Indirizzo", "Telefono", "Comune di Ricerca"])

        # Clean up
        if os.path.exists(filename):
            os.remove(filename)

if __name__ == "__main__":
    unittest.main()
