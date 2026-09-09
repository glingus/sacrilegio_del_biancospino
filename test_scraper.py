import os
import unittest
from pathlib import Path
import pandas as pd
from selenium.common.exceptions import StaleElementReferenceException
from paginebianche_scraper import (
    Contatto,
    NO_RESULTS_CONTAINER_XPATH,
    NO_RESULTS_TEXT_XPATH,
    RESULT_CARD_SELECTORS,
    SEARCH_BASE_URL,
    build_search_url,
    deduplicate_contatti,
    element_is_stale,
    find_result_cards,
    page_reports_no_results,
    save_to_excel,
    append_checkpoint_csv,
    validate_and_clean_phone
)


class FakeElement:
    """Minimal WebElement stand-in: exposes .text and is_enabled(), and can go stale."""

    def __init__(self, text="", stale=False):
        self._text = text
        self._stale = stale

    @property
    def text(self):
        if self._stale:
            raise StaleElementReferenceException("stale element")
        return self._text

    def is_enabled(self):
        if self._stale:
            raise StaleElementReferenceException("stale element")
        return True


class FakeDriver:
    """Returns canned elements per XPath; anything unmapped yields no matches."""

    def __init__(self, mapping=None):
        self._mapping = mapping or {}
        self.current_url = "https://www.paginebianche.it/ricerca?qs=x&dv=y"

    def find_elements(self, by, xpath):
        return self._mapping.get(xpath, [])


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


class TestSearchUrl(unittest.TestCase):

    def test_uses_ricerca_endpoint(self):
        url = build_search_url("Mario Rossi", "Suzzara")
        self.assertTrue(url.startswith("https://www.paginebianche.it/ricerca?"))
        self.assertIn("qs=Mario%20Rossi", url)
        self.assertIn("dv=Suzzara", url)

    def test_does_not_use_homepage_paths(self):
        # /persone and /cerca both serve the homepage rather than a result set.
        url = build_search_url("Giulia", "Suzzara")
        self.assertNotIn("/persone?", url)
        self.assertNotIn("/cerca?", url)

    def test_honours_custom_base_url(self):
        url = build_search_url("Giulia", "Suzzara", "https://example.test/altro")
        self.assertEqual(url, "https://example.test/altro?qs=Giulia&dv=Suzzara")

    def test_encodes_special_characters(self):
        url = build_search_url("Nicolò D'Amico", "Reggio nell'Emilia", SEARCH_BASE_URL)
        self.assertNotIn(" ", url)
        self.assertIn("%27", url)


class TestNoResultsDetection(unittest.TestCase):

    def test_container_alone_is_conclusive(self):
        # Unfamiliar copy inside a no-results container must still count as
        # "no results", not as a broken layout.
        driver = FakeDriver({
            NO_RESULTS_CONTAINER_XPATH: [
                FakeElement("Non abbiamo trovato nessuna corrispondenza")
            ]
        })
        self.assertTrue(page_reports_no_results(driver))

    def test_empty_container_is_conclusive(self):
        driver = FakeDriver({NO_RESULTS_CONTAINER_XPATH: [FakeElement("")]})
        self.assertTrue(page_reports_no_results(driver))

    def test_rendered_phrase_is_detected(self):
        driver = FakeDriver({
            NO_RESULTS_TEXT_XPATH: [FakeElement("Spiacenti, nessun risultato trovato")]
        })
        self.assertTrue(page_reports_no_results(driver))

    def test_non_rendered_phrase_is_ignored(self):
        # A hidden template or i18n bundle matches the XPath but renders no text.
        driver = FakeDriver({NO_RESULTS_TEXT_XPATH: [FakeElement("")]})
        self.assertFalse(page_reports_no_results(driver))

    def test_stale_element_does_not_propagate(self):
        driver = FakeDriver({
            NO_RESULTS_TEXT_XPATH: [FakeElement(stale=True), FakeElement("Nessun risultato")]
        })
        self.assertTrue(page_reports_no_results(driver))

    def test_only_stale_elements_yield_false(self):
        driver = FakeDriver({NO_RESULTS_TEXT_XPATH: [FakeElement(stale=True)]})
        self.assertFalse(page_reports_no_results(driver))

    def test_empty_page_yields_false(self):
        self.assertFalse(page_reports_no_results(FakeDriver()))


class TestResultCards(unittest.TestCase):

    def test_returns_first_matching_selector(self):
        expected = [FakeElement("Mario Rossi")]
        driver = FakeDriver({
            RESULT_CARD_SELECTORS[1]: expected,
            RESULT_CARD_SELECTORS[2]: [FakeElement("altro")],
        })
        self.assertIs(find_result_cards(driver)[0], expected[0])

    def test_returns_empty_list_when_no_cards(self):
        self.assertEqual(find_result_cards(FakeDriver()), [])

    def test_cards_and_no_results_marker_are_independent(self):
        # A results page can carry both; the scrape loop resolves this by asking
        # for cards first, so a stray "nessun risultato" widget cannot discard them.
        driver = FakeDriver({
            RESULT_CARD_SELECTORS[0]: [FakeElement("Mario Rossi")],
            NO_RESULTS_TEXT_XPATH: [FakeElement("Nessun risultato per la categoria")],
        })
        self.assertTrue(find_result_cards(driver))
        self.assertTrue(page_reports_no_results(driver))


class TestElementStaleness(unittest.TestCase):

    def test_live_element(self):
        self.assertFalse(element_is_stale(FakeElement("x")))

    def test_stale_element(self):
        self.assertTrue(element_is_stale(FakeElement(stale=True)))

    def test_identical_text_is_not_treated_as_same_page(self):
        # Pagination keys on DOM identity, so two pages opening with an
        # identically rendered entry no longer look like a failed page change.
        before = FakeElement("Mario Rossi - Via Roma 10", stale=True)
        after = FakeElement("Mario Rossi - Via Roma 10")
        self.assertEqual(after.text, "Mario Rossi - Via Roma 10")
        self.assertTrue(element_is_stale(before))


if __name__ == "__main__":
    unittest.main()
