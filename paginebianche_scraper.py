#!/usr/bin/env python3
"""
Pagine Bianche Web Scraper
--------------------------
Scrapes personal contact data from Pagine Bianche (Italy) using Selenium.
Exports search results and summary statistics to an Excel file with pandas and openpyxl.
"""

import csv
import re
import sys
import time
import random
import argparse
import logging
import urllib.parse
from pathlib import Path
from dataclasses import dataclass
from datetime import datetime

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException
)
from webdriver_manager.chrome import ChromeDriverManager

# Setup Logger
logger = logging.getLogger("paginebianche_scraper")

# Constants
MAX_PAGES_DEFAULT = 20
PHONE_REGEX = re.compile(r'^(?:\+?39)?(?:0\d{5,10}|3\d{8,9})$')


class BotBlockedException(Exception):
    """Exception raised when an anti-bot challenge or WAF block is detected."""
    pass


@dataclass
class Contatto:
    nome: str
    indirizzo: str
    telefono: str
    comune_ricerca: str

    def get_dedup_key(self) -> tuple[str, str, str]:
        """
        Returns a normalized tuple (nome, indirizzo, telefono) used for deduplication.
        """
        norm_nome = re.sub(r'\s+', ' ', self.nome.strip().lower())
        norm_indirizzo = re.sub(r'\s+', ' ', self.indirizzo.strip().lower())
        norm_telefono = re.sub(r'[^\d+]', '', self.telefono.strip())
        return (norm_nome, norm_indirizzo, norm_telefono)


def setup_logging(verbose: bool = False) -> None:
    """
    Configures application-wide logging with DEBUG or INFO levels.
    """
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )


def random_delay(min_sec: float = 1.5, max_sec: float = 3.5) -> None:
    """
    Introduces a random sleep delay to simulate human browsing behavior.
    """
    time.sleep(random.uniform(min_sec, max_sec))


def parse_arguments() -> argparse.Namespace:
    """
    Parses command line arguments. Returns Namespace with parameters.
    """
    parser = argparse.ArgumentParser(
        description="Scraper Pagine Bianche per ricerca persone fisiche e esportazione in Excel."
    )
    parser.add_argument("--nome", type=str, help="Nome e cognome da cercare (es. 'Mario Rossi')")
    parser.add_argument("--comuni", type=str, help="Lista dei comuni separati da virgola (es. 'Suzzara, Mantova')")
    parser.add_argument("--output", type=str, help="Nome del file di output (opzionale)")
    parser.add_argument("--output-dir", type=str, default=".", help="Directory di destinazione per file Excel e checkpoint")
    parser.add_argument("--max-pages", type=int, default=MAX_PAGES_DEFAULT, help="Numero massimo di pagine per comune (default: 20)")
    parser.add_argument("--no-headless", action="store_true", help="Esegui il browser in modalità visibile (non headless)")
    parser.add_argument("--verbose", action="store_true", help="Abilita log dettagliati (DEBUG)")
    return parser.parse_args()


def get_user_inputs(args: argparse.Namespace) -> tuple[str, list[str]]:
    """
    Retrieves inputs from argparse options or falls back to interactive CLI prompts.
    """
    target_name = args.nome
    if not target_name:
        print("=" * 60)
        print("      Pagine Bianche Scraper - Persona Fisica (Privati)")
        print("=" * 60)
        target_name = input("Quale nome/cognome stai cercando? ").strip()
        while not target_name:
            print("Il nome/cognome non può essere vuoto.")
            target_name = input("Quale nome/cognome stai cercando? ").strip()

    comuni_raw = args.comuni
    if not comuni_raw:
        comuni_raw = input(
            "In quale comune/comuni vuoi cercare? (se più di uno, separali da una virgola, es: Suzzara, Mantova): "
        ).strip()
        while not comuni_raw:
            print("Devi inserire almeno un comune.")
            comuni_raw = input(
                "In quale comune/comuni vuoi cercare? (se più di uno, separali da una virgola, es: Suzzara, Mantova): "
            ).strip()

    comuni_list = [c.strip() for c in comuni_raw.split(",") if c.strip()]
    return target_name, comuni_list


def setup_driver(headless: bool = True) -> webdriver.Chrome:
    """
    Configures and initializes Chrome WebDriver using webdriver_manager.
    """
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, Gecko) Chrome/122.0.0.0 Safari/537.36"
    )

    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.cookies_accepted = False  # Attach cookie acceptance state to driver instance
    return driver


def check_for_bot_block(driver: webdriver.Chrome) -> None:
    """
    Checks if an anti-bot challenge, captcha, or WAF block page is presented.
    Raises BotBlockedException if a block is detected.
    """
    title = driver.title.lower()
    page_source = driver.page_source.lower()

    block_indicators = ["captcha", "challenge", "verifica la tua identità", "access denied", "robot"]
    if any(ind in title for ind in block_indicators) or "awswafintegration" in page_source:
        logger.critical(f"Rilevato blocco anti-bot / WAF / Captcha (Titolo pagina: '{driver.title}'). Interruzione run.")
        raise BotBlockedException(f"Blocco anti-bot rilevato: {driver.title}")


def accept_cookies(driver: webdriver.Chrome, timeout: int = 3) -> bool:
    """
    Detects and clicks the cookie acceptance banner button if present using a single combined XPath.
    Uses driver attribute state to avoid re-checking after success.
    """
    if getattr(driver, 'cookies_accepted', False):
        return True

    combined_xpath = (
        "//button[contains(@id, 'iubenda-cs-accept-btn') or "
        "contains(@class, 'iubenda-cs-accept-btn') or "
        "@id='onetrust-accept-btn-handler'] | "
        "//button[contains(text(), 'Accetta') or contains(text(), 'ACCETTA')] | "
        "//a[contains(text(), 'Accetta') or contains(text(), 'ACCETTA')]"
    )

    try:
        wait = WebDriverWait(driver, timeout)
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, combined_xpath)))
        driver.execute_script("arguments[0].click();", btn)
        logger.info("[+] Banner cookie accettato.")
        driver.cookies_accepted = True
        random_delay(1.0, 2.0)
        return True
    except TimeoutException:
        logger.debug("Nessun banner cookie trovato entro il timeout.")
    except Exception as e:
        logger.warning(f"Avviso durante la gestione dei cookie: {e}", exc_info=True)
    return False


def search_municipality(driver: webdriver.Chrome, name: str, comune: str) -> None:
    """
    Navigates to Pagine Bianche search page for the given name and municipality
    and waits for results container or no-results element.
    """
    encoded_name = urllib.parse.quote(name)
    encoded_comune = urllib.parse.quote(comune)
    url = f"https://www.paginebianche.it/cerca?qs={encoded_name}&dv={encoded_comune}"

    logger.info(f"Ricerca per '{name}' a '{comune}' -> {url}")
    driver.get(url)

    check_for_bot_block(driver)
    accept_cookies(driver)

    results_or_no_results_xpath = (
        "//div[contains(@class, 'search-item') or "
        "contains(@class, 'item-listing') or "
        "contains(@class, 'search-itm') or "
        "contains(@class, 'no-results') or "
        "contains(@class, 'no-result') or "
        "//*[contains(text(), 'Nessun risultato')] or "
        "//*[contains(text(), 'non ha prodotto risultati')]]"
    )
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, results_or_no_results_xpath))
        )
    except TimeoutException:
        logger.error("Timeout in attesa del caricamento dei risultati della ricerca.")


def validate_and_clean_phone(raw_phone: str) -> str:
    """
    Extracts, normalizes to digits only, and validates Italian phone numbers using PHONE_REGEX.
    Returns cleaned string or empty string if invalid.
    """
    if not raw_phone:
        return ""

    candidates = re.findall(r'(\+?\d[\d\s\-\.\/]{5,}\d)', raw_phone)
    for cand in candidates:
        digits_only = re.sub(r'[^\d+]', '', cand)
        if PHONE_REGEX.match(digits_only):
            return digits_only

    digits_only = re.sub(r'[^\d+]', '', raw_phone)
    if PHONE_REGEX.match(digits_only):
        return digits_only

    return ""


def extract_card_data(driver: webdriver.Chrome, card_element, comune_ricerca: str) -> Contatto | None:
    """
    Extracts Name, Address, and Phone from a single card WebElement and returns Contatto or None.
    Uses href='tel:' as primary phone source, and waits via WebDriverWait after 'Mostra numero'.
    """
    # Name extraction
    nome = ""
    nome_selectors = [
        ".//h2",
        ".//h3",
        ".//*[contains(@class, 'title')]",
        ".//*[contains(@class, 'name')]",
        ".//a[contains(@class, 'header')]"
    ]
    for sel in nome_selectors:
        try:
            elem = card_element.find_element(By.XPATH, sel)
            text = elem.text.strip()
            if text:
                nome = text
                break
        except NoSuchElementException:
            continue

    if not nome:
        return None

    # Primary Phone extraction via href='tel:'
    telefono = ""
    try:
        tel_links = card_element.find_elements(By.XPATH, ".//a[contains(@href, 'tel:')]")
        for link in tel_links:
            href_val = link.get_attribute("href") or ""
            cleaned = validate_and_clean_phone(href_val)
            if cleaned:
                telefono = cleaned
                break
    except Exception:
        pass

    # If phone not found in href, try clicking "Mostra numero" if present
    if not telefono:
        try:
            phone_buttons = card_element.find_elements(
                By.XPATH,
                ".//button[contains(text(), 'Mostra') or contains(text(), 'numero') or contains(@class, 'phone')]"
                " | .//a[contains(text(), 'Mostra') or contains(text(), 'numero') or contains(@class, 'phone')]"
            )
            for btn in phone_buttons:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    # WebDriverWait for phone element to appear after click
                    try:
                        WebDriverWait(driver, 3).until(
                            lambda d: card_element.find_elements(By.XPATH, ".//a[contains(@href, 'tel:')] | .//*[contains(@class, 'phone') or contains(@class, 'tel')]")
                        )
                    except TimeoutException:
                        pass
                    break
        except Exception as e:
            logger.debug(f"Pulsante mostra numero non presente o non cliccabile: {e}")

        # Secondary Phone extraction from card text / tel elements
        phone_selectors = [
            ".//a[contains(@href, 'tel:')]",
            ".//*[contains(@class, 'phone') or contains(@class, 'tel')]"
        ]
        for sel in phone_selectors:
            try:
                elem = card_element.find_element(By.XPATH, sel)
                raw_text = elem.text.strip() or elem.get_attribute("href") or ""
                cleaned_phone = validate_and_clean_phone(raw_text)
                if cleaned_phone:
                    telefono = cleaned_phone
                    break
            except NoSuchElementException:
                continue

    # Address extraction
    indirizzo = ""
    address_selectors = [
        ".//*[contains(@class, 'address')]",
        ".//*[contains(@class, 'street')]",
        ".//*[contains(@class, 'location')]",
        ".//*[contains(@class, 'indirizzo')]"
    ]
    for sel in address_selectors:
        try:
            elem = card_element.find_element(By.XPATH, sel)
            text = elem.text.strip()
            if text:
                indirizzo = text
                break
        except NoSuchElementException:
            continue

    if not indirizzo:
        try:
            card_text = card_element.text.strip().split("\n")
            if len(card_text) > 1:
                for line in card_text[1:]:
                    if any(kw in line.lower() for kw in ['via', 'viale', 'piazza', 'corso', 'largo', 'strada']) or re.search(r'\b\d{5}\b', line):
                        indirizzo = line.strip()
                        break
        except StaleElementReferenceException:
            logger.debug("Stale element durante parsing del testo di fallback.")

    return Contatto(
        nome=nome,
        indirizzo=indirizzo,
        telefono=telefono,
        comune_ricerca=comune_ricerca
    )


def append_checkpoint_csv(checkpoint_filepath: Path, contatti: list[Contatto]) -> None:
    """
    Appends a list of Contatto objects to a CSV checkpoint file.
    Creates header if file does not exist.
    """
    if not contatti:
        return
    file_exists = checkpoint_filepath.exists()
    try:
        with open(checkpoint_filepath, mode="a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Nome", "Indirizzo", "Telefono", "Comune di Ricerca"])
            for c in contatti:
                writer.writerow([c.nome, c.indirizzo, c.telefono, c.comune_ricerca])
        logger.info(f"Appesi {len(contatti)} record al file di checkpoint '{checkpoint_filepath.name}'.")
    except Exception as e:
        logger.error(f"Errore durante la scrittura del checkpoint CSV '{checkpoint_filepath}': {e}", exc_info=True)


def scrape_results_for_municipality(
    driver: webdriver.Chrome,
    target_name: str,
    comune_ricerca: str,
    output_dir: Path,
    run_timestamp: str,
    max_pages: int = MAX_PAGES_DEFAULT
) -> list[Contatto]:
    """
    Scrapes all pages of search results for a single municipality.
    Appends checkpoint CSV per page.
    Distinguishes no-results (INFO) from broken layout (ERROR).
    """
    sanitized_name = re.sub(r'[^a-zA-Z0-9]', '_', target_name.strip().lower())
    sanitized_comune = re.sub(r'[^a-zA-Z0-9]', '_', comune_ricerca.strip().lower())
    checkpoint_file = output_dir / f"checkpoint_{sanitized_name}_{sanitized_comune}_{run_timestamp}.csv"

    contatti: list[Contatto] = []
    page_num = 1

    specific_card_selectors = [
        "//div[contains(@class, 'search-item')]",
        "//div[contains(@class, 'item-listing')]",
        "//div[contains(@class, 'search-itm')]"
    ]

    no_results_xpath = (
        "//div[contains(@class, 'no-results') or contains(@class, 'no-result')] | "
        "//*[contains(text(), 'Nessun risultato')] | "
        "//*[contains(text(), 'non ha prodotto risultati')]"
    )

    while page_num <= max_pages:
        logger.info(f"Estrazione pagina {page_num}/{max_pages} per comune: {comune_ricerca}...")
        random_delay(1.0, 2.0)
        check_for_bot_block(driver)

        # Check explicit no-results elements first
        no_results_elements = driver.find_elements(By.XPATH, no_results_xpath)
        if no_results_elements and page_num == 1:
            logger.info(f"Nessun risultato trovato per '{target_name}' a '{comune_ricerca}'.")
            break

        selected_selector = None
        card_elements = []

        for selector in specific_card_selectors:
            found = driver.find_elements(By.XPATH, selector)
            if found:
                selected_selector = selector
                card_elements = found
                break

        if not card_elements:
            if page_num == 1 and not no_results_elements:
                logger.error(f"Layout cambiato o selettori non validi per {comune_ricerca} a pagina {page_num}.")
            else:
                logger.info(f"Nessun'altra scheda trovata a pagina {page_num} per {comune_ricerca}.")
            break

        page_contatti: list[Contatto] = []
        for card_elem in card_elements:
            try:
                contatto = extract_card_data(driver, card_elem, comune_ricerca)
                if contatto:
                    page_contatti.append(contatto)
            except Exception as e:
                logger.error(f"Errore durante l'estrazione di una scheda per {comune_ricerca}: {e}", exc_info=True)
                continue

        logger.info(f"Estratte {len(page_contatti)} persone da pagina {page_num}.")
        contatti.extend(page_contatti)

        # Write per-page checkpoint CSV
        append_checkpoint_csv(checkpoint_file, page_contatti)

        # Pagination check
        next_button = None
        next_selectors = [
            "//a[contains(@class, 'pagination__next')]",
            "//a[@rel='next']",
            "//a[contains(text(), 'Successiva') or contains(text(), 'Prossima')]",
            "//li[contains(@class, 'next')]/a"
        ]

        for sel in next_selectors:
            try:
                elems = driver.find_elements(By.XPATH, sel)
                for elem in elems:
                    if elem.is_displayed() and elem.is_enabled():
                        next_button = elem
                        break
                if next_button:
                    break
            except Exception:
                continue

        if next_button:
            # Capture current first card text for page change verification
            first_card_text_before = card_elements[0].text.strip() if card_elements else ""
            logger.info(f"Passaggio alla pagina successiva ({page_num + 1})...")

            try:
                driver.execute_script("arguments[0].click();", next_button)
                page_num += 1

                # Wait for page content to update
                try:
                    WebDriverWait(driver, 10).until(
                        lambda d: (
                            len(d.find_elements(By.XPATH, selected_selector)) > 0 and
                            d.find_elements(By.XPATH, selected_selector)[0].text.strip() != first_card_text_before
                        )
                    )
                except TimeoutException:
                    logger.info("Timeout in attesa dell'aggiornamento del contenuto della nuova pagina. Interruzione paginazione.")
                    break

            except Exception as e:
                logger.error(f"Impossibile cliccare 'Pagina successiva': {e}", exc_info=True)
                break
        else:
            logger.info(f"Fine dei risultati paginati per {comune_ricerca}.")
            break

    if page_num > max_pages:
        logger.info(f"Raggiunto il limite massimo di pagine ({max_pages}) per {comune_ricerca}.")

    return contatti


def deduplicate_contatti(contatti: list[Contatto]) -> list[Contatto]:
    """
    Deduplicates contacts based on normalized (nome, indirizzo, telefono).
    """
    seen: set[tuple[str, str, str]] = set()
    unique_contatti: list[Contatto] = []
    for c in contatti:
        key = c.get_dedup_key()
        if key not in seen:
            seen.add(key)
            unique_contatti.append(c)
        else:
            logger.debug(f"Deduplicato contatto ripetuto: {c.nome} ({c.indirizzo})")
    logger.info(f"Deduplicazione completata: da {len(contatti)} a {len(unique_contatti)} contatti unici.")
    return unique_contatti


def filtra_rpo(contatti: list[Contatto]) -> list[Contatto]:
    """
    Stub di conformità per il filtraggio contro il Registro Pubblico delle Opposizioni (RPO).
    """
    logger.warning(
        "ATTENZIONE (Conformità RPO): I contatti non sono stati filtrati rispetto al "
        "Registro Pubblico delle Opposizioni. Prima di utilizzare i numeri per scopi commerciali "
        "o telemarketing, verificare l'iscrizione al RPO (D.P.R. n. 26/2022)."
    )
    return contatti


def save_to_excel(
    name: str,
    comuni_list: list[str],
    contatti: list[Contatto],
    output_dir: Path,
    run_timestamp: str,
    output_filename: str | None = None
) -> str:
    """
    Exports collected Contatto objects to an Excel file with two sheets:
    - Sheet 'Riepilogo': Summary table of searched municipalities and counts.
    - Sheet 'Dati': Full dataset (Nome, Indirizzo, Telefono, Comune di Ricerca).
    """
    if not output_filename:
        sanitized_name = re.sub(r'[^a-zA-Z0-9]', '_', name.strip().lower())
        filename_path = output_dir / f"ricerca_{sanitized_name}_{run_timestamp}.xlsx"
    else:
        filename_path = output_dir / output_filename

    summary_dict = {c: 0 for c in comuni_list}
    for c in contatti:
        comune = c.comune_ricerca
        if comune in summary_dict:
            summary_dict[comune] += 1
        else:
            summary_dict[comune] = 1

    df_riepilogo = pd.DataFrame([
        {"Comune di Ricerca": comune, "Numero Totale Persone Trovate": count}
        for comune, count in summary_dict.items()
    ])

    if contatti:
        raw_data = [
            {
                "Nome": c.nome,
                "Indirizzo": c.indirizzo,
                "Telefono": c.telefono,
                "Comune di Ricerca": c.comune_ricerca
            }
            for c in contatti
        ]
        df_dati = pd.DataFrame(raw_data)[["Nome", "Indirizzo", "Telefono", "Comune di Ricerca"]]
    else:
        df_dati = pd.DataFrame(columns=["Nome", "Indirizzo", "Telefono", "Comune di Ricerca"])

    with pd.ExcelWriter(filename_path, engine='openpyxl') as writer:
        df_riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)
        df_dati.to_excel(writer, sheet_name='Dati', index=False)

    logger.info(f"Esportazione completata con successo nel file: '{filename_path}'")
    return str(filename_path)


def main() -> None:
    args = parse_arguments()
    setup_logging(args.verbose)

    target_name, comuni_list = get_user_inputs(args)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    run_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    driver = None
    all_extracted_contatti: list[Contatto] = []

    try:
        headless = not args.no_headless
        driver = setup_driver(headless=headless)

        for comune in comuni_list:
            try:
                search_municipality(driver, target_name, comune)
                comune_results = scrape_results_for_municipality(
                    driver=driver,
                    target_name=target_name,
                    comune_ricerca=comune,
                    output_dir=output_dir,
                    run_timestamp=run_timestamp,
                    max_pages=args.max_pages
                )
                all_extracted_contatti.extend(comune_results)
            except BotBlockedException:
                logger.critical("Esecuzione interrotta per blocco anti-bot WAF/Captcha.")
                sys.exit(1)
            except Exception as e:
                logger.error(f"Errore durante l'elaborazione del comune '{comune}': {e}", exc_info=True)
                continue

    except BotBlockedException:
        logger.critical("Esecuzione interrotta per blocco anti-bot WAF/Captcha.")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Si è verificato un errore critico durante l'esecuzione del browser: {e}", exc_info=True)
    finally:
        if driver:
            driver.quit()

    unique_contatti = deduplicate_contatti(all_extracted_contatti)
    filtered_contatti = filtra_rpo(unique_contatti)
    save_to_excel(
        name=target_name,
        comuni_list=comuni_list,
        contatti=filtered_contatti,
        output_dir=output_dir,
        run_timestamp=run_timestamp,
        output_filename=args.output
    )


if __name__ == "__main__":
    main()
