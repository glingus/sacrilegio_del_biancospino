#!/usr/bin/env python3
"""
Pagine Bianche Web Scraper
--------------------------
Scrapes personal contact data from Pagine Bianche (Italy) using Selenium.
Exports search results and summary statistics to an Excel file with pandas and openpyxl.
"""

import csv
import re
import time
import random
import argparse
import logging
import urllib.parse
from dataclasses import dataclass
from datetime import datetime
from typing import List, Dict, Tuple, Set

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

# Session flag for cookie acceptance
COOKIES_ACCEPTED = False


@dataclass
class Contatto:
    nome: str
    indirizzo: str
    telefono: str
    comune_ricerca: str

    def get_dedup_key(self) -> Tuple[str, str, str]:
        """
        Returns a normalized tuple (nome, indirizzo, telefono) used for deduplication.
        """
        norm_nome = re.sub(r'\s+', ' ', self.nome.strip().lower())
        norm_indirizzo = re.sub(r'\s+', ' ', self.indirizzo.strip().lower())
        norm_telefono = re.sub(r'[^\d+]', '', self.telefono.strip())
        return (norm_nome, norm_indirizzo, norm_telefono)


def setup_logging(verbose: bool = False):
    """
    Configures application-wide logging with DEBUG or INFO levels.
    """
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )


def random_delay(min_sec: float = 1.5, max_sec: float = 3.5):
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
    parser.add_argument("--max-pages", type=int, default=MAX_PAGES_DEFAULT, help="Numero massimo di pagine per comune (default: 20)")
    parser.add_argument("--no-headless", action="store_true", help="Esegui il browser in modalità visibile (non headless)")
    parser.add_argument("--verbose", action="store_true", help="Abilita log dettagliati (DEBUG)")
    return parser.parse_args()


def get_user_inputs(args: argparse.Namespace) -> Tuple[str, List[str]]:
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
    return driver


def accept_cookies(driver: webdriver.Chrome, timeout: int = 3) -> bool:
    """
    Detects and clicks the cookie acceptance banner button if present using a single combined XPath.
    Uses a session flag to avoid re-checking after success.
    """
    global COOKIES_ACCEPTED
    if COOKIES_ACCEPTED:
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
        COOKIES_ACCEPTED = True
        random_delay(1.0, 2.0)
        return True
    except TimeoutException:
        logger.debug("Nessun banner cookie trovato entro il timeout.")
    except Exception as e:
        logger.warning(f"Avviso durante la gestione dei cookie: {e}", exc_info=True)
    return False


def search_municipality(driver: webdriver.Chrome, name: str, comune: str):
    """
    Navigates to Pagine Bianche search page for the given name and municipality
    and waits for results container.
    """
    encoded_name = urllib.parse.quote(name)
    encoded_comune = urllib.parse.quote(comune)
    url = f"https://www.paginebianche.it/cerca?qs={encoded_name}&dv={encoded_comune}"

    logger.info(f"Ricerca per '{name}' a '{comune}' -> {url}")
    driver.get(url)

    accept_cookies(driver)

    results_container_xpath = (
        "//div[contains(@class, 'search-item') or "
        "contains(@class, 'item-listing') or "
        "contains(@class, 'search-itm') or "
        "contains(@class, 'no-results') or "
        "contains(@class, 'no-result')]"
    )
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, results_container_xpath))
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

    if any(word in raw_phone.lower() for word in ['via', 'viale', 'piazza', 'corso', 'largo', 'strada', 'cap', 'milano', 'roma']):
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


def extract_card_data(driver: webdriver.Chrome, card_index: int, specific_selector: str) -> Dict[str, str]:
    """
    Extracts Name, Address, and Phone from a single result card re-queried by index.
    Clicks 'Mostra numero' if present, re-fetching elements to avoid StaleElementReferenceException.
    """
    cards = driver.find_elements(By.XPATH, specific_selector)
    if card_index >= len(cards):
        return {"Nome": "", "Indirizzo": "", "Telefono": ""}

    card = cards[card_index]

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
            elem = card.find_element(By.XPATH, sel)
            text = elem.text.strip()
            if text:
                nome = text
                break
        except NoSuchElementException:
            continue

    try:
        phone_buttons = card.find_elements(
            By.XPATH,
            ".//button[contains(text(), 'Mostra') or contains(text(), 'numero') or contains(@class, 'phone')]"
            " | .//a[contains(text(), 'Mostra') or contains(text(), 'numero') or contains(@class, 'phone')]"
        )
        for btn in phone_buttons:
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(0.5)
                cards = driver.find_elements(By.XPATH, specific_selector)
                if card_index < len(cards):
                    card = cards[card_index]
                break
    except Exception as e:
        logger.debug(f"Pulsante mostra numero non cliccabile o non presente: {e}")

    telefono = ""
    phone_selectors = [
        ".//*[contains(@class, 'phone') or contains(@class, 'tel')]",
        ".//a[contains(@href, 'tel:')]"
    ]
    for sel in phone_selectors:
        try:
            elem = card.find_element(By.XPATH, sel)
            raw_text = elem.text.strip() or elem.get_attribute("href") or ""
            cleaned_phone = validate_and_clean_phone(raw_text)
            if cleaned_phone:
                telefono = cleaned_phone
                break
        except NoSuchElementException:
            continue

    indirizzo = ""
    address_selectors = [
        ".//*[contains(@class, 'address')]",
        ".//*[contains(@class, 'street')]",
        ".//*[contains(@class, 'location')]",
        ".//*[contains(@class, 'indirizzo')]"
    ]
    for sel in address_selectors:
        try:
            elem = card.find_element(By.XPATH, sel)
            text = elem.text.strip()
            if text:
                indirizzo = text
                break
        except NoSuchElementException:
            continue

    if not indirizzo or not nome:
        try:
            card_text = card.text.strip().split("\n")
            if card_text:
                if not nome and len(card_text) > 0:
                    nome = card_text[0]
                if not indirizzo and len(card_text) > 1:
                    for line in card_text[1:]:
                        if any(kw in line.lower() for kw in ['via', 'viale', 'piazza', 'corso', 'largo', 'strada']) or re.search(r'\b\d{5}\b', line):
                            indirizzo = line.strip()
                            break
        except StaleElementReferenceException:
            logger.debug("Stale element intercettato durante il parsing testuale di fallback.")

    return {
        "Nome": nome,
        "Indirizzo": indirizzo,
        "Telefono": telefono
    }


def save_checkpoint_csv(comune: str, contatti: List[Contatto]):
    """
    Saves a CSV checkpoint file for the specified municipality.
    """
    sanitized_comune = re.sub(r'[^a-zA-Z0-9]', '_', comune.strip().lower())
    checkpoint_file = f"checkpoint_{sanitized_comune}.csv"
    try:
        with open(checkpoint_file, mode="w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Nome", "Indirizzo", "Telefono", "Comune di Ricerca"])
            for c in contatti:
                writer.writerow([c.nome, c.indirizzo, c.telefono, c.comune_ricerca])
        logger.info(f"Checkpoint salvato per {comune}: '{checkpoint_file}' ({len(contatti)} record).")
    except Exception as e:
        logger.error(f"Errore durante il salvataggio del checkpoint CSV per {comune}: {e}", exc_info=True)


def scrape_results_for_municipality(driver: webdriver.Chrome, comune_ricerca: str, max_pages: int = MAX_PAGES_DEFAULT) -> List[Contatto]:
    """
    Scrapes all pages of search results for a single municipality.
    Returns list of Contatto objects.
    """
    contatti = []
    page_num = 1
    previous_url = driver.current_url
    previous_first_card_text = ""

    specific_card_selectors = [
        "//div[contains(@class, 'search-item')]",
        "//div[contains(@class, 'item-listing')]",
        "//div[contains(@class, 'search-itm')]"
    ]

    while page_num <= max_pages:
        logger.info(f"Estrazione pagina {page_num}/{max_pages} per comune: {comune_ricerca}...")
        random_delay(1.0, 2.0)

        selected_selector = None
        cards_count = 0

        for selector in specific_card_selectors:
            found = driver.find_elements(By.XPATH, selector)
            if found:
                selected_selector = selector
                cards_count = len(found)
                break

        if not selected_selector or cards_count == 0:
            logger.error(f"Layout cambiato o nessun risultato trovato per {comune_ricerca} a pagina {page_num}.")
            break

        extracted_on_page = 0
        for i in range(cards_count):
            try:
                data = extract_card_data(driver, i, selected_selector)
                if data["Nome"]:
                    contatto = Contatto(
                        nome=data["Nome"],
                        indirizzo=data["Indirizzo"],
                        telefono=data["Telefono"],
                        comune_ricerca=comune_ricerca
                    )
                    contatti.append(contatto)
                    extracted_on_page += 1
            except Exception as e:
                logger.error(f"Errore durante l'estrazione della scheda {i} per {comune_ricerca}: {e}", exc_info=True)
                continue

        logger.info(f"Estratte {extracted_on_page} persone da pagina {page_num}.")

        try:
            cards = driver.find_elements(By.XPATH, selected_selector)
            current_first_card_text = cards[0].text.strip() if cards else ""
        except Exception:
            current_first_card_text = ""

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
            logger.info(f"Passaggio alla pagina successiva ({page_num + 1})...")
            try:
                driver.execute_script("arguments[0].click();", next_button)
                page_num += 1
                random_delay(2.0, 3.5)

                current_url = driver.current_url
                if current_url == previous_url and current_first_card_text == previous_first_card_text:
                    logger.info("Rilevato cambio pagina fallito (URL/contenuto immutato). Interruzione paginazione.")
                    break
                previous_url = current_url
                previous_first_card_text = current_first_card_text

            except Exception as e:
                logger.error(f"Impossibile cliccare 'Pagina successiva': {e}", exc_info=True)
                break
        else:
            logger.info(f"Fine dei risultati paginati per {comune_ricerca}.")
            break

    if page_num > max_pages:
        logger.info(f"Raggiunto il limite massimo di pagine ({max_pages}) per {comune_ricerca}.")

    save_checkpoint_csv(comune_ricerca, contatti)
    return contatti


def deduplicate_contatti(contatti: List[Contatto]) -> List[Contatto]:
    """
    Deduplicates contacts based on normalized (nome, indirizzo, telefono).
    """
    seen: Set[Tuple[str, str, str]] = set()
    unique_contatti: List[Contatto] = []
    for c in contatti:
        key = c.get_dedup_key()
        if key not in seen:
            seen.add(key)
            unique_contatti.append(c)
        else:
            logger.debug(f"Deduplicato contatto ripetuto: {c.nome} ({c.indirizzo})")
    logger.info(f"Deduplicazione completata: da {len(contatti)} a {len(unique_contatti)} contatti unici.")
    return unique_contatti


def filtra_rpo(contatti: List[Contatto]) -> List[Contatto]:
    """
    Stub di conformità per il filtraggio contro il Registro Pubblico delle Opposizioni (RPO).

    I numeri di telefono destinati a contatti o campagne commerciali / telemarketing
    devono essere obbligatoriamente verificati e filtrati rispetto al Registro Pubblico
    delle Opposizioni (D.P.R. n. 26/2022) prima del loro utilizzo.

    Attualmente la funzione restituisce tutti i contatti senza modifiche ed emette un
    warning nei log di conformità.
    """
    logger.warning(
        "ATTENZIONE (Conformità RPO): I contatti non sono stati filtrati rispetto al "
        "Registro Pubblico delle Opposizioni. Prima di utilizzare i numeri per scopi commerciali "
        "o telemarketing, verificare l'iscrizione al RPO (D.P.R. n. 26/2022)."
    )
    return contatti


def save_to_excel(name: str, comuni_list: List[str], contatti: List[Contatto], output_filename: str = None) -> str:
    """
    Exports collected Contatto objects to an Excel file with two sheets:
    - Sheet 'Riepilogo': Summary table of searched municipalities and counts.
    - Sheet 'Dati': Full dataset (Nome, Indirizzo, Telefono, Comune di Ricerca).
    """
    if not output_filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sanitized_name = re.sub(r'[^a-zA-Z0-9]', '_', name.strip().lower())
        filename = f"ricerca_{sanitized_name}_{timestamp}.xlsx"
    else:
        filename = output_filename

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

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)
        df_dati.to_excel(writer, sheet_name='Dati', index=False)

    logger.info(f"Esportazione completata con successo nel file: '{filename}'")
    return filename


def main():
    args = parse_arguments()
    setup_logging(args.verbose)

    target_name, comuni_list = get_user_inputs(args)

    driver = None
    all_extracted_contatti: List[Contatto] = []

    try:
        headless = not args.no_headless
        driver = setup_driver(headless=headless)

        for comune in comuni_list:
            try:
                search_municipality(driver, target_name, comune)
                comune_results = scrape_results_for_municipality(driver, comune, max_pages=args.max_pages)
                all_extracted_contatti.extend(comune_results)
            except Exception as e:
                logger.error(f"Errore durante l'elaborazione del comune '{comune}': {e}", exc_info=True)
                continue

    except Exception as e:
        logger.error(f"Si è verificato un errore critico durante l'esecuzione del browser: {e}", exc_info=True)
    finally:
        if driver:
            driver.quit()

    unique_contatti = deduplicate_contatti(all_extracted_contatti)
    filtered_contatti = filtra_rpo(unique_contatti)
    save_to_excel(target_name, comuni_list, filtered_contatti, output_filename=args.output)


if __name__ == "__main__":
    main()
