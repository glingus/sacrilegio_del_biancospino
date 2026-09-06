#!/usr/bin/env python3
"""
Pagine Bianche Web Scraper
--------------------------
Scrapes personal contact data from Pagine Bianche (Italy) using Selenium.
Exports search results and summary statistics to an Excel file with pandas and openpyxl.
"""

import os
import re
import sys
import time
import random
import urllib.parse
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException
)
from webdriver_manager.chrome import ChromeDriverManager


def random_delay(min_sec=1.5, max_sec=3.5):
    """
    Introduces a random sleep delay to simulate human browsing behavior.
    """
    time.sleep(random.uniform(min_sec, max_sec))


def get_user_inputs():
    """
    Prompts the user via CLI for search query parameters.
    Returns target_name and list of municipalities.
    """
    print("=" * 60)
    print("      Pagine Bianche Scraper - Persona Fisica (Privati)")
    print("=" * 60)

    target_name = input("Quale nome/cognome stai cercando? ").strip()
    while not target_name:
        print("Il nome/cognome non può essere vuoto.")
        target_name = input("Quale nome/cognome stai cercando? ").strip()

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


def setup_driver(headless=True):
    """
    Configures and initializes Chrome WebDriver using webdriver_manager.
    """
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
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
    driver.maximize_window()
    return driver


def accept_cookies(driver, timeout=5):
    """
    Detects and clicks the cookie acceptance banner button if present.
    """
    try:
        cookie_selectors = [
            "//button[contains(@id, 'iubenda-cs-accept-btn')]",
            "//button[contains(@class, 'iubenda-cs-accept-btn')]",
            "//button[contains(text(), 'Accetta tutti') or contains(text(), 'ACCETTA') or contains(text(), 'Accetta')]",
            "//a[contains(text(), 'Accetta') or contains(text(), 'ACCETTA')]",
            "//button[@id='onetrust-accept-btn-handler']"
        ]

        wait = WebDriverWait(driver, timeout)
        for selector in cookie_selectors:
            try:
                btn = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                driver.execute_script("arguments[0].click();", btn)
                print("[+] Banner cookie accettato.")
                random_delay(1.0, 2.0)
                return True
            except TimeoutException:
                continue
    except Exception:
        pass
    return False


def search_municipality(driver, name, comune):
    """
    Navigates to Pagine Bianche search page for the given name and municipality.
    """
    encoded_name = urllib.parse.quote(name)
    encoded_comune = urllib.parse.quote(comune)
    url = f"https://www.paginebianche.it/cerca?qs={encoded_name}&dv={encoded_comune}"

    print(f"\n[->] Ricerca per '{name}' a '{comune}' -> {url}")
    driver.get(url)
    random_delay(2.0, 4.0)
    accept_cookies(driver)


def click_show_phone_button(driver, card_element):
    """
    Attempts to click 'Mostra numero' button inside a result card if present.
    """
    try:
        phone_buttons = card_element.find_elements(
            By.XPATH,
            ".//button[contains(text(), 'Mostra') or contains(text(), 'numero') or contains(@class, 'phone')]"
            " | .//a[contains(text(), 'Mostra') or contains(text(), 'numero') or contains(@class, 'phone')]"
        )
        for btn in phone_buttons:
            if btn.is_displayed():
                try:
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.5)
                except Exception:
                    pass
    except Exception:
        pass


def extract_card_data(driver, card):
    """
    Extracts Name, Address, and Phone from a single result card element.
    Uses safe try/except fallback mechanisms.
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
            elem = card.find_element(By.XPATH, sel)
            text = elem.text.strip()
            if text:
                nome = text
                break
        except NoSuchElementException:
            continue

    # Try clicking "Mostra numero" if button exists
    click_show_phone_button(driver, card)

    # Phone extraction
    telefono = ""
    phone_selectors = [
        ".//*[contains(@class, 'phone') or contains(@class, 'tel')]",
        ".//a[contains(@href, 'tel:')]",
        ".//*[contains(text(), 'Tel') or contains(text(), '3') or contains(text(), '0')]"
    ]
    for sel in phone_selectors:
        try:
            elem = card.find_element(By.XPATH, sel)
            text = elem.text.strip()
            if text and ("0" in text or "3" in text or "+" in text):
                phone_match = re.search(r'(\+?\d[\d\s\-\.\/]{5,}\d)', text)
                if phone_match:
                    telefono = phone_match.group(1).strip()
                else:
                    telefono = text
                if telefono:
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
            elem = card.find_element(By.XPATH, sel)
            text = elem.text.strip()
            if text:
                indirizzo = text
                break
        except NoSuchElementException:
            continue

    # Fallback if specific address sub-selectors fail: parse card text lines
    if not indirizzo or not nome:
        card_text = card.text.strip().split("\n")
        if card_text:
            if not nome and len(card_text) > 0:
                nome = card_text[0]
            if not indirizzo and len(card_text) > 1:
                for line in card_text[1:]:
                    if any(kw in line.lower() for kw in ['via', 'viale', 'piazza', 'corso', 'largo', 'strada']) or re.search(r'\b\d{5}\b', line):
                        indirizzo = line.strip()
                        break

    return {
        "Nome": nome,
        "Indirizzo": indirizzo,
        "Telefono": telefono
    }


def scrape_results_for_municipality(driver, comune_ricerca):
    """
    Scrapes all pages of search results for a single municipality.
    Returns a list of extracted record dictionaries.
    """
    results = []
    page_num = 1

    while True:
        print(f"  [+] Estrazione pagina {page_num} per comune: {comune_ricerca}...")
        random_delay(1.0, 2.5)

        card_selectors = [
            "//div[contains(@class, 'search-item')]",
            "//div[contains(@class, 'item-listing')]",
            "//div[contains(@class, 'search-itm')]",
            "//article",
            "//div[contains(@class, 'result-item')]",
            "//div[contains(@class, 'card')]"
        ]

        cards = []
        for selector in card_selectors:
            found = driver.find_elements(By.XPATH, selector)
            if found:
                cards = found
                break

        if not cards:
            print(f"  [-] Nessun risultato trovato a pagina {page_num} per {comune_ricerca}.")
            break

        extracted_on_page = 0
        for card in cards:
            try:
                data = extract_card_data(driver, card)
                if data["Nome"]:
                    data["Comune di Ricerca"] = comune_ricerca
                    results.append(data)
                    extracted_on_page += 1
            except Exception as e:
                print(f"  [!] Errore durante l'estrazione di una scheda: {e}")
                continue

        print(f"  [v] Estratte {extracted_on_page} persone da pagina {page_num}.")

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
            print(f"  [->] Passaggio alla pagina successiva ({page_num + 1})...")
            try:
                driver.execute_script("arguments[0].click();", next_button)
                page_num += 1
                random_delay(2.0, 4.0)
            except Exception as e:
                print(f"  [!] Impossibile cliccare 'Pagina successiva': {e}")
                break
        else:
            print(f"  [*] Fine dei risultati paginati per {comune_ricerca}.")
            break

    return results


def save_to_excel(name, comuni_list, all_data):
    """
    Exports collected data to an Excel file with two sheets:
    - Sheet 'Riepilogo': Summary table of searched municipalities and counts.
    - Sheet 'Dati': Full dataset (Nome, Indirizzo, Telefono, Comune di Ricerca).
    """
    timestamp = datetime.now().strftime("%Y%m%d")
    sanitized_name = re.sub(r'[^a-zA-Z0-9]', '_', name.strip().lower())
    filename = f"ricerca_{sanitized_name}_{timestamp}.xlsx"

    # Calculate summary per municipality
    summary_dict = {c: 0 for c in comuni_list}
    for row in all_data:
        comune = row.get("Comune di Ricerca")
        if comune in summary_dict:
            summary_dict[comune] += 1
        else:
            summary_dict[comune] = 1

    df_riepilogo = pd.DataFrame([
        {"Comune di Ricerca": comune, "Numero Totale Persone Trovate": count}
        for comune, count in summary_dict.items()
    ])

    if all_data:
        df_dati = pd.DataFrame(all_data)[["Nome", "Indirizzo", "Telefono", "Comune di Ricerca"]]
    else:
        df_dati = pd.DataFrame(columns=["Nome", "Indirizzo", "Telefono", "Comune di Ricerca"])

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_riepilogo.to_excel(writer, sheet_name='Riepilogo', index=False)
        df_dati.to_excel(writer, sheet_name='Dati', index=False)

    print(f"\n[+] Esportazione completata con successo nel file: '{filename}'")
    return filename


def main():
    target_name, comuni_list = get_user_inputs()

    driver = None
    all_extracted_data = []

    try:
        driver = setup_driver(headless=True)

        for comune in comuni_list:
            search_municipality(driver, target_name, comune)
            comune_results = scrape_results_for_municipality(driver, comune)
            all_extracted_data.extend(comune_results)

    except Exception as e:
        print(f"\n[!] Si è verificato un errore durante l'esecuzione del browser: {e}")
    finally:
        if driver:
            driver.quit()

    save_to_excel(target_name, comuni_list, all_extracted_data)


if __name__ == "__main__":
    main()
