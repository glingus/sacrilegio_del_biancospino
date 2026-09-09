# Fixtures

## `paginebianche_homepage.html`

Response body served by `https://www.paginebianche.it/persone?qs=Giulia&dv=Suzzara`
(saved 2026-09-07). Despite the search parameters, the site answers with its
**homepage**, not a result set:

- `rel="canonical"` points at `https://www.paginebianche.it/`
- `<h1>` is "Numeri, indirizzi ed orari"
- it contains zero `search-item` / `item-listing` / `search-itm` cards
- it contains zero no-results markers

That combination — no cards *and* no no-results message — is what
`scrape_results_for_municipality` now reports as a layout error, so the run exits
non-zero instead of writing an empty spreadsheet and claiming success.

The same file carries the endpoint the scraper should use, in the site's own
JSON-LD `SearchAction`:

```json
"target": "https://www.paginebianche.it/ricerca?qs={search_term_string}"
```

Kept as a reference for the search endpoint and as a sample of a non-results
page. It is not loaded by the test suite.
