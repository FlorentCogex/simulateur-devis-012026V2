# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A static, no-build web app for COGEX (cabinet comptable) that lets staff simulate a fee quote ("devis d'honoraires") for a prospective client and generate a matching "lettre de mission" (engagement letter) document. There is no backend, no package.json, and no build/test tooling — everything is plain HTML/CSS/JS served as-is (e.g. via GitHub Pages or opened directly in a browser).

## Files

- `index.html` — the simulator itself. Single-page app with client-side routing done by toggling `.page` visibility (`goToPage()`), covering three "pages": home, `societes-page` (companies with share capital: SAS/SARL/SCI), `bnc-page` (professions libérales / no capital).
- `COGEX_Lettre_Mission_Societes_FINALE.html` / `COGEX_Lettre_Mission_BNC_FINALE.html` — the two engagement-letter documents (one per client type), opened in a new tab from the simulator and meant to be printed to PDF via `window.print()`.
- `Logo-COGEX-ok-vectoriseì-950-Ko.jpg`, `Logoneg.jpg`, `logo.jpg` — brand assets referenced by the HTML.

There is no dev server, linter, formatter, or test suite in this repo. To check a change, open the HTML file directly in a browser (or serve the folder with any static file server) and click through it.

## Architecture: data flow between the two documents

`index.html` and the two `COGEX_Lettre_Mission_*_FINALE.html` files are **not connected by a JS module or API** — they communicate purely through the browser's `localStorage`, keyed by plain string keys (`nom_client`, `formuleSelectionnee`, `honorairesComptables`, `honorairesSociale`, `pvAgo`, `registresLegaux`, `revenusFonciers`, `fraisChancellerie`, `totalAnnuel`, `nbSalaries`, `informatiqueAnnuel`, `informatiqueCommentaire`, plus company-specific fields like `capital`, `forme_juridique`, `dirigeants`, `qualite_dirigeant`, `adresse`, `activite`, `date_creation`, `date_debut`, `date_fin`). `genererLettre(type)` in `index.html` writes all of these before calling `window.open(...)` on the relevant letter file, which reads them back on load. When adding a new field to the quote, it must be written in `genererLettre()` in `index.html` **and** read on the corresponding key in the matching `COGEX_Lettre_Mission_*_FINALE.html` file — the two are easy to get out of sync since nothing enforces the shared key names.

Each client type (Sociétés vs BNC) duplicates its own set of element IDs, calculation functions, and pricing logic rather than sharing code (e.g. `calculerDevisSocietes()` / `calculerDevisBNC()`, `selectFormuleSocietes()` / `selectFormuleBNC()`). Changes to pricing/business logic typically need to be made in both places.

## Pricing logic (in `index.html`)

- `grillesSoc` / `grillesBNC` are CA-bracket tables (min/max annual revenue → monthly fee per formule: ESSENTIEL/PILOTAGE/PERFORMANCE for Sociétés, ESSENTIEL/PILOTAGE for BNC). `calculerTarifsSocietes()`/`calculerTarifsBNC()` pick the bracket from the CA input and populate the pricing cards.
- A user can instead pick "saisie manuelle" (`toggleManualMode`) to enter a custom formule name and monthly price directly, bypassing the grid.
- `calculerDevisSocietes()` / `calculerDevisBNC()` sum accounting fees (mensuel × 12), social fees (`nbSalaries × 31 × 12`), legal fees (PV d'AG + registres légaux, Sociétés only), IRPP, revenus fonciers (`nb biens × 150`), and an "informatique" line (custom monthly × 12), then add a flat 3% "frais de chancellerie" on the subtotal for the annual total.
- Secteur d'activité (`BTP`, `Carrosserie`, `CHR`) applies a +10% multiplier to Sociétés pricing.

## SIRET lookup (Pappers API)

`rechercherSiret()` calls the Pappers API directly from the browser using a hardcoded API key (`PAPPERS_API_KEY` in `index.html`). This is a client-exposed key with no backend proxy — the in-code error handling already anticipates CORS failures in the browser (the alert tells users to install a CORS-bypass extension). Keep this constraint in mind before suggesting "just call the API" fixes; a real fix would require a server-side proxy.

## Editing conventions already in use

- French UI strings and French variable/function names are the norm (`calculerDevis...`, `dossier`, `formule`, etc.) — follow this convention for consistency rather than switching to English.
- Money is formatted via `formatNumber()` using `toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })`; reuse it rather than reformatting inline.
- Styling is a single inline `<style>` block per file using CSS custom properties (`--bg-deep`, `--primary`, `--accent`, etc.) defined on `:root` — reuse existing variables/classes instead of introducing new colors.
