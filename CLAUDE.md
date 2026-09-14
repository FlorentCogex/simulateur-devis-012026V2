# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A static, no-build web app for COGEX (cabinet comptable) that lets staff simulate a fee quote ("devis d'honoraires") for a prospective client and generate a matching "lettre de mission" (engagement letter) document. There is no backend, no package.json, and no build/test tooling — everything is plain HTML/CSS/JS served as-is (e.g. via GitHub Pages or opened directly in a browser).

## Files

- `index.html` — the simulator itself. Single-page app with client-side routing done by toggling `.page` visibility (`goToPage()`), covering three "pages": home, `societes-page` (companies with share capital: SAS/SARL/SCI), `bnc-page` (professions libérales / no capital).
- `COGEX_Lettre_Mission_Societes_FINALE.html` / `COGEX_Lettre_Mission_BNC_FINALE.html` — the two engagement-letter documents (one per client type), opened in a new tab from the simulator and meant to be printed to PDF via `window.print()`.
- `COGEX_Lettre_Confrere.html` — the deontological "reprise de dossier" letter sent to a client's previous accountant, opened from either quote page via the "✉️ Courrier Confrère" button. Same `{{PLACEHOLDER}}` / `localStorage` pattern as the two mission letters, but keyed under a separate `confrere_*` prefix (`confrere_civilite`, `confrere_nom`, `confrere_cabinet`, `confrere_adresse`, `confrere_client_nom`, `confrere_dirigeant`, `confrere_date_debut`, `confrere_lieu_date`, `confrere_signataire_nom`, `confrere_signataire_titre`) so it never collides with the mission-letter keys.
- `logo-cogex.jpg`, `Logoneg.jpg`, `logo.jpg` — brand assets referenced by the HTML. (`logo-cogex.jpg` was renamed from a filename containing a decomposed-Unicode accented character, which 404'd depending on the OS/server — keep new asset filenames plain ASCII.)

There is no dev server, linter, formatter, or test suite in this repo. To check a change, open the HTML file directly in a browser (or serve the folder with any static file server) and click through it.

## Architecture: data flow between the documents

`index.html` and the three letter files (`COGEX_Lettre_Mission_*_FINALE.html`, `COGEX_Lettre_Confrere.html`) are **not connected by a JS module or API** — they communicate purely through the browser's `localStorage`, keyed by plain string keys (`nom_client`, `formuleSelectionnee`, `honorairesComptables`, `honorairesSociale`, `pvAgo`, `registresLegaux`, `revenusFonciers`, `fraisChancellerie`, `totalAnnuel`, `nbSalaries`, `informatiqueAnnuel`, `informatiqueCommentaire`, `lignesLibres` (JSON array), plus company-specific fields like `capital`, `forme_juridique`, `dirigeants`, `qualite_dirigeant`, `adresse`, `activite`, `date_creation`, `date_debut`, `date_fin`). `genererLettre(type)` in `index.html` writes all of these before calling `window.open(...)` on the relevant mission-letter file, which reads them back on load via `{{PLACEHOLDER}}` string replacement on `document.body.innerHTML`. When adding a new field to the quote, it must be written in `genererLettre()` **and** given a matching `{{PLACEHOLDER}}` + `replaceAll` line in **both** `COGEX_Lettre_Mission_*_FINALE.html` files — the three are easy to get out of sync since nothing enforces the shared key/placeholder names. `genererCourrierConfrere()` follows the same pattern independently for `COGEX_Lettre_Confrere.html`, under the `confrere_*` key prefix.

Each client type (Sociétés vs BNC) duplicates its own set of element IDs, calculation functions, and pricing logic rather than sharing code (e.g. `calculerDevisSocietes()` / `calculerDevisBNC()`, `selectFormuleSocietes()` / `selectFormuleBNC()`). Changes to pricing/business logic typically need to be made in both places. The saved-dossier field lists (`champsSocietes` / `champsBNC` arrays in `index.html`) must also be kept in sync with whatever input IDs exist per type — a field left out of these arrays silently won't be saved/restored by the dossier history feature below.

## Pricing logic (in `index.html`)

- `grillesSoc` / `grillesBNC` are CA-bracket tables (min/max annual revenue → monthly fee per formule: ESSENTIEL/PILOTAGE/PERFORMANCE for Sociétés, ESSENTIEL/PILOTAGE for BNC). `calculerTarifsSocietes()`/`calculerTarifsBNC()` pick the bracket from the CA input and populate the pricing cards (displayed highest-tier-first: PERFORMANCE → PILOTAGE → ESSENTIEL).
- A user can instead pick "saisie manuelle" (`toggleManualMode`) to enter a custom formule name and monthly price directly, bypassing the grid.
- `calculerDevisSocietes()` / `calculerDevisBNC()` sum accounting fees (mensuel × 12), social fees (`nbSalaries × tarifBulletin × 12`, where `tarifBulletin` is a per-dossier editable field defaulting to 31€), legal fees (PV d'AG + registres légaux, Sociétés only), IRPP, revenus fonciers (`nb biens × 150`), an "informatique" line (custom monthly × 12), and the sum of any "lignes libres" (see below), then add an editable "frais de chancellerie" percentage (default 3%, edited inline in the ticket sidebar via `#tauxChancellerieSocietes`/`BNC`) on the subtotal for the annual total.
- Secteur d'activité (`BTP`, `Carrosserie`, `CHR`) applies a +10% multiplier to Sociétés pricing.

## Lignes libres (free-form billing lines)

`lignesLibresSocietes` / `lignesLibresBNC` (JS arrays of `{description, montant}`) let staff add arbitrary one-off billing lines per dossier via "+ Ajouter une ligne" in the "Prestations Additionnelles" section. `renderLignesLibres(type)` re-renders the list from the array on every add/edit/remove; their sum feeds into `calculerDevisSocietes()`/`BNC()` and is also written to `localStorage` as JSON (`lignesLibres` key) so the mission letters can render them as extra `{{LIGNES_LIBRES}}` rows in the honoraires table.

## Dossiers sauvegardés (saved-quote history)

To avoid re-spending a Pappers API credit just to re-open/correct an already-generated quote, every dossier can be saved to a `dossiersCogex` JSON array in `localStorage` (list of `{id, type, savedAt, ...all champsSocietes/champsBNC field values, selFormule, prixBase, lignesLibres, totalAnnuel}`). "💾 Enregistrer le dossier" (`sauvegarderDossier`) upserts the current form into that array — a hidden `#dossierIdSocietes`/`#dossierIdBNC` field tracks whether the current session is already backed by a saved dossier (update in place) or not (create new). `genererLettre()` also silently upserts on every letter generation. "📂 Mes dossiers sauvegardés" on the home page opens a searchable list (`renderListeDossiers`) to reopen (`ouvrirDossier` → `appliquerDossier`, which repopulates every field, recalculates, and restores the selected formule/lignes libres) or delete a dossier. `resetData()` ("Nouveau Dossier") clears `localStorage` but explicitly preserves `dossiersCogex` first — don't remove that preservation when touching `resetData()`, or saved history gets wiped on every reset.

## SIRET lookup (Pappers API)

`rechercherSiret()` calls the Pappers API directly from the browser using a hardcoded API key (`PAPPERS_API_KEY` in `index.html`). This is a client-exposed key with no backend proxy — the in-code error handling already anticipates CORS failures in the browser (the alert tells users to install a CORS-bypass extension). Keep this constraint in mind before suggesting "just call the API" fixes; a real fix would require a server-side proxy.

## Editing conventions already in use

- French UI strings and French variable/function names are the norm (`calculerDevis...`, `dossier`, `formule`, etc.) — follow this convention for consistency rather than switching to English.
- Money is formatted via `formatNumber()` using `toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })`; reuse it rather than reformatting inline.
- Styling is a single inline `<style>` block per file using CSS custom properties (`--bg-deep`, `--primary`, `--accent`, etc.) defined on `:root` — reuse existing variables/classes instead of introducing new colors.
