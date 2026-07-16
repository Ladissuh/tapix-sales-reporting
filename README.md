# Tapix Sales Reporting Pipeline

Automatický týdenní generátor Sales Reporting reportu z HubSpotu.
Nahrazuje: 2 GitHub skripty + ruční doplňování Won/Lost + SharePoint Excel
formule. Výstup je hotový `.xlsx` s grafy, který Make.com nahraje na
SharePoint.

## 1. Založení repozitáře na GitHubu

1. Jdi na [github.com/new](https://github.com/new).
2. **Repository name:** `tapix-sales-reporting` (nebo cokoliv rozumného -
   ale ať je z názvu jasné, o co jde, kdyby si to za rok prohlížel někdo
   jiný).
3. **Visibility: Private.** Tohle je důležité - jsou tam obchodní data
   (i když v CSV, ne přímo v HubSpotu), takže veřejné repo rozhodně ne.
4. Nezaškrtávej "Add README" (README už máš tady) - nebo ho přidej a pak
   při nahrávání souborů ho přepiš tímhle.
5. Vytvoř repozitář.

## 2. Nahrání souborů

Nejjednodušší cesta bez příkazové řádky:
1. V novém (prázdném) repu klikni **"uploading an existing file"**.
2. Přetáhni tam celou tuhle složku (se zachováním struktury podsložek -
   GitHub webové rozhraní umí drag&drop celých složek v moderních
   prohlížečích; pokud ne, nahraj po jednotlivých podsložkách).
3. Commitni rovnou do `main` větve.

Případně přes Git (pokud preferuješ):
```bash
cd tapix-sales-reporting
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<tvůj-účet>/tapix-sales-reporting.git
git push -u origin main
```

## 3. Nastavení HubSpot tokenu (GitHub Secret)

1. V repu: **Settings → Secrets and variables → Actions**.
2. **New repository secret.**
3. Name: `HUBSPOT_TOKEN`
4. Value: tvůj HubSpot Private App token (`pat-...`) - stejný, jaký
   používaly i dnešní 2 skripty.
5. **Add secret.**

## 4. Povolení zápisu pro GitHub Actions

Workflow si commituje aktualizovaná data zpátky do repa, takže potřebuje
právo zápisu:
1. **Settings → Actions → General → Workflow permissions.**
2. Zaškrtni **"Read and write permissions"**.
3. Save.

(Bez tohohle kroku bude workflow padat na `git push` s chybou oprávnění.)

## 5. Kontrola/otestování

1. Záložka **Actions** → vyber workflow **"Weekly Sales Report"**.
2. **Run workflow** (tlačítko vpravo) → spustí se ručně, nemusíš čekat na
   pondělí.
3. Sleduj log - měl by stáhnout dealy, zapsat `data/deals_snapshots.csv`
   a `data/deals_closed_ledger.csv`, a vygenerovat
   `outputs/Sales_Reporting_2026.xlsx`.
4. Po doběhnutí by měl workflow soubory rovnou commitnout zpátky do repa -
   zkontroluj složky `data/` a `outputs/`.

## 6. Napojení na Make.com

Dosavadní Make scénář, který stahoval soubory z GitHubu přes GitHub API a
posílal je do SharePointu, teď stačí zjednodušit:
- Zdrojový soubor: `outputs/Sales_Reporting_<rok>.xlsx` (cesta v repu).
- Cílová akce v SharePointu: **nahradit/přepsat** existující soubor (žádné
  přepisování formulí ani buněk - jen upload hotového souboru).
- Zvaž spouštět Make scénář *po* doběhnutí GitHub Actions (např. přes
  webhook na konci workflow, nebo jen s dostatečným časovým odstupem v
  rozvrhu), ať nenahráváš starou verzi souboru.

## 7. Struktura repozitáře

```
tapix-sales-reporting/
├── .github/workflows/weekly_report.yml   # týdenní běh (pondělí) + manuální spuštění
├── config/
│   ├── owners.yaml                        # WHITELIST obchodníků + mapování jmen + cíle - uprav kdykoliv
│   └── stage_probabilities.yaml          # váhy fází - uprav kdykoliv
├── data/
│   ├── deals_snapshots.csv               # historie otevřené pipeline (roste každý týden)
│   └── deals_closed_ledger.csv           # Won/Lost ledger (nahrazuje ruční práci)
├── outputs/
│   └── Sales_Reporting_<rok>.xlsx        # finální report, generuje se při každém běhu
├── src/
│   ├── hubspot_client.py                 # veškerá komunikace s HubSpot API
│   ├── data_store.py                     # čtení/zápis historických CSV dat
│   ├── metrics.py                        # všechny výpočty (nahrazuje SharePoint formule)
│   ├── report_builder.py                 # vykreslení Excelu (styly, tabulky, grafy)
│   └── main.py                           # orchestrace celého běhu
├── requirements.txt
├── .env.example                          # pro lokální testování (zkopíruj jako .env)
└── .gitignore
```

## 8. Lokální testování (nepovinné)

```bash
cd tapix-sales-reporting
pip install -r requirements.txt
cp .env.example .env      # a doplň skutečný HUBSPOT_TOKEN
python src/main.py
```

## 9. Historický import (už hotovo, jen pro info)

Data z tvého dosavadního `Sales_reporting_2026.xlsx` (týdny 1–28, tj. od
4.1. do konce června) už jsou naimportovaná v `data/deals_snapshots.csv`
a `data/deals_closed_ledger.csv` - stačí je commitnout spolu se zbytkem
repa a GitHub Actions na ně od příštího pondělí jen přidá nové týdny.

Pár poznámek k importu:
- U historických Won/Lost řádků chybí **Company** (název firmy) - dosavadní
  ruční tabulka měla jen Company ID, ne název, a bez HubSpot tokenu jsem ho
  nemohl dohledat. Nové (automaticky stahované) Won/Lost řádky už název
  firmy mít budou. Pokud chceš historii dopočítat, spusť lokálně (s
  nastaveným `.env`): `python src/import_historical.py cesta/k/souboru.xlsx`
  - tentokrát se company názvy dotáhnou z HubSpotu automaticky.
- Historické stage částky (Tabulka 1) jsou importované jako souhrn za
  celou pipeline, ne po jednotlivých dealech - to je jediná informace,
  kterou starý systém uchovával. Od teď už se bude ukládat na úrovni
  jednotlivého dealu (přesnější podklad pro průměrnou velikost dealu,
  funnel apod.).
- Do reportu se z historie promítnou jen obchodníci uvedení v
  `config/owners.yaml` - `Ainaz` (už ve firmě není) se do importu vůbec
  nezahrnula.

## 10. Co dělat, když...

- **Přibude nový obchodník:** doplň ho do `config/owners.yaml` (hubspot_name
  musí přesně sedět se jménem v HubSpotu, display_name je název listu,
  annual_goal klidně `null`, pokud cíl ještě nemá). Kdo NENÍ v tomto
  souboru, se v reportu vůbec neobjeví - i kdyby v HubSpotu vlastnil dealy.
- **Změní se váhy fází nebo přibude nová stage:** uprav
  `config/stage_probabilities.yaml` - nic jiného se měnit nemusí.
- **Chceš zobrazovat víc/míň týdnů historie:** `DISPLAY_WEEKS` konstanta
  v `src/main.py`.
- **HubSpot pipeline nepoužívá konvenci `probability: 1.0/0.0` pro
  Closed Won/Lost:** uprav detekci v `hubspot_client.get_stage_metadata()`.
