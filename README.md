
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
