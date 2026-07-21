
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
