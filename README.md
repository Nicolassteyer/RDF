# Analyse RDF Flams

Application Streamlit locale pour analyser :
- un fichier personnes (`xlsx`, `xls` ou `csv`)
- un ou plusieurs rapports tickets (`html` / `htm`)

## Fonctionnalités
- import multiple de fichiers HTML
- détection des remises :
  - RDF DESSERT
  - RDF CAFE
  - RDF BOISSON
  - RDF ELSASSICH
- tableau de bord :
  - nombre de personnes par jour et par lot
  - nombre de remises par type
  - montant total des remises
  - montant total to pay des tickets RDF
  - ROI = `total to pay des tickets RDF ÷ montant total des remises RDF`
  - part des remises sur le total to pay
  - répartition des personnes par jour en restaurant à partir des dates présentes dans les tickets HTML
- filtres :
  - dates
  - types RDF
  - lots
  - fichiers HTML
  - recherche par note
- export Excel

## Installation
```bash
python -m venv .venv
```

### Windows PowerShell
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

### macOS / Linux
```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Colonnes attendues dans le fichier personnes
- `PRENOM`
- `NOM`
- `LOT`
- `DATE` ou `ADDED_TIME`

## Notes
- L’application ne conserve que les tickets contenant au moins une remise RDF ciblée.
- Le tableau `personnes_restaurant` est basé sur les dates détectées dans les tickets HTML chargés.
- En cas de doute sur la définition métier du ROI, adapte facilement la formule dans `compute_roi()` dans `app.py`.
