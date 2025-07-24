# Extraction des KPI Zendesk

Ce script Python extrait des indicateurs clés à partir de tickets Zendesk.

## Fonctionnalités

- Récupère uniquement les tickets des types : `problem`, `task`, `question`, `incident`
- Calcule les délais de première réponse et de résolution
- Analyse la satisfaction client
- Exporte les résultats dans un fichier Excel

## Installation

Installez les dépendances avec :
pip install requests openpyxl tqdm


## Configuration

Modifiez dans le script les variables suivantes avec vos infos Zendesk :

- SUBDOMAIN  
- EMAIL  
- API_TOKEN  

## Usage

Lancez le script avec :

python zendesk_extract_KPI


Suivez les instructions pour choisir la plage de dates.

---

*Ne partagez jamais vos clés API publiquement !*
