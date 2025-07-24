# Extraction des KPI Zendesk

Ce script Python extrait des indicateurs clés à partir de tickets Zendesk.

## Sommaire
- [Fonctionnalités](#fonctionnalités)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Avertissement](#avertissement)

## Fonctionnalités

- Récupère uniquement les tickets des types : `problem`, `task`, `question`, `incident`
- Calcule les délais de première réponse et de résolution
- Analyse la satisfaction client
- Exporte les résultats dans un fichier Excel

## Installation

Installez les dépendances, tapez CMD, puis dans l'invite de commandes tapez :
pip install requests openpyxl tqdm

Si vous n'avez pas Python installé, vous pouvez le télécharger ici :
https://www.python.org/downloads/


## Configuration

Avant d'exécuter le script, modifiez les variables suivantes dans `zendesk_extract_KPI.py` :

- `SUBDOMAIN` : votre sous-domaine Zendesk  
- `EMAIL` : votre email Zendesk  
- `API_TOKEN` : votre token API Zendesk (ne partagez jamais ce token publiquement)

## Usage

Lancez le script avec :

python zendesk_extract_KPI


Suivez les instructions pour choisir la plage de dates.

---

⚠️ Ne partagez jamais vos clés API publiquement !

![Texte alternatif](images/nom_de_ton_image.png)

