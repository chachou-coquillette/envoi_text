# envoi_text

Envoyer un texto à plusieurs téléphones via l'application **Phone Link** (Windows).

## Prérequis

- Windows 10 / 11
- L'application [Phone Link](https://www.microsoft.com/fr-fr/windows/sync-across-your-devices) installée et connectée à un téléphone Android avec les permissions SMS accordées.
- Python 3.11+

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

Éditez **`config.py`** pour personnaliser :

| Variable | Description | Valeur par défaut |
|---|---|---|
| `CONTACTS_FILE` | Chemin vers le fichier CSV des contacts | `contacts.csv` |
| `MESSAGE` | Message à envoyer | `"Bonjour, ceci est un message automatique."` |
| `IMAGE_FILE` | Chemin vers une image à joindre (optionnel, MMS) | `""` |
| `DELAY_BETWEEN_MESSAGES` | Délai (secondes) entre deux envois | `5` |
| `UI_TIMEOUT` | Délai d'attente de l'interface (secondes) | `10` |

## Fichier de contacts

Créez (ou modifiez) **`contacts.csv`** avec les colonnes `name` et `phone` :

```csv
name,phone
Alice Dupont,+33612345678
Bob Martin,+33623456789
```

## Utilisation

1. Assurez-vous que Phone Link est ouvert et que votre téléphone est connecté.
2. Lancez l'agent :

```bash
python agent.py
```

L'agent ouvrira automatiquement Phone Link si l'application n'est pas déjà lancée, puis enverra le message configuré à chaque contact du fichier CSV.
