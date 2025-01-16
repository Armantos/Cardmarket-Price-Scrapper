# Cardmarket Price Scraper

## Présentation

Ce script Python permet d'extraire les prix depuis Cardmarket à partir d'URL dans un tableau Excel ou Google Sheets en utilisant Selenium avec ChromeDriver. Il fonctionne pour tous les produits (booster, display, etc.) ainsi que pour les cartes à l'unité, et est compatible avec tous les TCG.

---

## Explication du fonctionnement

1. Ouvre un fichier Excel ou Google Sheets.
2. Lit une liste d'URL dans une colonne.
3. Lance automatiquement Google Chrome.
4. Se connecte à un compte Cardmarket.
5. Parcourt chaque URL et récupère :
   - Le premier prix et les frais de port pour les produits.
   - Le premier prix uniquement pour les cartes à l'unité.
6. Si une cellule de la colonne URL est vide, les prix seront renseignés à `0`.
7. Écrit les résultats dans le fichier Excel ou Google Sheets.

---

## Prérequis

- **Google Chrome** installé : [Lien de téléchargement](https://www.google.com/chrome/)
- **Python** installé : [Lien de téléchargement](https://www.python.org/downloads/)  
  ⚠️ N'oubliez pas de cocher la case "Ajouter au PATH" lors de l'installation de Python

---

## Installation

### Étape 1 : Télécharger le projet 

Clic droit > Ouvrir dans le terminal 
(ou Commande Prompt / Powershell dans la barre de recherche)

```bash
git clone https://github.com/Armantos/Cardmarket-Price-Scrapper
cd Cardmarket-Price-Scrapper
```

### Étape 2 : Installer les dépendances

```bash
pip install selenium webdriver-manager python-dotenv google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client openpyxl
```

---

## Paramétrage

### Configuration des colonnes

Dans votre fichier Excel ou Google Sheets, identifiez ou configurez 4 colonnes :

1. **Colonne des URL** : Contient les liens vers les produits sur Cardmarket.  
   Si une cellule est vide, les cases de résultat seront remplies à `0`.
2. **Colonne des prix** : Recevra le prix des produits.
3. **Colonne des frais de port** : Recevra les frais de port (ou `0` si non applicable).
4. **Colonne du total** : Calcul automatique du prix total (prix + frais de port).

Utilisez le fichier `test.xlsx` comme exemple.

### Configuration du fichier `.env`

1. Renommez le fichier `.env.example` en `.env`.
2. Modifiez le fichier `.env` avec un éditeur de texte. 
Modifier les identifiants et mot de passe du compte Cardmarket
Exemple :
   ```
   CARDMARKET_USERNAME=VotreNomUtilisateur
   CARDMARKET_PASSWORD=VotreMotDePasse
   ```

1. Définissez les paramètres de votre fichier Excel ou Google Sheets :
   ```
   SHEET_NAME=NomDeLaFeuille
   URL_COLUMN=ColonneDesURLs
   BUYING_PRICE_COLUMN=ColonneDesPrix
   SHIPPING_PRICE_COLUMN=ColonneDesFrais
   TOTAL_PRICE_COLUMN=ColonneTotal
   NUMBER_OF_URL_ROWS=NombreDeLignesÀTraiter
   ```

   **Note** : Si vous utilisez Google Sheets, configurez également l'API Google Sheets (voir ci-dessous).

---

### Note sur l'url

N'oubliez pas de configurer les paramètres du produit (langue, condition de la carte, pays du vendeur, etc.) et de cliquer sur "Filtrer" avant de copier l'URL dans la feuille. Par exemple :

```
https://www.cardmarket.com/en/Pokemon/Products/Singles/Evolving-Skies/Umbreon-VMAX-V3?language=2&minCondition=1
```

- `?language=2` : La carte est en français.  
- `minCondition=1` : La condition est **mint**.

---

### Utilisation avec Excel

1. Modifiez les paramètres suivants dans `.env` :
   ```
   SHEETS_OR_EXCEL=EXCEL
   EXCEL_NAME=NomDeVotreFichier.xlsx
   ```
2. Exemple de chemin pour `EXCEL_NAME` : `D:/Utilisateurs/John/Bureau/mon_excel.xlsx`.

### Utilisation avec Google Sheets

1. Modifiez les paramètres suivants dans `.env` :
   ```
   SHEETS_OR_EXCEL=SHEETS
   SPREADSHEET_ID=VotreIdentifiantGoogleSheet
   ```

   L'identifiant est visible dans l'URL de votre Google Sheet.  
   Exemple : `https://docs.google.com/spreadsheets/d/1gvDxxxxxxxxxxxxxxxxxxxxx`

2. Configurez l'API Google Sheets :
   - Si besoin, suivez ce [tutoriel vidéo](https://youtu.be/K6Vcfm7TA5U?t=214).
   - Activez l'API Google Sheets sur [Google Cloud Console](https://console.cloud.google.com/marketplace/product/google/sheets.googleapis.com).
   - Créez un compte de service, exemple `nomprenom@nom-du-projet.iam.gserviceaccount.com` et téléchargez la clé au format JSON.
   - Renommez ce fichier en `secrets.json` et placez-le dans le dossier à côte du script `cardmarket-price-scrapper.py`. Voir le fichier `secrets-example.json` comme exemple.
   -  Sur la page du Google Sheet, bouton partager (en haut à droite) et indiquer le mail de service crée précédemment
  

---

## Exécution du script

1. **Si vous utilisez Excel, assurez-vous que le fichier est bien fermé.**
2. Lancez la commande suivante dans le terminal pour extraire les prix :

```bash
python cardmarket-price-scrapper.py
```

3. Vous pouvez réduire la fenêtre Chrome, mais ne la fermez pas.

---

## Améliorations futures (TODO)

- Faire fonctionner le script en tâche de fond (sans ouvrir le navigateu, mode "headless").
- Héberger le script dans le cloud pour une exécution continue.
- Simplifier le script en utilisant l'API Cardmarket.
- Supprimer la variable `NUMBER_OF_URL_ROWS` et gérer automatiquement les lignes vides (sans la valeur, le script tourne à l'infini car il prend en compte les cases d'url vides).

---

## Remerciements

Remerciements à **Thomas** pour son aide dans la réalisation de ce script. Cimer chef 🙏

---
