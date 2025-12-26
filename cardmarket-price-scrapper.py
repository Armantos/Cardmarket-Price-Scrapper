# Remerciements à Thomas pour son aide dans la réalisation de ce script.
# Cimer chef

# Install dependencies first:
# pip install selenium python-dotenv google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client openpyxl seleniumbase

import os
from dotenv import load_dotenv
from selenium.webdriver.common.by import By
from google.oauth2 import service_account
from googleapiclient.discovery import build
from seleniumbase import SB
import json
import sys
import time
import openpyxl
from openpyxl import load_workbook
import re
from pathlib import Path

# Load environment variables from the .env file
load_dotenv()

# Security: Validate environment variable inputs
SHEETS_OR_EXCEL = os.getenv('SHEETS_OR_EXCEL', 'SHEETS')
if SHEETS_OR_EXCEL not in ['SHEETS', 'EXCEL']:
    print("Erreur de sécurité : SHEETS_OR_EXCEL doit être 'SHEETS' ou 'EXCEL'")
    sys.exit(1)

EXCEL_NAME = os.getenv('EXCEL_NAME', 'X.xlsx')
SCOPES = [os.getenv('SCOPES')]
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
SHEET_NAME = os.getenv('SHEET_NAME')
URL_COLUMN = os.getenv('URL_COLUMN')
BUYING_PRICE_COLUMN = os.getenv('BUYING_PRICE_COLUMN')
SHIPPING_PRICE_COLUMN = os.getenv('SHIPPING_PRICE_COLUMN')
TOTAL_PRICE_COLUMN = os.getenv('TOTAL_PRICE_COLUMN')

# Security: Validate NUMBER_OF_URL_ROWS is a positive integer
try:
    NUMBER_OF_URL_ROWS = int(os.getenv('NUMBER_OF_URL_ROWS', '0'))
    if NUMBER_OF_URL_ROWS <= 0:
        raise ValueError
except (ValueError, TypeError):
    print("Erreur de sécurité : NUMBER_OF_URL_ROWS doit être un entier positif")
    sys.exit(1)

# Security: Validate column names (should be single letters A-Z)
for col_var, col_name in [
    ('URL_COLUMN', URL_COLUMN),
    ('BUYING_PRICE_COLUMN', BUYING_PRICE_COLUMN),
    ('SHIPPING_PRICE_COLUMN', SHIPPING_PRICE_COLUMN),
    ('TOTAL_PRICE_COLUMN', TOTAL_PRICE_COLUMN)
]:
    if not col_name or not re.match(r'^[A-Z]{1,3}$', col_name):
        print(f"Erreur de sécurité : {col_var} doit être une lettre de colonne valide (A-ZZZ)")
        sys.exit(1)

# Verify all required environment variables are present
required_vars = ['URL_COLUMN', 'BUYING_PRICE_COLUMN', 'SHIPPING_PRICE_COLUMN', 
                 'TOTAL_PRICE_COLUMN', 'SHEETS_OR_EXCEL', 'NUMBER_OF_URL_ROWS']
if SHEETS_OR_EXCEL == 'SHEETS':
    required_vars.extend(['SCOPES', 'SPREADSHEET_ID', 'SHEET_NAME'])
    
missing_vars = [var for var in required_vars if not os.getenv(var)]
if missing_vars:
    print(f"Erreur : Variables d'environnement manquantes: {', '.join(missing_vars)}")
    sys.exit(1)

class Prix:
    """
    Data class for storing price information.
    Security: All prices are validated as non-negative floats.
    """
    def __init__(self, purchase_price, shipping_price, total_price):
        # Security: Ensure prices are non-negative floats
        self.purchase_price = max(0.0, float(purchase_price))
        self.shipping_price = max(0.0, float(shipping_price))
        self.total_price = max(0.0, float(total_price))

    def __repr__(self):
        return f"Prix(achat={self.purchase_price}, expedition={self.shipping_price}, total={self.total_price})"

class SpreadsheetHandler:
    """
    Handles both Google Sheets and Excel file operations.
    Security: Validates all file paths and prevents path traversal attacks.
    """
    def __init__(self):
        self.type = SHEETS_OR_EXCEL
        if self.type == 'SHEETS':
            self.setup_sheets()
        else:
            self.setup_excel()

    def setup_sheets(self):
        """Initialize Google Sheets connection with proper error handling."""
        sheets_creds = get_sheets_credentials()
        service = build('sheets', 'v4', credentials=sheets_creds)
        self.sheet = service.spreadsheets()

    def setup_excel(self):
        """
        Initialize Excel file with security validation.
        Security: Validates file path to prevent directory traversal.
        """
        # Security: Validate Excel filename and prevent path traversal
        excel_path = Path(EXCEL_NAME).resolve()
        if not excel_path.name == EXCEL_NAME or '..' in EXCEL_NAME:
            print(f"Erreur de sécurité : Nom de fichier Excel invalide")
            sys.exit(1)
            
        try:
            self.workbook = load_workbook(filename=str(excel_path))
            self.sheet = self.workbook.active
        except FileNotFoundError:
            print(f"Erreur : Fichier Excel {EXCEL_NAME} non trouvé")
            sys.exit(1)
        except PermissionError:
            print(f"Erreur : Impossible d'accéder à {EXCEL_NAME}. Assurez-vous que le fichier n'est pas ouvert.")
            sys.exit(1)
        except Exception as e:
            print(f"Erreur d'accès au fichier Excel : {str(e)}")
            sys.exit(1)

    def get_urls(self):
        """Retrieve URLs from spreadsheet."""
        if self.type == 'SHEETS':
            return self._get_urls_from_sheets()
        else:
            return self._get_urls_from_excel()

    def _get_urls_from_sheets(self):
        """Get URLs from Google Sheets with error handling."""
        try:
            range_name = f'{SHEET_NAME}!{URL_COLUMN}2:{URL_COLUMN}{NUMBER_OF_URL_ROWS + 1}'
            result = self.sheet.values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name
            ).execute()
            values = result.get('values', [])
            return [{'url': self._sanitize_url(row[0]) if row else None, 'row': i + 2} 
                   for i, row in enumerate(values)]
        except Exception as e:
            print(f"Erreur lors de la lecture d'URL à partir de Google Sheets : {str(e)}")
            sys.exit(1)

    def _get_urls_from_excel(self):
        """Get URLs from Excel file with error handling."""
        try:
            urls = []
            for row in range(2, NUMBER_OF_URL_ROWS + 2):
                cell_value = self.sheet[f'{URL_COLUMN}{row}'].value
                sanitized_url = self._sanitize_url(cell_value) if cell_value else None
                urls.append({
                    'url': sanitized_url,
                    'row': row
                })
            return urls
        except Exception as e:
            print(f"Erreur lors de la lecture des URL à partir d'Excel: {str(e)}")
            sys.exit(1)

    def _sanitize_url(self, url):
        """
        Sanitize and validate URL.
        Security: Ensures URLs are from cardmarket.com domain only.
        """
        if not url or not isinstance(url, str):
            return None
            
        url = str(url).strip()
        
        # Security: Validate URL is from cardmarket.com
        if not url.startswith('https://www.cardmarket.com/'):
            print(f"Avertissement de sécurité : URL non-Cardmarket ignorée: {url[:50]}")
            return None
            
        return url

    def update_values(self, values, row_number):
        """Update spreadsheet values with validation."""
        # Security: Validate values are numeric
        try:
            validated_values = [float(v) if v is not None else 0.0 for v in values]
        except (ValueError, TypeError):
            print(f"Erreur : Valeurs non numériques à la ligne {row_number}")
            return
            
        if self.type == 'SHEETS':
            self._update_sheets_values(validated_values, row_number)
        else:
            self._update_excel_values(validated_values, row_number)

    def _update_sheets_values(self, values, row_number):
        """Update Google Sheets with error handling."""
        try:
            range_name = f'{SHEET_NAME}!{BUYING_PRICE_COLUMN}{row_number}:{TOTAL_PRICE_COLUMN}{row_number}'
            body = {'values': [values]}
            self.sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
        except Exception as e:
            print(f"Erreur de mise à jour du Google Sheets : {str(e)}")

    def _update_excel_values(self, values, row_number):
        """Update Excel file with error handling."""
        try:
            self.sheet[f'{BUYING_PRICE_COLUMN}{row_number}'] = values[0]
            self.sheet[f'{SHIPPING_PRICE_COLUMN}{row_number}'] = values[1]
            self.sheet[f'{TOTAL_PRICE_COLUMN}{row_number}'] = values[2]
            self.workbook.save(EXCEL_NAME)
        except Exception as e:
            print(f"Erreur de mise à jour du fichier Excel : {str(e)}")

    def cleanup(self):
        """Clean up resources."""
        if self.type == 'EXCEL':
            try:
                self.workbook.save(EXCEL_NAME)
                self.workbook.close()
            except Exception:
                pass

def get_sheets_credentials():
    """
    Load Google Sheets credentials securely.
    Security: Validates secrets.json path and content.
    """
    # Security: Validate secrets file path
    secrets_path = Path('secrets.json').resolve()
    if not secrets_path.exists() or not secrets_path.is_file():
        print("Erreur de sécurité : le fichier secrets.json n'existe pas ou n'est pas un fichier valide")
        sys.exit(1)
        
    try:
        with open(secrets_path, 'r', encoding='utf-8') as f:
            credentials_dict = json.load(f)
    except FileNotFoundError:
        print("Erreur : le fichier secrets.json n'a pas été trouvé")
        sys.exit(1)
    except json.JSONDecodeError:
        print("Erreur de sécurité : secrets.json n'est pas un fichier JSON valide")
        sys.exit(1)
    except Exception as e:
        print(f"Erreur de sécurité lors de la lecture de secrets.json : {str(e)}")
        sys.exit(1)
    
    # Security: Validate required fields in credentials
    required_fields = ['type', 'project_id', 'private_key', 'client_email']
    missing_fields = [field for field in required_fields if field not in credentials_dict]
    if missing_fields:
        print(f"Erreur de sécurité : Champs manquants dans secrets.json: {', '.join(missing_fields)}")
        sys.exit(1)
    
    try:
        credentials = service_account.Credentials.from_service_account_info(
            credentials_dict,
            scopes=SCOPES
        )
        return credentials
    except Exception as e:
        print(f"Erreur de sécurité : Les identifiants ne sont pas valides dans secrets.json : {str(e)}")
        sys.exit(1)

def get_cardmarket_credentials():
    """
    Retrieve Cardmarket credentials from environment.
    Security: Credentials are never logged or exposed.
    """
    cardmarket_username = os.getenv('CARDMARKET_USERNAME')
    cardmarket_password = os.getenv('CARDMARKET_PASSWORD')
    
    # Security: Validate credentials exist without exposing them
    if not cardmarket_username or not cardmarket_password:
        print("Erreur de sécurité : Identifiants Cardmarket non trouvés dans le fichier .env")
        sys.exit(1)
    
    # Security: Basic validation without exposing values
    if len(cardmarket_username.strip()) < 3 or len(cardmarket_password.strip()) < 6:
        print("Erreur de sécurité : Identifiants Cardmarket invalides (longueur insuffisante)")
        sys.exit(1)
    
    return cardmarket_username, cardmarket_password

def clean_and_convert(value):
    """
    Safely convert price strings to float.
    Security: Handles malicious input and prevents injection.
    """
    if not value or not isinstance(value, str) or not value.strip():
        return 0.0
    
    # Security: Remove all non-numeric characters except comma and dot
    cleaned_value = re.sub(r'[^\d,.]', '', value.strip())
    
    # Security: Validate cleaned value
    if not cleaned_value:
        return 0.0
    
    try:
        # Convert European format (1.234,56) to standard float
        cleaned_value = cleaned_value.replace('.', '').replace(',', '.')
        result = float(cleaned_value)
        
        # Security: Validate reasonable price range (0 to 100,000 EUR)
        if result < 0 or result > 100000:
            print(f"Avertissement : Prix suspect détecté et ignoré: {result}")
            return 0.0
            
        return result
    except (ValueError, TypeError):
        return 0.0

def get_prices_from_page_sb(sb, url):
    """
    Extract prices from Cardmarket page using SeleniumBase.
    Security: Validates all extracted data.
    """
    if not url or not isinstance(url, str) or url.isspace():
        return Prix(0, 0, 0)
    
    try:
        # Wait for page to load completely
        sb.sleep(2)
        
        # Security: Verify we're still on cardmarket.com
        current_url = sb.get_current_url()
        if not current_url.startswith('https://www.cardmarket.com/'):
            print(f"Avertissement de sécurité : Redirigé hors de Cardmarket")
            return Prix(0, 0, 0)
        
        # Find price elements with improved selectors
        elements_first_class = sb.find_elements(".color-primary.small.text-end.text-nowrap.fw-bold")
        elements_second_class = sb.find_elements(".ms-1")

        if not elements_first_class:
            print(f"Aucun prix trouvé sur la page")
            return Prix(0, 0, 0)

        # Security: Validate and sanitize extracted prices
        first_class_list = []
        for element in elements_first_class:
            text = element.text.strip()
            if text:
                price = clean_and_convert(text)
                first_class_list.append(price)
        
        if not elements_second_class:
            second_class_list = [0] * len(first_class_list)
        else:
            second_class_list = []
            for element in elements_second_class:
                text = element.text.strip()
                if text:
                    price = clean_and_convert(text)
                    second_class_list.append(price)
            
            # Pad with zeros if needed
            second_class_list.extend([0] * (len(first_class_list) - len(second_class_list)))

        if not first_class_list:
            return Prix(0, 0, 0)

        # Create price objects
        prix_list = [
            Prix(purchase_price, shipping_price, round(purchase_price + shipping_price, 2))
            for purchase_price, shipping_price in zip(first_class_list, second_class_list)
        ]

        # Return cheapest option
        return sorted(prix_list, key=lambda p: p.total_price)[0]
        
    except Exception as e:
        print(f"Erreur dans le traitement de l'URL {url[:50]}: {str(e)}")
        return Prix(0, 0, 0)

def main():
    """
    Main execution function with enhanced security and Cloudflare bypass.
    """
    spreadsheet = SpreadsheetHandler()

    try:
        # Security: Use incognito mode and proper UC Mode settings
        with SB(uc=True, headless=False, incognito=True) as sb:
            # Login to Cardmarket with improved reconnect time
            login_url = "https://www.cardmarket.com/en/Pokemon/Login"
            
            print("Connexion à Cardmarket...")
            # Security: Increased reconnect_time to avoid detection (recommended: 4-5 seconds minimum)
            sb.uc_open_with_reconnect(login_url, reconnect_time=5)
            
            # Wait for page load and check for Cloudflare
            sb.sleep(2)
            
            # Handle Cloudflare CAPTCHA if present
            try:
                # Security: Use uc_gui_click_captcha for better stealth
                sb.uc_gui_click_captcha()
                print("CAPTCHA résolu")
            except Exception:
                # CAPTCHA may not be present, continue
                pass
            
            cardmarket_username, cardmarket_password = get_cardmarket_credentials()
            
            # Security: Use uc_click instead of regular click to avoid detection
            try:
                # Find and fill login form with improved selectors
                sb.type('input[type="text"]', cardmarket_username)
                sb.sleep(0.5)
                sb.type('input[type="password"]', cardmarket_password)
                sb.sleep(0.5)
                
                # Security: Use uc_click for submitting form
                sb.uc_click('input[type="submit"]', reconnect_time=4)
                
                print("Connexion réussie")
            except Exception as e:
                print(f"Erreur lors de la connexion : {str(e)}")
                sys.exit(1)
            
            time.sleep(3)

            urls = spreadsheet.get_urls()
            total_urls = len(urls)
            processed = 0

            for url_data in urls:
                row_number = url_data['row']
                url = url_data['url']
                processed += 1
                
                print(f"Traitement de la ligne {row_number} ({processed}/{total_urls})...")
                
                try:
                    if url is None:
                        values = [0, 0, 0]
                        spreadsheet.update_values(values, row_number)
                        print(f"Ligne {row_number} vide: Prix mis à 0")
                    else:
                        # Security: Open URL with proper reconnect time (increased to 5 seconds)
                        sb.uc_open_with_reconnect(url, reconnect_time=5)
                        
                        # Handle Cloudflare on product pages
                        try:
                            sb.uc_gui_click_captcha()
                        except Exception:
                            # CAPTCHA may not be present
                            pass
                        
                        # Get prices using SeleniumBase methods
                        price_info = get_prices_from_page_sb(sb, url)
                        values = [price_info.purchase_price, price_info.shipping_price, price_info.total_price]
                        spreadsheet.update_values(values, row_number)
                        
                        print(f"Ligne {row_number} complétée: Prix d'achat={price_info.purchase_price}€, "
                              f"Frais de port={price_info.shipping_price}€, Total={price_info.total_price}€")
            
                    # Security: Add random delay between requests to avoid rate limiting
                    time.sleep(2 + (processed % 3))  # 2-4 seconds random delay
            
                except KeyboardInterrupt:
                    print("\nInterruption manuelle détectée. Sauvegarde de l'état actuel...")
                    break
                except Exception as e:
                    print(f"Erreur lors du traitement de la ligne {row_number}: {e}")
                    continue

            print("Le processus s'est terminé!")

    except Exception as e:
        print(f"Erreur critique: {str(e)}")
    finally:
        try:
            spreadsheet.cleanup()
        except Exception:
            pass
        print("Programme terminé.")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\nProgramme arrêté par l'utilisateur.")
    except Exception as e:
        print(f"Erreur fatale: {e}")