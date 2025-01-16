# Remerciements à Thomas pour son aide dans la réalisation de ce script.
# Cimer chef

import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
import sys
import time
import openpyxl
from openpyxl import load_workbook

# Load environment variables from the .env file
load_dotenv()

SHEETS_OR_EXCEL = os.getenv('SHEETS_OR_EXCEL', 'SHEETS')
EXCEL_NAME = os.getenv('EXCEL_NAME', 'X.xlsx')
SCOPES = [os.getenv('SCOPES')]
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
SHEET_NAME = os.getenv('SHEET_NAME')
URL_COLUMN = os.getenv('URL_COLUMN')
BUYING_PRICE_COLUMN = os.getenv('BUYING_PRICE_COLUMN')
SHIPPING_PRICE_COLUMN = os.getenv('SHIPPING_PRICE_COLUMN')
TOTAL_PRICE_COLUMN = os.getenv('TOTAL_PRICE_COLUMN')
NUMBER_OF_URL_ROWS = int(os.getenv('NUMBER_OF_URL_ROWS'))

# Verify all required environment variables are present
required_vars = ['URL_COLUMN', 'BUYING_PRICE_COLUMN', 'SHIPPING_PRICE_COLUMN', 'TOTAL_PRICE_COLUMN', 'SHEETS_OR_EXCEL','NUMBER_OF_URL_ROWS']
if SHEETS_OR_EXCEL == 'SHEETS':
    required_vars.extend(['SCOPES', 'SPREADSHEET_ID', 'SHEET_NAME'])
missing_vars = [var for var in required_vars if not os.getenv(var)]
if missing_vars:
    print(f"Erreur : Variables d'environnement manquantes: {', '.join(missing_vars)}")
    sys.exit(1)

class Prix:
    def __init__(self, purchase_price, shipping_price, total_price):
        self.purchase_price = purchase_price
        self.shipping_price = shipping_price
        self.total_price = total_price

    def __repr__(self):
        return f"Prix(achat={self.purchase_price}, expedition={self.shipping_price}, total={self.total_price})"

class SpreadsheetHandler:
    def __init__(self):
        self.type = SHEETS_OR_EXCEL
        if self.type == 'SHEETS':
            self.setup_sheets()
        else:
            self.setup_excel()

    def setup_sheets(self):
        sheets_creds = get_sheets_credentials()
        service = build('sheets', 'v4', credentials=sheets_creds)
        self.sheet = service.spreadsheets()

    def setup_excel(self):
        try:
            self.workbook = load_workbook(filename=EXCEL_NAME)
            self.sheet = self.workbook.active
        except FileNotFoundError:
            print(f"Erreur : Fichier Excel {EXCEL_NAME} non trouvee")
            sys.exit(1)
        except PermissionError:
            print(f"Erreur : Impossible d'acceder a {EXCEL_NAME}. Assurez-vous que le fichier n'est pas ouvert.")
            sys.exit(1)
        except Exception as e:
            print(f"Erreur d'acces au fichier Excel : {str(e)}")
            sys.exit(1)

    def get_urls(self):
        if self.type == 'SHEETS':
            return self._get_urls_from_sheets()
        else:
            return self._get_urls_from_excel()

    def _get_urls_from_sheets(self):
        try:
            range_name = f'{SHEET_NAME}!{URL_COLUMN}2:{URL_COLUMN}{NUMBER_OF_URL_ROWS + 1}'
            result = self.sheet.values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name
            ).execute()
            values = result.get('values', [])
            return [{'url': row[0] if row else None, 'row': i + 2} 
                   for i, row in enumerate(values)]
        except Exception as e:
            print(f"Erreur lors de la lecture d'URL a partir de Google Sheets : {str(e)}")
            sys.exit(1)

    def _get_urls_from_excel(self):
        try:
            urls = []
            for row in range(2, NUMBER_OF_URL_ROWS + 2):
                cell_value = self.sheet[f'{URL_COLUMN}{row}'].value
                urls.append({
                    'url': cell_value if cell_value and str(cell_value).strip() else None,
                    'row': row
                })
            return urls
        except Exception as e:
            print(f"Erreur lors de la lecture des URL à partir d'Excel: {str(e)}")
            sys.exit(1)

    def update_values(self, values, row_number):
        if self.type == 'SHEETS':
            self._update_sheets_values(values, row_number)
        else:
            self._update_excel_values(values, row_number)

    def _update_sheets_values(self, values, row_number):
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
            print(f"Erreur de mise a jour du Google Sheets : {str(e)}")

    def _update_excel_values(self, values, row_number):
        try:
            self.sheet[f'{BUYING_PRICE_COLUMN}{row_number}'] = values[0]
            self.sheet[f'{SHIPPING_PRICE_COLUMN}{row_number}'] = values[1]
            self.sheet[f'{TOTAL_PRICE_COLUMN}{row_number}'] = values[2]
            self.workbook.save(EXCEL_NAME)
        except Exception as e:
            print(f"Erreur de mise a jour du fichier Excel : {str(e)}")

    def cleanup(self):
        if self.type == 'EXCEL':
            try:
                self.workbook.save(EXCEL_NAME)
                self.workbook.close()
            except Exception:
                pass

# Rest of the functions remain the same
def get_sheets_credentials():
    try:
        with open('secrets.json', 'r') as f:
            credentials_dict = json.load(f)
    except FileNotFoundError:
        print("Erreur : le fichier secrets.json n'a pas ete trouve")
        sys.exit(1)
    except json.JSONDecodeError:
        print("Erreur : secrets.json n'est pas un fichier JSON valide")
        sys.exit(1)
    
    try:
        credentials = service_account.Credentials.from_service_account_info(
            credentials_dict,
            scopes=SCOPES
        )
        return credentials
    except Exception as e:
        print(f"Erreur : Les identifiants ne sont pas valides dans secrets.json : {str(e)}")
        sys.exit(1)

def get_cardmarket_credentials():
    cardmarket_username = os.getenv('CARDMARKET_USERNAME')
    cardmarket_password = os.getenv('CARDMARKET_PASSWORD')
    
    if not cardmarket_username or not cardmarket_password:
        print("Erreur : Identifiants Cardmarket non trouvees dans le fichier .env")
        sys.exit(1)
    
    return cardmarket_username, cardmarket_password

def clean_and_convert(value):
    if not value or not value.strip():
        return 0.0
    cleaned_value = (
        value.strip()
        .replace("€", "")
        .replace(".", "")
        .replace(",", ".")
    )
    return float(cleaned_value)

def login_to_platform(driver, cardmarket_username, cardmarket_password):
    try:
        login_url = "https://www.cardmarket.com/en/Pokemon/Login"
        driver.get(login_url)
        driver.implicitly_wait(10)

        input_field_id = driver.find_element(By.XPATH, "/html/body/main/div[2]/div[2]/div/form/div/div[1]/div/input")
        input_field_id.send_keys(cardmarket_username)

        input_field_cardmarket_password = driver.find_element(By.XPATH, "/html/body/main/div[2]/div[2]/div/form/div/div[2]/div/input")
        input_field_cardmarket_password.send_keys(cardmarket_password)

        submit_button = driver.find_element(By.XPATH, "/html/body/main/div[2]/div[2]/div/form/div/div[3]/div/input")
        submit_button.click()
    except Exception as e:
        print(f"Erreur lors de la connexion a Cardmarket : {str(e)}")
        sys.exit(1)

def get_prices_from_page(driver, url):
    if not url or url.isspace():
        return Prix(0, 0, 0)
    try:
        driver.get(url)
        driver.implicitly_wait(10)

        elements_first_class = driver.find_elements(By.CLASS_NAME, "color-primary.small.text-end.text-nowrap.fw-bold")
        elements_second_class = driver.find_elements(By.CLASS_NAME, "ms-1")

        if not elements_first_class:
            return Prix(0, 0, 0)

        first_class_list = [clean_and_convert(element.text) for element in elements_first_class if element.text.strip()]
        
        if not elements_second_class:
            second_class_list = [0] * len(first_class_list)
        else:
            second_class_list = [clean_and_convert(element.text) for element in elements_second_class if element.text.strip()]
            second_class_list.extend([0] * (len(first_class_list) - len(second_class_list)))

        if not first_class_list:
            return Prix(0, 0, 0)

        prix_list = [
            Prix(purchase_price, shipping_price, round(purchase_price + shipping_price, 2))
            for purchase_price, shipping_price in zip(first_class_list, second_class_list)
        ]

        return sorted(prix_list, key=lambda p: p.total_price)[0]
    except Exception as e:
        print(f"Erreur dans le traitement de l'URL {url}: {str(e)}")
        return Prix(0, 0, 0)

def main():
    spreadsheet = SpreadsheetHandler()
    driver = None

    try:
        chrome_options = Options()
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument("--disable-logging")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-usb")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        prefs = {
            'profile.default_content_setting_values': {
                'notifications': 2,
                'automatic_downloads': 1
            },
            'profile.default_content_settings': {
                'popups': 2
            },
            'credentials_enable_service': False,
            'profile.password_manager_enabled': False
        }
        chrome_options.add_experimental_option('prefs', prefs)
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)

        cardmarket_username, cardmarket_password = get_cardmarket_credentials()
        login_to_platform(driver, cardmarket_username, cardmarket_password)

        urls = spreadsheet.get_urls()

        for url_data in urls:
            row_number = url_data['row']
            url = url_data['url']
            
            print(f"Traitement de la ligne {row_number}...")
            
            try:
                if url is None:
                    values = [0, 0, 0]
                    spreadsheet.update_values(values, row_number)
                    print(f"Ligne {row_number} vide: Prix mis à 0")
                else:
                    price_info = get_prices_from_page(driver, url)
                    values = [price_info.purchase_price, price_info.shipping_price, price_info.total_price]
                    spreadsheet.update_values(values, row_number)
                    print(f"Ligne {row_number} completee: Prix d'achat={price_info.purchase_price}€, "
                          f"Frais de port={price_info.shipping_price}€, Total={price_info.total_price}€")
        
                if SHEETS_OR_EXCEL == 'SHEETS':
                    time.sleep(1.5)
        
            except KeyboardInterrupt:
                print("\nInterruption manuelle détectée. Sauvegarde de l'état actuel...")
                break
            except Exception as e:
                print(f"Erreur lors du traitement de la ligne {row_number}: {e}")
                continue

        print("Le processus s'est terminé!")

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
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