"""
TEK SEFERLIK kurulum script'i -- Sheet1 uzerinde 5 adet Filter View olusturur:
Yesterday / This Week / This Month / This Quarter / This Year

Bunlar Create Date (E sutunu) uzerinden CUSTOM_FORMULA kriteriyle calisir,
TODAY()'e bagli oldugu icin her gun otomatik guncel kalir.

NOT: Bu script'in CUSTOM_FORMULA kriter davranisi canli bir spreadsheet'te
test edilmedi. Calistirdiktan sonra Data > Filter views menusunden 5 view'in
de goruldugunu ve dogru satirlari filtreledigini kontrol et. Sorun cikarsa,
README.md'deki "Manuel kurulum" adimlarini kullan (2 dakika suren, UI'dan
elle kurulum -- garanti calisir).

Calistirma: python setup_filter_views.py
Gerekli ortam degiskenleri: GOOGLE_SERVICE_ACCOUNT_JSON, SPREADSHEET_ID
"""

import os
import json

import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1ZaTNGfpbLvkR-E9WaMJjnFoFEs6HPIIFpwohQHg9xcU")
SHEET_NAME = os.environ.get("SHEET_NAME", "Sheet1")
NUM_COLUMNS = 27  # A..AA

# Create Date = E sutunu = index 4 (0-based)
CREATE_DATE_COL_INDEX = 4

VIEWS = [
    ("Yesterday", "=$E2>=TODAY()-1"),
    ("This Week", "=$E2>=TODAY()-WEEKDAY(TODAY(),3)"),
    ("This Month", "=$E2>=EOMONTH(TODAY(),-1)+1"),
    ("This Quarter", "=$E2>=DATE(YEAR(TODAY()),1+3*INT((MONTH(TODAY())-1)/3),1)"),
    ("This Year", "=$E2>=DATE(YEAR(TODAY()),1,1)"),
]


def main():
    creds_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(SHEET_NAME)

    requests_batch = []
    for title, formula in VIEWS:
        requests_batch.append({
            "addFilterView": {
                "filter": {
                    "title": title,
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 0,
                        "startColumnIndex": 0,
                        "endColumnIndex": NUM_COLUMNS,
                    },
                    "criteria": {
                        str(CREATE_DATE_COL_INDEX): {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [{"userEnteredValue": formula}],
                            }
                        }
                    },
                }
            }
        })

    sh.batch_update({"requests": requests_batch})
    print(f"{len(VIEWS)} filter view olusturuldu: {', '.join(v[0] for v in VIEWS)}")
    print("Data > Filter views menusunden kontrol et.")


if __name__ == "__main__":
    main()
