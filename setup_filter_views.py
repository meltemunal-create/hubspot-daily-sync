"""
TEK SEFERLIK kurulum script'i -- Sheet1 uzerinde filter view'lari olusturur:

  Today, Yesterday, This Week, Last Week, This Month, Last Month,
  This Quarter, Last Quarter, This Year, Last Year,
  January, February, March, April, May, June, July, August, September,
  October, November, December, Q1, Q2, Q3, Q4

Toplam 26 view. Her granularite icin hem "bulunulan donem" hem "bir onceki
donem" var. Aylar ve adlandirilmis ceyrekler (Q1-Q4) her zaman ICINDE
BULUNULAN YILA gore hesaplanir (YEAR(TODAY()) uzerinden) -- yani 2027
geldiginde script'i tekrar calistirmadan otomatik o yila kayar.

Hepsi Create Date (E sutunu) uzerinden CUSTOM_FORMULA kriteriyle calisir.

NOT: Eger daha once eski/bozuk filter view'lar olusturduysan, bu script'i
calistirmadan once Data > Filter views'tan hepsini silmelisin -- yoksa
ayni isimle ikinci bir kopyasi olusabilir / cakisabilir.

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

AYLAR = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def month_formula(month_number):
    """month_number: 1-12. Icinde bulunulan yilin o ayi. EDATE yil donusunu
    (ör. Aralik -> bir sonraki yilin Ocak'i) otomatik dogru hesaplar."""
    start = f"DATE(YEAR(TODAY()),{month_number},1)"
    end = f"EDATE(DATE(YEAR(TODAY()),{month_number},1),1)"
    return f"=AND($E2>={start},$E2<{end})"


def quarter_formula(quarter_number):
    """quarter_number: 1-4. Icinde bulunulan yilin o ceyregi."""
    start_month = 3 * (quarter_number - 1) + 1
    start = f"DATE(YEAR(TODAY()),{start_month},1)"
    end = f"EDATE(DATE(YEAR(TODAY()),{start_month},1),3)"
    return f"=AND($E2>={start},$E2<{end})"


VIEWS = [
    ("Today", "=AND($E2>=TODAY(),$E2<TODAY()+1)"),
    ("Yesterday", "=AND($E2>=TODAY()-1,$E2<TODAY())"),
    ("This Week", "=AND($E2>=TODAY()-WEEKDAY(TODAY(),3),$E2<TODAY()-WEEKDAY(TODAY(),3)+7)"),
    ("Last Week", "=AND($E2>=TODAY()-WEEKDAY(TODAY(),3)-7,$E2<TODAY()-WEEKDAY(TODAY(),3))"),
    ("This Month", "=AND($E2>=DATE(YEAR(TODAY()),MONTH(TODAY()),1),$E2<EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))"),
    ("Last Month", "=AND($E2>=EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),-1),$E2<DATE(YEAR(TODAY()),MONTH(TODAY()),1))"),
    ("This Quarter", "=AND($E2>=DATE(YEAR(TODAY()),1+3*INT((MONTH(TODAY())-1)/3),1),$E2<EDATE(DATE(YEAR(TODAY()),1+3*INT((MONTH(TODAY())-1)/3),1),3))"),
    ("Last Quarter", "=AND($E2>=EDATE(DATE(YEAR(TODAY()),1+3*INT((MONTH(TODAY())-1)/3),1),-3),$E2<DATE(YEAR(TODAY()),1+3*INT((MONTH(TODAY())-1)/3),1))"),
    ("This Year", "=AND($E2>=DATE(YEAR(TODAY()),1,1),$E2<DATE(YEAR(TODAY())+1,1,1))"),
    ("Last Year", "=AND($E2>=DATE(YEAR(TODAY())-1,1,1),$E2<DATE(YEAR(TODAY()),1,1))"),
]
VIEWS += [(ay, month_formula(i + 1)) for i, ay in enumerate(AYLAR)]
VIEWS += [(f"Q{q}", quarter_formula(q)) for q in range(1, 5)]

# Google Sheets'in "Change view" menusu, olusturma sirasina bakmadan HER ZAMAN
# alfabetik siralar -- bunu kapatan bir ayar yok. Baslara numara koyup
# alfabetik siralamayi kronolojik siraya zorluyoruz (01, 02, 03... "This/Last"
# harflerinden once gelir).
VIEWS = [(f"{i+1:02d} - {title}", formula) for i, (title, formula) in enumerate(VIEWS)]


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
