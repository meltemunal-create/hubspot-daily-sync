"""
TEK SEFERLIK kurulum script'i -- spreadsheet'teki HER "veri sekmesi"nde
(basligi "Create Date" olan bir E sutunu bulunan her sekme -- Sheet1, HubSpot
liste sekmesi, ileride eklenecek her turlu benzer sekme) asagidaki filter
view'lari otomatik olusturur:

  Today, Yesterday, This Week, Last Week, This Month, Last Month,
  This Quarter, Last Quarter, This Year, Last Year,
  January, February, March, April, May, June, July, August, September,
  October, November, December, Q1, Q2, Q3, Q4

Sekme ismini hardcode etmiyoruz -- SHEET_NAME degistiyse veya yeni bir HubSpot
liste sekmesi eklendiyse, bu script tekrar calistirildiginda otomatik onu da
bulur ve filtreleri kurar.

Toplam 26 view / sekme. Aylar ve adlandirilmis ceyrekler (Q1-Q4) her zaman
ICINDE BULUNULAN YILA gore hesaplanir (YEAR(TODAY()) uzerinden).

NOT: Eger bir sekmede daha once eski/bozuk filter view'lar olusturduysan, bu
script'i calistirmadan once o sekmede Data > Filter views'tan hepsini
silmelisin -- yoksa ayni isimle ikinci bir kopyasi olusabilir / cakisabilir.

Calistirma: python setup_filter_views.py
Gerekli ortam degiskenleri: GOOGLE_SERVICE_ACCOUNT_JSON, SPREADSHEET_ID
"""

import os
import json

import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1ZaTNGfpbLvkR-E9WaMJjnFoFEs6HPIIFpwohQHg9xcU")
NUM_COLUMNS = 27  # A..AA

# Create Date = E sutunu = index 4 (0-based)
CREATE_DATE_COL_INDEX = 4
CREATE_DATE_HEADER = "Create Date"

AYLAR = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def month_formula(month_number):
    start = f"DATE(YEAR(TODAY()),{month_number},1)"
    end = f"EDATE(DATE(YEAR(TODAY()),{month_number},1),1)"
    return f"=AND($E2>={start},$E2<{end})"


def quarter_formula(quarter_number):
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

# Google Sheets'in "Change view" menusu HER ZAMAN alfabetik siralar -- baslara
# numara koyup bu siralamayi kronolojik siraya zorluyoruz.
VIEWS = [(f"{i+1:02d} - {title}", formula) for i, (title, formula) in enumerate(VIEWS)]


def find_data_sheets(sh):
    """Basliginda (1. satir, E sutunu) 'Create Date' yazan her sekmeyi bulur."""
    data_sheets = []
    for ws in sh.worksheets():
        header = ws.row_values(1)
        if len(header) > CREATE_DATE_COL_INDEX and header[CREATE_DATE_COL_INDEX] == CREATE_DATE_HEADER:
            data_sheets.append(ws)
    return data_sheets


def build_requests_for_sheet(ws):
    return [
        {
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
        }
        for title, formula in VIEWS
    ]


def main():
    creds_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    data_sheets = find_data_sheets(sh)
    if not data_sheets:
        print("Hicbir sekmede 'Create Date' basligi (E sutunu) bulunamadi -- "
              "once sync.py'yi calistirip veriyi olusturman gerekebilir.")
        return

    print(f"{len(data_sheets)} veri sekmesi bulundu: {', '.join(ws.title for ws in data_sheets)}")

    requests_batch = []
    for ws in data_sheets:
        requests_batch.extend(build_requests_for_sheet(ws))

    sh.batch_update({"requests": requests_batch})
    print(f"Her sekmeye {len(VIEWS)} filter view kuruldu "
          f"(toplam {len(requests_batch)} view, {len(data_sheets)} sekme).")


if __name__ == "__main__":
    main()
