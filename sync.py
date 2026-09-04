"""
HubSpot -> Google Sheets günlük senkronizasyon
ActionCoach Turkiye - Genel Lead Takip Excel'i

Her calistiginda TUM HubSpot contact'larini ceker, is mantigini uygular,
Sheet1'i BASTAN YAZAR (bulk overwrite), A ve B sutunlarini (manuel notlar)
korur, en yeni Create Date en usте olacak sekilde siralar.

Calistirma: GitHub Actions -> gunluk cron (bkz. .github/workflows/daily-sync.yml)
"""

import os
import re
import sys
import time
import html
import json
import logging
from datetime import datetime, timezone, date as date_cls
from zoneinfo import ZoneInfo

import requests
import gspread
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------

HUBSPOT_TOKEN = os.environ["HUBSPOT_PRIVATE_APP_TOKEN"]
GOOGLE_SA_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]  # tum JSON icerigi, secret olarak

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1ZaTNGfpbLvkR-E9WaMJjnFoFEs6HPIIFpwohQHg9xcU")
SHEET_NAME = os.environ.get("SHEET_NAME", "Sheet1")
LIST_ID = os.environ.get("HUBSPOT_LIST_ID", "6526")  # ayri sekmede gosterilecek HubSpot listesi

TZ = ZoneInfo("Europe/Istanbul")
DATE_FMT = "%d.%m.%Y"          # goruntuleme formati (hucre number-format olarak uygulanir)
SHEETS_EPOCH = date_cls(1899, 12, 30)  # Google Sheets'in kendi tarih sayma baslangici


def to_sheets_serial(d):
    """Python date -> Sheets'in gercek tarih hucresi olarak tanidigi tam sayi.
    Locale'e / string parse'a bagimli degil, garanti calisir."""
    return (d - SHEETS_EPOCH).days

SALES_PIPELINE_LABEL = "Sales Pipeline"

HUBSPOT_BASE = "https://api.hubapi.com"
HEADERS = {
    "Authorization": f"Bearer {HUBSPOT_TOKEN}",
    "Content-Type": "application/json",
}

# Sutun sirasi (A..AA) -- write_sheet() bu sirayla yazar
COLUMNS = [
    "AC NOTLAR",                     # A - manuel, dokunulmaz
    "Ekip Notlari",                  # B - manuel, dokunulmaz
    "Source",                        # C
    "Doldurdugu Ilk Form",           # D
    "Create Date",                   # E
    "HubSpot Bilgileri (Ozet)",      # F
    "Musteri Durumu",                # G
    "First Name",                    # H
    "Last Name",                     # I
    "Company Name",                  # J
    "Segment",                       # K
    "Pozisyon",                      # L
    "Sirket Yapisi",                 # M
    "Yillik Ciro",                   # N
    "Calisan Sayisi",                # O
    "Sektor",                        # P
    "Katildigi Webinarlar",          # Q
    "Katildigi GrowthCLUB Oturumlari",  # R
    "Katildigi Planlama Oturumlari", # S
    "Notes / Call Notes",            # T
    "Deal Stage",                    # U
    "Phone Number",                  # V
    "Email",                         # W
    "Marketing Contact Status",      # X
    "Unsubscribe",                   # Y
    "MasterCLASS Process",           # Z
    "Record ID",                     # AA
]

MANUAL_COL_INDEXES = {0, 1}  # A, B -- guncellemede mevcut degeri koru

CONTACT_PROPERTIES = [
    "firstname", "lastname", "company", "phone", "email", "createdate",
    "first_conversion_event_name",
    "pozisyon", "job_title",
    "sirket_yaps",
    "yillik_ciro",
    "cal_san_say_s_",
    "sektor_v2", "faaliyet_gosterdiginiz_sektor",
    "kat_ld_g__webinarlar",
    "kat_ld_g__tum_growthclub_oturumlar_",
    "katldg_planlama_oturumlar",
    "masterclass_process",
    "hs_marketable_status",
    "hs_email_optout",
    "hs_analytics_source",
    "hs_analytics_source_data_1",
    "hs_analytics_source_data_2",
    "hs_object_id",
]

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("hubspot-sync")


# ----------------------------------------------------------------------------
# HTTP helper (retry + basic rate-limit handling)
# ----------------------------------------------------------------------------

def hs_request(method, path, **kwargs):
    url = f"{HUBSPOT_BASE}{path}"
    for attempt in range(5):
        resp = requests.request(method, url, headers=HEADERS, timeout=30, **kwargs)
        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", 2))
            log.warning(f"Rate limited, {wait}s bekleniyor...")
            time.sleep(wait)
            continue
        if resp.status_code >= 500:
            log.warning(f"HubSpot 5xx ({resp.status_code}), tekrar deneniyor...")
            time.sleep(2 ** attempt)
            continue
        resp.raise_for_status()
        return resp.json()
    raise RuntimeError(f"HubSpot request basarisiz: {method} {path}")


# ----------------------------------------------------------------------------
# 1) TUM CONTACT'LARI CEK (ID-cursor pagination)
# ----------------------------------------------------------------------------

def fetch_all_contacts():
    contacts = []
    after = None
    while True:
        params = {
            "limit": 100,
            "properties": ",".join(CONTACT_PROPERTIES),
        }
        if after:
            params["after"] = after
        data = hs_request("GET", "/crm/v3/objects/contacts", params=params)
        contacts.extend(data.get("results", []))
        paging = data.get("paging", {})
        after = paging.get("next", {}).get("after")
        if not after:
            break
        log.info(f"{len(contacts)} contact cekildi, devam ediyor...")
    log.info(f"Toplam {len(contacts)} contact cekildi.")
    return contacts


def fetch_list_membership_ids(list_id):
    """HubSpot listesindeki (statik veya dinamik fark etmez) tum record ID'leri doner."""
    ids = set()
    after = None
    while True:
        params = {"limit": 100}
        if after:
            params["after"] = after
        data = hs_request("GET", f"/crm/v3/lists/{list_id}/memberships", params=params)
        for item in data.get("results", []):
            ids.add(str(item.get("recordId")))
        after = data.get("paging", {}).get("next", {}).get("after")
        if not after:
            break
    log.info(f"Liste {list_id}: {len(ids)} uye bulundu.")
    return ids


def fetch_list_name(list_id, fallback=None):
    try:
        data = hs_request("GET", f"/crm/v3/lists/{list_id}")
        name = data.get("name") or (data.get("list") or {}).get("name")
        return name or fallback or f"HubSpot List {list_id}"
    except requests.HTTPError as e:
        log.warning(f"Liste adi alinamadi ({list_id}): {e}")
        return fallback or f"HubSpot List {list_id}"


# ----------------------------------------------------------------------------
# 2) ASSOCIATIONS: notes / calls / deals (batch)
# ----------------------------------------------------------------------------
# NOT: Onceki Colab projesinde (France migration) batch association okumasi
# task'lar icin ~%60 basarili cikmisti, tekil v4 cagrisina donulmustu. Burada
# once batch deniyoruz; ilk gercek calistirmada oran dusuk cikarsa fallback
# olarak tekil cagriya (chunk'lar halinde, thread'li) donebiliriz.

def fetch_associations_batch(contact_ids, to_object_type):
    """contact_id -> [associated_object_id, ...] dict'i doner."""
    result = {cid: [] for cid in contact_ids}
    chunk_size = 100
    for i in range(0, len(contact_ids), chunk_size):
        chunk = contact_ids[i:i + chunk_size]
        body = {"inputs": [{"id": cid} for cid in chunk]}
        try:
            data = hs_request(
                "POST",
                f"/crm/v4/associations/contacts/{to_object_type}/batch/read",
                json=body,
            )
            for item in data.get("results", []):
                from_id = item["from"]["id"]
                to_ids = [t["toObjectId"] for t in item.get("to", [])]
                result[from_id] = to_ids
        except requests.HTTPError as e:
            log.warning(f"Batch association hatasi ({to_object_type}): {e}")
    return result


def batch_read_objects(object_type, object_ids, properties):
    """Verilen ID listesindeki object'leri (note/call/deal) toplu okur."""
    out = {}
    chunk_size = 100
    ids = list(set(object_ids))
    for i in range(0, len(ids), chunk_size):
        chunk = ids[i:i + chunk_size]
        body = {
            "properties": properties,
            "inputs": [{"id": oid} for oid in chunk],
        }
        data = hs_request("POST", f"/crm/v3/objects/{object_type}/batch/read", json=body)
        for obj in data.get("results", []):
            out[obj["id"]] = obj.get("properties", {})
    return out


# ----------------------------------------------------------------------------
# 3) SALES PIPELINE stage-id -> stage-adi eslemesi
# ----------------------------------------------------------------------------

def fetch_sales_pipeline_stage_map():
    data = hs_request("GET", "/crm/v3/pipelines/deals")
    for pipeline in data.get("results", []):
        if pipeline.get("label") == SALES_PIPELINE_LABEL:
            return (
                pipeline["id"],
                {stage["id"]: stage["label"] for stage in pipeline.get("stages", [])},
            )
    log.warning(f"'{SALES_PIPELINE_LABEL}' adinda bir pipeline bulunamadi.")
    return None, {}


def fetch_deal_stage_for_contacts(contact_ids, sales_pipeline_id, stage_map):
    """Her contact icin, Sales Pipeline'daki en son guncellenen deal'in stage adini doner."""
    assoc = fetch_associations_batch(contact_ids, "deals")
    all_deal_ids = [d for ids in assoc.values() for d in ids]
    if not all_deal_ids:
        return {cid: "" for cid in contact_ids}

    deals = batch_read_objects(
        "deals", all_deal_ids,
        properties=["pipeline", "dealstage", "hs_lastmodifieddate"],
    )

    result = {}
    for cid, deal_ids in assoc.items():
        candidates = []
        for did in deal_ids:
            props = deals.get(did)
            if not props:
                continue
            if props.get("pipeline") != sales_pipeline_id:
                continue
            candidates.append(props)
        if not candidates:
            result[cid] = ""
            continue
        candidates.sort(key=lambda p: p.get("hs_lastmodifieddate") or "", reverse=True)
        latest = candidates[0]
        result[cid] = stage_map.get(latest.get("dealstage"), latest.get("dealstage", ""))
    return result


# ----------------------------------------------------------------------------
# 4) NOTES / CALLS temizleme + birlestirme
# ----------------------------------------------------------------------------

NOISE_PATTERNS = [
    r"zoom\.us\S*", r"meeting recording", r"cloud recording", r"recording link",
    r"kay[ıi]t linki", r"toplant[ıi] kayd[ıi]", r"transcript", r"passcode", r"meeting id",
]
NOISE_RE = re.compile("|".join(NOISE_PATTERNS), re.IGNORECASE)
HTML_TAG_RE = re.compile(r"<[^>]+>")


def clean_note_text(raw):
    if not raw:
        return ""
    text = html.unescape(raw)
    text = HTML_TAG_RE.sub(" ", text)
    text = NOISE_RE.sub("", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_for_dedup(text):
    text = text.lower()
    # Turkce karakterleri de kapsayan noktalama/ozel karakter temizligi
    text = re.sub(r"[^\w\s]", "", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def to_local_date(hs_props, owners_map, entity_id, owner_field="hs_created_by",
                   date_field="hs_timestamp"):
    ts = hs_props.get(date_field)
    dt_str = format_hs_date(ts)
    owner_id = hs_props.get(owner_field)
    owner_name = owners_map.get(str(owner_id), owner_id or "-")
    return dt_str, owner_name


def format_hs_date(raw, storage=False):
    """
    HubSpot tarihleri hem ISO string hem epoch-ms string olarak gelebiliyor.

    storage=True  -> Sheets tarih-serial'i (int): E sutununa YAZILACAK gercek
                      deger. Bu sayede hucre GERCEKTEN tarih tipinde olur,
                      filter view'lardaki TODAY() karsilastirmalari calisir.
    storage=False -> dd.MM.yyyy string: sadece Notes/Call metni gibi
                      duz-metin baglamlarda kullanilir (Sheets hucresi degil).
    """
    if not raw:
        return None if storage else ""
    try:
        if str(raw).isdigit():
            dt = datetime.fromtimestamp(int(raw) / 1000, tz=timezone.utc)
        else:
            dt = datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
        dt_local = dt.astimezone(TZ)
        if storage:
            return to_sheets_serial(dt_local.date())
        return dt_local.strftime(DATE_FMT)
    except (ValueError, TypeError):
        return None if storage else ""


def build_notes_and_calls_text(note_props_list, call_props_list, owners_map, max_chars=45000):
    """
    T sutunu: butun note+call kayitlari, format:
      dd.MM.yyyy | Kullanici Adi | NOTE/CALL
      metin
    Yeniden eskiye, dedup'li (normalize edilmis metne gore), max 45.000 karakter.
    """
    entries = []  # (datetime_obj, formatted_block, normalized_text)

    for props in note_props_list:
        raw_text = clean_note_text(props.get("hs_note_body", ""))
        if not raw_text:
            continue
        date_str = format_hs_date(props.get("hs_timestamp"))
        owner_name = owners_map.get(str(props.get("hubspot_owner_id")), "-")
        block = f"{date_str} | {owner_name} | NOTE\n{raw_text}"
        entries.append((props.get("hs_timestamp") or "0", block, normalize_for_dedup(raw_text)))

    for props in call_props_list:
        raw_text = clean_note_text(props.get("hs_call_body", "") or props.get("hs_call_title", ""))
        if not raw_text:
            continue
        date_str = format_hs_date(props.get("hs_timestamp"))
        owner_name = owners_map.get(str(props.get("hubspot_owner_id")), "-")
        block = f"{date_str} | {owner_name} | CALL\n{raw_text}"
        entries.append((props.get("hs_timestamp") or "0", block, normalize_for_dedup(raw_text)))

    # yeniden eskiye sirala
    entries.sort(key=lambda e: e[0], reverse=True)

    seen = set()
    final_blocks = []
    for _, block, norm in entries:
        if norm in seen:
            continue
        seen.add(norm)
        final_blocks.append(block)

    full_text = "\n\n".join(final_blocks)
    if len(full_text) > max_chars:
        full_text = full_text[:max_chars]
    return full_text, final_blocks  # final_blocks C-sutunu ozeti icin de kullanilacak


def build_ozet_c_column(webinar_list, gc_list, planning_list, note_blocks_top5):
    """
    F sutunu (HubSpot Bilgileri Ozet), 4 satir, renklendirme write_sheet() icinde uygulanir.
    Not: webinar_list / gc_list / planning_list -- ham property degerinden parse edilmis
    (tarih, isim) listeleri. Property'nin gercek formatini ilk canli veriyle DOGRULAMAMIZ lazim.
    """
    def summarize(items, label):
        count = len(items)
        last = items[-1] if items else "-"
        return f"Katildigi {label} Sayisi: {count} | Son Katildigi {label}: {last}"

    lines = [
        summarize(webinar_list, "Webinar"),
        summarize(gc_list, "GC"),
        summarize(planning_list, "Planlama"),
    ]

    note_summary_parts = []
    for block in note_blocks_top5[:5]:
        # block = "tarih | isim | NOTE\nmetin" -- ilk 240 karaktere kirp
        snippet = block.replace("\n", " ")[:240]
        note_summary_parts.append(snippet)
    lines.append("Not Ozeti: " + (" || ".join(note_summary_parts) if note_summary_parts else "-"))

    return "\n".join(lines)


def parse_multivalue_property(raw):
    """
    kat_ld_g__webinarlar vb. alanlar HubSpot'ta genelde ';' ile ayrilmis
    checkbox/multi-select degerleri olarak gelir.
    VARSAYIM: format "Isim1;Isim2;..." -- ilk gercek veriyle dogrulanacak.
    """
    if not raw:
        return []
    return [v.strip() for v in str(raw).split(";") if v.strip()]


# ----------------------------------------------------------------------------
# 5) SOURCE mantigi
# ----------------------------------------------------------------------------

KEYWORD_PLATFORM = [
    ("facebook", "Facebook"), ("meta", "Facebook"),
    ("instagram", "Instagram"),
    ("linkedin", "LinkedIn"),
    ("whatsapp", "WhatsApp"),
    ("youtube", "Youtube"),
]
PAID_HINTS = ["cpc", "paid", "ppc", "cpm", "reklam"]
DOMAIN_RE = re.compile(r"[a-z0-9-]+\.[a-z]{2,}", re.IGNORECASE)


# Bazi platformlarda organik lead akisi is modelimizde hic yok (hepsi reklam) --
# HubSpot'un kendi organic/paid tespiti bu platformlar icin guvenilir degil,
# bu yuzden burada is kuraliyla eziyoruz.
PLATFORMS_ALWAYS_PAID = {"facebook"}

# Facebook ve Instagram reklamlari ayni Meta Ads Manager'dan gidiyor -- ayri
# "Instagram Ads" diye bir platform yok, ikisi de "Meta Ads" altinda birlesir.
META_PLATFORMS = {"facebook", "instagram"}


def ads_label(platform):
    if (platform or "").strip().lower() in META_PLATFORMS:
        return "Meta Ads"
    return f"{platform} Ads"


def compute_source(props):
    top = (props.get("hs_analytics_source") or "").upper()
    d1 = props.get("hs_analytics_source_data_1") or ""
    d2 = props.get("hs_analytics_source_data_2") or ""

    if top == "PAID_SOCIAL":
        return ads_label(d1 or "Bilinmiyor")
    if top in ("ORGANIC_SOCIAL", "SOCIAL_MEDIA"):
        platform = d1 or "Bilinmiyor"
        if platform.strip().lower() in PLATFORMS_ALWAYS_PAID:
            return ads_label(platform)
        return f"Organic Social - {platform}"
    if top == "PAID_SEARCH":
        return "Paid Search - Google"
    if top == "ORGANIC_SEARCH":
        return f"Organic Search - {d2 or 'Google'}"
    if top == "EMAIL_MARKETING":
        return "Marketing Email"
    if top == "REFERRALS":
        return f"Referrals - {d1 or 'Bilinmiyor'}"
    if top == "DIRECT_TRAFFIC":
        return "Direct Traffic"
    if top == "OFFLINE":
        # HubSpot arayuzunden elle eklenen kayitlar burada "CRM_UI" olarak
        # gelir, Drill-down 2'de ekleyen kisinin adi/emaili bulunur.
        if "crm_ui" in (d1 or "").lower().replace(" ", "_"):
            adder = re.sub(r"\s*\(.*?\)\s*$", "", d2 or "").strip()
            return f"Manually Added - {adder}" if adder else "Manually Added"
        return f"Offline Sources - {d1 or 'Bilinmiyor'}"

    if top == "OTHER_CAMPAIGNS":
        combined = f"{d1} {d2}".lower()

        if "refcode" in combined or "referral" in combined:
            return "Referral"

        for keyword, platform in KEYWORD_PLATFORM:
            if keyword in combined:
                is_paid = (
                    any(hint in combined for hint in PAID_HINTS)
                    or platform.lower() in PLATFORMS_ALWAYS_PAID
                )
                # Artik TUM platformlar icin ayni ayrim: reklamsa "{Platform} Ads",
                # degilse "Organic Social - {Platform}" (sadece Facebook/Instagram'a
                # ozel degil -- LinkedIn/WhatsApp/Youtube de ayni mantiga tabi).
                return ads_label(platform) if is_paid else f"Organic Social - {platform}"

        if "google" in combined:
            is_paid = any(hint in combined for hint in PAID_HINTS)
            return "Paid Search - Google" if is_paid else "Organic Search - Google"

        if "webinar" in combined and "qr" in combined:
            return "Webinar QR"
        if "zoom" in combined:
            return "Direct Zoom"

        domain_match = DOMAIN_RE.search(combined)
        if domain_match:
            return f"Referrals - {domain_match.group(0)}"

        return "Direct Traffic"

    return "Direct Traffic"


# ----------------------------------------------------------------------------
# 6) MUSTERI DURUMU + SEGMENT
# ----------------------------------------------------------------------------

MUSTERI_RENK = {
    "aktif":     {"bg": "#B7D7A8", "fg": "#276419", "bold": True},
    "sureçte_degil": {"bg": "#FFD966", "fg": "#B45F06", "bold": True},
    "degil":     {"bg": "#F4CCCC", "fg": "#990000", "bold": True},
}


def compute_musteri_durumu(masterclass_process):
    val = (masterclass_process or "").strip()
    if not val:
        return "Musteri Degil", MUSTERI_RENK["degil"]

    low = val.lower()
    if "aktif" in low and not any(x in low for x in ["aktif değil", "aktif degil", "pasif", "inaktif", "mezun"]):
        return f"Musteri Aktif Surecte ({val})", MUSTERI_RENK["aktif"]
    if "mezun" in low:
        return f"Musteri Surecte Degil Mezun ({val})", MUSTERI_RENK["sureçte_degil"]
    return f"Musteri Surecte Degil ({val})", MUSTERI_RENK["sureçte_degil"]


CIRO_PUAN = [
    (["0-999b", "1m-4.9m"], 70),
    (["5m-24.9m", "25m-50m"], 140),
    (["51m-100m", "101m-500m", "500m+"], 210),
]
STEP_UP_POZISYONLAR = {"çalışan", "calisan", "şu an çalışmıyorum", "su an calismiyorum"}


def compute_segment(yillik_ciro, pozisyon):
    if pozisyon and pozisyon.strip().lower() in STEP_UP_POZISYONLAR:
        return "Step Up"

    ciro_key = (yillik_ciro or "").strip().lower().replace(" ", "")
    puan = 0
    for keys, p in CIRO_PUAN:
        if ciro_key in keys:
            puan = p
            break

    if puan <= 149:
        return "Step Up"
    elif puan <= 249:
        return "Power Up"
    else:
        return "Scale Up"


# ----------------------------------------------------------------------------
# 7) GOOGLE SHEETS baglantisi + yazma
# ----------------------------------------------------------------------------

def connect_spreadsheet():
    creds_info = json.loads(GOOGLE_SA_JSON)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


def get_or_create_worksheet(sh, title):
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=2000, cols=len(COLUMNS))


def read_manual_columns(ws):
    """
    Mevcut sheet'teki A/B (manuel notlar) sutunlarini Record ID (AA, son sutun)
    ile eslestirip donduruyor -- yeni yazimda bu degerler korunacak.
    """
    all_values = ws.get_all_values()
    if len(all_values) < 2:
        return {}
    header = all_values[0]
    try:
        record_id_idx = header.index("Record ID")
    except ValueError:
        return {}

    manual_by_record_id = {}
    for row in all_values[1:]:
        if len(row) <= record_id_idx:
            continue
        record_id = row[record_id_idx]
        if not record_id:
            continue
        a_val = row[0] if len(row) > 0 else ""
        b_val = row[1] if len(row) > 1 else ""
        manual_by_record_id[record_id] = (a_val, b_val)
    return manual_by_record_id


def write_sheet(ws, rows, color_map):
    """
    rows: [[col_A, col_B, ...col_AA], ...] -- E (Create Date) DESC sirali olmali.
    color_map: {row_index_0based: {"F": {...}, "G": {...}}} -- F/G renklendirme
    """
    ws.clear()
    # RAW: Sheets girilen degeri YORUMLAMADAN oldugu gibi yazar -- telefon
    # numarasi gibi alanlarda basdaki "0"in veya Record ID gibi uzun sayilarin
    # yanlislikla sayiya cevrilip bozulmasini onler. Tek gercek sayisal deger
    # olan Create Date (E) zaten yukarida Python int (Sheets serial) olarak
    # hazirlandi, RAW ile de doğru sekilde gercek tarih hucresi olarak yazilir.
    ws.update([COLUMNS] + rows, value_input_option="RAW")

    ws.freeze(rows=1, cols=8)

    # E sutunu (Create Date): deger ISO (yyyy-MM-dd) yazildi, GORUNUMU dd.MM.yyyy
    # olarak ayarliyoruz -- hem gercek tarih hucresi (filter view'lar calisir)
    # hem de istenen goruntu formati.
    ws.format("E2:E", {"numberFormat": {"type": "DATE", "pattern": "dd.MM.yyyy"}})

    # Renklendirme (F: Ozet, G: Musteri Durumu) -- batch_format ile
    requests_batch = []
    for row_idx, colors in color_map.items():
        sheet_row = row_idx + 2  # +1 header, +1 1-indexed
        for col_letter, style in colors.items():
            requests_batch.append({
                "range": f"{col_letter}{sheet_row}",
                "format": {
                    "backgroundColor": _hex_to_rgb(style.get("bg")),
                    "textFormat": {
                        "foregroundColor": _hex_to_rgb(style.get("fg", "#000000")),
                        "bold": style.get("bold", False),
                    },
                    "wrapStrategy": "WRAP",
                },
            })
    if requests_batch:
        ws.batch_format(requests_batch)


def _hex_to_rgb(hex_color):
    hex_color = (hex_color or "#FFFFFF").lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return {"red": r, "green": g, "blue": b}


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------

def assemble_and_write(ws, rows):
    """
    rows: [(create_date_raw, row_list, musteri_renk), ...]
    En yeni Create Date en usте olacak sekilde siralar, mevcut A/B manuel
    notlarini (Record ID eslesmesiyle) korur, ve sheet'e yazar.
    """
    rows = sorted(rows, key=lambda r: r[0] or "0", reverse=True)
    manual_by_record_id = read_manual_columns(ws)

    final_rows = []
    color_map = {}
    for idx, (_, row, musteri_renk) in enumerate(rows):
        record_id = row[-1]
        a_val, b_val = manual_by_record_id.get(record_id, ("", ""))
        row[0] = a_val
        row[1] = b_val
        final_rows.append(row)
        color_map[idx] = {"G": musteri_renk}

    write_sheet(ws, final_rows, color_map)
    return len(final_rows)


def main():
    log.info("HubSpot contact'lari cekiliyor...")
    contacts = fetch_all_contacts()
    contact_ids = [c["id"] for c in contacts]

    log.info("Sales Pipeline stage haritasi cekiliyor...")
    sales_pipeline_id, stage_map = fetch_sales_pipeline_stage_map()

    log.info("Deal stage'ler cekiliyor...")
    deal_stage_by_contact = fetch_deal_stage_for_contacts(contact_ids, sales_pipeline_id, stage_map)

    log.info("Note/Call association'lari cekiliyor...")
    note_assoc = fetch_associations_batch(contact_ids, "notes")
    call_assoc = fetch_associations_batch(contact_ids, "calls")
    all_note_ids = [n for ids in note_assoc.values() for n in ids]
    all_call_ids = [c for ids in call_assoc.values() for c in ids]
    notes_data = batch_read_objects("notes", all_note_ids, ["hs_note_body", "hs_timestamp", "hubspot_owner_id"])
    calls_data = batch_read_objects("calls", all_call_ids, ["hs_call_body", "hs_call_title", "hs_timestamp", "hubspot_owner_id"])

    # TODO: owners.read izni varsa /crm/v3/owners ile gercek isimleri cek.
    # Simdilik bos -- fallback olarak owner ID ham haliyle yaziliyor.
    owners_map = {}

    rows = []

    for contact in contacts:
        props = contact.get("properties", {})
        record_id = contact["id"]

        source = compute_source(props)
        ilk_form = props.get("first_conversion_event_name") or "-"
        create_date_raw = props.get("createdate")
        create_date = format_hs_date(create_date_raw, storage=True) or ""

        webinar_list = parse_multivalue_property(props.get("kat_ld_g__webinarlar"))
        gc_list = parse_multivalue_property(props.get("kat_ld_g__tum_growthclub_oturumlar_"))
        planning_list = parse_multivalue_property(props.get("katldg_planlama_oturumlar"))

        note_props = [notes_data[nid] for nid in note_assoc.get(record_id, []) if nid in notes_data]
        call_props = [calls_data[cid] for cid in call_assoc.get(record_id, []) if cid in calls_data]
        notes_text, note_blocks = build_notes_and_calls_text(note_props, call_props, owners_map)

        ozet = build_ozet_c_column(webinar_list, gc_list, planning_list, note_blocks)
        musteri_durumu, musteri_renk = compute_musteri_durumu(props.get("masterclass_process"))

        pozisyon = props.get("pozisyon") or props.get("job_title") or ""
        sektor = props.get("sektor_v2") or props.get("faaliyet_gosterdiginiz_sektor") or ""
        segment = compute_segment(props.get("yillik_ciro"), pozisyon)

        row = [
            "",  # A - manuel (asagida merge edilecek)
            "",  # B - manuel (asagida merge edilecek)
            source,
            ilk_form,
            create_date,
            ozet,
            musteri_durumu,
            props.get("firstname") or "",
            props.get("lastname") or "",
            props.get("company") or "",
            segment,
            pozisyon,
            props.get("sirket_yaps") or "",
            props.get("yillik_ciro") or "",
            props.get("cal_san_say_s_") or "",
            sektor,
            "; ".join(webinar_list),
            "; ".join(gc_list),
            "; ".join(planning_list),
            notes_text,
            deal_stage_by_contact.get(record_id, ""),
            props.get("phone") or "",
            props.get("email") or "",
            props.get("hs_marketable_status") or "",
            props.get("hs_email_optout") or "",
            props.get("masterclass_process") or "",
            record_id,
        ]
        rows.append((create_date_raw or "0", row, musteri_renk))

    # En yeni Create Date en usте
    rows.sort(key=lambda r: r[0] or "0", reverse=True)

    sh = connect_spreadsheet()

    ws_all = get_or_create_worksheet(sh, SHEET_NAME)
    n = assemble_and_write(ws_all, rows)
    log.info(f"{SHEET_NAME}: {n} satir yazildi.")

    log.info(f"HubSpot listesi ({LIST_ID}) uyeleri cekiliyor...")
    list_member_ids = fetch_list_membership_ids(LIST_ID)
    list_name = fetch_list_name(LIST_ID)
    list_rows = [r for r in rows if r[1][-1] in list_member_ids]

    ws_list = get_or_create_worksheet(sh, list_name)
    n_list = assemble_and_write(ws_list, list_rows)
    log.info(f"{list_name}: {n_list} satir yazildi.")

    log.info("Tamamlandi.")


if __name__ == "__main__":
    main()
