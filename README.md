# HubSpot -> Google Sheets Gunluk Senkronizasyon

ActionCoach Turkiye genel lead takip Excel'i. Her gun HubSpot'taki TUM
contact'lari ceker, Sheet1'i bastan yazar (A/B manuel notlar korunur), en
yeni Create Date en usте olacak sekilde siralar.

## Kurulum

### 1) Repo secrets (Settings > Secrets and variables > Actions)

| Secret adi | Icerik |
|---|---|
| `HUBSPOT_PRIVATE_APP_TOKEN` | HubSpot Private App token'i |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Service account'un tum JSON key dosyasi (tek satirda, ic Ice) |
| `SPREADSHEET_ID` | `1ZaTNGfpbLvkR-E9WaMJjnFoFEs6HPIIFpwohQHg9xcU` |

### 2) HubSpot Private App izinleri

- `crm.objects.contacts.read`
- `crm.objects.notes.read`
- `crm.objects.calls.read`
- `crm.objects.deals.read`
- `crm.schemas.deals.read` (pipeline/stage listesi icin)
- (varsa) `crm.objects.owners.read` -- not/call yazan kisi isimlerini cozmek icin

### 3) Google Service Account

1. GCP Console'da bir service account olustur, JSON key indir.
2. Bu JSON'un tamamini `GOOGLE_SERVICE_ACCOUNT_JSON` secret'ina yapistir.
3. Service account'un email'ini (ornek: `xxx@yyy.iam.gserviceaccount.com`)
   hedef Google Sheet'e **Editor** olarak paylas -- bu adim atlanirsa script
   "permission denied" hatasi verir.

### 4) Filter view'lari kur (tek seferlik)

```
python setup_filter_views.py
```

Bu, Yesterday / This Week / This Month / This Quarter / This Year adinda 5
filter view olusturur. **Test edilmedi** -- calistirdiktan sonra Google
Sheets'te Data > Filter views'a bak, dogru satirlari filtreliyor mu kontrol
et.

**Sorun cikarsa manuel kurulum (garanti calisir, ~2 dk):**
1. Sheets'te Data > Filter views > Create new filter view
2. Isim ver (ornek: "This Month")
3. E sutunu basligina tikla > Filter by condition > Custom formula is
4. Formul: `=$E2>=EOMONTH(TODAY(),-1)+1` (diger view'lar icin `setup_filter_views.py`
   icindeki VIEWS listesindeki formulleri kullan)
5. Save

### 5) Ilk test calistirmasi

Actions sekmesinden workflow'u **workflow_dispatch** ile elle tetikle (cron'u
beklemeden). Sheet1'e veri dustugunu gordukten sonra gunluk cron'a guvenebilirsin.

---

## Netlesmemis / varsayima dayanan noktalar

Ilk gercek calistirmada MUTLAKA kontrol edilmesi gerekenler:

1. **Multi-value alanlar** (`kat_ld_g__webinarlar`, `kat_ld_g__tum_growthclub_oturumlar_`,
   `katldg_planlama_oturumlar`) -- script bunlarin `;` ile ayrilmis liste oldugunu
   varsayiyor (`parse_multivalue_property`). Gercek format farkliysa (ornek:
   virgul, veya tarih+isim birlikte tek string) bu fonksiyonu guncellememiz lazim.

2. **F sutunu (Ozet) icindeki "Son Katildigi X"** -- yukaridaki liste sadece
   isim/etiket iceriyorsa "en son" bilgisini dogru veremeyiz (tarih bilgisi
   yok). Eger bu alanlarda tarih de varsa, parse fonksiyonunu (tarih, isim)
   ciftine gore genisletip gercekten en yeniyi secmemiz lazim.

3. **Batch association guvenilirligi** -- Fransa migrasyon projesinde task
   association'lari icin batch okuma ~%60 basariliydi, tekil v4 cagrisina
   donulmustu. Bu script once batch deniyor (`fetch_associations_batch`);
   ilk gercek calistirmada notes/calls/deals sayilari HubSpot arayuzundekiyle
   uyusmuyorsa, bu fonksiyonu chunk'lanmis tekil cagriya cevirmemiz gerekebilir.

4. **Owner isimleri** -- `owners_map` şu an bos (`{}`), yani Notes/Call
   Notes'ta yazan kisi ID olarak gorunecek, isim olarak degil.
   `crm.objects.owners.read` izni varsa `/crm/v3/owners` endpoint'inden
   `owners_map = {str(owner["id"]): f'{owner["firstName"]} {owner["lastName"]}' ...}`
   seklinde doldurulmasi lazim -- bu fonksiyon henuz yazilmadi.

5. **"F: Ozet" hucre-ici renklendirme** -- şu an sadece G (Musteri Durumu)
   hucresi tek renkle boyaniyor. F sutunundaki 4 satirin HER BIRINE FARKLI
   renk vermek (mavi/yesil/turuncu/mor) Sheets API'de "rich text" (per-run
   formatting) gerektiriyor -- bu, `updateCells` + `textFormatRuns` ile ayri
   bir implementasyon istiyor, henuz eklenmedi.

6. **Kredi/kart/kota** -- HubSpot API Call Usage ekraninda gunluk limit
   625.000, su an %15 kullanimda -- bu script gunluk 1 kez calistigi surece
   sorun cikmaz.

Bu 6 maddeyi ilk canli veriyle birlikte gozden gecirip duzeltelim.
