import os, re, json, io, base64, zipfile, datetime as dt
from typing import Dict, List, Any, Optional, Tuple, Set

import requests
from PIL import Image

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

# EasyOCR
import easyocr
import numpy as np


# ============================================================
# ENV / CONFIG
# ============================================================

def env_str(name: str, default: str = "") -> str:
    v = os.environ.get(name, default)
    if v is None:
        return default
    return str(v).strip()

def env_int(name: str, default: int) -> int:
    s = env_str(name, "")
    if s == "":
        return default
    try:
        return int(s)
    except Exception:
        return default

def today_ymd_utc() -> str:
    return dt.date.today().strftime("%Y-%m-%d")


WORKER_BASE_URL    = env_str("WORKER_BASE_URL")
EXPORT_TOKEN       = env_str("EXPORT_TOKEN")
TELEGRAM_TOKEN     = env_str("TELEGRAM_TOKEN")

BREVO_API_KEY      = env_str("BREVO_API_KEY")
BREVO_SENDER_EMAIL = env_str("BREVO_SENDER_EMAIL", "no-reply@example.com")
BREVO_SENDER_NAME  = env_str("BREVO_SENDER_NAME",  "Prospection Bot")

MAIL_ROUTING_JSON  = env_str("MAIL_ROUTING_JSON", "")

GOOGLE_PLACES_API_KEY = env_str("GOOGLE_PLACES_API_KEY", "")
GEMINI_API_KEY        = env_str("GEMINI_API_KEY", "")  # plus utilisé ici (on simplifie)

SEND_MODE    = env_str("SEND_MODE", "individual").lower()  # individual | agency_manager | admin
RUN_DATE     = env_str("RUN_DATE", today_ymd_utc())
AGENCY       = env_str("AGENCY", "").upper()
INITIALS     = env_str("INITIALS", "").upper()

MAX_OCR_IMAGES   = env_int("MAX_OCR_IMAGES", 50)
MAX_PHOTO_IMAGES = env_int("MAX_PHOTO_IMAGES", 15)

OUT_DIR = env_str("OUT_DIR", "out").strip() or "out"
if os.path.exists(OUT_DIR) and not os.path.isdir(OUT_DIR):
    OUT_DIR = "exports"
os.makedirs(OUT_DIR, exist_ok=True)

VALID_AGENCIES = {"GR", "VR", "GRS", "SLS"}

# ============================================================
# REGEX
# ============================================================

EMAIL_RE_TEXT = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"(https?://[^\s]+|www\.[^\s]+)", re.I)
PHONE_RE = re.compile(r"(\+33|0)\s*[1-9](?:[\s\.-]*\d{2}){4}")
CP_RE = re.compile(r"\b\d{5}\b")

SIRET_RE = re.compile(r"\b\d{14}\b")
SIREN_RE = re.compile(r"\b\d{9}\b")

EMAIL_VALID = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

HEADERS = [
    "date","agency","initials","name","address","postal_code","city",
    "siret","naf","dirigeant","interlocuteur","contact_firstname","contact_lastname",
    "phone","phone2","email","website","resume","commande"
]

# ============================================================
# ROUTING
# ============================================================

DEFAULT_ROUTING = {
    "agencies": {
        "GR": {"manager": {"initials": "JL", "email": "jennifer.laurens@ras-interim.fr"},
               "commercial": {"initials": "CZ", "email": "celine.zunarelli@ras-interim.fr"}},
        "VR": {"manager": {"initials": "JB", "email": "jelena.carrasso@ras-interim.fr"},
               "commercial": {"initials": "LB", "email": "laura.berthet@ras-interim.fr"}},
        "GRS": {"manager": {"initials": "PV", "email": "pauline.vieira@ras-interim.fr"},
                "commercial": {"initials": "ST", "email": "severine.thevenin@ras-interim.fr"}},
        "SLS": {"manager": {"initials": "AC", "email": "aurelie.curt@ras-interim.fr"},
                "commercial": {"initials": "AC", "email": "aurelie.curt@ras-interim.fr"}},
    },
    "users": {
        "JL": "jennifer.laurens@ras-interim.fr",
        "CZ": "celine.zunarelli@ras-interim.fr",
        "JB": "jelena.carrasso@ras-interim.fr",
        "LB": "laura.berthet@ras-interim.fr",
        "PV": "pauline.vieira@ras-interim.fr",
        "ST": "severine.thevenin@ras-interim.fr",
        "AC": "aurelie.curt@ras-interim.fr",
        "SL": "samuel.lengue@ras-interim.fr",
    },
    "admin": {"initials": "SL", "email": "samuel.lengue@ras-interim.fr"},
}

def load_routing() -> Dict[str, Any]:
    if MAIL_ROUTING_JSON:
        try:
            data = json.loads(MAIL_ROUTING_JSON)
            if isinstance(data, dict) and data:
                return data
        except Exception:
            pass
    return DEFAULT_ROUTING

ROUTING = load_routing()

def clean_email(s: str) -> str:
    return re.sub(r"\s+", "", (s or "").strip())

def is_valid_email(s: str) -> bool:
    s = clean_email(s)
    return bool(EMAIL_VALID.match(s))

def email_for_initials(initials: str) -> Optional[str]:
    initials = (initials or "").upper().strip()
    if not initials:
        return None
    users = ROUTING.get("users") or {}
    if initials in users:
        em = clean_email(users[initials])
        return em if is_valid_email(em) else None
    agencies = ROUTING.get("agencies") or {}
    for _, cfg in agencies.items():
        for role in ("manager","commercial"):
            r = (cfg or {}).get(role) or {}
            if (r.get("initials") or "").upper() == initials:
                em2 = clean_email(r.get("email") or "")
                return em2 if is_valid_email(em2) else None
    return None


# ============================================================
# HTTP HELPERS
# ============================================================

def worker_dump(kind: str, date: str) -> List[Dict[str, Any]]:
    url = f"{WORKER_BASE_URL.rstrip('/')}/dump?date={date}&kind={kind}"
    r = requests.get(url, headers={"X-Export-Token": EXPORT_TOKEN}, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"/dump failed {kind} {r.status_code}: {r.text[:400]}")
    out: List[Dict[str, Any]] = []
    for ln in [x.strip() for x in r.text.splitlines() if x.strip()]:
        try:
            out.append(json.loads(ln))
        except Exception:
            pass
    return out

def tg_get_file_url(file_id: str) -> Optional[str]:
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile"
    r = requests.post(url, json={"file_id": file_id}, timeout=20)
    j = r.json()
    if not j.get("ok"):
        return None
    fp = j["result"]["file_path"]
    return f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{fp}"

def download_bytes(url: str) -> bytes:
    r = requests.get(url, timeout=40)
    r.raise_for_status()
    return r.content


# ============================================================
# NORMALISATION
# ============================================================

def normalize_phone(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[^\d+]", "", s)
    s = s.replace("+33", "0") if s.startswith("+33") else s
    return s

def pick_best_url(text: str) -> str:
    m = URL_RE.search(text or "")
    if not m:
        return ""
    u = m.group(0).strip().rstrip(").,;")
    if u.lower().startswith("www."):
        u = "http://" + u
    return u

def extract_domain(email: str) -> str:
    email = (email or "").strip().lower()
    if "@" not in email:
        return ""
    dom = email.split("@", 1)[1]
    dom = dom.split("/", 1)[0]
    return dom

def domain_to_company(dom: str) -> str:
    """
    petavit.com -> PETAVIT
    ras-interim.fr -> RAS INTERIM (simplifié)
    """
    dom = (dom or "").lower().strip()
    if not dom:
        return ""
    dom = dom.replace("www.", "")
    core = dom.split(":")[0].split("/")[0]
    core = core.split(".")[0]  # avant .fr/.com
    core = core.replace("-", " ").replace("_", " ")
    core = re.sub(r"\s+", " ", core).strip()
    return core.upper()

def split_human_name(full: str) -> Tuple[str, str]:
    s = re.sub(r"\s+", " ", (full or "").strip())
    if not s:
        return "", ""
    parts = s.split(" ")
    if len(parts) == 1:
        return "", parts[0]
    return parts[0], " ".join(parts[1:])

def ensure_contact_fallback(d: Dict[str, Any]) -> Dict[str, Any]:
    if not (d.get("interlocuteur") or "").strip():
        fn = (d.get("contact_firstname") or "").strip()
        ln = (d.get("contact_lastname") or "").strip()
        combo = f"{fn} {ln}".strip()
        if combo:
            d["interlocuteur"] = combo
    return d


# ============================================================
# EasyOCR (1 instance)
# ============================================================

READER = None

def get_reader():
    global READER
    if READER is None:
        # CPU only
        READER = easyocr.Reader(['fr','en'], gpu=False)
    return READER

def easyocr_text(img_bytes: bytes) -> str:
    if not img_bytes:
        return ""
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    arr = np.array(im)
    reader = get_reader()
    parts = reader.readtext(arr, detail=0, paragraph=True)
    txt = " ".join([p.strip() for p in parts if p and str(p).strip()])
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


# ============================================================
# CP -> Ville (API Adresse)
# ============================================================

def city_from_postal(cp: str) -> str:
    cp = (cp or "").strip()
    if not cp:
        return ""
    try:
        # API adresse (simple & rapide)
        r = requests.get(
            "https://api-adresse.data.gouv.fr/search/",
            params={"q": cp, "type": "municipality", "limit": 1},
            timeout=10
        )
        j = r.json()
        feats = j.get("features") or []
        if not feats:
            return ""
        props = feats[0].get("properties") or {}
        return (props.get("city") or "").strip()
    except Exception:
        return ""


# ============================================================
# API GOUV
# ============================================================

def search_gouv_company(name: str, city: str) -> Dict[str, str]:
    name = (name or "").strip()
    city = (city or "").strip()
    if not name:
        return {}
    try:
        q = f"{name} {city}".strip()
        url = "https://recherche-entreprises.api.gouv.fr/search"
        r = requests.get(url, params={"q": q, "page": 1, "per_page": 1}, timeout=15)
        j = r.json()
        res = (j.get("results") or [])
        if not res:
            return {}
        e = res[0]
        siege = e.get("siege") or {}
        naf = (e.get("activite_principale") or e.get("naf") or "").replace(".","").replace(" ","").upper()
        return {
            "name": e.get("nom_raison_sociale") or e.get("denomination") or name,
            "siret": (siege.get("siret") or ""),
            "naf": naf,
            "address": (siege.get("adresse") or siege.get("libelle_voie") or ""),
            "postal_code": (siege.get("code_postal") or ""),
            "city": (siege.get("libelle_commune") or city),
            "dirigeant": "",  # on évite de complexifier ici
        }
    except Exception:
        return {}


# ============================================================
# Google Places (optionnel)
# ============================================================

def places_enrich(name: str, city: str) -> Dict[str, str]:
    if not GOOGLE_PLACES_API_KEY or not name:
        return {}
    try:
        r1 = requests.post(
            "https://places.googleapis.com/v1/places:searchText",
            headers={
                "Content-Type":"application/json",
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask":"places.id,places.websiteUri",
            },
            json={
                "textQuery": f"{name} {city}".strip(),
                "maxResultCount": 1,
                "languageCode":"fr",
                "regionCode":"FR",
            },
            timeout=15
        )
        j1 = r1.json()
        p = (j1.get("places") or [None])[0]
        if not p or not p.get("id"):
            return {}
        pid = p["id"]
        website = (p.get("websiteUri") or "").strip()
        r2 = requests.get(
            f"https://places.googleapis.com/v1/places/{pid}",
            headers={
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask":"nationalPhoneNumber,internationalPhoneNumber,websiteUri",
            },
            timeout=15
        )
        j2 = r2.json()
        phone = j2.get("nationalPhoneNumber") or j2.get("internationalPhoneNumber") or ""
        website2 = (j2.get("websiteUri") or website).strip()
        return {"phone": normalize_phone(phone), "website": website2}
    except Exception:
        return {}


# ============================================================
# Email scrape (optionnel)
# ============================================================

def scrape_email_from_site(url: str) -> str:
    if not url:
        return ""
    try:
        r = requests.get(url, headers={"user-agent":"Mozilla/5.0 (ProspectionBot)"}, timeout=8)
        if r.status_code >= 300:
            return ""
        html = r.text[:200000]
        matches = EMAIL_RE_TEXT.findall(html)
        if not matches:
            return ""
        filt = [m for m in matches if not re.search(r"no-?reply", m, re.I)]
        return (filt[0] if filt else matches[0]).strip().lower()
    except Exception:
        return ""


# ============================================================
# MEDIA ZIP
# ============================================================

def build_media_zip(date: str, agency: str, initials: str,
                    photos: List[Dict[str, Any]], cards: List[Dict[str, Any]]) -> Optional[Tuple[str, bytes]]:
    photos_u = [p for p in photos if (p.get("agency") == agency and (p.get("user") or "").upper() == initials)]
    cards_u  = [c for c in cards  if (c.get("agency") == agency and (c.get("user") or "").upper() == initials)]
    if not photos_u and not cards_u:
        return None

    photos_u = photos_u[:MAX_PHOTO_IMAGES]
    cards_u  = cards_u[:MAX_OCR_IMAGES]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        lines = ["type;file;date;agency;initials;city;comment;geo_lat;geo_lon"]

        for i, p in enumerate(photos_u, start=1):
            fid = p.get("file_id") or ""
            url = tg_get_file_url(fid)
            if not url:
                continue
            img = download_bytes(url)
            fname = f"photos/{date}_{agency}_{initials}_{i:02d}.jpg"
            z.writestr(fname, img)

            geo = p.get("geo") or {}
            comment = p.get("comment") or p.get("meeting") or ""
            lines.append(
                f"photo;{fname};{date};{agency};{initials};{(p.get('city') or '')};"
                f"{comment};{geo.get('lat') or ''};{geo.get('lon') or ''}"
            )

        for i, c in enumerate(cards_u, start=1):
            fid = c.get("file_id") or ""
            url = tg_get_file_url(fid)
            if not url:
                continue
            img = download_bytes(url)
            fname = f"cards/{date}_{agency}_{initials}_{i:02d}.jpg"
            z.writestr(fname, img)
            lines.append(f"card;{fname};{date};{agency};{initials};;{(c.get('comment') or '')};;")

        z.writestr("index.csv", ("\n".join(lines)).encode("utf-8"))

    zip_name = f"MEDIA_{date}_{agency}_{initials}.zip"
    return zip_name, buf.getvalue()


# ============================================================
# BREVO EMAIL
# ============================================================

def brevo_send_email(to_email: str, subject: str, html: str, attachments: Optional[List[Tuple[str, bytes]]] = None):
    if not BREVO_API_KEY:
        raise RuntimeError("BREVO_API_KEY missing")
    to_email = clean_email(to_email)
    if not is_valid_email(to_email):
        raise RuntimeError(f"Invalid recipient email: '{to_email}'")

    payload = {
        "sender": {"name": BREVO_SENDER_NAME, "email": BREVO_SENDER_EMAIL},
        "to": [{"email": to_email}],
        "subject": subject,
        "htmlContent": html,
    }

    if attachments:
        payload["attachment"] = []
        for filename, content in attachments:
            payload["attachment"].append({
                "name": filename,
                "content": base64.b64encode(content).decode("ascii")
            })

    r = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={
            "accept": "application/json",
            "api-key": BREVO_API_KEY,
            "content-type": "application/json"
        },
        data=json.dumps(payload),
        timeout=40
    )

    if r.status_code >= 300:
        raise RuntimeError(f"Brevo send failed {r.status_code}: {r.text[:400]}")


# ============================================================
# EXCEL
# ============================================================

def autosize(ws):
    for col in range(1, ws.max_column + 1):
        max_len = 10
        for row in range(1, ws.max_row + 1):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            s = str(v)
            if len(s) > 80:
                s = s[:80]
            max_len = max(max_len, len(s))
            max_len = min(max_len, 60)
        ws.column_dimensions[get_column_letter(col)].width = max_len

def record_score(d: Dict[str, Any]) -> int:
    fields = ["name","address","postal_code","city","siret","naf","dirigeant","interlocuteur",
              "contact_firstname","contact_lastname","phone","phone2","email","website","resume","commande"]
    return sum(1 for k in fields if str(d.get(k,"") or "").strip() != "")

def to_row(d: Dict[str, Any]) -> List[Any]:
    return [d.get(h, "") for h in HEADERS]

def build_excel_one_sheet(date: str, rows: List[Dict[str, Any]], suffix: str) -> str:
    rows_sorted = sorted(rows, key=lambda x: record_score(x), reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"

    ws.append(HEADERS)
    for r in rows_sorted:
        ws.append(to_row(r))

    for c in range(1, len(HEADERS)+1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    autosize(ws)

    filename = os.path.join(OUT_DIR, f"PROSPECTION_{date}_{suffix}.xlsx")
    wb.save(filename)
    return filename


# ============================================================
# UNIFY SOURCES
# ============================================================

def unify_from_prospect(p: Dict[str, Any]) -> Dict[str, Any]:
    d = {h: "" for h in HEADERS}
    d.update({
        "date": p.get("date",""),
        "agency": (p.get("agency") or "").upper(),
        "initials": (p.get("initials") or "").upper(),
        "name": p.get("name",""),
        "address": p.get("address",""),
        "postal_code": p.get("postal_code",""),
        "city": p.get("city",""),
        "siret": p.get("siret",""),
        "naf": p.get("naf",""),
        "dirigeant": p.get("dirigeant",""),
        "interlocuteur": p.get("interlocuteur",""),
        "contact_firstname": p.get("contact_firstname",""),
        "contact_lastname": p.get("contact_lastname",""),
        "phone": p.get("phone",""),
        "phone2": p.get("phone2",""),
        "email": p.get("email",""),
        "website": p.get("website",""),
        "resume": p.get("resume",""),
        "commande": p.get("commande",""),
    })
    return ensure_contact_fallback(d)

def extract_from_text_card(txt: str) -> Dict[str, str]:
    txt = txt or ""
    out = {"email":"", "phone":"", "phone2":"", "postal_code":"", "website":"", "siret":""}

    emails = EMAIL_RE_TEXT.findall(txt)
    if emails:
        out["email"] = emails[0].lower()

    phones = PHONE_RE.findall(txt)
    # PHONE_RE.findall renvoie des tuples; on refait search global pour récupérer la chaîne
    ph = PHONE_RE.findall(txt)
    ph_list = [normalize_phone(m.group(0)) for m in PHONE_RE.finditer(txt)]
    if ph_list:
        out["phone"] = ph_list[0]
        if len(ph_list) > 1:
            out["phone2"] = ph_list[1]

    cp = CP_RE.search(txt)
    if cp:
        out["postal_code"] = cp.group(0)

    url = pick_best_url(txt)
    out["website"] = url

    siret = SIRET_RE.search(txt)
    if siret:
        out["siret"] = siret.group(0)

    return out

def unify_from_card(date: str, agency: str, initials: str, comment: str, img_bytes: bytes) -> Dict[str, Any]:
    base = {h: "" for h in HEADERS}
    base.update({
        "date": date,
        "agency": agency,
        "initials": initials,
        "resume": (comment or "").strip(),
    })

    if not img_bytes:
        return base

    txt = easyocr_text(img_bytes)
    feats = extract_from_text_card(txt)

    email = feats.get("email","")
    domain = extract_domain(email)
    company_guess = domain_to_company(domain)

    cp = feats.get("postal_code","")
    city = city_from_postal(cp)

    # Recherche gouv = entreprise déduite + ville
    gouv = search_gouv_company(company_guess, city) if company_guess else {}

    # Places = complément
    place = places_enrich(gouv.get("name") or company_guess, gouv.get("city") or city)

    website = feats.get("website") or place.get("website") or ""
    if website and not email:
        email = scrape_email_from_site(website) or ""

    # contact name: on prend 2 mots capitalisés les plus plausibles (ultra simple)
    contact = ""
    # Heuristique: chercher deux mots avec majuscules
    m = re.search(r"\b([A-ZÉÈÊËÀÂÄÎÏÔÖÛÜÇ][a-zéèêëàâäîïôöûüç]+)\s+([A-ZÉÈÊËÀÂÄÎÏÔÖÛÜÇ][a-zéèêëàâäîïôöûüç]+)\b", txt)
    if m:
        contact = f"{m.group(1)} {m.group(2)}".strip()
    fn, ln = split_human_name(contact)

    row = dict(base)
    row.update({
        "name": gouv.get("name") or company_guess or "",
        "address": gouv.get("address") or "",
        "postal_code": gouv.get("postal_code") or cp or "",
        "city": gouv.get("city") or city or "",

        "siret": gouv.get("siret") or feats.get("siret") or "",
        "naf": gouv.get("naf") or "",
        "dirigeant": gouv.get("dirigeant") or "",

        "interlocuteur": contact,
        "contact_firstname": fn,
        "contact_lastname": ln,

        "phone": feats.get("phone") or place.get("phone") or "",
        "phone2": feats.get("phone2") or "",
        "email": email or "",
        "website": website or "",
    })
    return ensure_contact_fallback(row)

def unify_from_photo(date: str, agency: str, initials: str, city_hint: str, comment: str, img_bytes: bytes) -> Dict[str, Any]:
    # Photo entreprise: on ne fait plus Gemini ici (simple).
    # On essaie juste OCR -> domaine/email -> gouv+places
    base = {h: "" for h in HEADERS}
    base.update({
        "date": date,
        "agency": agency,
        "initials": initials,
        "resume": (comment or "").strip(),
        "city": (city_hint or "").strip(),
    })

    if not img_bytes:
        return base

    txt = easyocr_text(img_bytes)
    feats = extract_from_text_card(txt)

    email = feats.get("email","")
    domain = extract_domain(email)
    company_guess = domain_to_company(domain)

    cp = feats.get("postal_code","")
    city = base.get("city") or city_from_postal(cp)

    gouv = search_gouv_company(company_guess, city) if company_guess else {}
    place = places_enrich(gouv.get("name") or company_guess, gouv.get("city") or city)

    website = feats.get("website") or place.get("website") or ""
    if website and not email:
        email = scrape_email_from_site(website) or ""

    row = dict(base)
    row.update({
        "name": gouv.get("name") or company_guess or "",
        "address": gouv.get("address") or "",
        "postal_code": gouv.get("postal_code") or cp or "",
        "city": gouv.get("city") or city or "",
        "siret": gouv.get("siret") or feats.get("siret") or "",
        "naf": gouv.get("naf") or "",
        "phone": feats.get("phone") or place.get("phone") or "",
        "email": email or "",
        "website": website or "",
    })
    return row


# ============================================================
# SEND LOGIC
# ============================================================

def uniq_users(records: List[Dict[str, Any]]) -> List[Tuple[str, str]]:
    seen = set()
    out: List[Tuple[str, str]] = []
    for r in records:
        ag = (r.get("agency") or "").upper()
        ini = (r.get("initials") or r.get("user") or "").upper()
        if ag in VALID_AGENCIES and ini:
            k = (ag, ini)
            if k not in seen:
                seen.add(k)
                out.append(k)
    return out

def send_individual_pack(date: str, agency: str, initials: str,
                         prospects: List[Dict[str, Any]],
                         photos: List[Dict[str, Any]],
                         cards: List[Dict[str, Any]]):

    to_email = email_for_initials(initials)
    if not to_email:
        print(f"[WARN] No email for initials={initials}, skip.")
        return

    p_u = [p for p in prospects if (p.get("agency") == agency and (p.get("initials") or "").upper() == initials)]
    ph_u = [ph for ph in photos if (ph.get("agency") == agency and (ph.get("user") or "").upper() == initials)]
    ca_u = [ca for ca in cards  if (ca.get("agency") == agency and (ca.get("user") or "").upper() == initials)]

    rows: List[Dict[str, Any]] = []
    rows.extend([unify_from_prospect(p) for p in p_u])

    for ph in ph_u[:MAX_PHOTO_IMAGES]:
        fid = ph.get("file_id") or ""
        url = tg_get_file_url(fid) if fid else None
        city_hint = ph.get("city") or ""
        comment = ph.get("comment") or ph.get("meeting") or ""
        try:
            img = download_bytes(url) if url else b""
        except Exception:
            img = b""
        rows.append(unify_from_photo(date, agency, initials, city_hint, comment, img))

    for ca in ca_u[:MAX_OCR_IMAGES]:
        fid = ca.get("file_id") or ""
        url = tg_get_file_url(fid) if fid else None
        comment = ca.get("comment") or ""
        try:
            img = download_bytes(url) if url else b""
        except Exception:
            img = b""
        rows.append(unify_from_card(date, agency, initials, comment, img))

    if not rows:
        print(f"[INFO] No rows for {agency}/{initials}, skip.")
        return

    xlsx = build_excel_one_sheet(date, rows, f"INDIV_{agency}_{initials}")

    attachments: List[Tuple[str, bytes]] = []
    with open(xlsx, "rb") as f:
        attachments.append((os.path.basename(xlsx), f.read()))

    media = build_media_zip(date, agency, initials, photos, cards)
    if media:
        attachments.append(media)

    subject = f"Prospection {date} — {agency}/{initials} (Excel + médias)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Export prospection du <b>{date}</b> pour <b>{agency}/{initials}</b>.</p>"
        f"<ul>"
        f"<li>Excel 1 onglet (trié par complétude)</li>"
        f"<li>ZIP médias (photos + cartes)</li>"
        f"</ul>"
        f"<p>— Bot Prospection</p>"
    )
    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] sent to {to_email} ({agency}/{initials})")

def send_agency_manager_pack(date: str, agency: str, prospects: List[Dict[str, Any]]):
    agencies_cfg = ROUTING.get("agencies") or {}
    cfg = agencies_cfg.get(agency) or {}
    manager = (cfg.get("manager") or {})
    to_email = clean_email(manager.get("email") or "")
    if not is_valid_email(to_email):
        return

    p_ag = [p for p in prospects if (p.get("agency") == agency)]
    rows = [unify_from_prospect(p) for p in p_ag]
    if not rows:
        return

    xlsx = build_excel_one_sheet(date, rows, f"AGENCE_{agency}")
    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    brevo_send_email(
        to_email,
        f"Prospection {date} — Agence {agency}",
        f"<p>Consolidé {date} agence <b>{agency}</b> (1 onglet).</p>",
        attachments=attachments
    )

def send_admin_pack(date: str, prospects: List[Dict[str, Any]]):
    admin = ROUTING.get("admin") or {}
    to_email = clean_email(admin.get("email") or "")
    if not is_valid_email(to_email):
        return

    rows = [unify_from_prospect(p) for p in prospects]
    if not rows:
        return

    xlsx = build_excel_one_sheet(date, rows, "ADMIN_ALL")
    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    brevo_send_email(
        to_email,
        f"Prospection {date} — ADMIN",
        f"<p>Consolidé global {date} (1 onglet).</p>",
        attachments=attachments
    )


# ============================================================
# MAIN
# ============================================================

def main():
    if not WORKER_BASE_URL or not EXPORT_TOKEN or not TELEGRAM_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN / TELEGRAM_TOKEN")

    run_date = RUN_DATE if RUN_DATE else today_ymd_utc()

    print(f"🚀 export_and_mail.py — mode={SEND_MODE} date={run_date} agency={AGENCY} initials={INITIALS}")
    print(f"📦 OUT_DIR={OUT_DIR}")
    print(f"🔑 EasyOCR=ON | Places={'ON' if GOOGLE_PLACES_API_KEY else 'OFF'}")

    prospects = worker_dump("prospects", run_date)
    photos    = worker_dump("photos",    run_date)
    cards     = worker_dump("cards",     run_date)

    if SEND_MODE == "individual":
        if AGENCY not in VALID_AGENCIES:
            raise RuntimeError("AGENCY required (GR|VR|GRS|SLS) for mode=individual")
        if not INITIALS:
            raise RuntimeError("INITIALS required for mode=individual")
        send_individual_pack(run_date, AGENCY, INITIALS, prospects, photos, cards)
        return

    if SEND_MODE == "agency_manager":
        for ag in sorted(VALID_AGENCIES):
            send_agency_manager_pack(run_date, ag, prospects)
        media_users: Set[Tuple[str, str]] = set(uniq_users(photos) + uniq_users(cards))
        for ag, ini in sorted(media_users):
            send_individual_pack(run_date, ag, ini, prospects, photos, cards)
        return

    if SEND_MODE == "admin":
        send_admin_pack(run_date, prospects)
        media_users: Set[Tuple[str, str]] = set(uniq_users(photos) + uniq_users(cards))
        for ag, ini in sorted(media_users):
            send_individual_pack(run_date, ag, ini, prospects, photos, cards)
        return

    raise RuntimeError(f"Unknown SEND_MODE={SEND_MODE}")

if __name__ == "__main__":
    main()