import os
import re
import json
import io
import base64
import zipfile
import datetime as dt
from typing import Dict, List, Any, Optional, Tuple, Set

import requests
from PIL import Image
import pytesseract

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment


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
GEMINI_API_KEY        = env_str("GEMINI_API_KEY", "")

SEND_MODE    = env_str("SEND_MODE", "individual").lower()  # individual | agency_manager | admin
RUN_DATE     = env_str("RUN_DATE", today_ymd_utc())
AGENCY       = env_str("AGENCY", "").upper()
INITIALS     = env_str("INITIALS", "").upper()

MAX_OCR_IMAGES   = env_int("MAX_OCR_IMAGES", 50)
MAX_PHOTO_IMAGES = env_int("MAX_PHOTO_IMAGES", 15)

OUT_DIR = env_str("OUT_DIR", "out").strip() or "out"
if os.path.exists(OUT_DIR) and not os.path.isdir(OUT_DIR):
    print(f"[WARN] OUT_DIR='{OUT_DIR}' existe mais n'est pas un dossier. Fallback -> 'exports'")
    OUT_DIR = "exports"
os.makedirs(OUT_DIR, exist_ok=True)

VALID_AGENCIES = {"GR", "VR", "GRS", "SLS"}

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
PHONE_RE = re.compile(r"(\+33|0)\s*[1-9](?:[\s\.-]*\d{2}){4}")
EMAIL_IN_TEXT_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"(https?://[^\s)]+|www\.[^\s)]+)", re.I)
CP_RE = re.compile(r"\b(0[1-9]\d{3}|[1-8]\d{4}|9[0-5]\d{3}|97\d{3}|98\d{3})\b")

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
        except Exception as e:
            print(f"[WARN] MAIL_ROUTING_JSON invalid JSON, fallback default. err={e}")
    return DEFAULT_ROUTING

ROUTING = load_routing()

def clean_email(v) -> str:
    if v is None:
        return ""
    if isinstance(v, dict):
        v = v.get("email") or ""
    if not isinstance(v, str):
        v = str(v)
    v = v.strip()
    v = re.sub(r"\s+", "", v)
    return v

def is_valid_email(s: str) -> bool:
    s = clean_email(s)
    return bool(EMAIL_RE.match(s))

def email_for_initials(initials: str) -> Optional[str]:
    initials = (initials or "").upper().strip()
    if not initials:
        return None

    users = ROUTING.get("users") or {}
    if initials in users:
        em = clean_email(users.get(initials))
        if is_valid_email(em):
            return em

    agencies = ROUTING.get("agencies") or {}
    for _, cfg in agencies.items():
        for role in ("manager", "commercial"):
            r = (cfg or {}).get(role) or {}
            if (r.get("initials") or "").upper() == initials:
                em2 = clean_email(r.get("email") or "")
                return em2 if is_valid_email(em2) else None
    return None


# ============================================================
# HTTP HELPERS
# ============================================================

def worker_dump(kind: str, date: str) -> List[Dict[str, Any]]:
    if not WORKER_BASE_URL or not EXPORT_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN")
    url = f"{WORKER_BASE_URL.rstrip('/')}/dump?date={date}&kind={kind}"
    r = requests.get(url, headers={"X-Export-Token": EXPORT_TOKEN}, timeout=45)
    if r.status_code != 200:
        raise RuntimeError(f"/dump failed {kind} {r.status_code}: {r.text[:400]}")
    lines = [ln.strip() for ln in r.text.splitlines() if ln.strip()]
    out: List[Dict[str, Any]] = []
    for ln in lines:
        try:
            out.append(json.loads(ln))
        except Exception:
            pass
    return out

def tg_get_file_url(file_id: str) -> Optional[str]:
    if not file_id:
        return None
    if not TELEGRAM_TOKEN:
        raise RuntimeError("Missing TELEGRAM_TOKEN")
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile"
    r = requests.post(url, json={"file_id": file_id}, timeout=25)
    j = r.json()
    if not j.get("ok"):
        return None
    fp = j["result"]["file_path"]
    return f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{fp}"

def download_bytes(url: str) -> bytes:
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content


# ============================================================
# TEXT HELPERS
# ============================================================

def normalize_phone(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    return re.sub(r"[^\d+]", "", s)

def best_phone(text: str) -> str:
    if not text:
        return ""
    m = PHONE_RE.search(text)
    return normalize_phone(m.group(0)) if m else ""

def best_email(text: str) -> str:
    if not text:
        return ""
    m = EMAIL_IN_TEXT_RE.search(text)
    return (m.group(0).strip().lower()) if m else ""

def extract_urls(text: str) -> List[str]:
    if not text:
        return []
    out, seen = [], set()
    for m in URL_RE.finditer(text):
        u = m.group(0).strip()
        if u.lower().startswith("www."):
            u = "https://" + u
        k = u.lower()
        if k not in seen:
            seen.add(k)
            out.append(u)
    return out

def extract_postal_code(text: str) -> str:
    if not text:
        return ""
    cps = CP_RE.findall(text)
    return cps[0] if cps else ""

def domain_from_email(email: str) -> str:
    email = (email or "").strip().lower()
    if "@" not in email:
        return ""
    dom = email.split("@", 1)[1].strip()
    dom = re.sub(r"[^a-z0-9\.\-]", "", dom)
    if dom in {"gmail.com","outlook.com","hotmail.com","yahoo.com","yahoo.fr","icloud.com","free.fr","orange.fr","laposte.net"}:
        return ""
    return dom

def domain_from_url(url: str) -> str:
    if not url:
        return ""
    u = url.lower().strip()
    u = re.sub(r"^https?://", "", u)
    u = u.split("/", 1)[0]
    u = u.split(":", 1)[0]
    u = re.sub(r"[^a-z0-9\.\-]", "", u)
    if u.startswith("www."):
        u = u[4:]
    return u

def brand_from_domain(dom: str) -> str:
    dom = (dom or "").strip().lower()
    if not dom:
        return ""
    dom = dom.split(".", 1)[0]
    dom = dom.replace("-", " ").replace("_", " ")
    dom = re.sub(r"\s+", " ", dom).strip()
    return dom

def deduce_company(company_raw: str, email: str, website: str) -> str:
    company_raw = (company_raw or "").strip()
    dom_email = domain_from_email(email)
    dom_web = domain_from_url(website)
    if dom_email:
        return brand_from_domain(dom_email)
    if dom_web:
        return brand_from_domain(dom_web)
    return company_raw

def split_human_name(full: str) -> Tuple[str, str]:
    s = (full or "").strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return "", ""
    parts = s.split(" ")
    if len(parts) == 1:
        return "", parts[0]
    return parts[0], " ".join(parts[1:])

def ensure_contact_fallback(d: Dict[str, Any]) -> Dict[str, Any]:
    fn = (d.get("contact_firstname") or "").strip()
    ln = (d.get("contact_lastname") or "").strip()
    if (not fn and not ln) and (d.get("dirigeant") or "").strip():
        f, l = split_human_name(d.get("dirigeant") or "")
        d["contact_firstname"] = d.get("contact_firstname") or f
        d["contact_lastname"]  = d.get("contact_lastname") or (l or d.get("dirigeant") or "")
    if not (d.get("interlocuteur") or "").strip():
        combo = f"{(d.get('contact_firstname') or '').strip()} {(d.get('contact_lastname') or '').strip()}".strip()
        if combo:
            d["interlocuteur"] = combo
    return d


# ============================================================
# OCR
# ============================================================

def ocr_image_bytes(img_bytes: bytes) -> str:
    if not img_bytes:
        return ""
    try:
        im = Image.open(io.BytesIO(img_bytes))
        if im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        txt = pytesseract.image_to_string(im, lang="fra+eng")
        return (txt or "").strip()
    except Exception:
        return ""


# ============================================================
# GEMINI (image-first)
# ============================================================

def _guess_mime(img_bytes: bytes) -> str:
    try:
        im = Image.open(io.BytesIO(img_bytes))
        fmt = (im.format or "").upper()
        if fmt == "PNG":
            return "image/png"
        if fmt in ("JPG", "JPEG"):
            return "image/jpeg"
    except Exception:
        pass
    return "image/jpeg"

def gemini_vision_json(img_bytes: bytes, prompt: str, max_tokens: int = 650) -> Dict[str, Any]:
    if not GEMINI_API_KEY or not img_bytes:
        return {}
    endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
    mime = _guess_mime(img_bytes)
    b64 = base64.b64encode(img_bytes).decode("ascii")

    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inlineData": {"mimeType": mime, "data": b64}}
            ]
        }],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": max_tokens}
    }

    try:
        r = requests.post(endpoint, params={"key": GEMINI_API_KEY}, json=payload, timeout=60)
        if r.status_code >= 300:
            return {}
        j = r.json()
        parts = (((j.get("candidates") or [None])[0] or {}).get("content") or {}).get("parts") or []
        raw = ""
        for p in parts:
            if isinstance(p, dict) and p.get("text"):
                raw += p["text"]
        raw = (raw or "").strip()
        raw = re.sub(r"^```json", "", raw, flags=re.I).strip()
        raw = re.sub(r"```$", "", raw).strip()
        m = re.search(r"\{.*\}", raw, re.S)
        if not m:
            return {}
        return json.loads(m.group(0))
    except Exception:
        return {}

def gemini_extract_business_card(img_bytes: bytes) -> Dict[str, str]:
    prompt = (
        "Tu es un extracteur de carte de visite.\n"
        "Lis l'IMAGE (ne te base PAS sur un OCR).\n"
        "Retourne STRICTEMENT un JSON avec les clés EXACTES:\n"
        "name, company, title, email, phone, website, postal_code, city, address\n"
        "Règles:\n"
        "- Si inconnu: \"\"\n"
        "- email en minuscules\n"
        "- postal_code: 5 chiffres\n"
    )
    d = gemini_vision_json(img_bytes, prompt, max_tokens=700) or {}
    return {
        "name": str(d.get("name") or "").strip(),
        "company": str(d.get("company") or "").strip(),
        "title": str(d.get("title") or "").strip(),
        "email": str(d.get("email") or "").strip().lower(),
        "phone": normalize_phone(str(d.get("phone") or "")),
        "website": str(d.get("website") or "").strip(),
        "postal_code": str(d.get("postal_code") or "").strip(),
        "city": str(d.get("city") or "").strip(),
        "address": str(d.get("address") or "").strip(),
    }

def gemini_extract_facade_logo(img_bytes: bytes) -> Dict[str, str]:
    prompt = (
        "Tu analyses une photo de prospection (façade, enseigne, logo).\n"
        "Lis l'IMAGE.\n"
        "Retourne STRICTEMENT un JSON avec les clés:\n"
        "company, city\n"
        "Règles: si inconnu -> \"\".\n"
    )
    d = gemini_vision_json(img_bytes, prompt, max_tokens=350) or {}
    return {"company": str(d.get("company") or "").strip(), "city": str(d.get("city") or "").strip()}

print(f"🔑 Gemini={'ON' if GEMINI_API_KEY else 'OFF'} | Places={'ON' if GOOGLE_PLACES_API_KEY else 'OFF'}")
# ============================================================
# ENRICH: GOUV + PLACES + SCRAPE
# ============================================================

def search_gouv_company(name: str, city: str = "") -> Dict[str, str]:
    name = (name or "").strip()
    if not name:
        return {}
    try:
        q = name if not city else f"{name} {city}".strip()
        url = f"https://recherche-entreprises.api.gouv.fr/search?q={requests.utils.quote(q)}&page=1&per_page=1"
        r = requests.get(url, headers={"accept": "application/json"}, timeout=20)
        j = r.json()
        res = (j.get("results") or [])
        if not res:
            return {}
        e = res[0]
        siege = e.get("siege") or {}

        dirigeant = ""
        arr = e.get("dirigeants") or e.get("representants") or []
        if arr:
            first = arr[0]
            if isinstance(first, str):
                dirigeant = first
            elif isinstance(first, dict):
                p = first.get("personne") or first
                prenom = p.get("prenom") or p.get("prenoms") or ""
                nom = p.get("nom") or p.get("nom_usage") or ""
                dirigeant = f"{prenom} {nom}".strip()

        naf = (e.get("activite_principale") or e.get("naf") or "").replace(".", "").replace(" ", "").upper()
        return {
            "name": e.get("nom_raison_sociale") or e.get("denomination") or name,
            "siret": (siege.get("siret") or ""),
            "naf": naf,
            "address": (siege.get("adresse") or siege.get("libelle_voie") or ""),
            "postal_code": (siege.get("code_postal") or ""),
            "city": (siege.get("libelle_commune") or city),
            "dirigeant": dirigeant,
        }
    except Exception:
        return {}

def places_enrich(name: str, city: str = "") -> Dict[str, str]:
    if not GOOGLE_PLACES_API_KEY or not name:
        return {}
    try:
        query = name if not city else f"{name} {city}".strip()
        r1 = requests.post(
            "https://places.googleapis.com/v1/places:searchText",
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "places.id,places.websiteUri",
            },
            json={
                "textQuery": query,
                "maxResultCount": 1,
                "languageCode": "fr",
                "regionCode": "FR",
            },
            timeout=20
        )
        j1 = r1.json()
        p = (j1.get("places") or [None])[0]
        if not p or not p.get("id"):
            return {}
        pid = p["id"]
        website = p.get("websiteUri") or ""

        r2 = requests.get(
            f"https://places.googleapis.com/v1/places/{pid}",
            headers={
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "nationalPhoneNumber,internationalPhoneNumber,websiteUri,formattedAddress",
            },
            timeout=20
        )
        j2 = r2.json()
        phone = j2.get("nationalPhoneNumber") or j2.get("internationalPhoneNumber") or ""
        website2 = j2.get("websiteUri") or website
        addr = (j2.get("formattedAddress") or "").strip()

        return {"phone": normalize_phone(phone), "website": (website2 or "").strip(), "address": addr}
    except Exception:
        return {}

def scrape_email_from_site(url: str) -> str:
    if not url:
        return ""
    try:
        r = requests.get(url, headers={"user-agent": "Mozilla/5.0 (ProspectionBot)"}, timeout=10)
        if r.status_code >= 300:
            return ""
        html = r.text[:250000]
        matches = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}", html)
        if not matches:
            return ""
        filtered = [m for m in matches if not re.search(r"no-?reply", m, re.I)]
        return (filtered[0] if filtered else matches[0]).strip().lower()
    except Exception:
        return ""


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
# UNIFY
# ============================================================

def unify_from_prospect(p: Dict[str, Any]) -> Dict[str, Any]:
    d = {
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
    }
    return ensure_contact_fallback(d)

def unify_from_card(date: str, agency: str, initials: str, comment: str, img_bytes: bytes) -> Dict[str, Any]:
    base = {
        "date": date, "agency": agency, "initials": initials,
        "name": "", "address": "", "postal_code": "", "city": "",
        "siret": "", "naf": "", "dirigeant": "",
        "interlocuteur": "", "contact_firstname": "", "contact_lastname": "",
        "phone": "", "phone2": "", "email": "", "website": "",
        "resume": (comment or "").strip(),
        "commande": "",
    }
    if not img_bytes:
        return base

    # 1) Gemini image-first
    vis = gemini_extract_business_card(img_bytes) if GEMINI_API_KEY else {
        "name":"","company":"","title":"","email":"","phone":"","website":"","postal_code":"","city":"","address":""
    }

    # 2) OCR complément
    ocr_txt = ocr_image_bytes(img_bytes)
    email_ocr = best_email(ocr_txt)
    phone_ocr = best_phone(ocr_txt)
    urls_ocr = extract_urls(ocr_txt)
    cp_ocr = extract_postal_code(ocr_txt)

    email = (vis.get("email") or email_ocr or "").strip().lower()
    phone = normalize_phone(vis.get("phone") or phone_ocr or "")
    website = (vis.get("website") or (urls_ocr[0] if urls_ocr else "") or "").strip()

    postal_code = (vis.get("postal_code") or cp_ocr or "").strip()
    city = (vis.get("city") or "").strip()
    address = (vis.get("address") or "").strip()

    company_guess = deduce_company(vis.get("company") or "", email, website)

    # 3) API gouv: sans ville -> avec ville si besoin
    gouv = search_gouv_company(company_guess, "") if company_guess else {}
    if not gouv and company_guess and city:
        gouv = search_gouv_company(company_guess, city)

    # 4) Places
    place = {}
    if company_guess:
        place = places_enrich(gouv.get("name") or company_guess, city) if city else places_enrich(gouv.get("name") or company_guess, "")
        place = place or {}

    final_website = (website or place.get("website") or "").strip()
    final_email = email or (scrape_email_from_site(final_website) if final_website else "")
    final_phone = phone or place.get("phone") or ""

    full_name = (vis.get("name") or "").strip()
    fn, ln = split_human_name(full_name)

    row = dict(base)
    row.update({
        "name": (gouv.get("name") or company_guess or "").strip(),
        "address": (gouv.get("address") or address or place.get("address") or "").strip(),
        "postal_code": (gouv.get("postal_code") or postal_code or "").strip(),
        "city": (gouv.get("city") or city or "").strip(),
        "siret": (gouv.get("siret") or "").strip(),
        "naf": (gouv.get("naf") or "").strip(),
        "dirigeant": (gouv.get("dirigeant") or "").strip(),

        "interlocuteur": full_name or "",
        "contact_firstname": fn,
        "contact_lastname": ln,

        "phone": final_phone,
        "email": (final_email or "").strip().lower(),
        "website": final_website,
    })
    return ensure_contact_fallback(row)

def unify_from_photo(date: str, agency: str, initials: str, city_hint: str, comment: str, img_bytes: bytes) -> Dict[str, Any]:
    base = {
        "date": date, "agency": agency, "initials": initials,
        "name": "", "address": "", "postal_code": "", "city": (city_hint or "").strip(),
        "siret": "", "naf": "", "dirigeant": "",
        "interlocuteur": "", "contact_firstname": "", "contact_lastname": "",
        "phone": "", "phone2": "", "email": "", "website": "",
        "resume": (comment or "").strip(),
        "commande": "",
    }
    if not img_bytes:
        return base

    vis = gemini_extract_facade_logo(img_bytes) if GEMINI_API_KEY else {"company":"", "city":""}
    company_raw = (vis.get("company") or "").strip()
    city = (vis.get("city") or "").strip() or (city_hint or "").strip()

    ocr_txt = ocr_image_bytes(img_bytes)
    email_ocr = best_email(ocr_txt)
    urls_ocr = extract_urls(ocr_txt)
    website = (urls_ocr[0] if urls_ocr else "").strip()

    company_guess = deduce_company(company_raw, email_ocr, website)
    if not company_guess:
        return base

    gouv = search_gouv_company(company_guess, "")
    if not gouv and city:
        gouv = search_gouv_company(company_guess, city)

    place = places_enrich(gouv.get("name") or company_guess, city) if GOOGLE_PLACES_API_KEY else {}
    final_website = (website or place.get("website") or "").strip()
    final_email = (email_ocr or "").strip().lower() or (scrape_email_from_site(final_website) if final_website else "")

    row = dict(base)
    row.update({
        "name": (gouv.get("name") or company_guess).strip(),
        "address": (gouv.get("address") or place.get("address") or "").strip(),
        "postal_code": (gouv.get("postal_code") or "").strip(),
        "city": (gouv.get("city") or city).strip(),
        "siret": (gouv.get("siret") or "").strip(),
        "naf": (gouv.get("naf") or "").strip(),
        "dirigeant": (gouv.get("dirigeant") or "").strip(),
        "phone": (place.get("phone") or "").strip(),
        "website": final_website,
        "email": final_email,
    })
    return ensure_contact_fallback(row)


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
        timeout=60
    )
    if r.status_code >= 300:
        raise RuntimeError(f"Brevo send failed {r.status_code}: {r.text[:400]}")


# ============================================================
# SEND MODES
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
        print(f"[WARN] No email for initials={initials}, skip individual pack.")
        return

    initials = initials.upper().strip()

    p_u = [p for p in prospects if (p.get("agency") == agency and (p.get("initials") or "").upper() == initials)]
    ph_u = [ph for ph in photos if (ph.get("agency") == agency and (ph.get("user") or "").upper() == initials)]
    ca_u = [ca for ca in cards  if (ca.get("agency") == agency and (ca.get("user") or "").upper() == initials)]

    rows: List[Dict[str, Any]] = []
    rows.extend([unify_from_prospect(p) for p in p_u])

    for ph in ph_u[:MAX_PHOTO_IMAGES]:
        fid = ph.get("file_id") or ""
        url = tg_get_file_url(fid)
        city_hint = ph.get("city") or ""
        comment = ph.get("comment") or ph.get("meeting") or ""
        try:
            img = download_bytes(url) if url else b""
        except Exception:
            img = b""
        rows.append(unify_from_photo(date, agency, initials, city_hint, comment, img))

    for ca in ca_u[:MAX_OCR_IMAGES]:
        fid = ca.get("file_id") or ""
        url = tg_get_file_url(fid)
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
        f"<p>Voici ton export de prospection du <b>{date}</b> pour <b>{agency}/{initials}</b>.</p>"
        f"<ul><li><b>Excel</b> (1 onglet, trié par complétude)</li>"
        f"<li><b>ZIP médias</b> (photos + cartes)</li></ul>"
        f"<p>— Bot Prospection</p>"
    )
    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] individual pack sent to {to_email} ({agency}/{initials})")

def send_agency_manager_pack(date: str, agency: str,
                             prospects: List[Dict[str, Any]],
                             photos: List[Dict[str, Any]],
                             cards: List[Dict[str, Any]]):

    agencies_cfg = ROUTING.get("agencies") or {}
    cfg = agencies_cfg.get(agency) or {}
    manager = (cfg.get("manager") or {})
    to_email = clean_email(manager.get("email") or "")
    if not is_valid_email(to_email):
        print(f"[WARN] No valid manager email for agency={agency} ({to_email})")
        return

    # Consolidé agence = prospects + cartes + photos (tout le monde)
    rows: List[Dict[str, Any]] = []
    p_ag = [p for p in prospects if (p.get("agency") == agency)]
    rows.extend([unify_from_prospect(p) for p in p_ag])

    ph_ag = [ph for ph in photos if (ph.get("agency") == agency)]
    ca_ag = [ca for ca in cards  if (ca.get("agency") == agency)]

    for ph in ph_ag[:200]:
        fid = ph.get("file_id") or ""
        url = tg_get_file_url(fid)
        city_hint = ph.get("city") or ""
        comment = ph.get("comment") or ph.get("meeting") or ""
        try:
            img = download_bytes(url) if url else b""
        except Exception:
            img = b""
        rows.append(unify_from_photo(date, agency, (ph.get("user") or "").upper(), city_hint, comment, img))

    for ca in ca_ag[:300]:
        fid = ca.get("file_id") or ""
        url = tg_get_file_url(fid)
        comment = ca.get("comment") or ""
        try:
            img = download_bytes(url) if url else b""
        except Exception:
            img = b""
        rows.append(unify_from_card(date, agency, (ca.get("user") or "").upper(), comment, img))

    if not rows:
        print(f"[INFO] No rows for agency={agency}, skip manager mail.")
        return

    xlsx = build_excel_one_sheet(date, rows, f"AGENCE_{agency}")

    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    brevo_send_email(
        to_email,
        f"Prospection {date} — Agence {agency} (consolidé)",
        f"<p>Bonjour,</p><p>Consolidé {date} agence <b>{agency}</b> (1 onglet).</p><p>— Bot Prospection</p>",
        attachments=attachments
    )
    print(f"[OK] agency manager pack sent to {to_email} (agency={agency})")

def send_admin_pack(date: str,
                    prospects: List[Dict[str, Any]],
                    photos: List[Dict[str, Any]],
                    cards: List[Dict[str, Any]]):

    admin = ROUTING.get("admin") or {}
    to_email = clean_email(admin.get("email") or "")
    if not is_valid_email(to_email):
        print(f"[WARN] No valid admin email configured ({to_email})")
        return

    rows: List[Dict[str, Any]] = []
    rows.extend([unify_from_prospect(p) for p in prospects])

    for ph in photos[:400]:
        fid = ph.get("file_id") or ""
        url = tg_get_file_url(fid)
        city_hint = ph.get("city") or ""
        comment = ph.get("comment") or ph.get("meeting") or ""
        agency = (ph.get("agency") or "").upper()
        initials = (ph.get("user") or "").upper()
        try:
            img = download_bytes(url) if url else b""
        except Exception:
            img = b""
        rows.append(unify_from_photo(date, agency, initials, city_hint, comment, img))

    for ca in cards[:600]:
        fid = ca.get("file_id") or ""
        url = tg_get_file_url(fid)
        comment = ca.get("comment") or ""
        agency = (ca.get("agency") or "").upper()
        initials = (ca.get("user") or "").upper()
        try:
            img = download_bytes(url) if url else b""
        except Exception:
            img = b""
        rows.append(unify_from_card(date, agency, initials, comment, img))

    if not rows:
        print(f"[INFO] No rows for admin date={date}, skip.")
        return

    xlsx = build_excel_one_sheet(date, rows, "ADMIN_ALL")

    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    brevo_send_email(
        to_email,
        f"Prospection {date} — ADMIN (toutes agences)",
        f"<p>Bonjour,</p><p>Consolidé global {date} (1 onglet).</p><p>— Bot Prospection</p>",
        attachments=attachments
    )
    print(f"[OK] admin pack sent to {to_email}")


# ============================================================
# MAIN
# ============================================================

def main():
    if not WORKER_BASE_URL or not EXPORT_TOKEN or not TELEGRAM_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN / TELEGRAM_TOKEN")

    run_date = RUN_DATE if RUN_DATE else today_ymd_utc()

    print(f"🚀 export_and_mail.py — mode={SEND_MODE} date={run_date} agency={AGENCY} initials={INITIALS}")
    print(f"📦 OUT_DIR={OUT_DIR}")
    print(f"🔑 Gemini={'ON' if GEMINI_API_KEY else 'OFF'} | Places={'ON' if GOOGLE_PLACES_API_KEY else 'OFF'}")

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
            send_agency_manager_pack(run_date, ag, prospects, photos, cards)
        return

    if SEND_MODE == "admin":
        send_admin_pack(run_date, prospects, photos, cards)
        return

    raise RuntimeError(f"Unknown SEND_MODE={SEND_MODE}")

if __name__ == "__main__":
    main()