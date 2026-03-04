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

MAIL_ROUTING_JSON  = env_str("MAIL_ROUTING_JSON", "")  # optional

GOOGLE_PLACES_API_KEY = env_str("GOOGLE_PLACES_API_KEY", "")
GEMINI_API_KEY        = env_str("GEMINI_API_KEY", "")

SEND_MODE    = env_str("SEND_MODE", "agency_manager").lower()  # individual | agency_manager | admin
RUN_DATE     = env_str("RUN_DATE", today_ymd_utc())
AGENCY       = env_str("AGENCY", "").upper()
INITIALS     = env_str("INITIALS", "").upper()

MAX_OCR_IMAGES   = env_int("MAX_OCR_IMAGES", 50)
MAX_PHOTO_IMAGES = env_int("MAX_PHOTO_IMAGES", 50)

OUT_DIR = env_str("OUT_DIR", ".")
VALID_AGENCIES = {"GR", "VR", "GRS", "SLS"}

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EMAIL_IN_TEXT_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(\+33|0)\s*[1-9](?:[\s\.-]*\d{2}){4}")


# ============================================================
# ROUTING (fallback + override JSON)
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
        "JL": {"email":"jennifer.laurens@ras-interim.fr","agency":"GR","role":"manager"},
        "CZ": {"email":"celine.zunarelli@ras-interim.fr","agency":"GR","role":"commercial"},
        "JB": {"email":"jelena.carrasso@ras-interim.fr","agency":"VR","role":"manager"},
        "LB": {"email":"laura.berthet@ras-interim.fr","agency":"VR","role":"commercial"},
        "PV": {"email":"pauline.vieira@ras-interim.fr","agency":"GRS","role":"manager"},
        "ST": {"email":"severine.thevenin@ras-interim.fr","agency":"GRS","role":"commercial"},
        "AC": {"email":"aurelie.curt@ras-interim.fr","agency":"SLS","role":"manager"},
        "SL": {"email":"samuel.lengue@ras-interim.fr","agency":"ADMIN","role":"admin"},
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

def clean_email(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", "", s)
    return s

def is_valid_email(s: str) -> bool:
    s = clean_email(s)
    return bool(EMAIL_RE.match(s))

def routing_user_email(initials: str) -> Optional[str]:
    initials = (initials or "").upper().strip()
    users = ROUTING.get("users") or {}
    if initials in users:
        v = users[initials]
        if isinstance(v, str):
            em = clean_email(v)
            return em if is_valid_email(em) else None
        if isinstance(v, dict):
            em = clean_email(v.get("email") or "")
            return em if is_valid_email(em) else None
    return None

def email_for_initials(initials: str) -> Optional[str]:
    initials = (initials or "").upper().strip()
    if not initials:
        return None
    em = routing_user_email(initials)
    if em:
        return em

    agencies = ROUTING.get("agencies") or {}
    for _, cfg in agencies.items():
        for role in ("manager", "commercial"):
            r = (cfg or {}).get(role) or {}
            if (r.get("initials") or "").upper() == initials:
                em2 = clean_email(r.get("email") or "")
                return em2 if is_valid_email(em2) else None
    return None

def manager_email_for_agency(agency: str) -> Optional[str]:
    agencies_cfg = ROUTING.get("agencies") or {}
    cfg = agencies_cfg.get(agency) or {}
    manager = (cfg.get("manager") or {})
    to_email = clean_email(manager.get("email") or "")
    return to_email if is_valid_email(to_email) else None

def admin_email() -> Optional[str]:
    adm = ROUTING.get("admin") or {}
    to_email = clean_email(adm.get("email") or "")
    return to_email if is_valid_email(to_email) else None


# ============================================================
# HTTP helpers
# ============================================================

def worker_dump(kind: str, date: str) -> List[Dict[str, Any]]:
    if not WORKER_BASE_URL or not EXPORT_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN")
    url = f"{WORKER_BASE_URL.rstrip('/')}/dump?date={date}&kind={kind}"
    r = requests.get(url, headers={"X-Export-Token": EXPORT_TOKEN}, timeout=25)
    if r.status_code != 200:
        raise RuntimeError(f"/dump failed {kind} {r.status_code}: {r.text[:200]}")
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
# Parsing / normalisation
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

def split_human_name(full: str) -> Tuple[str, str]:
    s = (full or "").strip()
    if not s:
        return "", ""
    parts = re.split(r"\s+", s)
    if len(parts) == 1:
        return "", parts[0]
    return parts[0], " ".join(parts[1:])

def normalize_naf(naf: str) -> str:
    return (naf or "").replace(".", "").replace(" ", "").upper().strip()


# ============================================================
# Enrichissement API entreprises + Places + scrape email
# ============================================================

def api_gouv_search(name: str, city: str) -> Dict[str, str]:
    name = (name or "").strip()
    city = (city or "").strip()
    if not name:
        return {}

    try:
        url = "https://recherche-entreprises.api.gouv.fr/search"
        q = f"{name} {city}".strip()
        r = requests.get(url, params={"q": q, "page": 1, "per_page": 1}, timeout=20)
        j = r.json()
        res = (j.get("results") or [])
        if not res:
            return {}
        e = res[0]
        siege = e.get("siege") or {}
        out = {
            "name": e.get("nom_raison_sociale") or e.get("denomination") or "",
            "siret": siege.get("siret") or "",
            "naf": e.get("activite_principale") or e.get("naf") or "",
            "address": siege.get("adresse") or siege.get("libelle_voie") or "",
            "postal_code": siege.get("code_postal") or "",
            "city": siege.get("libelle_commune") or "",
        }
        # dirigeant (best effort)
        dirigeants = e.get("dirigeants") or e.get("representants") or []
        dirigeant = ""
        if dirigeants:
            d0 = dirigeants[0]
            if isinstance(d0, str):
                dirigeant = d0
            elif isinstance(d0, dict):
                p = d0.get("personne") or d0
                prenom = p.get("prenom") or p.get("prenoms") or ""
                nom = p.get("nom") or p.get("nom_usage") or ""
                dirigeant = f"{prenom} {nom}".strip()
        out["dirigeant"] = dirigeant
        out["naf"] = normalize_naf(out["naf"])
        return out
    except Exception:
        return {}

def places_search_text(query: str) -> Optional[str]:
    if not GOOGLE_PLACES_API_KEY:
        return None
    try:
        r = requests.post(
            "https://places.googleapis.com/v1/places:searchText",
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "places.id",
            },
            json={
                "textQuery": query.strip(),
                "maxResultCount": 1,
                "languageCode": "fr",
                "regionCode": "FR",
            },
            timeout=20
        )
        j = r.json()
        p = (j.get("places") or [None])[0]
        if not p or not p.get("id"):
            return None
        return p["id"]
    except Exception:
        return None

def places_get(place_id: str) -> Dict[str, str]:
    if not GOOGLE_PLACES_API_KEY or not place_id:
        return {}
    try:
        r = requests.get(
            f"https://places.googleapis.com/v1/places/{place_id}",
            headers={
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "nationalPhoneNumber,internationalPhoneNumber,websiteUri,displayName",
            },
            timeout=20
        )
        j = r.json()
        return {
            "phone": normalize_phone(j.get("nationalPhoneNumber") or j.get("internationalPhoneNumber") or ""),
            "website": (j.get("websiteUri") or "").strip(),
            "displayName": ((j.get("displayName") or {}).get("text") or "").strip(),
        }
    except Exception:
        return {}

def scrape_email_from_site(url: str) -> str:
    if not url:
        return ""
    try:
        r = requests.get(url, timeout=6, headers={"User-Agent": "Mozilla/5.0 (ProspectionBot/1.0)"})
        if r.status_code >= 400:
            return ""
        matches = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", r.text)
        if not matches:
            return ""
        filtered = [m for m in matches if not re.search(r"no-?reply", m, re.I)]
        return (filtered[0] if filtered else matches[0]).strip().lower()
    except Exception:
        return ""


# ============================================================
# OCR + GEMINI VISION (Gemini first, OCR fallback)
# ============================================================

def ocr_image_bytes(img_bytes: bytes) -> str:
    try:
        im = Image.open(io.BytesIO(img_bytes))
        if im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        txt = pytesseract.image_to_string(im, lang="fra+eng")
        return (txt or "").strip()
    except Exception:
        return ""

GEMINI_VISION_PROMPT = """
Tu analyses une image qui peut être:
- une carte de visite (contact)
- une façade/enseigne/logo d'entreprise (entreprise)

Retourne STRICTEMENT un JSON (pas de texte autour) avec ces clés exactes:
company, person, title, phone, email, website, city

Règles:
- Si inconnu: chaîne vide.
- phone: format brut tel que vu, sans inventer.
- city: si visible/indicatif, sinon vide.
"""

def gemini_vision_extract(image_bytes: bytes) -> Dict[str, str]:
    out = {"company":"", "person":"", "title":"", "phone":"", "email":"", "website":"", "city":""}
    if not GEMINI_API_KEY:
        return out
    try:
        endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
        img64 = base64.b64encode(image_bytes).decode("ascii")
        r = requests.post(
            endpoint,
            params={"key": GEMINI_API_KEY},
            json={
                "contents":[{"parts":[
                    {"text": GEMINI_VISION_PROMPT},
                    {"inline_data":{"mime_type":"image/jpeg","data": img64}}
                ]}],
                "generationConfig":{"temperature":0.1,"maxOutputTokens":300},
            },
            timeout=25
        )
        j = r.json()
        parts = (((j.get("candidates") or [None])[0] or {}).get("content") or {}).get("parts") or []
        raw = ""
        for p in parts:
            if isinstance(p, dict) and p.get("text"):
                raw += p["text"]
        raw = raw.strip()
        m = re.search(r"\{.*\}", raw, re.S)
        if not m:
            return out
        data = json.loads(m.group(0))
        for k in out.keys():
            v = (data.get(k) or "").strip()
            out[k] = v
        out["email"] = out["email"].lower().strip()
        out["phone"] = normalize_phone(out["phone"])
        return out
    except Exception:
        return out


# ============================================================
# Single-sheet Excel (exact columns)
# ============================================================

HEADERS = [
    "date","agency","initials","name","address","postal_code","city","siret","naf","dirigeant",
    "interlocuteur","contact_firstname","contact_lastname","phone","phone2","email","website","resume","commande"
]

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

def build_single_sheet_xlsx(path: str, rows: List[Dict[str, Any]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"
    ws.append(HEADERS)

    for r in rows:
        ws.append([r.get(h, "") for h in HEADERS])

    # header style
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    autosize(ws)
    wb.save(path)

def completeness_score(r: Dict[str, Any]) -> int:
    # score sur les champs métier, pas sur date/agency/initials
    fields = [h for h in HEADERS if h not in ("date","agency","initials")]
    return sum(1 for h in fields if (r.get(h) or "").strip() != "")

def normalize_row_defaults(r: Dict[str, Any]) -> Dict[str, Any]:
    for h in HEADERS:
        if h not in r:
            r[h] = ""
        if r[h] is None:
            r[h] = ""
    # normalisations
    r["email"] = (r["email"] or "").strip().lower()
    r["naf"] = normalize_naf(r["naf"] or "")
    r["phone"] = normalize_phone(r["phone"] or "")
    r["phone2"] = normalize_phone(r["phone2"] or "")
    return r


# ============================================================
# Media ZIP (cards + photos)
# ============================================================

def build_media_zip(date: str, scope_label: str,
                    photos: List[Dict[str, Any]], cards: List[Dict[str, Any]],
                    filter_agency: Optional[str]=None, filter_initials: Optional[str]=None) -> Optional[Tuple[str, bytes]]:
    def match(rec):
        if filter_agency and (rec.get("agency") or "").upper() != filter_agency:
            return False
        ini = (rec.get("user") or rec.get("initials") or "").upper()
        if filter_initials and ini != filter_initials:
            return False
        return True

    ph = [p for p in photos if match(p)]
    ca = [c for c in cards if match(c)]
    if not ph and not ca:
        return None

    ph = ph[:MAX_PHOTO_IMAGES]
    ca = ca[:MAX_OCR_IMAGES]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for i, p in enumerate(ph, start=1):
            fid = p.get("file_id") or ""
            url = tg_get_file_url(fid)
            if not url:
                continue
            img = download_bytes(url)
            fname = f"photos/{date}_{scope_label}_photo_{i:03d}.jpg"
            z.writestr(fname, img)

        for i, c in enumerate(ca, start=1):
            fid = c.get("file_id") or ""
            url = tg_get_file_url(fid)
            if not url:
                continue
            img = download_bytes(url)
            fname = f"cards/{date}_{scope_label}_card_{i:03d}.jpg"
            z.writestr(fname, img)

    zip_name = f"MEDIA_{date}_{scope_label}.zip"
    return zip_name, buf.getvalue()


# ============================================================
# Brevo email
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
        payload["attachment"] = [{
            "name": filename,
            "content": base64.b64encode(content).decode("ascii")
        } for filename, content in attachments]

    r = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={"accept": "application/json","api-key": BREVO_API_KEY,"content-type": "application/json"},
        data=json.dumps(payload),
        timeout=40
    )
    if r.status_code >= 300:
        raise RuntimeError(f"Brevo send failed {r.status_code}: {r.text[:300]}")


# ============================================================
# Transform sources -> unified rows
# ============================================================

def unify_from_prospect(p: Dict[str, Any]) -> Dict[str, Any]:
    r = {
        "date": p.get("date", RUN_DATE),
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
    return normalize_row_defaults(r)

def unify_from_image_record(rec: Dict[str, Any], kind: str) -> Dict[str, Any]:
    """
    kind = 'card' or 'photo'
    - Gemini Vision first
    - OCR fallback for missing fields
    - Enrich: API gouv + Places + scrape
    """
    agency = (rec.get("agency") or "").upper()
    initials = (rec.get("user") or rec.get("initials") or "").upper()
    comment = (rec.get("comment") or rec.get("meeting") or "") or ""

    file_id = rec.get("file_id") or ""
    url = tg_get_file_url(file_id) if file_id else None
    img = b""
    if url:
        try:
            img = download_bytes(url)
        except Exception:
            img = b""

    g = gemini_vision_extract(img) if img else {"company":"","person":"","title":"","phone":"","email":"","website":"","city":""}

    # OCR fallback (si manque)
    ocr_txt = ""
    if img and (not g.get("email") or not g.get("phone") or not g.get("company")):
        ocr_txt = ocr_image_bytes(img)

    o_email = best_email(ocr_txt) if ocr_txt else ""
    o_phone = best_phone(ocr_txt) if ocr_txt else ""

    company = (g.get("company") or "").strip()
    city_hint = (g.get("city") or "").strip()

    # Si photo lot : on a rec.city (= secteur) utile
    if kind == "photo":
        city_fallback = (rec.get("city") or "").strip()
        if not city_hint:
            city_hint = city_fallback

    # email/phone fallback
    email = (g.get("email") or "").strip().lower() or o_email
    phone = normalize_phone(g.get("phone") or "") or normalize_phone(o_phone)

    # website from Gemini
    website = (g.get("website") or "").strip()

    # Enrich entreprise via API gouv
    enr = api_gouv_search(company, city_hint)
    name_final = enr.get("name") or company
    city_final = enr.get("city") or city_hint
    addr = enr.get("address") or ""
    cp = enr.get("postal_code") or ""
    siret = enr.get("siret") or ""
    naf = enr.get("naf") or ""
    dirigeant = enr.get("dirigeant") or ""

    # Google Places (phone/website)
    place_id = places_search_text(f"{name_final} {city_final}".strip()) if name_final else None
    pl = places_get(place_id) if place_id else {}
    if not phone:
        phone = pl.get("phone") or ""
    if not website:
        website = pl.get("website") or ""

    # Scrape email si besoin
    if not email and website:
        email = scrape_email_from_site(website)

    # Interlocuteur depuis Gemini (person)
    interlocuteur = (g.get("person") or "").strip()
    cf, cl = split_human_name(interlocuteur) if interlocuteur else ("","")

    # row final
    r = {
        "date": rec.get("date", RUN_DATE),
        "agency": agency,
        "initials": initials,
        "name": name_final or "",
        "address": addr,
        "postal_code": cp,
        "city": city_final or "",
        "siret": siret,
        "naf": naf,
        "dirigeant": dirigeant,
        "interlocuteur": interlocuteur,
        "contact_firstname": cf,
        "contact_lastname": cl,
        "phone": phone,
        "phone2": "",
        "email": email,
        "website": website,
        "resume": comment,
        "commande": "",
    }
    return normalize_row_defaults(r)

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


# ============================================================
# BUILD + SEND
# ============================================================

def build_excel_single(date: str, rows: List[Dict[str, Any]], suffix: str) -> str:
    # sort by completeness desc
    rows_scored = [(completeness_score(r), r) for r in rows]
    rows_scored.sort(key=lambda x: x[0], reverse=True)
    rows_sorted = [r for _, r in rows_scored]

    filename = os.path.join(OUT_DIR, f"PROSPECTION_{date}_{suffix}.xlsx")
    build_single_sheet_xlsx(filename, rows_sorted)
    return filename

def send_email_with_attachments(to_email: str, subject: str, html: str, attachments: List[Tuple[str, bytes]]):
    brevo_send_email(to_email, subject, html, attachments=attachments)

def send_agency_manager_pack(date: str, agency: str,
                             prospects: List[Dict[str, Any]],
                             photos: List[Dict[str, Any]],
                             cards: List[Dict[str, Any]],
                             closes: List[Dict[str, Any]]):

    to_email = manager_email_for_agency(agency)
    if not to_email:
        print(f"[WARN] No manager email for agency={agency}")
        return

    rows: List[Dict[str, Any]] = []
    # prospects already enriched by worker
    for p in prospects:
        if (p.get("agency") or "").upper() == agency:
            rows.append(unify_from_prospect(p))

    # add image-derived prospects for that agency
    for ph in photos:
        if (ph.get("agency") or "").upper() == agency:
            rows.append(unify_from_image_record(ph, "photo"))

    for ca in cards:
        if (ca.get("agency") or "").upper() == agency:
            rows.append(unify_from_image_record(ca, "card"))

    if not rows:
        print(f"[INFO] No data for agency={agency}, skip manager pack.")
        return

    xlsx = build_excel_single(date, rows, f"AGENCE_{agency}")
    attachments: List[Tuple[str, bytes]] = []
    with open(xlsx, "rb") as f:
        attachments.append((os.path.basename(xlsx), f.read()))

    # ZIP média agence (utile au manager)
    media = build_media_zip(date, f"AGENCE_{agency}", photos, cards, filter_agency=agency, filter_initials=None)
    if media:
        attachments.append(media)

    # closes summary (in email body)
    c_ag = [c for c in closes if (c.get("agency") or "").upper() == agency]
    close_lines = ""
    if c_ag:
        close_lines = "<br><br><b>Clôtures</b><br>" + "<br>".join(
            [f"{escape_html(c.get('initials'))}: clients={escape_html(c.get('visits_clients'))}, prospects={escape_html(c.get('visits_prospects'))}, cmd={escape_html(c.get('commandes'))}"
             for c in c_ag]
        )

    subject = f"Prospection {date} — Agence {agency} (Excel + médias)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Voici l’export du <b>{date}</b> pour l’agence <b>{agency}</b> :</p>"
        f"<ul><li>Excel (1 onglet, trié par complétude)</li><li>ZIP médias (photos + cartes)</li></ul>"
        f"{close_lines}"
        f"<p>— Bot Prospection</p>"
    )

    send_email_with_attachments(to_email, subject, html, attachments)
    print(f"[OK] agency pack sent to {to_email} ({agency})")

def escape_html(x):
    return str(x or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def send_admin_pack(date: str, prospects: List[Dict[str, Any]], photos: List[Dict[str, Any]], cards: List[Dict[str, Any]], closes: List[Dict[str, Any]]):
    to_email = admin_email()
    if not to_email:
        print("[WARN] No admin email configured")
        return

    rows: List[Dict[str, Any]] = []
    for p in prospects:
        rows.append(unify_from_prospect(p))
    for ph in photos:
        rows.append(unify_from_image_record(ph, "photo"))
    for ca in cards:
        rows.append(unify_from_image_record(ca, "card"))

    if not rows:
        print("[INFO] No data for admin, skip.")
        return

    xlsx = build_excel_single(date, rows, "ADMIN_ALL")
    attachments: List[Tuple[str, bytes]] = []
    with open(xlsx, "rb") as f:
        attachments.append((os.path.basename(xlsx), f.read()))

    media = build_media_zip(date, "ADMIN_ALL", photos, cards, None, None)
    if media:
        attachments.append(media)

    subject = f"Prospection {date} — ADMIN (toutes agences)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Export global du <b>{date}</b> (toutes agences) :</p>"
        f"<ul><li>Excel (1 onglet, trié par complétude)</li><li>ZIP médias (photos + cartes)</li></ul>"
        f"<p>— Bot Prospection</p>"
    )

    send_email_with_attachments(to_email, subject, html, attachments)
    print(f"[OK] admin pack sent to {to_email}")


# ============================================================
# MAIN
# ============================================================

def main():
    if not WORKER_BASE_URL or not EXPORT_TOKEN or not TELEGRAM_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN / TELEGRAM_TOKEN")

    print(f"🚀 export_and_mail.py — mode={SEND_MODE} date={RUN_DATE} agency={AGENCY} initials={INITIALS}")

    prospects = worker_dump("prospects", RUN_DATE)
    closes    = worker_dump("closes",    RUN_DATE)
    photos    = worker_dump("photos",    RUN_DATE)
    cards     = worker_dump("cards",     RUN_DATE)

    # -------- individual
    if SEND_MODE == "individual":
        if AGENCY not in VALID_AGENCIES:
            raise RuntimeError("AGENCY required (GR|VR|GRS|SLS) for mode=individual")
        if not INITIALS:
            raise RuntimeError("INITIALS required for mode=individual")

        to_email = email_for_initials(INITIALS)
        if not to_email:
            raise RuntimeError(f"No email for initials={INITIALS}")

        rows: List[Dict[str, Any]] = []

        for p in prospects:
            if (p.get("agency") or "").upper() == AGENCY and (p.get("initials") or "").upper() == INITIALS:
                rows.append(unify_from_prospect(p))
        for ph in photos:
            if (ph.get("agency") or "").upper() == AGENCY and (ph.get("user") or "").upper() == INITIALS:
                rows.append(unify_from_image_record(ph, "photo"))
        for ca in cards:
            if (ca.get("agency") or "").upper() == AGENCY and (ca.get("user") or "").upper() == INITIALS:
                rows.append(unify_from_image_record(ca, "card"))

        if not rows:
            print("[INFO] No rows to send (individual).")
            return

        xlsx = build_excel_single(RUN_DATE, rows, f"INDIV_{AGENCY}_{INITIALS}")
        attachments: List[Tuple[str, bytes]] = []
        with open(xlsx, "rb") as f:
            attachments.append((os.path.basename(xlsx), f.read()))

        media = build_media_zip(RUN_DATE, f"{AGENCY}_{INITIALS}", photos, cards, filter_agency=AGENCY, filter_initials=INITIALS)
        if media:
            attachments.append(media)

        subject = f"Prospection {RUN_DATE} — {AGENCY}/{INITIALS} (Excel + médias)"
        html = (
            f"<p>Bonjour,</p>"
            f"<p>Voici ton export du <b>{RUN_DATE}</b> pour <b>{AGENCY}/{INITIALS}</b>.</p>"
            f"<ul><li>Excel (1 onglet, trié par complétude)</li><li>ZIP médias (photos + cartes)</li></ul>"
            f"<p>— Bot Prospection</p>"
        )
        brevo_send_email(to_email, subject, html, attachments=attachments)
        print(f"[OK] individual pack sent to {to_email}")
        return

    # -------- agency_manager
    if SEND_MODE == "agency_manager":
        for ag in sorted(VALID_AGENCIES):
            send_agency_manager_pack(RUN_DATE, ag, prospects, photos, cards, closes)
        return

    # -------- admin
    if SEND_MODE == "admin":
        send_admin_pack(RUN_DATE, prospects, photos, cards, closes)
        return

    raise RuntimeError(f"Unknown SEND_MODE={SEND_MODE}")


if __name__ == "__main__":
    main()
