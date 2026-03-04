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

MAIL_ROUTING_JSON  = env_str("MAIL_ROUTING_JSON", "")  # override possible

GOOGLE_PLACES_API_KEY = env_str("GOOGLE_PLACES_API_KEY", "")
GEMINI_API_KEY        = env_str("GEMINI_API_KEY", "")

# Modes inchangés, mais Excel = 1 onglet unique
SEND_MODE    = env_str("SEND_MODE", "individual").lower()  # individual | agency_manager | admin
RUN_DATE     = env_str("RUN_DATE", today_ymd_utc())
AGENCY       = env_str("AGENCY", "").upper()               # GR|VR|GRS|SLS
INITIALS     = env_str("INITIALS", "").upper()             # JL etc.

MAX_OCR_IMAGES   = env_int("MAX_OCR_IMAGES", 50)
MAX_PHOTO_IMAGES = env_int("MAX_PHOTO_IMAGES", 15)

OUT_DIR = env_str("OUT_DIR", ".")
MEDIA_DIR = os.path.join(OUT_DIR, "media")
os.makedirs(MEDIA_DIR, exist_ok=True)

VALID_AGENCIES = {"GR", "VR", "GRS", "SLS"}

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EMAIL_IN_TEXT_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(\+33|0)\s*[1-9](?:[\s\.-]*\d{2}){4}")
URL_RE = re.compile(r"(https?://[^\s]+|www\.[^\s]+)", re.I)


# ============================================================
# ROUTING (fallback intégré + override secret JSON)
# ============================================================

DEFAULT_ROUTING = {
    "agencies": {
        "GR": {
            "manager": {"initials": "JL", "email": "jennifer.laurens@ras-interim.fr"},
            "commercial": {"initials": "CZ", "email": "celine.zunarelli@ras-interim.fr"},
        },
        "VR": {
            "manager": {"initials": "JB", "email": "jelena.carrasso@ras-interim.fr"},
            "commercial": {"initials": "LB", "email": "laura.berthet@ras-interim.fr"},
        },
        "GRS": {
            "manager": {"initials": "PV", "email": "pauline.vieira@ras-interim.fr"},
            "commercial": {"initials": "ST", "email": "severine.thevenin@ras-interim.fr"},
        },
        "SLS": {
            "manager": {"initials": "AC", "email": "aurelie.curt@ras-interim.fr"},
            "commercial": {"initials": "AC", "email": "aurelie.curt@ras-interim.fr"},
        },
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
    for ag, cfg in agencies.items():
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
    r = requests.get(url, headers={"X-Export-Token": EXPORT_TOKEN}, timeout=30)
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
# NORMALISATION / REGEX FALLBACK
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

def best_website(text: str) -> str:
    if not text:
        return ""
    m = URL_RE.search(text)
    if not m:
        return ""
    u = m.group(0).strip()
    if u.lower().startswith("www."):
        u = "http://" + u
    return u

def norm_naf(naf: str) -> str:
    return (naf or "").strip().upper().replace(".", "").replace(" ", "")


# ============================================================
# ENRICH: API GOUV (recherche-entreprises) + Places + scrape site
# ============================================================

def api_gouv_search(name: str, city: str) -> Optional[Dict[str, Any]]:
    q = " ".join([x for x in [name, city] if x]).strip()
    if len(q) < 2:
        return None
    url = "https://recherche-entreprises.api.gouv.fr/search"
    try:
        r = requests.get(url, params={"q": q, "page": 1, "per_page": 3}, headers={"accept": "application/json"}, timeout=15)
        if r.status_code != 200:
            return None
        j = r.json()
        results = j.get("results") or []
        if not results:
            return None

        # simple scoring: prefer exact-ish name tokens + city match
        def score(res: Dict[str, Any]) -> int:
            s = 0
            nom = (res.get("nom_raison_sociale") or res.get("denomination") or "").lower()
            siege = res.get("siege") or {}
            ville = (siege.get("libelle_commune") or "").lower()
            tokens = [t for t in re.split(r"\s+", (name or "").lower()) if len(t) >= 3]
            for t in tokens:
                if t in nom:
                    s += 2
            if city and city.lower() in ville:
                s += 3
            if siege.get("siret"):
                s += 1
            return s

        best = sorted(results, key=score, reverse=True)[0]
        siege = best.get("siege") or {}
        dirigeants = best.get("dirigeants") or best.get("representants") or []
        dirigeant = ""
        if dirigeants:
            d0 = dirigeants[0]
            if isinstance(d0, str):
                dirigeant = d0
            elif isinstance(d0, dict):
                p = d0.get("personne") or d0
                prenom = p.get("prenom") or p.get("prenoms") or ""
                nom = p.get("nom") or p.get("nom_usage") or p.get("nomNaissance") or ""
                dirigeant = (prenom + " " + nom).strip() or (p.get("denomination") or "")

        return {
            "name": best.get("nom_raison_sociale") or best.get("denomination") or "",
            "siret": siege.get("siret") or "",
            "naf": best.get("activite_principale") or best.get("naf") or "",
            "address": siege.get("adresse") or siege.get("libelle_voie") or "",
            "postal_code": siege.get("code_postal") or "",
            "city": siege.get("libelle_commune") or "",
            "dirigeant": dirigeant,
        }
    except Exception:
        return None

def places_retry(name: str, city: str) -> Dict[str, str]:
    if not GOOGLE_PLACES_API_KEY or not name:
        return {"phone": "", "website": ""}
    try:
        r1 = requests.post(
            "https://places.googleapis.com/v1/places:searchText",
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "places.id,places.websiteUri",
            },
            json={
                "textQuery": f"{name} {city}".strip(),
                "maxResultCount": 1,
                "languageCode": "fr",
                "regionCode": "FR",
            },
            timeout=15
        )
        j1 = r1.json()
        p = (j1.get("places") or [None])[0]
        if not p or not p.get("id"):
            return {"phone": "", "website": p.get("websiteUri") if isinstance(p, dict) else ""}

        place_id = p["id"]
        website = p.get("websiteUri") or ""

        r2 = requests.get(
            f"https://places.googleapis.com/v1/places/{place_id}",
            headers={
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "nationalPhoneNumber,internationalPhoneNumber,websiteUri",
            },
            timeout=15
        )
        j2 = r2.json()
        phone = j2.get("nationalPhoneNumber") or j2.get("internationalPhoneNumber") or ""
        website = j2.get("websiteUri") or website
        return {"phone": normalize_phone(phone), "website": website or ""}
    except Exception:
        return {"phone": "", "website": ""}

def scrape_email_from_site(url: str) -> str:
    if not url:
        return ""
    try:
        headers = {"user-agent": "Mozilla/5.0 (compatible; ProspectionBot/3.0)"}
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code >= 300:
            return ""
        html = r.text[:250_000]  # simple limit
        matches = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}", html)
        if not matches:
            return ""
        matches = [m for m in matches if not re.search(r"no-?reply", m, re.I)]
        return (matches[0] if matches else "").strip().lower()
    except Exception:
        return ""


# ============================================================
# OCR (fallback) + GEMINI VISION (first)
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

def gemini_vision_extract(image_bytes: bytes, hint: str = "") -> Dict[str, str]:
    """
    Gemini Vision FIRST.
    Retourne des champs "prospection" génériques.
    """
    out = {
        "company": "",
        "address": "",
        "postal_code": "",
        "city": "",
        "contact_name": "",
        "title": "",
        "phone": "",
        "email": "",
        "website": "",
    }
    if not GEMINI_API_KEY or not image_bytes:
        return out

    try:
        b64 = base64.b64encode(image_bytes).decode("ascii")
        endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"

        prompt = (
            "Tu es un extracteur de données de prospection à partir d'une IMAGE (carte de visite OU photo d'enseigne/logo).\n"
            "Retourne STRICTEMENT un JSON (pas de texte autour) avec les clés:\n"
            "company, address, postal_code, city, contact_name, title, phone, email, website\n"
            "Règles:\n"
            "- Si non visible, mets une chaîne vide.\n"
            "- Si c'est une enseigne/logo: company doit être rempli si possible.\n"
            "- phone en format FR si possible.\n"
        )
        if hint:
            prompt += f"\nContexte (facultatif): {hint}\n"

        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {"inlineData": {"mimeType": "image/jpeg", "data": b64}}
                ]
            }],
            "generationConfig": {"temperature": 0.1, "maxOutputTokens": 400}
        }

        r = requests.post(endpoint, params={"key": GEMINI_API_KEY}, json=payload, timeout=30)
        j = r.json()
        parts = (((j.get("candidates") or [None])[0] or {}).get("content") or {}).get("parts") or []
        raw = ""
        for p in parts:
            if isinstance(p, dict) and p.get("text"):
                raw += p["text"]
        raw = (raw or "").strip()

        m = re.search(r"\{.*\}", raw, re.S)
        if not m:
            return out
        data = json.loads(m.group(0))

        out["company"] = (data.get("company") or "").strip()
        out["address"] = (data.get("address") or "").strip()
        out["postal_code"] = (data.get("postal_code") or "").strip()
        out["city"] = (data.get("city") or "").strip()
        out["contact_name"] = (data.get("contact_name") or "").strip()
        out["title"] = (data.get("title") or "").strip()
        out["phone"] = normalize_phone(data.get("phone") or "")
        out["email"] = (data.get("email") or "").strip().lower()
        out["website"] = (data.get("website") or "").strip()
        return out
    except Exception:
        return out


# ============================================================
# MEDIA ZIP (inchangé)
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
        raise RuntimeError(f"Brevo send failed {r.status_code}: {r.text[:300]}")


# ============================================================
# UNIQUE USERS helper
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


# ============================================================
# FINAL EXCEL FORMAT (1 onglet)
# ============================================================

FINAL_HEADERS = [
    "date", "agency", "initials",
    "name", "address", "postal_code", "city",
    "siret", "naf", "dirigeant",
    "interlocuteur", "contact_firstname", "contact_lastname",
    "phone", "phone2", "email", "website",
    "resume", "commande",
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

def make_wb() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "PROSPECTION"
    return wb

def completion_score(row: Dict[str, Any]) -> int:
    # plus il y a de champs remplis, plus le score est haut
    keys = ["name","address","postal_code","city","siret","naf","dirigeant","interlocuteur","contact_firstname","contact_lastname","phone","email","website","resume","commande"]
    s = 0
    for k in keys:
        if str(row.get(k, "") or "").strip():
            s += 1
    return s

def row_to_list(row: Dict[str, Any]) -> List[Any]:
    return [row.get(h, "") for h in FINAL_HEADERS]


# ============================================================
# TRANSFORM: prospects -> unified rows
# ============================================================

def split_human_name(full: str) -> Tuple[str, str]:
    s = (full or "").strip()
    if not s:
        return ("", "")
    parts = re.split(r"\s+", s)
    if len(parts) == 1:
        return ("", parts[0])
    return (parts[0], " ".join(parts[1:]))

def unify_from_prospect(p: Dict[str, Any]) -> Dict[str, Any]:
    # prospects KV = déjà structuré
    return {
        "date": p.get("date",""),
        "agency": (p.get("agency") or "").upper(),
        "initials": (p.get("initials") or "").upper(),
        "name": p.get("name",""),
        "address": p.get("address",""),
        "postal_code": p.get("postal_code",""),
        "city": p.get("city",""),
        "siret": p.get("siret",""),
        "naf": norm_naf(p.get("naf","")),
        "dirigeant": p.get("dirigeant",""),
        "interlocuteur": p.get("interlocuteur",""),
        "contact_firstname": p.get("contact_firstname",""),
        "contact_lastname": p.get("contact_lastname",""),
        "phone": normalize_phone(p.get("phone","")),
        "phone2": normalize_phone(p.get("phone2","")),
        "email": (p.get("email","") or "").strip().lower(),
        "website": p.get("website",""),
        "resume": p.get("resume",""),
        "commande": p.get("commande",""),
    }


# ============================================================
# TRANSFORM: card/photo -> unified rows using Gemini->OCR->regex->enrich
# ============================================================

def fill_if_empty(dst: Dict[str, Any], key: str, val: str):
    if not str(dst.get(key, "") or "").strip() and str(val or "").strip():
        dst[key] = val

def enrich_row_with_apis(row: Dict[str, Any]) -> Dict[str, Any]:
    name = str(row.get("name","") or "").strip()
    city = str(row.get("city","") or "").strip()

    # 1) API Gouv
    if name:
        g = api_gouv_search(name, city)
        if g:
            fill_if_empty(row, "name", g.get("name",""))
            fill_if_empty(row, "siret", g.get("siret",""))
            fill_if_empty(row, "naf", norm_naf(g.get("naf","")))
            fill_if_empty(row, "address", g.get("address",""))
            fill_if_empty(row, "postal_code", g.get("postal_code",""))
            fill_if_empty(row, "city", g.get("city",""))
            fill_if_empty(row, "dirigeant", g.get("dirigeant",""))

    # 2) Places (phone/website)
    if name and (not row.get("phone") or not row.get("website")):
        pl = places_retry(name, row.get("city",""))
        fill_if_empty(row, "phone", pl.get("phone",""))
        fill_if_empty(row, "website", pl.get("website",""))

    # 3) scrape email
    if row.get("website") and not row.get("email"):
        em = scrape_email_from_site(str(row.get("website")))
        fill_if_empty(row, "email", em)

    return row

def unify_from_card(date: str, agency: str, initials: str, c: Dict[str, Any]) -> Dict[str, Any]:
    fid = c.get("file_id") or ""
    comment = c.get("comment") or ""

    base = {
        "date": date,
        "agency": agency,
        "initials": initials,
        "name": "", "address": "", "postal_code": "", "city": "",
        "siret": "", "naf": "", "dirigeant": "",
        "interlocuteur": "", "contact_firstname": "", "contact_lastname": "",
        "phone": "", "phone2": "", "email": "", "website": "",
        "resume": (comment or ""),  # on stocke le commentaire en "resume"
        "commande": "",
    }

    url = tg_get_file_url(fid) if fid else None
    if not url:
        return base

    img = download_bytes(url)

    # 1) Gemini Vision FIRST
    g = gemini_vision_extract(img, hint="Carte de visite")
    fill_if_empty(base, "name", g.get("company",""))
    fill_if_empty(base, "address", g.get("address",""))
    fill_if_empty(base, "postal_code", g.get("postal_code",""))
    fill_if_empty(base, "city", g.get("city",""))
    fill_if_empty(base, "phone", g.get("phone",""))
    fill_if_empty(base, "email", g.get("email",""))
    fill_if_empty(base, "website", g.get("website",""))

    # contact
    contact = (g.get("contact_name") or "").strip()
    if contact and not base.get("interlocuteur"):
        base["interlocuteur"] = contact
        fn, ln = split_human_name(contact)
        fill_if_empty(base, "contact_firstname", fn)
        fill_if_empty(base, "contact_lastname", ln)

    # 2) OCR fallback (si incomplet)
    need_more = (not base["email"]) or (not base["phone"]) or (not base["website"]) or (not base["name"])
    if need_more:
        ocr_txt = ocr_image_bytes(img)
        fill_if_empty(base, "email", best_email(ocr_txt))
        fill_if_empty(base, "phone", best_phone(ocr_txt))
        fill_if_empty(base, "website", best_website(ocr_txt))

        # tentative company via OCR: première ligne non vide
        if not base["name"]:
            lines = [ln.strip() for ln in ocr_txt.splitlines() if ln.strip()]
            if lines:
                fill_if_empty(base, "name", lines[0][:180])

    # 3) Regex already done above; now enrich
    base = enrich_row_with_apis(base)
    return base

def unify_from_photo(date: str, agency: str, initials: str, p: Dict[str, Any]) -> Dict[str, Any]:
    fid = p.get("file_id") or ""
    comment = p.get("comment") or p.get("meeting") or ""
    city_hint = (p.get("city") or "").strip()

    base = {
        "date": date,
        "agency": agency,
        "initials": initials,
        "name": "", "address": "", "postal_code": "", "city": city_hint,
        "siret": "", "naf": "", "dirigeant": "",
        "interlocuteur": "", "contact_firstname": "", "contact_lastname": "",
        "phone": "", "phone2": "", "email": "", "website": "",
        "resume": (comment or ""),
        "commande": "",
    }

    url = tg_get_file_url(fid) if fid else None
    if not url:
        return base

    img = download_bytes(url)

    # 1) Gemini Vision FIRST
    g = gemini_vision_extract(img, hint=f"Photo enseigne/logo. Ville/secteur: {city_hint}")
    fill_if_empty(base, "name", g.get("company",""))
    fill_if_empty(base, "address", g.get("address",""))
    fill_if_empty(base, "postal_code", g.get("postal_code",""))
    fill_if_empty(base, "city", g.get("city",""))
    fill_if_empty(base, "phone", g.get("phone",""))
    fill_if_empty(base, "email", g.get("email",""))
    fill_if_empty(base, "website", g.get("website",""))

    # 2) OCR fallback
    need_more = (not base["name"]) or (not base["email"]) or (not base["phone"]) or (not base["website"])
    if need_more:
        ocr_txt = ocr_image_bytes(img)
        fill_if_empty(base, "email", best_email(ocr_txt))
        fill_if_empty(base, "phone", best_phone(ocr_txt))
        fill_if_empty(base, "website", best_website(ocr_txt))
        if not base["name"]:
            lines = [ln.strip() for ln in ocr_txt.splitlines() if ln.strip()]
            if lines:
                fill_if_empty(base, "name", lines[0][:180])

    # 3) enrich
    base = enrich_row_with_apis(base)
    return base


# ============================================================
# BUILD EXCEL (1 sheet) + sorting by completion
# ============================================================

def build_excel_single(date: str, rows: List[Dict[str, Any]], title_suffix: str) -> str:
    # sort by completeness desc
    rows_sorted = sorted(rows, key=completion_score, reverse=True)

    wb = make_wb()
    ws = wb.active

    ws.append(FINAL_HEADERS)
    for r in rows_sorted:
        ws.append(row_to_list(r))

    for c in range(1, len(FINAL_HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    autosize(ws)

    filename = os.path.join(OUT_DIR, f"PROSPECTION_{date}_{title_suffix}.xlsx")
    wb.save(filename)
    return filename


# ============================================================
# SEND LOGIC
# ============================================================

def send_individual_pack(date: str, agency: str, initials: str,
                         prospects: List[Dict[str, Any]],
                         photos: List[Dict[str, Any]],
                         cards: List[Dict[str, Any]]):

    to_email = email_for_initials(initials)
    if not to_email:
        print(f"[WARN] No email for initials={initials}, skip individual pack.")
        return

    initials = initials.upper().strip()

    p_u  = [p  for p  in prospects if (p.get("agency") == agency and (p.get("initials") or "").upper() == initials)]
    ph_u = [ph for ph in photos    if (ph.get("agency") == agency and (ph.get("user") or "").upper() == initials)]
    ca_u = [ca for ca in cards     if (ca.get("agency") == agency and (ca.get("user") or "").upper() == initials)]

    # Build unified rows
    rows: List[Dict[str, Any]] = []
    for p in p_u:
        rows.append(unify_from_prospect(p))

    # photos/cards -> turned into prospects-like rows
    for ph in ph_u[:MAX_PHOTO_IMAGES]:
        rows.append(unify_from_photo(date, agency, initials, ph))

    for ca in ca_u[:MAX_OCR_IMAGES]:
        rows.append(unify_from_card(date, agency, initials, ca))

    if not rows:
        print(f"[INFO] No activity for {agency}/{initials}, skip individual pack.")
        return

    xlsx = build_excel_single(date, rows, f"INDIV_{agency}_{initials}")

    attachments: List[Tuple[str, bytes]] = []
    with open(xlsx, "rb") as f:
        attachments.append((os.path.basename(xlsx), f.read()))

    media = build_media_zip(date, agency, initials, photos, cards)
    if media:
        attachments.append(media)

    subject = f"Prospection {date} — {agency}/{initials} (Excel unique + médias)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Voici ton export de prospection du <b>{date}</b> pour <b>{agency}/{initials}</b>.</p>"
        f"<ul>"
        f"<li>Excel (1 seul onglet, trié par complétude)</li>"
        f"<li>Zip médias (photos + cartes) si présents</li>"
        f"</ul>"
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

    # scope agency
    p_ag  = [p for p in prospects if (p.get("agency") == agency)]
    ph_ag = [ph for ph in photos    if (ph.get("agency") == agency)]
    ca_ag = [ca for ca in cards     if (ca.get("agency") == agency)]

    rows: List[Dict[str, Any]] = []
    for p in p_ag:
        rows.append(unify_from_prospect(p))
    # managers also get photo/card as rows (no media zip attached)
    for ph in ph_ag[:MAX_PHOTO_IMAGES]:
        ini = (ph.get("user") or "").upper() or ""
        rows.append(unify_from_photo(date, agency, ini or "?", ph))
    for ca in ca_ag[:MAX_OCR_IMAGES]:
        ini = (ca.get("user") or "").upper() or ""
        rows.append(unify_from_card(date, agency, ini or "?", ca))

    if not rows:
        print(f"[INFO] No rows for agency manager pack {agency}, skip.")
        return

    xlsx = build_excel_single(date, rows, f"AGENCE_{agency}")

    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    subject = f"Prospection {date} — Agence {agency} (Excel unique)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Voici le consolidé prospection du <b>{date}</b> pour l’agence <b>{agency}</b>.</p>"
        f"<p>(Excel unique, trié par complétude. Médias envoyés individuellement aux preneurs.)</p>"
        f"<p>— Bot Prospection</p>"
    )

    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] agency manager pack sent to {to_email} (agency={agency})")


def send_admin_pack(date: str, prospects: List[Dict[str, Any]], photos: List[Dict[str, Any]], cards: List[Dict[str, Any]]):
    admin = ROUTING.get("admin") or {}
    to_email = clean_email(admin.get("email") or "")

    if not is_valid_email(to_email):
        print(f"[WARN] No valid admin email configured ({to_email})")
        return

    rows: List[Dict[str, Any]] = []
    for p in prospects:
        rows.append(unify_from_prospect(p))
    for ph in photos[:MAX_PHOTO_IMAGES]:
        ag = (ph.get("agency") or "").upper()
        ini = (ph.get("user") or "").upper() or ""
        rows.append(unify_from_photo(date, ag, ini or "?", ph))
    for ca in cards[:MAX_OCR_IMAGES]:
        ag = (ca.get("agency") or "").upper()
        ini = (ca.get("user") or "").upper() or ""
        rows.append(unify_from_card(date, ag, ini or "?", ca))

    if not rows:
        print("[INFO] No rows for admin pack, skip.")
        return

    xlsx = build_excel_single(date, rows, "ADMIN_ALL")

    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    subject = f"Prospection {date} — ADMIN (Excel unique)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Voici le consolidé global du <b>{date}</b> (toutes agences / tous collaborateurs).</p>"
        f"<p>(Excel unique, trié par complétude. Médias envoyés individuellement aux preneurs.)</p>"
        f"<p>— Bot Prospection</p>"
    )

    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] admin pack sent to {to_email}")


# ============================================================
# MAIN
# ============================================================

def main():
    if not WORKER_BASE_URL or not EXPORT_TOKEN or not TELEGRAM_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN / TELEGRAM_TOKEN")

    print(f"🚀 export_and_mail.py — mode={SEND_MODE} date={RUN_DATE} agency={AGENCY} initials={INITIALS}")

    prospects = worker_dump("prospects", RUN_DATE)
    photos    = worker_dump("photos",    RUN_DATE)
    cards     = worker_dump("cards",     RUN_DATE)

    # MODE: individual
    if SEND_MODE == "individual":
        if AGENCY not in VALID_AGENCIES:
            raise RuntimeError("AGENCY required (GR|VR|GRS|SLS) for mode=individual")
        if not INITIALS:
            raise RuntimeError("INITIALS required for mode=individual")

        send_individual_pack(RUN_DATE, AGENCY, INITIALS, prospects, photos, cards)
        return

    # MODE: agency_manager
    if SEND_MODE == "agency_manager":
        for ag in sorted(VALID_AGENCIES):
            send_agency_manager_pack(RUN_DATE, ag, prospects, photos, cards)

        # packs individuels aux preneurs (media + excel)
        media_users: Set[Tuple[str, str]] = set(uniq_users(photos) + uniq_users(cards) + uniq_users(prospects))
        for ag, ini in sorted(media_users):
            send_individual_pack(RUN_DATE, ag, ini, prospects, photos, cards)
        return

    # MODE: admin
    if SEND_MODE == "admin":
        send_admin_pack(RUN_DATE, prospects, photos, cards)

        media_users: Set[Tuple[str, str]] = set(uniq_users(photos) + uniq_users(cards) + uniq_users(prospects))
        for ag, ini in sorted(media_users):
            send_individual_pack(RUN_DATE, ag, ini, prospects, photos, cards)
        return

    raise RuntimeError(f"Unknown SEND_MODE={SEND_MODE}")


if __name__ == "__main__":
    main()