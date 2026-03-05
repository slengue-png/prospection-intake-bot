import os
import re
import json
import io
import base64
import zipfile
import datetime as dt
from typing import Dict, List, Any, Optional, Tuple, Set

import requests
from PIL import Image, ImageOps, ImageFilter

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

# debug (tu peux mettre DEBUG=1 en env pour générer un jsonl avec détails extraction)
DEBUG = env_str("DEBUG", "0") == "1"

VALID_AGENCIES = {"GR", "VR", "GRS", "SLS"}

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EMAIL_IN_TEXT_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(\+33|0)\s*[1-9](?:[\s\.-]*\d{2}){4}")
SIRET_RE = re.compile(r"\b\d{14}\b")
POSTAL_RE = re.compile(r"\b(\d{5})\b")

URL_RE = re.compile(r"\bhttps?://[^\s]+|\bwww\.[^\s]+", re.I)

# 1 seul onglet (format cible)
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

def clean_email(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", "", s)
    return s

def is_valid_email(s: str) -> bool:
    s = clean_email(s)
    return bool(EMAIL_RE.match(s))

def email_for_initials(initials: str) -> Optional[str]:
    initials = (initials or "").upper().strip()
    if not initials:
        return None

    users = ROUTING.get("users") or {}
    if initials in users:
        em = clean_email(users[initials] if isinstance(users[initials], str) else (users[initials] or {}).get("email",""))
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
    if not WORKER_BASE_URL or not EXPORT_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN")
    url = f"{WORKER_BASE_URL.rstrip('/')}/dump?date={date}&kind={kind}"
    r = requests.get(url, headers={"X-Export-Token": EXPORT_TOKEN}, timeout=30)
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
    r = requests.post(url, json={"file_id": file_id}, timeout=20)
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
# NORMALISATION / HELPERS
# ============================================================

def normalize_phone(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    s = re.sub(r"[^\d+]", "", s)
    # convert +33 -> 0 if possible
    if s.startswith("+33") and len(s) >= 11:
        s = "0" + s[3:]
    return s

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

def best_url(text: str) -> str:
    if not text:
        return ""
    m = URL_RE.search(text)
    if not m:
        return ""
    u = m.group(0).strip().rstrip(").,;")
    if u.lower().startswith("www."):
        u = "http://" + u
    return u

def best_siret(text: str) -> str:
    if not text:
        return ""
    m = SIRET_RE.search(re.sub(r"\s+", "", text))
    return m.group(0) if m else ""

def best_postal_city(text: str) -> Tuple[str, str]:
    if not text:
        return "", ""
    # cherche pattern "38340 VOREPPE"
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines:
        m = re.search(r"\b(\d{5})\s+([A-Za-zÀ-ÿ'\- ]{2,})\b", ln)
        if m:
            pc = m.group(1)
            city = m.group(2).strip()
            # évite les trucs trop longs
            if 2 <= len(city) <= 40:
                return pc, city
    # fallback
    m2 = POSTAL_RE.search(text)
    return (m2.group(1), "") if m2 else ("","")

def split_human_name(full: str) -> Tuple[str,str]:
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
# ENRICH: API GOUV + PLACES + SCRAPE EMAIL
# ============================================================

def search_gouv_company(name: str, city: str) -> Dict[str, str]:
    if not name:
        return {}
    try:
        q = f"{name} {city}".strip()
        url = f"https://recherche-entreprises.api.gouv.fr/search?q={requests.utils.quote(q)}&page=1&per_page=1"
        r = requests.get(url, headers={"accept":"application/json"}, timeout=15)
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

        naf = (e.get("activite_principale") or e.get("naf") or "").replace(".","").replace(" ","").upper()

        addr = (
            siege.get("adresse")
            or " ".join([str(x) for x in [siege.get("numero_voie"), siege.get("type_voie"), siege.get("libelle_voie")] if x])
            or ""
        ).strip()

        return {
            "name": e.get("nom_raison_sociale") or e.get("denomination") or name,
            "siret": (siege.get("siret") or ""),
            "naf": naf,
            "address": addr,
            "postal_code": (siege.get("code_postal") or ""),
            "city": (siege.get("libelle_commune") or city),
            "dirigeant": dirigeant,
        }
    except Exception:
        return {}

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
        website = p.get("websiteUri") or ""
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
        website2 = j2.get("websiteUri") or website
        return {"phone": normalize_phone(phone), "website": website2}
    except Exception:
        return {}

def scrape_email_from_site(url: str) -> str:
    if not url:
        return ""
    try:
        r = requests.get(url, headers={"user-agent":"Mozilla/5.0 (ProspectionBot)"}, timeout=10, allow_redirects=True)
        if r.status_code >= 300:
            return ""
        html = r.text[:250000]
        matches = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}", html)
        if not matches:
            return ""
        filtered = [m for m in matches if not re.search(r"no-?reply|example\.com", m, re.I)]
        return (filtered[0] if filtered else matches[0]).strip().lower()
    except Exception:
        return ""


# ============================================================
# OCR (multi-pass + preprocessing)
# ============================================================

def _pil_open(img_bytes: bytes) -> Optional[Image.Image]:
    try:
        im = Image.open(io.BytesIO(img_bytes))
        if im.mode not in ("RGB","L"):
            im = im.convert("RGB")
        return im
    except Exception:
        return None

def _preprocess_variants(im: Image.Image) -> List[Image.Image]:
    variants: List[Image.Image] = []

    # base
    base = im.copy()
    variants.append(base)

    # grayscale + autocontrast
    g = ImageOps.grayscale(im)
    g = ImageOps.autocontrast(g)
    variants.append(g)

    # upscale x2 + sharpen
    up = im.resize((im.size[0]*2, im.size[1]*2))
    up = up.filter(ImageFilter.SHARPEN)
    variants.append(up)

    # grayscale + threshold
    gt = ImageOps.grayscale(im)
    gt = ImageOps.autocontrast(gt)
    gt = gt.point(lambda p: 255 if p > 160 else 0)
    variants.append(gt)

    return variants

def ocr_image_bytes(img_bytes: bytes) -> str:
    if not img_bytes:
        return ""
    im = _pil_open(img_bytes)
    if im is None:
        return ""
    texts: List[str] = []
    config = "--oem 3 --psm 6"
    for v in _preprocess_variants(im):
        try:
            txt = pytesseract.image_to_string(v, lang="fra+eng", config=config)
            txt = (txt or "").strip()
            if txt:
                texts.append(txt)
        except Exception:
            pass
    # garde le plus "long" (souvent le plus informatif)
    texts = sorted(texts, key=lambda s: len(s), reverse=True)
    return texts[0] if texts else ""


# ============================================================
# GEMINI (Vision + Normalisation)
# ============================================================

def _guess_mime(img_bytes: bytes) -> str:
    try:
        im = Image.open(io.BytesIO(img_bytes))
        fmt = (im.format or "").upper()
        if fmt == "PNG":
            return "image/png"
        if fmt in ("JPG","JPEG"):
            return "image/jpeg"
    except Exception:
        pass
    return "image/jpeg"

def gemini_vision_json(img_bytes: bytes, prompt: str) -> Dict[str, Any]:
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
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 700}
    }

    try:
        r = requests.post(endpoint, params={"key": GEMINI_API_KEY}, json=payload, timeout=40)
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

def gemini_text_json(prompt: str) -> Dict[str, Any]:
    if not GEMINI_API_KEY or not prompt.strip():
        return {}
    endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 700}
    }
    try:
        r = requests.post(endpoint, params={"key": GEMINI_API_KEY}, json=payload, timeout=40)
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

def gemini_extract_business_card(img_bytes: bytes) -> Dict[str, Any]:
    prompt = (
        "RÔLE: Tu es un extracteur de carte de visite (France).\n"
        "Tu dois lire l'IMAGE et retourner STRICTEMENT un JSON valide (aucun texte autour).\n"
        "Clés EXACTES:\n"
        "{"
        "\"person_fullname\":\"\","
        "\"person_title\":\"\","
        "\"company_name\":\"\","
        "\"company_address\":\"\","
        "\"postal_code\":\"\","
        "\"city\":\"\","
        "\"phone\":\"\","
        "\"phone2\":\"\","
        "\"email\":\"\","
        "\"website\":\"\","
        "\"siret\":\"\""
        "}\n"
        "Règles:\n"
        "- Si inconnu -> \"\"\n"
        "- email en minuscules\n"
        "- phone en format FR si possible (0XXXXXXXXX)\n"
        "- siret = 14 chiffres si présent\n"
        "- N'invente rien.\n"
    )
    return gemini_vision_json(img_bytes, prompt) or {}

def gemini_extract_facade(img_bytes: bytes) -> Dict[str, Any]:
    prompt = (
        "RÔLE: Tu analyses une photo de prospection (façade/enseigne/logo/devanture).\n"
        "Retourne STRICTEMENT un JSON valide (aucun texte autour) avec clés EXACTES:\n"
        "{"
        "\"company_name\":\"\","
        "\"city\":\"\""
        "}\n"
        "Règles:\n"
        "- Si tu n'es pas sûr -> \"\"\n"
        "- N'invente pas.\n"
    )
    return gemini_vision_json(img_bytes, prompt) or {}

def gemini_normalize(structured: Dict[str, Any], ocr_text: str, kind: str) -> Dict[str, Any]:
    """
    kind = 'card' ou 'photo'
    On force Gemini à NORMALISER (pas inventer).
    """
    if not GEMINI_API_KEY:
        return {}

    # limite taille OCR
    ocr_trim = (ocr_text or "").strip()
    if len(ocr_trim) > 6000:
        ocr_trim = ocr_trim[:6000]

    prompt = (
        f"RÔLE: Tu es un normaliseur de données de prospection (France).\n"
        f"OBJECTIF: produire un JSON FINAL propre, cohérent, sans invention.\n"
        f"TYPE: {kind}\n\n"
        f"JSON_CANDIDAT:\n{json.dumps(structured, ensure_ascii=False)}\n\n"
        f"TEXTE_OCR_BRUT (peut être bruité):\n{ocr_trim}\n\n"
        "INSTRUCTIONS STRICTES:\n"
        "- Retourne STRICTEMENT un JSON valide, aucune phrase autour.\n"
        "- Ne rajoute pas de champs.\n"
        "- Si une info est douteuse -> laisse \"\".\n"
        "- Nettoie: espaces, majuscules abusives, URL (garde https:// si possible), phone au format 0XXXXXXXXX.\n"
        "- postal_code = 5 chiffres.\n"
        "- city = nom de ville sans CP.\n"
        "- company_name = raison sociale (pas le nom de personne).\n"
    )
    out = gemini_text_json(prompt) or {}
    return out if isinstance(out, dict) else {}


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
# ONE-SHEET EXCEL
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
# PIPELINES (Card / Photo)
# ============================================================

def _debug_write(obj: Dict[str, Any], path: str):
    if not DEBUG:
        return
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(obj, ensure_ascii=False) + "\n")
    except Exception:
        pass

def _merge_keep(existing: Dict[str, Any], incoming: Dict[str, Any]):
    for k, v in (incoming or {}).items():
        if not str(existing.get(k,"") or "").strip() and str(v or "").strip():
            existing[k] = v

def pipeline_card(img_bytes: bytes) -> Tuple[Dict[str, Any], str]:
    """
    Retourne (structured, ocr_text)
    structured: dict proche "carte" (champs card)
    """
    structured = {
        "person_fullname": "",
        "person_title": "",
        "company_name": "",
        "company_address": "",
        "postal_code": "",
        "city": "",
        "phone": "",
        "phone2": "",
        "email": "",
        "website": "",
        "siret": ""
    }

    # 1) Gemini Vision sur image (prioritaire)
    vis = gemini_extract_business_card(img_bytes) if GEMINI_API_KEY else {}
    if isinstance(vis, dict):
        # nettoyages mini
        vis["email"] = (vis.get("email") or "").strip().lower()
        vis["phone"] = normalize_phone(vis.get("phone") or "")
        vis["phone2"] = normalize_phone(vis.get("phone2") or "")
        vis["website"] = (vis.get("website") or "").strip()
        vis["siret"] = re.sub(r"\D", "", str(vis.get("siret") or ""))[:14]
        _merge_keep(structured, vis)

    # 2) OCR multi-pass (toujours), pour compléter
    ocr_txt = ocr_image_bytes(img_bytes)
    if ocr_txt:
        # email / phone / url / siret / cp+city
        if not structured["email"]:
            structured["email"] = best_email(ocr_txt)
        if not structured["phone"]:
            structured["phone"] = best_phone(ocr_txt)
        if not structured["website"]:
            structured["website"] = best_url(ocr_txt)
        if not structured["siret"]:
            structured["siret"] = best_siret(ocr_txt)

        if not structured["postal_code"] or not structured["city"]:
            pc, city = best_postal_city(ocr_txt)
            if not structured["postal_code"]:
                structured["postal_code"] = pc
            if not structured["city"]:
                structured["city"] = city

        # tentative adresse : si on a cp, on prend la ligne qui le contient comme base
        if not structured["company_address"] and structured["postal_code"]:
            for ln in [l.strip() for l in ocr_txt.splitlines() if l.strip()]:
                if structured["postal_code"] in ln:
                    structured["company_address"] = ln.strip()
                    break

    return structured, ocr_txt

def pipeline_photo(img_bytes: bytes, city_hint: str) -> Tuple[Dict[str, Any], str]:
    structured = {"company_name": "", "city": (city_hint or "").strip()}
    vis = gemini_extract_facade(img_bytes) if GEMINI_API_KEY else {}
    if isinstance(vis, dict):
        _merge_keep(structured, {"company_name": (vis.get("company_name") or "").strip(),
                                 "city": (vis.get("city") or "").strip()})
    ocr_txt = ocr_image_bytes(img_bytes)
    if ocr_txt:
        # si ville manquante -> tente cp+ville
        if not structured["city"]:
            _, city = best_postal_city(ocr_txt)
            structured["city"] = city
    return structured, ocr_txt


# ============================================================
# UNIFY -> ONE ROW MODEL
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

    # 1) Vision + OCR (complément)
    structured, ocr_txt = pipeline_card(img_bytes)

    # 2) Enrichissements (API gouv + places + scrape)
    company = (structured.get("company_name") or "").strip()
    city = (structured.get("city") or "").strip()

    # si pas de company mais un email domaine -> tente company via domaine (option simple)
    # (on garde soft pour ne pas inventer)
    if not company and structured.get("website"):
        company = company  # rien ici pour l’instant

    gouv = search_gouv_company(company, city) if company else {}
    places = places_enrich(gouv.get("name") or company, gouv.get("city") or city)

    website = (structured.get("website") or places.get("website") or "").strip()
    email_site = scrape_email_from_site(website) if website and not structured.get("email") else ""

    # 3) Normalisation Gemini (texte) : réconcilie structured + OCR
    normalized = {}
    if GEMINI_API_KEY:
        normalized = gemini_normalize(structured, ocr_txt, kind="card")

    final_card = dict(structured)
    if isinstance(normalized, dict) and normalized:
        # on remplace par le normalisé (mais sans perdre tel/email déjà ok)
        final_card.update(normalized)

    # re-nettoyage
    final_card["email"] = (final_card.get("email") or "").strip().lower()
    final_card["phone"] = normalize_phone(final_card.get("phone") or "")
    final_card["phone2"] = normalize_phone(final_card.get("phone2") or "")
    final_card["website"] = (final_card.get("website") or "").strip()
    final_card["siret"] = re.sub(r"\D", "", str(final_card.get("siret") or ""))[:14]

    # merge enrichissements sans écraser
    # adresse / cp / ville / siret / naf / dirigeant
    def keep_if_empty(key, val):
        if not str(final_card.get(key,"") or "").strip() and str(val or "").strip():
            final_card[key] = val

    keep_if_empty("company_name", gouv.get("name"))
    keep_if_empty("company_address", gouv.get("address"))
    keep_if_empty("postal_code", gouv.get("postal_code"))
    keep_if_empty("city", gouv.get("city"))
    keep_if_empty("siret", gouv.get("siret"))
    keep_if_empty("naf", gouv.get("naf"))
    keep_if_empty("dirigeant", gouv.get("dirigeant"))

    keep_if_empty("phone", places.get("phone"))
    keep_if_empty("website", places.get("website"))
    keep_if_empty("email", email_site)

    # 4) Mapping -> ton format Excel
    fullname = (final_card.get("person_fullname") or "").strip()
    fn, ln = split_human_name(fullname)

    row = dict(base)
    row.update({
        "name": final_card.get("company_name") or "",
        "address": final_card.get("company_address") or "",
        "postal_code": final_card.get("postal_code") or "",
        "city": final_card.get("city") or "",
        "siret": final_card.get("siret") or "",
        "naf": final_card.get("naf") or "",
        "dirigeant": final_card.get("dirigeant") or "",
        "interlocuteur": fullname or "",
        "contact_firstname": fn,
        "contact_lastname": ln,
        "phone": final_card.get("phone") or "",
        "phone2": final_card.get("phone2") or "",
        "email": final_card.get("email") or "",
        "website": final_card.get("website") or "",
    })

    if DEBUG:
        _debug_write({
            "kind": "card",
            "date": date, "agency": agency, "initials": initials,
            "structured_initial": structured,
            "normalized": normalized,
            "gouv": gouv,
            "places": places,
            "row": row,
            "ocr_excerpt": (ocr_txt[:800] if ocr_txt else "")
        }, os.path.join(OUT_DIR, f"debug_{date}_{agency}_{initials}.jsonl"))

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

    structured, ocr_txt = pipeline_photo(img_bytes, city_hint=city_hint)
    # enrich
    company = (structured.get("company_name") or "").strip()
    city = (structured.get("city") or "").strip()

    if not company:
        return base

    gouv = search_gouv_company(company, city)
    places = places_enrich(gouv.get("name") or company, gouv.get("city") or city)
    website = (places.get("website") or "").strip()
    email_site = scrape_email_from_site(website) if website else ""

    row = dict(base)
    row.update({
        "name": gouv.get("name") or company,
        "address": gouv.get("address") or "",
        "postal_code": gouv.get("postal_code") or "",
        "city": gouv.get("city") or city,
        "siret": gouv.get("siret") or "",
        "naf": gouv.get("naf") or "",
        "dirigeant": gouv.get("dirigeant") or "",
        "phone": places.get("phone") or "",
        "website": website,
        "email": email_site or "",
    })

    if DEBUG:
        _debug_write({
            "kind": "photo",
            "date": date, "agency": agency, "initials": initials,
            "structured": structured,
            "gouv": gouv,
            "places": places,
            "row": row,
            "ocr_excerpt": (ocr_txt[:800] if ocr_txt else "")
        }, os.path.join(OUT_DIR, f"debug_{date}_{agency}_{initials}.jsonl"))

    return ensure_contact_fallback(row)


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
        if not url:
            rows.append(unify_from_photo(date, agency, initials, city_hint, comment, b""))
            continue
        try:
            img = download_bytes(url)
            rows.append(unify_from_photo(date, agency, initials, city_hint, comment, img))
        except Exception:
            rows.append(unify_from_photo(date, agency, initials, city_hint, comment, b""))

    for ca in ca_u[:MAX_OCR_IMAGES]:
        fid = ca.get("file_id") or ""
        url = tg_get_file_url(fid)
        comment = ca.get("comment") or ""
        if not url:
            rows.append(unify_from_card(date, agency, initials, comment, b""))
            continue
        try:
            img = download_bytes(url)
            rows.append(unify_from_card(date, agency, initials, comment, img))
        except Exception:
            rows.append(unify_from_card(date, agency, initials, comment, b""))

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
        f"<ul>"
        f"<li><b>Excel</b> (1 onglet, trié par complétude)</li>"
        f"<li><b>ZIP médias</b> (photos + cartes)</li>"
        f"</ul>"
        f"<p>— Bot Prospection</p>"
    )

    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] individual pack sent to {to_email} ({agency}/{initials})")

def send_agency_manager_pack(date: str, agency: str, prospects: List[Dict[str, Any]]):
    agencies_cfg = ROUTING.get("agencies") or {}
    cfg = agencies_cfg.get(agency) or {}
    manager = (cfg.get("manager") or {})
    to_email = clean_email(manager.get("email") or "")
    if not is_valid_email(to_email):
        print(f"[WARN] No valid manager email for agency={agency} ({to_email})")
        return

    p_ag = [p for p in prospects if (p.get("agency") == agency)]
    rows = [unify_from_prospect(p) for p in p_ag]
    if not rows:
        print(f"[INFO] No prospect rows for agency={agency}, skip manager mail.")
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

def send_admin_pack(date: str, prospects: List[Dict[str, Any]]):
    admin = ROUTING.get("admin") or {}
    to_email = clean_email(admin.get("email") or "")
    if not is_valid_email(to_email):
        print(f"[WARN] No valid admin email configured ({to_email})")
        return

    rows = [unify_from_prospect(p) for p in prospects]
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
    print(f"🔑 Gemini={'ON' if GEMINI_API_KEY else 'OFF'} | Places={'ON' if GOOGLE_PLACES_API_KEY else 'OFF'} | DEBUG={'ON' if DEBUG else 'OFF'}")

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