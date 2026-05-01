"""
export_and_mail.py — v2.0
==========================
Script GitHub Actions : export et envoi mail des fiches de prospection.

Flux :
  1. Dump KV via /dump (NDJSON) → prospects, closes, photos, cards
  2. OCR hybride sur images (Gemini + EasyOCR + Tesseract)
  3. Enrichissement (API Gouv + Google Places + scraping email)
  4. Déduplication avancée multi-critères
  5. Construction Excel via excel_export.py (xlsx consultation + xls CRM)
  6. Envoi email via Brevo selon le mode :
       individual       → commercial reçoit ses fiches
       agency_manager   → responsable reçoit toutes les fiches de l'agence
       admin            → manager général reçoit toutes les agences

Variables d'environnement requises :
  WORKER_BASE_URL, EXPORT_TOKEN, TELEGRAM_TOKEN
  BREVO_API_KEY, BREVO_SENDER_EMAIL, BREVO_SENDER_NAME
  MAIL_ROUTING_JSON (optionnel, sinon routing par défaut)
  GOOGLE_PLACES_API_KEY (optionnel)
  GEMINI_API_KEY (optionnel)
  SEND_MODE   : individual | agency_manager | admin
  RUN_DATE    : YYYY-MM-DD (optionnel, défaut = aujourd'hui Paris)
  AGENCY      : GR|VR|GRS|SLS (requis si mode=individual ou agency_manager)
  INITIALS    : ex CZ (requis si mode=individual)
"""

import os
import re
import json
import io
import base64
import hashlib
import datetime
import zoneinfo
import logging
from typing import Any, Dict, List, Optional, Tuple

import requests
from PIL import Image
import pytesseract
import numpy as np
import cv2
import easyocr
from openpyxl import Workbook

# Import module Excel (même répertoire)
from excel_export import (
    export_commercial,
    export_agency_manager,
    export_admin,
    build_email_attachments,
)

# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


# ============================================================
# ENV / CONFIG
# ============================================================

def env_str(name: str, default: str = "") -> str:
    return str(os.environ.get(name, default) or default).strip()


def env_int(name: str, default: int) -> int:
    try:
        return int(env_str(name, str(default)))
    except (ValueError, TypeError):
        return default


def paris_today() -> str:
    """Date du jour en timezone Europe/Paris — corrige le bug UTC de l'ébauche."""
    try:
        tz = zoneinfo.ZoneInfo("Europe/Paris")
        return datetime.datetime.now(tz).strftime("%Y-%m-%d")
    except Exception:
        return datetime.date.today().strftime("%Y-%m-%d")


WORKER_BASE_URL   = env_str("WORKER_BASE_URL")
EXPORT_TOKEN      = env_str("EXPORT_TOKEN")
TELEGRAM_TOKEN    = env_str("TELEGRAM_TOKEN")

BREVO_API_KEY     = env_str("BREVO_API_KEY")
BREVO_SENDER_EMAIL = env_str("BREVO_SENDER_EMAIL", "no-reply@example.com")
BREVO_SENDER_NAME  = env_str("BREVO_SENDER_NAME", "Prospection Bot")

MAIL_ROUTING_JSON = env_str("MAIL_ROUTING_JSON", "")
GOOGLE_PLACES_API_KEY = env_str("GOOGLE_PLACES_API_KEY", "")
GEMINI_API_KEY    = env_str("GEMINI_API_KEY", "")

SEND_MODE  = env_str("SEND_MODE", "individual").lower()
RUN_DATE   = env_str("RUN_DATE", paris_today())
AGENCY     = env_str("AGENCY", "").upper()
INITIALS   = env_str("INITIALS", "").upper()

MAX_OCR_IMAGES   = env_int("MAX_OCR_IMAGES", 50)
MAX_PHOTO_IMAGES = env_int("MAX_PHOTO_IMAGES", 15)

OUT_DIR = env_str("OUT_DIR", "out") or "out"
if os.path.exists(OUT_DIR) and not os.path.isdir(OUT_DIR):
    logger.warning("OUT_DIR='%s' existe mais n'est pas un dossier → fallback 'exports'", OUT_DIR)
    OUT_DIR = "exports"
os.makedirs(OUT_DIR, exist_ok=True)

VALID_AGENCIES = {"GR", "VR", "GRS", "SLS"}

# Regex
EMAIL_RE         = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EMAIL_IN_TEXT_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE         = re.compile(r"(?:(?:\+33|0)[\s\.\-]*[1-9](?:[\s\.\-]*\d{2}){4})")
URL_RE           = re.compile(r"(https?://[^\s)]+|www\.[^\s)]+)", re.I)
CP_RE            = re.compile(r"\b((?:0[1-9]|[1-8]\d|9[0-5])\s?\d{3}|97\d{3}|98\d{3})\b")

DISPOSABLE_DOMAINS = {"tempmail.com", "guerrillamail.com", "yopmail.com", "10minutemail.com"}
FREE_EMAIL_DOMAINS = {
    "gmail.com", "outlook.com", "hotmail.com", "yahoo.com", "yahoo.fr",
    "icloud.com", "free.fr", "orange.fr", "laposte.net", "live.fr",
}

# Qualification
QUALIF_LABELS = {"chaud": "🔥 Chaud", "tiede": "📅 Tiède", "froid": "❄️ Froid"}


# ============================================================
# ROUTING
# ============================================================

DEFAULT_ROUTING = {
    "agencies": {
        "GR":  {"manager": {"initials": "JL", "email": "jennifer.laurens@ras-interim.fr"},
                "commercial": {"initials": "CZ", "email": "celine.zunarelli@ras-interim.fr"}},
        "VR":  {"manager": {"initials": "JB", "email": "jelena.carrasso@ras-interim.fr"},
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
            logger.warning("MAIL_ROUTING_JSON invalide, fallback routing par défaut : %s", e)
    return DEFAULT_ROUTING


ROUTING = load_routing()


# ============================================================
# STATS
# ============================================================

STATS: Dict[str, int] = {
    "worker_dump_calls": 0,
    "telegram_downloads": 0,
    "gemini_calls": 0,
    "gouv_calls": 0,
    "places_calls": 0,
    "scrape_calls": 0,
    "emails_sent": 0,
    "ocr_easyocr": 0,
    "ocr_tesseract": 0,
    "dedupe_merged": 0,
    "cards_processed": 0,
    "photos_processed": 0,
    "rows_exported": 0,
}


# ============================================================
# CACHES
# ============================================================

GOUV_CACHE:        Dict[str, Any]   = {}
PLACES_CACHE:      Dict[str, Any]   = {}
EMAIL_CACHE:       Dict[str, str]   = {}
GEMINI_CARD_CACHE: Dict[str, Any]   = {}
GEMINI_FAC_CACHE:  Dict[str, Any]   = {}
TG_PATH_CACHE:     Dict[str, str]   = {}
TG_BYTES_CACHE:    Dict[str, bytes] = {}
OCR_TEXT_CACHE:    Dict[str, str]   = {}


# ============================================================
# HELPERS GÉNÉRAUX
# ============================================================

def safe_strip(v: Any) -> str:
    return re.sub(r"\s+", " ", str(v or "").strip())


def normalize_spaces(v: str) -> str:
    return re.sub(r"\s+", " ", str(v or "").strip())


def normalize_cp(cp: str) -> str:
    return re.sub(r"\s+", "", str(cp or "").strip())


def sha1_hex(s: str) -> str:
    return hashlib.sha1(str(s or "").encode("utf-8", errors="ignore")).hexdigest()


def image_hash(img_bytes: bytes) -> str:
    return hashlib.sha1(img_bytes or b"").hexdigest()


def clean_email(v: Any) -> str:
    if isinstance(v, dict):
        v = v.get("email", "")
    v = re.sub(r"\s+", "", str(v or "").strip().lower())
    return v


def is_valid_email(s: str) -> bool:
    s = clean_email(s)
    if not s or len(s) > 254:
        return False
    if not EMAIL_RE.match(s):
        return False
    domain = (s.split("@", 1)[1] or "").lower()
    return domain not in DISPOSABLE_DOMAINS


def normalize_email(email: str) -> str:
    e = clean_email(email)
    return e if is_valid_email(e) else ""


def normalize_phone(s: str) -> str:
    s = str(s or "").strip()
    if not s:
        return ""
    s = re.sub(r"[^\d+]", "", s)
    if s.startswith("+33"):
        s = "0" + s[3:]
    if not re.fullmatch(r"0\d{9}", s):
        return ""
    return s


def extract_phones(text: str) -> List[str]:
    out, seen = [], set()
    for m in PHONE_RE.finditer(str(text or "")):
        p = normalize_phone(m.group(0))
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out


def split_phones(phones: List[str]) -> Tuple[str, str]:
    fixed, mobile, other = [], [], []
    seen = set()
    for p in phones:
        pn = normalize_phone(p)
        if not pn or pn in seen:
            continue
        seen.add(pn)
        if pn.startswith(("01", "02", "03", "04", "05")):
            fixed.append(pn)
        elif pn.startswith(("06", "07")):
            mobile.append(pn)
        else:
            other.append(pn)
    phone = fixed[0] if fixed else (mobile[0] if mobile else (other[0] if other else ""))
    phone2 = ""
    if phone:
        candidates = (fixed + mobile + other)
        remaining = [x for x in candidates if x != phone]
        phone2 = remaining[0] if remaining else ""
    if phone and phone2 and phone == phone2:
        phone2 = ""
    return phone, phone2


def validate_siret(raw: str) -> str:
    s = re.sub(r"\D", "", str(raw or "").strip())
    if len(s) != 14:
        return ""
    total = 0
    for i in range(14):
        d = int(s[13 - i])
        if i % 2 == 1:
            d *= 2
        if d > 9:
            d -= 9
        total += d
    return s if total % 10 == 0 else ""


def normalize_naf(naf: str) -> str:
    return str(naf or "").strip().upper().replace(".", "").replace(" ", "")


def normalize_for_match(s: str) -> str:
    s = str(s or "").lower()
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a")
    s = s.replace("â", "a").replace("î", "i").replace("ô", "o").replace("ù", "u")
    s = s.replace("û", "u").replace("ç", "c")
    s = re.sub(r"\b(sas|sasu|sarl|eurl|sa|sci|snc|societe|groupe|compagnie|cie)\b", " ", s)
    return re.sub(r"[^a-z0-9]", "", s).strip()


def split_human_name(full: str) -> Tuple[str, str]:
    s = normalize_spaces(full)
    if not s:
        return "", ""
    parts = s.split(" ")
    if len(parts) == 1:
        return "", parts[0]
    return parts[0], " ".join(parts[1:])


def ensure_contact_fallback(d: Dict[str, Any]) -> Dict[str, Any]:
    fn = str(d.get("contact_firstname") or "").strip()
    ln = str(d.get("contact_lastname") or "").strip()
    interlocuteur = str(d.get("interlocuteur") or "").strip()
    dirigeant = str(d.get("dirigeant") or "").strip()
    if interlocuteur:
        f, l = split_human_name(interlocuteur)
        if not fn and f:
            d["contact_firstname"] = f
        if not ln and l:
            d["contact_lastname"] = l
    fn = str(d.get("contact_firstname") or "").strip()
    ln = str(d.get("contact_lastname") or "").strip()
    if not fn and not ln and dirigeant:
        f, l = split_human_name(dirigeant)
        d["contact_firstname"] = f or ""
        d["contact_lastname"] = l or dirigeant
    if not str(d.get("interlocuteur") or "").strip():
        combo = f"{str(d.get('contact_firstname') or '').strip()} {str(d.get('contact_lastname') or '').strip()}".strip()
        if combo:
            d["interlocuteur"] = combo
        elif dirigeant:
            d["interlocuteur"] = dirigeant
    return d


def strip_postal_city(addr: str, postal: str, city: str) -> str:
    a = normalize_spaces(addr)
    if not a:
        return ""
    cp = normalize_spaces(postal)
    ct = normalize_spaces(city)
    if cp and ct:
        tail = f"{cp} {ct}".strip()
        a = re.sub(rf"(?:\s|-)?{re.escape(tail)}\s*$", "", a, flags=re.IGNORECASE).strip()
    if cp:
        a = re.sub(rf"(?:\s|-)?{re.escape(cp)}\s*$", "", a).strip()
    return normalize_spaces(a)


def uppercase_fields(row: Dict[str, Any]) -> Dict[str, Any]:
    for k in ["name", "address", "city", "dirigeant", "interlocuteur"]:
        row[k] = normalize_spaces(str(row.get(k, "") or "")).upper()
    return row


def email_for_initials(initials: str) -> Optional[str]:
    ini = (initials or "").upper().strip()
    if not ini:
        return None
    users = ROUTING.get("users") or {}
    if ini in users:
        em = clean_email(users[ini])
        return em if is_valid_email(em) else None
    agencies = ROUTING.get("agencies") or {}
    for cfg in agencies.values():
        for role in ("manager", "commercial"):
            r = (cfg or {}).get(role) or {}
            if (r.get("initials") or "").upper() == ini:
                em = clean_email(r.get("email") or "")
                return em if is_valid_email(em) else None
    return None


# ============================================================
# EASYOCR
# ============================================================

_EASYOCR_READER = None


def get_easyocr():
    global _EASYOCR_READER
    if _EASYOCR_READER is None:
        try:
            _EASYOCR_READER = easyocr.Reader(["fr", "en"], gpu=False)
            logger.info("EasyOCR initialisé")
        except Exception as e:
            logger.warning("EasyOCR indisponible : %s", e)
            _EASYOCR_READER = False
    return _EASYOCR_READER


# ============================================================
# HTTP — WORKER DUMP
# ============================================================

def worker_dump(kind: str, date: str) -> List[Dict[str, Any]]:
    """Récupère les enregistrements KV depuis le Worker (format NDJSON)."""
    if not WORKER_BASE_URL or not EXPORT_TOKEN:
        raise RuntimeError("WORKER_BASE_URL / EXPORT_TOKEN manquants")
    STATS["worker_dump_calls"] += 1
    url = f"{WORKER_BASE_URL.rstrip('/')}/dump?date={date}&kind={kind}"
    r = requests.get(
        url,
        headers={"X-Export-Token": EXPORT_TOKEN},
        timeout=60,
    )
    if r.status_code != 200:
        raise RuntimeError(f"/dump {kind} → {r.status_code}: {r.text[:300]}")
    out: List[Dict[str, Any]] = []
    for line in r.text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            out.append(json.loads(line))
        except Exception:
            pass
    logger.info("dump %s : %d enregistrements", kind, len(out))
    return out


# ============================================================
# TELEGRAM — TÉLÉCHARGEMENT IMAGES
# ============================================================

def tg_get_file_path(file_id: str) -> Optional[str]:
    if not file_id:
        return None
    if file_id in TG_PATH_CACHE:
        return TG_PATH_CACHE[file_id]
    if not TELEGRAM_TOKEN:
        raise RuntimeError("TELEGRAM_TOKEN manquant")
    r = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile",
        json={"file_id": file_id},
        timeout=20,
    )
    r.raise_for_status()
    j = r.json()
    if not j.get("ok"):
        return None
    fp = j["result"]["file_path"]
    TG_PATH_CACHE[file_id] = fp
    return fp


def tg_download(file_id: str) -> bytes:
    if not file_id:
        return b""
    if file_id in TG_BYTES_CACHE:
        return TG_BYTES_CACHE[file_id]
    fp = tg_get_file_path(file_id)
    if not fp:
        return b""
    STATS["telegram_downloads"] += 1
    r = requests.get(
        f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{fp}",
        timeout=60,
    )
    r.raise_for_status()
    TG_BYTES_CACHE[file_id] = r.content
    return r.content


# ============================================================
# OCR HYBRIDE — PREPROCESSING + TESSERACT + EASYOCR
# ============================================================

def pil_to_cv(img: Image.Image) -> np.ndarray:
    arr = np.array(img)
    if len(arr.shape) == 2:
        return arr
    return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)


def preprocess_variants(img_bytes: bytes) -> List[Image.Image]:
    variants = []
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    except Exception:
        return variants
    variants.append(img)
    try:
        variants.append(img.resize((img.width * 2, img.height * 2)))
    except Exception:
        pass
    try:
        gray = np.array(img.convert("L"))
        variants.append(Image.fromarray(cv2.equalizeHist(gray)))
    except Exception:
        pass
    try:
        gray = np.array(img.convert("L"))
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        variants.append(Image.fromarray(th))
    except Exception:
        pass
    try:
        gray = np.array(img.convert("L"))
        ad = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 15)
        variants.append(Image.fromarray(ad))
    except Exception:
        pass
    return variants


def ocr_easyocr(img: Image.Image) -> str:
    reader = get_easyocr()
    if not reader:
        return ""
    try:
        STATS["ocr_easyocr"] += 1
        results = reader.readtext(pil_to_cv(img), detail=0, paragraph=False)
        return "\n".join(str(x).strip() for x in results if str(x).strip())
    except Exception:
        return ""


def ocr_tesseract(img: Image.Image) -> str:
    try:
        STATS["ocr_tesseract"] += 1
        t1 = pytesseract.image_to_string(img, lang="fra+eng", config="--psm 6") or ""
        t2 = pytesseract.image_to_string(img, lang="fra+eng", config="--psm 11") or ""
        return "\n".join(t for t in [t1, t2] if t.strip())
    except Exception:
        return ""


def merge_text_blocks(texts: List[str]) -> str:
    lines, seen = [], set()
    for txt in texts:
        for ln in str(txt or "").splitlines():
            ln = re.sub(r"\s+", " ", ln).strip()
            if not ln:
                continue
            k = ln.lower()
            if k not in seen:
                seen.add(k)
                lines.append(ln)
    return "\n".join(lines)


def ocr_image(img_bytes: bytes, light: bool = False) -> str:
    if not img_bytes:
        return ""
    key = f"ocr:{image_hash(img_bytes)}:{'light' if light else 'full'}"
    if key in OCR_TEXT_CACHE:
        return OCR_TEXT_CACHE[key]
    variants = preprocess_variants(img_bytes)
    base = variants[:2] if light else variants
    texts = []
    for img in base:
        t1 = ocr_easyocr(img)
        if t1:
            texts.append(t1)
        t2 = ocr_tesseract(img)
        if t2:
            texts.append(t2)
    merged = merge_text_blocks(texts)
    OCR_TEXT_CACHE[key] = merged
    return merged


# ============================================================
# GEMINI OCR
# ============================================================

def _guess_mime(img_bytes: bytes) -> str:
    try:
        img = Image.open(io.BytesIO(img_bytes))
        fmt = (img.format or "").upper()
        if fmt == "PNG":
            return "image/png"
        if fmt in ("JPG", "JPEG"):
            return "image/jpeg"
    except Exception:
        pass
    return "image/jpeg"


def gemini_vision_json(img_bytes: bytes, prompt: str, max_tokens: int = 700) -> Dict[str, Any]:
    if not GEMINI_API_KEY or not img_bytes:
        return {}
    endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
    mime = _guess_mime(img_bytes)
    b64  = base64.b64encode(img_bytes).decode("ascii")
    payload = {
        "contents": [{"parts": [{"text": prompt}, {"inlineData": {"mimeType": mime, "data": b64}}]}],
        "generationConfig": {"temperature": 0.05, "maxOutputTokens": max_tokens},
    }
    try:
        STATS["gemini_calls"] += 1
        r = requests.post(
            endpoint,
            params={"key": GEMINI_API_KEY},
            json=payload,
            timeout=60,
        )
        if r.status_code >= 300:
            logger.warning("Gemini HTTP %d", r.status_code)
            return {}
        j = r.json()
        parts = ((j.get("candidates") or [{}])[0].get("content") or {}).get("parts") or []
        raw = "".join(p.get("text", "") for p in parts).strip()
        raw = re.sub(r"^```json", "", raw, flags=re.I).strip()
        raw = re.sub(r"```$", "", raw).strip()
        m = re.search(r"\{.*\}", raw, re.S)
        if not m:
            return {}
        return json.loads(m.group(0))
    except Exception as e:
        logger.warning("Gemini vision error : %s", e)
        return {}


# Liste noire réponses Gemini invalides
_GEMINI_FAILURES = {"je ne peux pas", "impossible", "illisible", "inconnu", "n/a",
                    "not available", "cannot", "unable", "sorry", "désolé"}


def sanitize_gemini(d: Dict[str, Any]) -> Dict[str, Any]:
    out = {}
    for k, v in (d or {}).items():
        s = str(v or "").lower().strip()
        out[k] = "" if any(f in s for f in _GEMINI_FAILURES) else v
    return out


def gemini_extract_card(img_bytes: bytes) -> Dict[str, Any]:
    key = image_hash(img_bytes)
    if key in GEMINI_CARD_CACHE:
        return GEMINI_CARD_CACHE[key]
    prompt = (
        "Tu es un extracteur de carte de visite. Lis l'IMAGE et retourne STRICTEMENT un JSON : "
        "name, company, title, contact_civility (M. ou Mme sinon vide), email, phone, phone2, "
        "website, postal_code, city, address, confidence (low/medium/high). "
        "Vide si inconnu. Ne pas inventer."
    )
    d = sanitize_gemini(gemini_vision_json(img_bytes, prompt, 700))
    GEMINI_CARD_CACHE[key] = d
    return d


def gemini_extract_facade(img_bytes: bytes) -> Dict[str, Any]:
    key = image_hash(img_bytes)
    if key in GEMINI_FAC_CACHE:
        return GEMINI_FAC_CACHE[key]
    prompt = (
        "Tu analyses une photo de façade/enseigne. Retourne STRICTEMENT un JSON : "
        "company (nom exact sur l'enseigne), activity (secteur max 60 car.), "
        "phone, website, city, confidence (low/medium/high). "
        "Si plusieurs enseignes, prendre la plus grande. Ne pas inventer."
    )
    d = sanitize_gemini(gemini_vision_json(img_bytes, prompt, 400))
    GEMINI_FAC_CACHE[key] = d
    return d


def gemini_structure_text(raw_text: str, mode: str = "card") -> Dict[str, Any]:
    """Structuration par Gemini sur texte brut (pas image = 10x moins cher)."""
    if not GEMINI_API_KEY or not raw_text:
        return {}
    if mode == "card":
        prompt = (
            f"Texte OCR d'une carte de visite:\n\n{raw_text}\n\n"
            "Retourne STRICTEMENT un JSON: name, company, title, contact_civility, "
            "email, phone, phone2, website, postal_code, city, address, confidence (low/medium/high). "
            "Vide si inconnu. Ne pas inventer."
        )
    else:
        prompt = (
            f"Texte OCR d'une enseigne/façade:\n\n{raw_text}\n\n"
            "Retourne STRICTEMENT un JSON: company, activity, phone, website, city, "
            "confidence (low/medium/high). Vide si inconnu."
        )
    endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.05, "maxOutputTokens": 500},
    }
    try:
        STATS["gemini_calls"] += 1
        r = requests.post(endpoint, params={"key": GEMINI_API_KEY}, json=payload, timeout=30)
        if r.status_code >= 300:
            return {}
        j = r.json()
        parts = ((j.get("candidates") or [{}])[0].get("content") or {}).get("parts") or []
        raw = "".join(p.get("text", "") for p in parts).strip()
        raw = re.sub(r"^```json", "", raw, flags=re.I).replace("```", "").strip()
        m = re.search(r"\{.*\}", raw, re.S)
        if not m:
            return {}
        return sanitize_gemini(json.loads(m.group(0)))
    except Exception:
        return {}


def merge_ocr_results(gemini: Dict, vision_structured: Dict, fallbacks: Dict) -> Dict[str, Any]:
    """Fusionne résultats Gemini image + Gemini texte (Vision) avec score confiance."""
    fields = ["name", "company", "title", "contact_civility",
              "email", "phone", "phone2", "website", "city", "address", "postal_code", "activity"]
    result: Dict[str, Any] = {}
    conf:   Dict[str, str] = {}
    for f in fields:
        g = str((gemini or {}).get(f, "") or "").strip()
        v = str((vision_structured or {}).get(f, "") or "").strip()
        if g and v and g.lower() == v.lower():
            result[f] = g
            conf[f] = "high"
        elif g and v:
            result[f] = g if len(g) >= len(v) else v
            conf[f] = "medium"
        else:
            result[f] = g or v or ""
            conf[f] = "medium" if (g or v) else "low"
    if not result.get("phone") and fallbacks.get("phone"):
        result["phone"] = fallbacks["phone"]
        conf["phone"] = "medium"
    if not result.get("email") and fallbacks.get("email"):
        result["email"] = fallbacks["email"]
        conf["email"] = "medium"
    high = sum(1 for c in conf.values() if c == "high")
    filled = sum(1 for v in result.values() if v)
    result["_confidence"] = "high" if high >= 3 else ("medium" if filled >= 3 else "low")
    return result


# ============================================================
# PIPELINE CARTE DE VISITE — DOUBLE PASSE
# ============================================================

def extract_card_fields(img_bytes: bytes) -> Dict[str, Any]:
    """OCR hybride : Gemini image + Tesseract/EasyOCR + Gemini texte."""
    if not img_bytes:
        return {}

    # Passe 1 : Gemini sur l'image
    gemini = gemini_extract_card(img_bytes) if GEMINI_API_KEY else {}
    STATS["cards_processed"] += 1

    # Passe 2 : OCR local (allégé si Gemini a bien fonctionné)
    light = bool(gemini.get("name") or gemini.get("company") or gemini.get("email"))
    ocr_text = ocr_image(img_bytes, light=light)

    # Structuration du texte brut par Gemini (sans image)
    vision_structured = gemini_structure_text(ocr_text, "card") if ocr_text else {}

    # Extraction regex directe (plus fiable que Gemini sur email/tel)
    phones  = extract_phones(ocr_text)
    emails  = [clean_email(m.group(0)) for m in EMAIL_IN_TEXT_RE.finditer(ocr_text)
                if is_valid_email(clean_email(m.group(0)))]
    cps     = CP_RE.findall(ocr_text)
    fallbacks = {
        "phone": phones[0] if phones else "",
        "email": emails[0] if emails else "",
        "postal_code": normalize_cp(cps[0]) if cps else "",
    }

    return merge_ocr_results(gemini, vision_structured, fallbacks)


# ============================================================
# PIPELINE FAÇADE
# ============================================================

def extract_facade_fields(img_bytes: bytes, city_hint: str = "") -> Dict[str, Any]:
    if not img_bytes:
        return {}
    gemini = gemini_extract_facade(img_bytes) if GEMINI_API_KEY else {}
    STATS["photos_processed"] += 1

    light = bool(gemini.get("company"))
    ocr_text = ocr_image(img_bytes, light=light)
    vision_structured = gemini_structure_text(ocr_text, "facade") if ocr_text and not gemini.get("company") else {}

    phones = extract_phones(ocr_text)
    urls   = URL_RE.findall(ocr_text)
    fallbacks = {
        "phone":   phones[0] if phones else "",
        "website": urls[0] if urls else "",
    }

    merged = merge_ocr_results(gemini, vision_structured, fallbacks)
    if not merged.get("city") and city_hint:
        merged["city"] = city_hint
    return merged


# ============================================================
# ENRICHISSEMENT ENTREPRISE
# ============================================================

def search_gouv(query: str, per_page: int = 5) -> List[Dict[str, Any]]:
    q = normalize_spaces(query)
    if not q:
        return []
    key = sha1_hex(f"gouv:{q}:{per_page}")
    if key in GOUV_CACHE:
        return GOUV_CACHE[key]
    try:
        STATS["gouv_calls"] += 1
        r = requests.get(
            "https://recherche-entreprises.api.gouv.fr/search",
            params={"q": q, "page": 1, "per_page": per_page},
            headers={"accept": "application/json"},
            timeout=15,
        )
        r.raise_for_status()
        results = r.json().get("results") or []
        GOUV_CACHE[key] = results
        return results
    except Exception as e:
        logger.warning("API Gouv error [%s] : %s", q, e)
        return []


def gouv_to_flat(e: Dict) -> Dict[str, Any]:
    siege = e.get("siege") or {}
    dirs_ = e.get("dirigeants") or e.get("representants") or []
    dirigeant = ""
    if dirs_:
        first = dirs_[0]
        if isinstance(first, str):
            dirigeant = first
        elif isinstance(first, dict):
            p = first.get("personne") or first
            dirigeant = normalize_spaces(
                f"{p.get('prenom', '')} {p.get('nom', '') or p.get('nom_usage', '')}".strip()
            )
    naf = normalize_naf(e.get("activite_principale") or e.get("naf") or "")
    tranche = e.get("tranche_effectif_salarie") or siege.get("tranche_effectif_salarie") or ""
    return {
        "name":          normalize_spaces(e.get("nom_raison_sociale") or e.get("denomination") or ""),
        "siret":         validate_siret(siege.get("siret") or ""),
        "naf":           naf,
        "address":       normalize_spaces(siege.get("adresse") or siege.get("libelle_voie") or ""),
        "postal_code":   normalize_cp(siege.get("code_postal") or ""),
        "city":          normalize_spaces(siege.get("libelle_commune") or ""),
        "dirigeant":     dirigeant,
        "effectif":      _effectif_label(tranche),
        "date_creation": str(e.get("date_creation") or ""),
        "capital":       str(e.get("capital") or ""),
    }


_EFFECTIF_MAP = {
    "NN": "Non employeur", "00": "0 salarié", "01": "1-2 salariés", "02": "3-5 salariés",
    "03": "6-9 salariés", "11": "10-19 salariés", "12": "20-49 salariés",
    "21": "50-99 salariés", "22": "100-199 salariés", "31": "200-249 salariés",
    "32": "250-499 salariés", "41": "500-999 salariés",
}


def _effectif_label(code: str) -> str:
    return _EFFECTIF_MAP.get(str(code or ""), str(code or ""))


def places_enrich(name: str, city: str = "") -> Dict[str, str]:
    if not GOOGLE_PLACES_API_KEY or not name:
        return {"phone": "", "website": "", "address": ""}
    key = sha1_hex(f"places:{name}|{city}")
    if key in PLACES_CACHE:
        return PLACES_CACHE[key]
    try:
        STATS["places_calls"] += 1
        r1 = requests.post(
            "https://places.googleapis.com/v1/places:searchText",
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "places.id,places.websiteUri",
            },
            json={"textQuery": f"{name} {city}".strip(), "maxResultCount": 1,
                  "languageCode": "fr", "regionCode": "FR"},
            timeout=15,
        )
        r1.raise_for_status()
        p = (r1.json().get("places") or [None])[0]
        if not p:
            PLACES_CACHE[key] = {"phone": "", "website": "", "address": ""}
            return PLACES_CACHE[key]
        pid     = p.get("id", "")
        website = normalize_spaces(p.get("websiteUri", ""))
        phone = address = ""
        if pid:
            r2 = requests.get(
                f"https://places.googleapis.com/v1/places/{pid}",
                headers={
                    "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                    "X-Goog-FieldMask": "nationalPhoneNumber,internationalPhoneNumber,websiteUri,formattedAddress",
                },
                timeout=15,
            )
            r2.raise_for_status()
            j2      = r2.json()
            phone   = normalize_phone(j2.get("nationalPhoneNumber") or j2.get("internationalPhoneNumber") or "")
            website = normalize_spaces(j2.get("websiteUri") or website)
            address = normalize_spaces(j2.get("formattedAddress") or "")
        result = {"phone": phone, "website": website, "address": address}
        PLACES_CACHE[key] = result
        return result
    except Exception as e:
        logger.warning("Places error [%s] : %s", name, e)
        return {"phone": "", "website": "", "address": ""}


def scrape_email(url: str) -> str:
    if not url:
        return ""
    try:
        from urllib.parse import urlparse
        host = urlparse(url).hostname or ""
        host = host.replace("www.", "")
    except Exception:
        return ""
    if host in EMAIL_CACHE:
        cached = EMAIL_CACHE[host]
        return "" if cached == "none" else cached

    candidates = [f"https://{host}/contact", f"https://{host}"]
    for attempt_url in candidates:
        try:
            STATS["scrape_calls"] += 1
            r = requests.get(
                attempt_url,
                headers={"user-agent": "Mozilla/5.0 (ProspectionBot/2.0)"},
                timeout=8,
                allow_redirects=True,
            )
            if r.status_code >= 300:
                continue
            html = r.text[:150_000]
            matches = [clean_email(m) for m in EMAIL_IN_TEXT_RE.findall(html) if is_valid_email(clean_email(m))]
            matches = [m for m in matches if not re.search(r"no-?reply", m, re.I)]
            if not matches:
                continue
            same = [m for m in matches if host in m]
            result = same[0] if same else matches[0]
            EMAIL_CACHE[host] = result
            return result
        except Exception:
            continue
    EMAIL_CACHE[host] = "none"
    return ""


def enrich_company(name: str, city: str, flat_gouv: Optional[Dict] = None) -> Dict[str, Any]:
    """Enrichit une entreprise : Places + email scraping."""
    place = places_enrich(name, city)
    website = place.get("website") or (flat_gouv or {}).get("website") or ""
    email   = scrape_email(website) if website else ""
    base = dict(flat_gouv or {})
    base.update({
        "website": website,
        "phone":   place.get("phone") or base.get("phone") or "",
        "email":   email,
        "address": strip_postal_city(
            place.get("address") or base.get("address") or "",
            base.get("postal_code") or "",
            base.get("city") or city,
        ),
    })
    return base


# ============================================================
# SCORE DE COMPLÉTUDE
# ============================================================

def record_score(d: Dict[str, Any]) -> int:
    weights = {
        "name": 5, "address": 3, "postal_code": 2, "city": 2,
        "siret": 8, "naf": 2, "activity_summary": 2,
        "dirigeant": 2, "interlocuteur": 4,
        "phone": 3, "phone2": 2, "email": 4, "website": 2,
        "resume": 3, "commande": 3,
        "qualification": 2, "effectif": 1, "date_creation": 1,
    }
    score = 0
    for k, w in weights.items():
        v = str(d.get(k) or "").strip()
        if not v:
            continue
        if k == "siret":
            score += w if validate_siret(v) else 0
        elif k == "email":
            score += w if is_valid_email(v) else 0
        elif k in ("phone", "phone2"):
            score += w if normalize_phone(v) else 0
        else:
            score += w
    return score


# ============================================================
# UNIFY — construction d'une ligne prospect normalisée
# ============================================================

def get_first(d: Dict, *keys: str) -> str:
    for k in keys:
        v = d.get(k)
        if v and str(v).strip():
            return str(v).strip()
    return ""


def unify_prospect(p: Dict[str, Any]) -> Dict[str, Any]:
    """Construit une ligne normalisée depuis un enregistrement prospect KV."""
    name    = get_first(p, "name", "company", "nom")
    naf     = normalize_naf(get_first(p, "naf", "activite_principale"))
    website = get_first(p, "website", "site_web")

    # Activité : libellé Gouv ou activity_summary du prospect
    activity = get_first(p, "activity_summary", "secteur_activite", "activite")

    phones = []
    for k in ["phone", "phone2", "telephone", "mobile"]:
        pn = normalize_phone(str(p.get(k) or ""))
        if pn and pn not in phones:
            phones.append(pn)
    phone, phone2 = split_phones(phones)

    raw_email = get_first(p, "email", "mail")
    email = normalize_email(raw_email)

    d = {
        "date":              get_first(p, "date"),
        "agency":            (get_first(p, "agency") or "").upper(),
        "initials":          (get_first(p, "initials", "user_initials") or "").upper(),
        "name":              name,
        "address":           strip_postal_city(
                                 get_first(p, "address", "adresse"),
                                 get_first(p, "postal_code", "code_postal"),
                                 get_first(p, "city", "ville"),
                             ),
        "postal_code":       normalize_cp(get_first(p, "postal_code", "code_postal")),
        "city":              normalize_spaces(get_first(p, "city", "ville")),
        "siret":             validate_siret(get_first(p, "siret")),
        "naf":               naf,
        "activity_summary":  activity,
        "effectif":          get_first(p, "effectif"),
        "date_creation":     get_first(p, "date_creation"),
        "capital":           get_first(p, "capital"),
        "dirigeant":         normalize_spaces(get_first(p, "dirigeant")),
        "interlocuteur":     normalize_spaces(get_first(p, "interlocuteur", "contact_name")),
        "contact_civility":  get_first(p, "contact_civility"),
        "contact_firstname": normalize_spaces(get_first(p, "contact_firstname", "firstname")),
        "contact_lastname":  normalize_spaces(get_first(p, "contact_lastname", "lastname")),
        "phone":             phone,
        "phone2":            phone2,
        "email":             email,
        "website":           website,
        "resume":            get_first(p, "resume", "meeting", "comment"),
        "commande":          get_first(p, "commande", "order"),
        "qualification":     get_first(p, "qualification"),
        "card_photo_url":    get_first(p, "card_photo_url"),
        "card_photo_file_id": get_first(p, "card_photo_file_id"),
        "facade_photo_url":  get_first(p, "facade_photo_url"),
        "facade_photo_file_id": get_first(p, "facade_photo_file_id"),
        "_confidence":       int(p.get("_confidence") or record_score(p)),
        "_source_type":      "prospect",
    }

    if d["phone"] and d["phone2"] and d["phone"] == d["phone2"]:
        d["phone2"] = ""

    d = ensure_contact_fallback(d)
    return uppercase_fields(d)


def unify_from_card(card: Dict[str, Any], date: str, img_bytes: bytes) -> Optional[Dict[str, Any]]:
    """Construit une ligne depuis un lot carte."""
    if not img_bytes:
        return None
    fields = extract_card_fields(img_bytes)
    if not fields:
        return None
    agency   = (card.get("agency") or "").upper()
    initials = (card.get("user") or card.get("initials") or "").upper()
    comment  = str(card.get("comment") or "").strip()
    company  = normalize_spaces(fields.get("company") or "")
    city     = normalize_spaces(fields.get("city") or "")
    # Enrichissement Gouv + Places
    gouv_results = search_gouv(f"{company} {city}".strip(), 1) if company else []
    flat_gouv = gouv_to_flat(gouv_results[0]) if gouv_results else {}
    enriched  = enrich_company(company or flat_gouv.get("name", ""), city or flat_gouv.get("city", ""), flat_gouv)
    phones = extract_phones(
        " ".join(filter(None, [fields.get("phone"), fields.get("phone2"), enriched.get("phone")]))
    )
    phone, phone2 = split_phones(phones)
    d = {
        "date":              date,
        "agency":            agency,
        "initials":          initials,
        "name":              normalize_spaces(company or enriched.get("name") or ""),
        "address":           strip_postal_city(
                                 fields.get("address") or enriched.get("address") or "",
                                 fields.get("postal_code") or enriched.get("postal_code") or "",
                                 city or enriched.get("city") or "",
                             ),
        "postal_code":       normalize_cp(fields.get("postal_code") or enriched.get("postal_code") or ""),
        "city":              normalize_spaces(city or enriched.get("city") or ""),
        "siret":             validate_siret(enriched.get("siret") or ""),
        "naf":               normalize_naf(enriched.get("naf") or ""),
        "activity_summary":  normalize_spaces(enriched.get("activity_summary") or ""),
        "effectif":          enriched.get("effectif") or "",
        "date_creation":     enriched.get("date_creation") or "",
        "capital":           enriched.get("capital") or "",
        "dirigeant":         normalize_spaces(enriched.get("dirigeant") or ""),
        "interlocuteur":     normalize_spaces(fields.get("name") or ""),
        "contact_civility":  fields.get("contact_civility") or "",
        "contact_firstname": "",
        "contact_lastname":  "",
        "phone":             phone,
        "phone2":            phone2,
        "email":             normalize_email(fields.get("email") or enriched.get("email") or ""),
        "website":           fields.get("website") or enriched.get("website") or "",
        "resume":            comment,
        "commande":          "",
        "qualification":     "",
        "card_photo_url":    "",
        "card_photo_file_id": card.get("file_id") or "",
        "facade_photo_url":  "",
        "facade_photo_file_id": "",
        "_confidence":       0,
        "_source_type":      "lot_card",
    }
    d = ensure_contact_fallback(d)
    d = uppercase_fields(d)
    d["_confidence"] = record_score(d)
    if d["_confidence"] < 30:
        logger.info("Carte ignorée (score %d < 30) : %s", d["_confidence"], d["name"])
        return None
    return d


def unify_from_photo(photo: Dict[str, Any], date: str, img_bytes: bytes) -> Optional[Dict[str, Any]]:
    """Construit une ligne depuis un lot photo façade."""
    if not img_bytes:
        return None
    city_hint = normalize_spaces(photo.get("city") or "")
    fields    = extract_facade_fields(img_bytes, city_hint)
    if not fields:
        return None
    agency   = (photo.get("agency") or "").upper()
    initials = (photo.get("user") or photo.get("initials") or "").upper()
    comment  = str(photo.get("comment") or photo.get("meeting") or "").strip()
    company  = normalize_spaces(fields.get("company") or "")
    city     = normalize_spaces(fields.get("city") or city_hint)
    if not company:
        return None
    gouv_results = search_gouv(f"{company} {city}".strip(), 1) if company else []
    flat_gouv = gouv_to_flat(gouv_results[0]) if gouv_results else {}
    enriched  = enrich_company(company or flat_gouv.get("name", ""), city or flat_gouv.get("city", ""), flat_gouv)
    phone, phone2 = split_phones([
        fields.get("phone") or "", enriched.get("phone") or "",
    ])
    d = {
        "date":              date,
        "agency":            agency,
        "initials":          initials,
        "name":              normalize_spaces(company or enriched.get("name") or ""),
        "address":           strip_postal_city(
                                 enriched.get("address") or "",
                                 enriched.get("postal_code") or "",
                                 city or enriched.get("city") or "",
                             ),
        "postal_code":       normalize_cp(enriched.get("postal_code") or ""),
        "city":              normalize_spaces(city or enriched.get("city") or ""),
        "siret":             validate_siret(enriched.get("siret") or ""),
        "naf":               normalize_naf(enriched.get("naf") or ""),
        "activity_summary":  normalize_spaces(fields.get("activity") or enriched.get("activity_summary") or ""),
        "effectif":          enriched.get("effectif") or "",
        "date_creation":     enriched.get("date_creation") or "",
        "capital":           enriched.get("capital") or "",
        "dirigeant":         normalize_spaces(enriched.get("dirigeant") or ""),
        "interlocuteur":     "",
        "contact_civility":  "",
        "contact_firstname": "",
        "contact_lastname":  "",
        "phone":             phone,
        "phone2":            phone2,
        "email":             normalize_email(enriched.get("email") or ""),
        "website":           fields.get("website") or enriched.get("website") or "",
        "resume":            comment,
        "commande":          "",
        "qualification":     "",
        "card_photo_url":    "",
        "card_photo_file_id": "",
        "facade_photo_url":  "",
        "facade_photo_file_id": photo.get("file_id") or "",
        "_confidence":       0,
        "_source_type":      "lot_photo",
    }
    d = ensure_contact_fallback(d)
    d = uppercase_fields(d)
    d["_confidence"] = record_score(d)
    if d["_confidence"] < 32:
        logger.info("Photo ignorée (score %d < 32) : %s", d["_confidence"], d["name"])
        return None
    return d


# ============================================================
# DÉDUPLICATION
# ============================================================

def _dup_score(target: Dict, existing: Dict) -> int:
    score = 0
    ts = validate_siret(target.get("siret") or "")
    es = validate_siret(existing.get("siret") or "")
    if ts and es and ts == es:
        score += 100
    tn = normalize_for_match(target.get("name") or "")
    en = normalize_for_match(existing.get("name") or "")
    if tn and en and tn == en:
        score += 40
    elif tn and en and (tn in en or en in tn):
        score += 25
    tc = normalize_for_match(target.get("city") or "")
    ec = normalize_for_match(existing.get("city") or "")
    if tc and ec and tc == ec:
        score += 15
    tp, ep = normalize_phone(target.get("phone") or ""), normalize_phone(existing.get("phone") or "")
    if tp and ep and tp == ep:
        score += 16
    tm, em = normalize_email(target.get("email") or ""), normalize_email(existing.get("email") or "")
    if tm and em and tm == em:
        score += 24
    return score


def _dedup_key(r: Dict) -> str:
    siret = validate_siret(r.get("siret") or "")
    if siret:
        return f"SIRET|{siret}"
    name = normalize_for_match(r.get("name") or "")
    city = normalize_for_match(r.get("city") or "")
    cp   = normalize_cp(r.get("postal_code") or "")
    if name and city and cp:
        return f"NCP|{name}|{city}|{cp}"
    if name and city:
        return f"NC|{name}|{city}"
    email = normalize_email(r.get("email") or "")
    if email:
        return f"EMAIL|{email}"
    return ""


def _merge_rows(primary: Dict, secondary: Dict) -> Dict:
    merged = dict(primary)
    protected = {"resume", "commande", "qualification"}
    for k, v2 in secondary.items():
        if k.startswith("_"):
            continue
        v1 = str(merged.get(k) or "").strip()
        v2s = str(v2 or "").strip()
        if k in protected and v1:
            continue
        if not v1 and v2s:
            merged[k] = v2
        elif v1 and v2s:
            if k == "email" and is_valid_email(v2s) and not is_valid_email(v1):
                merged[k] = v2
            elif k == "siret" and validate_siret(v2s) and not validate_siret(v1):
                merged[k] = v2
            elif k in ("address", "interlocuteur") and len(v2s) > len(v1):
                merged[k] = v2
    merged["_confidence"] = record_score(merged)
    return ensure_contact_fallback(merged)


def deduplicate(rows: List[Dict]) -> List[Dict]:
    if not rows:
        return []
    buckets: Dict[str, List[Dict]] = {}
    loose:   List[Dict] = []
    for r in rows:
        key = _dedup_key(r)
        if key:
            buckets.setdefault(key, []).append(r)
        else:
            loose.append(r)
    merged_groups: List[Dict] = []
    for group in buckets.values():
        consumed = [False] * len(group)
        for i in range(len(group)):
            if consumed[i]:
                continue
            base = dict(group[i])
            consumed[i] = True
            for j in range(i + 1, len(group)):
                if consumed[j]:
                    continue
                if _dup_score(base, group[j]) >= 28:
                    base = _merge_rows(base, group[j])
                    consumed[j] = True
                    STATS["dedupe_merged"] += 1
            merged_groups.append(base)
    out = list(merged_groups)
    for r in loose:
        attached = False
        for idx, ex in enumerate(out):
            if _dup_score(r, ex) >= 34:
                out[idx] = _merge_rows(ex, r)
                STATS["dedupe_merged"] += 1
                attached = True
                break
        if not attached:
            out.append(r)
    return out


# ============================================================
# CONSOLIDATION
# ============================================================

def consolidate(
    date: str,
    prospects: List[Dict],
    photos: List[Dict],
    cards: List[Dict],
    agency_filter: str = "",
    initials_filter: str = "",
) -> List[Dict]:
    agency_filter   = agency_filter.upper().strip()
    initials_filter = initials_filter.upper().strip()

    def match_agency(r: Dict) -> bool:
        ag = str(r.get("agency") or "").upper()
        return not agency_filter or ag == agency_filter

    def match_initials(r: Dict) -> bool:
        ini = str(r.get("initials") or r.get("user") or "").upper()
        return not initials_filter or ini == initials_filter

    rows: List[Dict] = []

    # 1. Prospects existants enrichis
    for p in prospects:
        if not match_agency(p) or not match_initials(p):
            continue
        rows.append(unify_prospect(p))

    # 2. Lots cartes (standalone — pas liés à un prospect existant)
    prospect_card_ids = {
        str(p.get("card_photo_file_id") or "") for p in prospects
        if p.get("card_photo_file_id")
    }
    standalone_cards = [
        c for c in cards
        if match_agency(c) and match_initials(c)
        and str(c.get("file_id") or "") not in prospect_card_ids
    ][:MAX_OCR_IMAGES]

    for card in standalone_cards:
        fid = card.get("file_id") or ""
        if not fid:
            continue
        try:
            img = tg_download(fid)
        except Exception as e:
            logger.warning("Téléchargement carte %s : %s", fid, e)
            continue
        row = unify_from_card(card, date, img)
        if row:
            rows.append(row)

    # 3. Lots photos façades
    prospect_photo_ids = {
        str(p.get("facade_photo_file_id") or "") for p in prospects
        if p.get("facade_photo_file_id")
    }
    standalone_photos = [
        ph for ph in photos
        if match_agency(ph) and match_initials(ph)
        and str(ph.get("file_id") or "") not in prospect_photo_ids
    ][:MAX_PHOTO_IMAGES]

    for photo in standalone_photos:
        fid = photo.get("file_id") or ""
        if not fid:
            continue
        try:
            img = tg_download(fid)
        except Exception as e:
            logger.warning("Téléchargement photo %s : %s", fid, e)
            continue
        row = unify_from_photo(photo, date, img)
        if row:
            rows.append(row)

    # Nettoyage + déduplication
    rows = [r for r in rows if any(str(v or "").strip() for v in r.values())]
    for r in rows:
        r["date"]    = r.get("date") or date
        r["agency"]  = (r.get("agency") or agency_filter or "").upper()
        r["initials"] = (r.get("initials") or initials_filter or "").upper()

    rows = deduplicate(rows)

    # Tri par score décroissant
    rows.sort(key=lambda r: r.get("_confidence", 0), reverse=True)
    STATS["rows_exported"] += len(rows)
    return rows


# ============================================================
# RÉSUMÉ MANAGER TELEGRAM (si MANAGER_CHAT_ID configuré)
# ============================================================

def build_manager_summary(
    date: str,
    rows_by_agency: Dict[str, List[Dict]],
    close_by_agency: Dict[str, List[Dict]],
) -> str:
    lines = [f"📊 <b>Bilan prospection — {date}</b>", ""]
    total_all = 0
    for agency in sorted(rows_by_agency.keys()):
        rows = rows_by_agency[agency]
        if not rows:
            continue
        by_ini: Dict[str, List] = {}
        for r in rows:
            ini = str(r.get("initials") or "?")
            by_ini.setdefault(ini, []).append(r)
        lines.append(f"🏷️ <b>{agency}</b>")
        total_agency = 0
        for ini, ini_rows in sorted(by_ini.items()):
            hot  = sum(1 for r in ini_rows if r.get("qualification") == "chaud")
            warm = sum(1 for r in ini_rows if r.get("qualification") == "tiede")
            cold = sum(1 for r in ini_rows if r.get("qualification") == "froid")
            total = len(ini_rows)
            total_agency += total
            closed = any(
                str(c.get("initials") or "").upper() == ini
                for c in close_by_agency.get(agency, [])
            )
            close_icon = "✅" if closed else "⚠️ pas de clôture"
            lines.append(
                f"  👤 {ini} : {total} fiche(s) · "
                f"{hot} 🔥 · {warm} 📅 · {cold} ❄️ — {close_icon}"
            )
        lines.append(f"  Total agence : <b>{total_agency}</b>")
        total_all += total_agency
    lines += ["", f"<b>Total global : {total_all} nouveaux prospects</b>",
              "Export CRM : ✅ déclenché"]
    return "\n".join(lines)


def send_telegram_summary(chat_id: str, text: str) -> bool:
    if not TELEGRAM_TOKEN or not chat_id:
        return False
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
            timeout=15,
        )
        return r.status_code == 200
    except Exception as e:
        logger.warning("Envoi résumé Telegram : %s", e)
        return False


# ============================================================
# BREVO EMAIL
# ============================================================

def brevo_send(
    to_email: str,
    subject: str,
    html: str,
    attachments: Optional[List[Tuple[str, bytes]]] = None,
) -> None:
    if not BREVO_API_KEY:
        raise RuntimeError("BREVO_API_KEY manquant")
    to_email = clean_email(to_email)
    if not is_valid_email(to_email):
        raise ValueError(f"Email destinataire invalide : '{to_email}'")
    payload: Dict[str, Any] = {
        "sender":      {"name": BREVO_SENDER_NAME, "email": BREVO_SENDER_EMAIL},
        "to":          [{"email": to_email}],
        "subject":     subject,
        "htmlContent": html,
    }
    if attachments:
        payload["attachment"] = [
            {"name": name, "content": base64.b64encode(content).decode("ascii")}
            for name, content in attachments
        ]
    r = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={
            "accept": "application/json",
            "api-key": BREVO_API_KEY,
            "content-type": "application/json",
        },
        data=json.dumps(payload),
        timeout=60,
    )
    if r.status_code >= 300:
        raise RuntimeError(f"Brevo {r.status_code}: {r.text[:300]}")
    STATS["emails_sent"] += 1
    logger.info("Email envoyé à %s", to_email)


def _email_html(title: str, body_lines: List[str]) -> str:
    body = "".join(f"<p>{line}</p>" for line in body_lines)
    return (
        f"<html><body style='font-family:Arial,sans-serif'>"
        f"<h2>{title}</h2>{body}"
        f"<hr><p style='color:#888;font-size:12px'>Bot Prospection — export automatique</p>"
        f"</body></html>"
    )


# ============================================================
# SEND MODES
# ============================================================

def send_individual_pack(
    date: str,
    agency: str,
    initials: str,
    prospects: List[Dict],
    photos: List[Dict],
    cards: List[Dict],
) -> None:
    to_email = email_for_initials(initials)
    if not to_email:
        logger.warning("Pas d'email pour initiales=%s, skip", initials)
        return

    rows = consolidate(date, prospects, photos, cards,
                       agency_filter=agency, initials_filter=initials)
    rows = [r for r in rows if any(str(v or "").strip() for v in r.values())]
    if not rows:
        logger.info("Aucune ligne pour %s/%s", agency, initials)
        return

    path_xlsx, path_xls = export_commercial(date, agency, initials, rows, OUT_DIR)
    attachments = build_email_attachments(path_xlsx, path_xls)

    qualifs = {
        "chaud": sum(1 for r in rows if r.get("qualification") == "chaud"),
        "tiede": sum(1 for r in rows if r.get("qualification") == "tiede"),
        "froid": sum(1 for r in rows if r.get("qualification") == "froid"),
    }
    html = _email_html(
        f"Prospection {date} — {agency}/{initials}",
        [
            f"Bonjour,",
            f"Voici ton export de prospection du <b>{date}</b> pour <b>{agency}/{initials}</b>.",
            f"<b>{len(rows)}</b> fiche(s) : "
            f"{qualifs['chaud']} 🔥 chaud · {qualifs['tiede']} 📅 tiède · {qualifs['froid']} ❄️ froid",
            "Pièces jointes : <b>Excel consultation</b> (xlsx) + <b>Excel import CRM</b> (xls)",
        ],
    )
    brevo_send(to_email, f"Prospection {date} — {agency}/{initials}", html, attachments)
    logger.info("Pack individuel envoyé → %s (%s/%s, %d fiches)", to_email, agency, initials, len(rows))


def send_agency_pack(
    date: str,
    agency: str,
    prospects: List[Dict],
    photos: List[Dict],
    cards: List[Dict],
) -> List[Dict]:
    agencies_cfg = (ROUTING.get("agencies") or {})
    cfg     = agencies_cfg.get(agency) or {}
    manager = (cfg.get("manager") or {})
    to_email = clean_email(manager.get("email") or "")
    if not is_valid_email(to_email):
        logger.warning("Pas d'email manager pour agence=%s", agency)
        return []

    rows = consolidate(date, prospects, photos, cards, agency_filter=agency)
    rows = [r for r in rows if any(str(v or "").strip() for v in r.values())]
    if not rows:
        logger.info("Aucune ligne pour agence=%s", agency)
        return rows

    path_xlsx, path_xls = export_agency_manager(date, agency, rows, OUT_DIR)
    attachments = build_email_attachments(path_xlsx, path_xls)

    html = _email_html(
        f"Prospection {date} — Agence {agency}",
        [
            "Bonjour,",
            f"Voici le consolidé de prospection du <b>{date}</b> pour l'agence <b>{agency}</b>.",
            f"<b>{len(rows)}</b> fiche(s) au total.",
        ],
    )
    brevo_send(to_email, f"Prospection {date} — Agence {agency}", html, attachments)
    logger.info("Pack agence envoyé → %s (%s, %d fiches)", to_email, agency, len(rows))
    return rows


def send_admin_pack(
    date: str,
    prospects: List[Dict],
    photos: List[Dict],
    cards: List[Dict],
    rows_by_agency: Optional[Dict[str, List[Dict]]] = None,
    close_by_agency: Optional[Dict[str, List[Dict]]] = None,
) -> None:
    admin    = ROUTING.get("admin") or {}
    to_email = clean_email(admin.get("email") or "")
    if not is_valid_email(to_email):
        logger.warning("Pas d'email admin configuré")
        return

    all_rows = consolidate(date, prospects, photos, cards)
    all_rows = [r for r in all_rows if any(str(v or "").strip() for v in r.values())]
    if not all_rows:
        logger.info("Aucune ligne pour admin date=%s", date)
        return

    path_xlsx, path_xls = export_admin(date, all_rows, OUT_DIR)
    attachments = build_email_attachments(path_xlsx, path_xls)

    html = _email_html(
        f"Prospection {date} — ADMIN (toutes agences)",
        [
            "Bonjour,",
            f"Voici le consolidé global de prospection du <b>{date}</b> pour toutes les agences.",
            f"<b>{len(all_rows)}</b> fiche(s) au total.",
            "Pièces jointes : Excel multi-onglets (xlsx) + Excel import CRM (xls)",
        ],
    )
    brevo_send(to_email, f"Prospection {date} — ADMIN", html, attachments)
    logger.info("Pack admin envoyé → %s (%d fiches)", to_email, len(all_rows))

    # Résumé Telegram manager général
    manager_chat_id = env_str("MANAGER_CHAT_ID", "")
    if manager_chat_id and rows_by_agency and close_by_agency is not None:
        summary = build_manager_summary(date, rows_by_agency, close_by_agency)
        if send_telegram_summary(manager_chat_id, summary):
            logger.info("Résumé Telegram envoyé → chat_id %s", manager_chat_id)


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    if not WORKER_BASE_URL or not EXPORT_TOKEN or not TELEGRAM_TOKEN:
        raise RuntimeError("Variables manquantes : WORKER_BASE_URL / EXPORT_TOKEN / TELEGRAM_TOKEN")

    date = RUN_DATE or paris_today()
    logger.info("🚀 export_and_mail.py v2.0 — mode=%s date=%s agency=%s initials=%s",
                SEND_MODE, date, AGENCY, INITIALS)
    logger.info("📦 OUT_DIR=%s | Gemini=%s | Places=%s",
                OUT_DIR, "ON" if GEMINI_API_KEY else "OFF",
                "ON" if GOOGLE_PLACES_API_KEY else "OFF")

    # Dump KV
    prospects = worker_dump("prospects", date)
    photos    = worker_dump("photos",    date)
    cards     = worker_dump("cards",     date)
    closes    = worker_dump("closes",    date)

    # ── Mode individual ───────────────────────────────
    if SEND_MODE == "individual":
        if AGENCY not in VALID_AGENCIES:
            raise RuntimeError(f"AGENCY requis (GR|VR|GRS|SLS) pour mode=individual — reçu : '{AGENCY}'")
        if not INITIALS:
            raise RuntimeError("INITIALS requis pour mode=individual")
        send_individual_pack(date, AGENCY, INITIALS, prospects, photos, cards)
        logger.info("📊 Stats : %s", json.dumps(STATS, ensure_ascii=False))
        return

    # ── Mode agency_manager ───────────────────────────
    if SEND_MODE == "agency_manager":
        agency = AGENCY if AGENCY in VALID_AGENCIES else None
        targets = [agency] if agency else sorted(VALID_AGENCIES)
        for ag in targets:
            send_agency_pack(date, ag, prospects, photos, cards)
        logger.info("📊 Stats : %s", json.dumps(STATS, ensure_ascii=False))
        return

    # ── Mode admin ────────────────────────────────────
    if SEND_MODE == "admin":
        rows_by_agency: Dict[str, List[Dict]] = {}
        close_by_agency: Dict[str, List[Dict]] = {}
        for ag in sorted(VALID_AGENCIES):
            ag_rows = consolidate(date, prospects, photos, cards, agency_filter=ag)
            rows_by_agency[ag] = ag_rows
            close_by_agency[ag] = [c for c in closes if str(c.get("agency") or "").upper() == ag]
        # Envoi pack agence à chaque manager
        for ag in sorted(VALID_AGENCIES):
            send_agency_pack(date, ag, prospects, photos, cards)
        # Envoi pack admin global
        send_admin_pack(date, prospects, photos, cards, rows_by_agency, close_by_agency)
        logger.info("📊 Stats : %s", json.dumps(STATS, ensure_ascii=False))
        return

    raise RuntimeError(f"SEND_MODE inconnu : '{SEND_MODE}' — valeurs : individual | agency_manager | admin")


if __name__ == "__main__":
    main()
