# ============================================================
# BLOC 1 / 6 — IMPORTS / CONFIG / ROUTING / CACHES / HELPERS BASE
# ============================================================

import os
import re
import json
import io
import base64
import zipfile
import hashlib
import datetime as dt
from typing import Dict, List, Any, Optional, Tuple

import requests
from PIL import Image
import pytesseract
import numpy as np
import cv2
import easyocr

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


def paris_ymd_fallback() -> str:
    return dt.date.today().strftime("%Y-%m-%d")


WORKER_BASE_URL = env_str("WORKER_BASE_URL")
EXPORT_TOKEN = env_str("EXPORT_TOKEN")
TELEGRAM_TOKEN = env_str("TELEGRAM_TOKEN")

BREVO_API_KEY = env_str("BREVO_API_KEY")
BREVO_SENDER_EMAIL = env_str("BREVO_SENDER_EMAIL", "no-reply@example.com")
BREVO_SENDER_NAME = env_str("BREVO_SENDER_NAME", "Prospection Bot")

MAIL_ROUTING_JSON = env_str("MAIL_ROUTING_JSON", "")

GOOGLE_PLACES_API_KEY = env_str("GOOGLE_PLACES_API_KEY", "")
GEMINI_API_KEY = env_str("GEMINI_API_KEY", "")

SEND_MODE = env_str("SEND_MODE", "individual").lower()
RUN_DATE = env_str("RUN_DATE", paris_ymd_fallback())
AGENCY = env_str("AGENCY", "").upper()
INITIALS = env_str("INITIALS", "").upper()

MAX_OCR_IMAGES = env_int("MAX_OCR_IMAGES", 50)
MAX_PHOTO_IMAGES = env_int("MAX_PHOTO_IMAGES", 15)

OUT_DIR = env_str("OUT_DIR", "out").strip() or "out"
if os.path.exists(OUT_DIR) and not os.path.isdir(OUT_DIR):
    print(f"[WARN] OUT_DIR='{OUT_DIR}' existe mais n'est pas un dossier. Fallback -> 'exports'")
    OUT_DIR = "exports"
os.makedirs(OUT_DIR, exist_ok=True)

VALID_AGENCIES = {"GR", "VR", "GRS", "SLS"}


# ============================================================
# REGEX / CONSTANTES
# ============================================================

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EMAIL_IN_TEXT_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:(?:\+33|0)[\s\.-]*[1-9](?:[\s\.-]*\d{2}){4})")
URL_RE = re.compile(r"(https?://[^\s)]+|www\.[^\s)]+)", re.I)
CP_RE = re.compile(r"\b((?:0[1-9]|[1-8]\d|9[0-5])\s?\d{3}|97\d{3}|98\d{3})\b")
SIRET_RE = re.compile(r"\b\d{14}\b")
TOKEN_SPLIT_RE = re.compile(r"[,\|;/]+")

DISPOSABLE_DOMAINS = {
    "tempmail.com",
    "guerrillamail.com",
    "yopmail.com",
    "10minutemail.com",
}

FREE_MAIL_DOMAINS = {
    "gmail.com",
    "outlook.com",
    "hotmail.com",
    "yahoo.com",
    "yahoo.fr",
    "icloud.com",
    "free.fr",
    "orange.fr",
    "laposte.net",
    "wanadoo.fr",
    "live.fr",
}

CORPORATE_WORDS = {
    "sas", "sasu", "sarl", "eurl", "sa", "sci", "snc",
    "societe", "groupe", "compagnie", "cie", "holding", "services"
}

COMMON_CITY_HINTS = {
    "grenoble", "voiron", "voreppe", "moirans", "lyon", "valence", "chambery",
    "annecy", "saint-etienne", "bourgoin", "vienne", "tullins", "coublevie",
    "fontaine", "meylan", "echirolles", "sassenage", "saint-egreve"
}

COMMON_ACTIVITY_HINTS = {
    "transport", "interim", "menuiserie", "serrurerie", "resine", "dechets",
    "environnement", "formation", "logistique", "industrie", "btp", "travaux",
    "restauration", "garage", "pneus", "recyclage", "chauffeur", "benne",
    "aquatique", "hotel", "commerce", "collectivite", "maintenance"
}

HEADERS = [
    "date",
    "agency",
    "initials",
    "name",
    "address",
    "postal_code",
    "city",
    "siret",
    "naf",
    "activity_summary",
    "dirigeant",
    "interlocuteur",
    "contact_firstname",
    "contact_lastname",
    "phone",
    "phone2",
    "email",
    "website",
    "resume",
    "commande",
]


# ============================================================
# ROUTING EMAIL
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


def clean_email(v) -> str:
    if v is None:
        return ""
    if isinstance(v, dict):
        v = v.get("email") or ""
    if not isinstance(v, str):
        v = str(v)
    v = v.strip().lower()
    v = re.sub(r"\s+", "", v)
    return v


def is_valid_email(s: str) -> bool:
    s = clean_email(s)
    if not s or len(s) > 254:
        return False
    if not EMAIL_RE.match(s):
        return False
    domain = s.split("@", 1)[1]
    if domain in DISPOSABLE_DOMAINS:
        return False
    return True


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
# RUNTIME STATS / CACHES
# ============================================================

STATS = {
    "worker_dump_calls": 0,
    "telegram_getfile_calls": 0,
    "telegram_file_downloads": 0,
    "gemini_vision_calls": 0,
    "gemini_activity_calls": 0,
    "gouv_calls": 0,
    "places_calls": 0,
    "scrape_calls": 0,
    "emails_sent": 0,
    "ocr_easyocr_calls": 0,
    "ocr_tesseract_calls": 0,
    "prospect_card_links_found": 0,
    "prospect_card_links_missing": 0,
    "dedupe_groups": 0,
    "dedupe_merged": 0,
    "lot_cards_promoted": 0,
    "lot_photos_promoted": 0,
}

GOUV_CACHE: Dict[str, Dict[str, Any]] = {}
PLACES_CACHE: Dict[str, Dict[str, str]] = {}
EMAIL_SCRAPE_CACHE: Dict[str, str] = {}
GEMINI_ACTIVITY_CACHE: Dict[str, str] = {}
GEMINI_CARD_CACHE: Dict[str, Dict[str, str]] = {}
GEMINI_FACADE_CACHE: Dict[str, Dict[str, str]] = {}
TG_FILE_PATH_CACHE: Dict[str, str] = {}
TG_FILE_BYTES_CACHE: Dict[str, bytes] = {}
OCR_TEXT_CACHE: Dict[str, str] = {}
PROSPECT_CARD_ANALYSIS_CACHE: Dict[str, Dict[str, Any]] = {}
SEARCH_MATCH_CACHE: Dict[str, Dict[str, Any]] = {}


# ============================================================
# EASYOCR READER
# ============================================================

EASYOCR_READER = None


def get_easyocr_reader():
    global EASYOCR_READER
    if EASYOCR_READER is None:
        try:
            EASYOCR_READER = easyocr.Reader(["fr", "en"], gpu=False)
        except Exception as e:
            print(f"[WARN] EasyOCR unavailable: {e}")
            EASYOCR_READER = False
    return EASYOCR_READER


# ============================================================
# HELPERS GÉNÉRAUX
# ============================================================

def safe_strip(v: Any) -> str:
    return re.sub(r"\s+", " ", str(v or "").strip())


def sha1_text(value: str) -> str:
    return hashlib.sha1((value or "").encode("utf-8", errors="ignore")).hexdigest()


def image_hash_bytes(img_bytes: bytes) -> str:
    return hashlib.sha1(img_bytes or b"").hexdigest()


def upper_or_empty(v: str) -> str:
    return safe_strip(v).upper()


def null_if_empty(v: str) -> str:
    v = safe_strip(v)
    return v if v else ""


def unique_keep_order(items: List[str]) -> List[str]:
    out = []
    seen = set()
    for item in items:
        key = safe_strip(item).lower()
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(safe_strip(item))
    return out


def parse_json_lines(text: str) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for ln in (text or "").splitlines():
        ln = ln.strip()
        if not ln:
            continue
        try:
            j = json.loads(ln)
            if isinstance(j, dict):
                out.append(j)
        except Exception:
            continue
    return out


def debug_preview(d: Dict[str, Any], keys: List[str]) -> str:
    parts = []
    for k in keys:
        v = safe_strip(d.get(k, ""))
        if v:
            parts.append(f"{k}={v}")
    return " | ".join(parts)


def normalize_spaces_inline(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def tokenize_search_input(raw: str) -> List[str]:
    raw = normalize_spaces_inline(raw)
    if not raw:
        return []
    raw = TOKEN_SPLIT_RE.sub(" ", raw)
    toks = [t.strip() for t in raw.split(" ") if t.strip()]
    return toks


def looks_like_phone_token(token: str) -> bool:
    return bool(normalize_phone_basic(token))


def looks_like_cp_token(token: str) -> bool:
    t = re.sub(r"\s+", "", token or "")
    return bool(CP_RE.fullmatch(t))


def looks_like_email_token(token: str) -> bool:
    return is_valid_email(token)


def normalize_phone_basic(s: str) -> str:
    s = safe_strip(s)
    if not s:
        return ""
    s = re.sub(r"[^\d+]", "", s)
    if s.startswith("+33"):
        s = "0" + s[3:]
    if not re.fullmatch(r"0\d{9}", s or ""):
        return ""
    return s


def domain_from_email(email: str) -> str:
    email = clean_email(email)
    if "@" not in email:
        return ""
    dom = email.split("@", 1)[1].strip()
    dom = re.sub(r"[^a-z0-9.\-]", "", dom)
    return dom


def domain_from_url(url: str) -> str:
    if not url:
        return ""
    u = url.lower().strip()
    u = re.sub(r"^https?://", "", u)
    u = u.split("/", 1)[0]
    u = u.split(":", 1)[0]
    u = re.sub(r"[^a-z0-9.\-]", "", u)
    if u.startswith("www."):
        u = u[4:]
    return u


def brand_from_domain(dom: str) -> str:
    dom = safe_strip(dom).lower()
    if not dom:
        return ""
    dom = dom.split(".", 1)[0]
    dom = dom.replace("-", " ").replace("_", " ")
    dom = re.sub(r"\s+", " ", dom).strip()
    return dom


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
        # ============================================================
# BLOC 2 / 6 — NORMALISATION / VALIDATION / MATCHING / DEDUPE
# ============================================================

def normalize_text(s: str) -> str:
    s = safe_strip(s).lower()
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e").replace("ë", "e")
    s = s.replace("à", "a").replace("â", "a")
    s = s.replace("î", "i").replace("ï", "i")
    s = s.replace("ô", "o")
    s = s.replace("ù", "u").replace("û", "u").replace("ü", "u")
    s = s.replace("ç", "c")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_company_name(s: str) -> str:
    s = normalize_text(s)
    parts = [p for p in s.split() if p not in CORPORATE_WORDS]
    return " ".join(parts).strip()


def normalize_upper(v: str) -> str:
    return safe_strip(v).upper()


def validate_siret(raw: str) -> str:
    siret = re.sub(r"\D", "", safe_strip(raw))
    if len(siret) != 14 or not siret.isdigit():
        return ""

    total = 0
    for i in range(14):
        digit = int(siret[13 - i])
        if i % 2 == 1:
            digit *= 2
        if digit > 9:
            digit -= 9
        total += digit

    return siret if total % 10 == 0 else ""


def normalize_phone(s: str) -> str:
    s = normalize_phone_basic(s)
    return s


def is_fixed_phone(phone: str) -> bool:
    p = normalize_phone(phone)
    return bool(p and re.fullmatch(r"0[1-5]\d{8}", p))


def is_mobile_phone(phone: str) -> bool:
    p = normalize_phone(phone)
    return bool(p and re.fullmatch(r"0[67]\d{8}", p))


def extract_all_phones(text: str) -> List[str]:
    if not text:
        return []
    out = []
    seen = set()
    for m in PHONE_RE.finditer(text):
        p = normalize_phone(m.group(0))
        if not p:
            continue
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def split_phones_by_priority(phones: List[str]) -> Tuple[str, str]:
    fixed = []
    mobile = []
    other = []

    for p in phones:
        pn = normalize_phone(p)
        if not pn:
            continue
        if is_fixed_phone(pn):
            if pn not in fixed:
                fixed.append(pn)
        elif is_mobile_phone(pn):
            if pn not in mobile:
                mobile.append(pn)
        else:
            if pn not in other:
                other.append(pn)

    phone = ""
    phone2 = ""

    if fixed:
        phone = fixed[0]
        if mobile:
            phone2 = mobile[0]
        elif len(fixed) > 1:
            phone2 = fixed[1]
        elif other:
            phone2 = other[0]
    elif mobile:
        phone = mobile[0]
        if len(mobile) > 1:
            phone2 = mobile[1]
        elif other:
            phone2 = other[0]
    elif other:
        phone = other[0]
        if len(other) > 1:
            phone2 = other[1]

    if phone and phone2 and phone == phone2:
        phone2 = ""

    return phone, phone2


def best_email(text: str) -> str:
    if not text:
        return ""
    candidates = [clean_email(m.group(0)) for m in EMAIL_IN_TEXT_RE.finditer(text)]
    for em in candidates:
        if is_valid_email(em):
            return em
    return ""


def extract_all_emails(text: str) -> List[str]:
    if not text:
        return []
    out = []
    seen = set()
    for m in EMAIL_IN_TEXT_RE.finditer(text):
        em = clean_email(m.group(0))
        if not is_valid_email(em):
            continue
        if em not in seen:
            seen.add(em)
            out.append(em)
    return out


def extract_urls(text: str) -> List[str]:
    if not text:
        return []
    out, seen = [], set()
    for m in URL_RE.finditer(text):
        u = m.group(0).strip()
        if u.lower().startswith("www."):
            u = "https://" + u
        u = u.rstrip(".,;:")
        k = u.lower()
        if k not in seen:
            seen.add(k)
            out.append(u)
    return out


def normalize_cp(cp: str) -> str:
    return re.sub(r"\s+", "", safe_strip(cp))


def extract_postal_code(text: str) -> str:
    if not text:
        return ""
    cps = CP_RE.findall(text)
    return normalize_cp(cps[0]) if cps else ""


def split_human_name(full: str) -> Tuple[str, str]:
    s = safe_strip(full)
    if not s:
        return "", ""
    parts = s.split(" ")
    if len(parts) == 1:
        return "", parts[0]
    return parts[0], " ".join(parts[1:])


def is_probable_person_name(s: str) -> bool:
    s = normalize_spaces_inline(s)
    if not s:
        return False

    low = normalize_text(s)

    if "@" in s or "www" in low:
        return False
    if re.search(r"\d", s):
        return False
    if len(low.split()) < 2 or len(low.split()) > 4:
        return False

    bad_exact = {"alt gr", "ait gr", "alt qr", "ait qr", "sas", "sarl", "sa", "eurl"}
    if low in bad_exact:
        return False

    bad_tokens = {
        "restaurant", "responsable", "commerce", "collectivite", "territoire",
        "recyclage", "valorisation", "dechets", "drh", "developpement", "rh",
        "onyx", "auvergne", "rhone", "alpes", "sas", "sarl", "eurl", "sa",
        "service", "environnement", "gestion", "cedex", "centr", "alp",
        "experience", "luxe", "scannez", "scanner", "savoir", "plus"
    }

    words = low.split()
    if any(w in bad_tokens for w in words):
        return False

    return sum(1 for w in words if len(w) >= 2) >= 2


def strip_postal_city_from_address(addr: str, postal_code: str, city: str) -> str:
    addr = safe_strip(addr)
    if not addr:
        return ""

    pc = safe_strip(postal_code)
    ct = safe_strip(city)
    addr = re.sub(r"\s+", " ", addr)

    if pc and ct:
        tail = f"{pc} {ct}".strip()
        addr = re.sub(rf"(?:\s|-)?{re.escape(tail)}\s*$", "", addr, flags=re.IGNORECASE).strip()

    if pc:
        addr = re.sub(rf"(?:\s|-)?{re.escape(pc)}\s*$", "", addr).strip()

    return re.sub(r"\s+", " ", addr).strip()


def normalize_row_address(row: Dict[str, Any]) -> Dict[str, Any]:
    row["address"] = strip_postal_city_from_address(
        row.get("address", ""),
        row.get("postal_code", ""),
        row.get("city", ""),
    )
    return row


def extract_local_address_from_text(text: str) -> Tuple[str, str, str]:
    if not text:
        return "", "", ""

    lines = [normalize_spaces_inline(ln) for ln in text.splitlines()]
    lines = [ln for ln in lines if ln]

    best_addr = ""
    best_cp = ""
    best_city = ""

    for i, ln in enumerate(lines):
        m = CP_RE.search(ln)
        if not m:
            continue

        cp = normalize_cp(m.group(1))
        before = ln[:m.start()].strip(" -,:;")
        after = ln[m.end():].strip(" -,:;")

        addr_line = ln
        if len(before) < 4 and i > 0:
            prev = lines[i - 1]
            if not CP_RE.search(prev):
                addr_line = f"{prev} {ln}".strip()

        city = after[:80].strip()
        addr_line = normalize_spaces_inline(addr_line)
        city = normalize_spaces_inline(city)

        if cp:
            best_addr, best_cp, best_city = addr_line, cp, city
            break

    return best_addr, best_cp, best_city


def ensure_contact_fallback(d: Dict[str, Any]) -> Dict[str, Any]:
    fn = safe_strip(d.get("contact_firstname", ""))
    ln = safe_strip(d.get("contact_lastname", ""))
    interlocuteur = safe_strip(d.get("interlocuteur", ""))
    dirigeant = safe_strip(d.get("dirigeant", ""))

    if interlocuteur:
        f, l = split_human_name(interlocuteur)
        if not fn and f:
            d["contact_firstname"] = f
        if not ln and l:
            d["contact_lastname"] = l

    fn = safe_strip(d.get("contact_firstname", ""))
    ln = safe_strip(d.get("contact_lastname", ""))

    if (not fn and not ln) and dirigeant:
        f, l = split_human_name(dirigeant)
        d["contact_firstname"] = f or ""
        d["contact_lastname"] = l or dirigeant

    if not safe_strip(d.get("interlocuteur", "")):
        combo = f"{safe_strip(d.get('contact_firstname', ''))} {safe_strip(d.get('contact_lastname', ''))}".strip()
        if combo:
            d["interlocuteur"] = combo
        elif dirigeant:
            d["interlocuteur"] = dirigeant

    return d


def clean_ocr_line(line: str) -> str:
    s = normalize_spaces_inline(line)
    s = s.replace("|", " ").replace("•", " ").replace("—", " ").replace("–", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_ocr_lines(text: str) -> List[str]:
    bad_contains = [
        "scannez-moi",
        "scannez moi",
        "en savoir plus",
        "qr code",
    ]
    out = []
    seen = set()

    for raw in (text or "").splitlines():
        ln = clean_ocr_line(raw)
        if not ln:
            continue

        low = normalize_text(ln)
        if any(x in low for x in bad_contains):
            continue

        if low in seen:
            continue
        seen.add(low)
        out.append(ln)

    return out


def looks_like_company_line(line: str) -> bool:
    ln = clean_ocr_line(line)
    if not ln:
        return False

    low = normalize_text(ln)

    if "@" in ln:
        return False
    if URL_RE.search(ln):
        return False
    if PHONE_RE.search(ln):
        return False
    if CP_RE.search(ln):
        return False

    bad_words = {
        "directeur", "responsable", "rh", "site", "manager",
        "portable", "port", "tel", "telephone", "mail", "email",
        "scannez", "scanner", "savoir", "experience", "luxe"
    }
    words = set(low.split())
    if words & bad_words:
        return False

    if len(ln) < 3:
        return False
    if len(ln) > 60:
        return False

    return True


def detect_company_from_lines(lines: List[str], gemini_company: str = "", website: str = "", email: str = "") -> str:
    if gemini_company and len(safe_strip(gemini_company)) >= 3:
        return normalize_spaces_inline(gemini_company)

    dom_web = brand_from_domain(domain_from_url(website))
    dom_email = brand_from_domain(domain_from_email(email))

    candidates = []
    for i, ln in enumerate(lines[:10]):
        if not looks_like_company_line(ln):
            continue

        low = normalize_company_name(ln)
        score = 0

        if i <= 1:
            score += 25
        if ln == ln.upper():
            score += 20
        if len(ln.split()) <= 3:
            score += 15
        if len(ln) <= 25:
            score += 10
        if dom_web and normalize_company_name(dom_web) in low:
            score += 30
        if dom_email and normalize_company_name(dom_email) in low:
            score += 30
        if len(ln.split()) in [2, 3]:
            score += 10

        candidates.append((score, ln))

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return normalize_spaces_inline(candidates[0][1])

    for cand in [dom_web, dom_email]:
        cand = normalize_spaces_inline(cand)
        if len(cand) >= 3:
            return cand

    return ""


def detect_person_name_from_lines(lines: List[str], company_line: str = "") -> str:
    company_n = normalize_company_name(company_line)

    for ln in lines[:12]:
        if not is_probable_person_name(ln):
            continue
        if company_n and normalize_company_name(ln) == company_n:
            continue
        return normalize_spaces_inline(ln)

    return ""


def clean_address_candidate(addr: str) -> str:
    a = normalize_spaces_inline(addr)
    if not a:
        return ""

    low = normalize_text(a)

    if "@" in a:
        return ""
    if URL_RE.search(a):
        return ""
    if PHONE_RE.search(a):
        return ""
    if low.startswith("port"):
        return ""
    if low.startswith("tel"):
        return ""
    if "portable" in low:
        return ""

    return a


def deduce_company(company_raw: str, email: str, website: str) -> str:
    company_raw = normalize_spaces_inline(company_raw)
    dom_email = brand_from_domain(domain_from_email(email))
    dom_web = brand_from_domain(domain_from_url(website))

    for cand in [company_raw, dom_web, dom_email]:
        cand = normalize_spaces_inline(cand)
        if len(cand) >= 4:
            return cand

    return company_raw


def choose_final_company_name(
    gouv_name: str,
    gemini_company: str,
    detected_company_line: str,
    person_name: str,
    website: str,
    email: str
) -> str:
    person_n = normalize_company_name(person_name)
    dom_web = brand_from_domain(domain_from_url(website))
    dom_mail = brand_from_domain(domain_from_email(email))

    candidates = [
        normalize_spaces_inline(gouv_name),
        normalize_spaces_inline(gemini_company),
        normalize_spaces_inline(detected_company_line),
    ]

    scored = []

    for cand in candidates:
        if not cand:
            continue

        cand_norm = normalize_company_name(cand)
        if not cand_norm:
            continue

        if person_n and cand_norm == person_n:
            continue
        if is_probable_person_name(cand):
            continue

        score = 0

        if gouv_name and safe_strip(cand) == safe_strip(gouv_name):
            score += 100

        if dom_web and normalize_company_name(dom_web) in cand_norm:
            score += 30
        if dom_mail and normalize_company_name(dom_mail) in cand_norm:
            score += 30

        wc = len(cand.split())
        if wc >= 2:
            score += 15
        if wc == 1:
            score -= 10
        if len(cand) <= 5:
            score -= 15

        scored.append((score, cand))

    if scored:
        scored.sort(key=lambda x: x[0], reverse=True)
        return normalize_spaces_inline(scored[0][1])

    for cand in [dom_web, dom_mail]:
        cand = normalize_spaces_inline(cand)
        if len(cand) >= 3:
            return cand

    return ""


# ============================================================
# RECHERCHE MULTI-SIGNAUX
# ============================================================

def classify_search_tokens(raw: str) -> Dict[str, Any]:
    toks = tokenize_search_input(raw)

    phones = []
    cps = []
    emails = []
    exact_hints = []
    others = []

    for tok in toks:
        if looks_like_email_token(tok):
            emails.append(clean_email(tok))
        elif looks_like_phone_token(tok):
            phones.append(normalize_phone(tok))
        elif looks_like_cp_token(tok):
            cps.append(normalize_cp(tok))
        else:
            others.append(tok)

    city_hints = []
    activity_hints = []
    name_parts = []

    for tok in others:
        low = normalize_text(tok)
        if low in COMMON_CITY_HINTS:
            city_hints.append(tok)
        elif low in COMMON_ACTIVITY_HINTS:
            activity_hints.append(tok)
        else:
            name_parts.append(tok)

    out = {
        "raw": normalize_spaces_inline(raw),
        "phones": unique_keep_order(phones),
        "postal_codes": unique_keep_order(cps),
        "emails": unique_keep_order(emails),
        "city_hints": unique_keep_order(city_hints),
        "activity_hints": unique_keep_order(activity_hints),
        "name_parts": unique_keep_order(name_parts),
        "token_count": len(toks),
    }

    out["signal_count"] = sum(
        1 for bucket in [
            out["phones"],
            out["postal_codes"],
            out["emails"],
            out["city_hints"],
            out["activity_hints"],
            out["name_parts"],
        ] if bucket
    )

    return out


def has_minimum_search_signals(parsed: Dict[str, Any]) -> bool:
    return int(parsed.get("signal_count", 0)) >= 2


def build_search_queries(parsed: Dict[str, Any]) -> List[str]:
    queries = []

    name = " ".join(parsed.get("name_parts", []))
    city = " ".join(parsed.get("city_hints", []))
    activity = " ".join(parsed.get("activity_hints", []))
    cp = " ".join(parsed.get("postal_codes", []))
    phone = " ".join(parsed.get("phones", []))
    email = " ".join(parsed.get("emails", []))

    combos = [
        f"{name} {city}",
        f"{name} {cp}",
        f"{name} {activity}",
        f"{city} {activity}",
        f"{phone} {city}",
        f"{phone} {activity}",
        f"{cp} {phone}",
        f"{email} {city}",
        f"{name} {email}",
        parsed.get("raw", ""),
    ]

    for q in combos:
        q = normalize_spaces_inline(q)
        if q and q not in queries:
            queries.append(q)

    return queries[:8]


def score_text_overlap(a: str, b: str) -> int:
    na = set(normalize_company_name(a).split())
    nb = set(normalize_company_name(b).split())
    if not na or not nb:
        return 0
    common = na & nb
    return min(len(common) * 12, 36)


def score_city_match(city_hint: str, city_value: str) -> int:
    a = normalize_text(city_hint)
    b = normalize_text(city_value)
    if not a or not b:
        return 0
    if a == b:
        return 30
    if a in b or b in a:
        return 15
    return 0


def score_cp_match(cp_hint: str, cp_value: str) -> int:
    a = normalize_cp(cp_hint)
    b = normalize_cp(cp_value)
    if not a or not b:
        return 0
    if a == b:
        return 25
    return 0


def score_domain_match(email: str, website: str, company_name: str) -> int:
    dom_email = brand_from_domain(domain_from_email(email))
    dom_web = brand_from_domain(domain_from_url(website))
    comp = normalize_company_name(company_name)

    score = 0
    for dom in [dom_email, dom_web]:
        if dom and normalize_company_name(dom) in comp:
            score += 20
    return score


def score_phone_presence(phone_hint: str, text_blob: str) -> int:
    ph = normalize_phone(phone_hint)
    if not ph:
        return 0
    blob = re.sub(r"\D", "", text_blob or "")
    if ph and re.sub(r"\D", "", ph) in blob:
        return 35
    return 0


def safe_record_text_blob(d: Dict[str, Any]) -> str:
    vals = [
        d.get("name", ""),
        d.get("address", ""),
        d.get("postal_code", ""),
        d.get("city", ""),
        d.get("phone", ""),
        d.get("phone2", ""),
        d.get("email", ""),
        d.get("website", ""),
        d.get("dirigeant", ""),
        d.get("interlocuteur", ""),
        d.get("naf", ""),
    ]
    return " | ".join([safe_strip(v) for v in vals if safe_strip(v)])


def score_candidate_from_signals(parsed: Dict[str, Any], candidate: Dict[str, Any]) -> int:
    score = 0

    name_hint = " ".join(parsed.get("name_parts", []))
    city_hint = " ".join(parsed.get("city_hints", []))
    activity_hint = " ".join(parsed.get("activity_hints", []))
    cp_hint = " ".join(parsed.get("postal_codes", []))
    email_hint = " ".join(parsed.get("emails", []))
    phone_hint = " ".join(parsed.get("phones", []))

    name_val = candidate.get("name", "") or candidate.get("nom", "")
    city_val = candidate.get("city", "") or candidate.get("ville", "")
    cp_val = candidate.get("postal_code", "") or candidate.get("codePostal", "")
    naf_val = candidate.get("naf", "")
    website_val = candidate.get("website", "")
    text_blob = safe_record_text_blob(candidate)

    if name_hint:
        score += score_text_overlap(name_hint, name_val)

    if city_hint:
        score += score_city_match(city_hint, city_val)

    if cp_hint:
        score += score_cp_match(cp_hint, cp_val)

    if email_hint:
        score += score_domain_match(email_hint, website_val, name_val)

    if phone_hint:
        score += score_phone_presence(phone_hint, text_blob)

    if activity_hint:
        act = normalize_text(activity_hint)
        blob = normalize_text(f"{name_val} {naf_val} {text_blob}")
        if act and act in blob:
            score += 18

    if validate_siret(candidate.get("siret", "")):
        score += 5

    return score


# ============================================================
# DEDUPE AVANCÉE
# ============================================================

def record_score(d: Dict[str, Any]) -> int:
    fields = [
        "name", "address", "postal_code", "city", "siret", "naf", "activity_summary",
        "dirigeant", "interlocuteur", "contact_firstname", "contact_lastname",
        "phone", "phone2", "email", "website", "resume", "commande"
    ]
    score = sum(1 for k in fields if safe_strip(d.get(k, "")))
    if validate_siret(d.get("siret", "")):
        score += 2
    if safe_strip(d.get("email", "")):
        score += 1
    if safe_strip(d.get("phone", "")):
        score += 1
    return score


def dedupe_key_strict(r: Dict[str, Any]) -> str:
    siret = validate_siret(r.get("siret", "") or "")
    if siret:
        return f"SIRET|{siret}"

    return "|".join([
        normalize_text(r.get("agency", "")),
        normalize_text(r.get("initials", "")),
        normalize_company_name(r.get("name", "")),
        normalize_text(r.get("email", "")),
        normalize_text(r.get("phone", "")),
        normalize_text(r.get("city", "")),
    ])


def dedupe_key_soft(r: Dict[str, Any]) -> str:
    return "|".join([
        normalize_text(r.get("agency", "")),
        normalize_company_name(r.get("name", "")),
        normalize_text(r.get("city", "")),
        normalize_text(domain_from_email(r.get("email", ""))),
        normalize_text(domain_from_url(r.get("website", ""))),
    ])


def merge_rows_keep_best(primary: Dict[str, Any], secondary: Dict[str, Any]) -> Dict[str, Any]:
    out = dict(primary)

    preserve_user_fields = {"resume", "commande"}

    for k, v in secondary.items():
        if k in preserve_user_fields:
            if not safe_strip(out.get(k, "")) and safe_strip(v):
                out[k] = v
            continue

        if not safe_strip(out.get(k, "")) and safe_strip(v):
            out[k] = v

    out = ensure_contact_fallback(out)
    out = normalize_row_address(out)
    return out


def dedupe_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not rows:
        return []

    strict_map: Dict[str, Dict[str, Any]] = {}
    soft_groups: Dict[str, List[Dict[str, Any]]] = {}

    for r in rows:
        k = dedupe_key_strict(r)
        if k.strip("|"):
            prev = strict_map.get(k)
            if prev is None:
                strict_map[k] = r
            else:
                if record_score(r) > record_score(prev):
                    strict_map[k] = merge_rows_keep_best(r, prev)
                else:
                    strict_map[k] = merge_rows_keep_best(prev, r)
        else:
            sk = dedupe_key_soft(r)
            soft_groups.setdefault(sk, []).append(r)

    merged = list(strict_map.values())

    for sk, group in soft_groups.items():
        if not sk.strip("|"):
            merged.extend(group)
            continue

        STATS["dedupe_groups"] += 1
        best = sorted(group, key=record_score, reverse=True)[0]
        for other in group[1:]:
            best = merge_rows_keep_best(best, other)
            STATS["dedupe_merged"] += 1
        merged.append(best)

    final_map: Dict[str, Dict[str, Any]] = {}
    for r in merged:
        k = dedupe_key_strict(r) or dedupe_key_soft(r) or sha1_text(json.dumps(r, ensure_ascii=False, sort_keys=True))
        prev = final_map.get(k)
        if prev is None:
            final_map[k] = r
        else:
            if record_score(r) > record_score(prev):
                final_map[k] = merge_rows_keep_best(r, prev)
            else:
                final_map[k] = merge_rows_keep_best(prev, r)

    return list(final_map.values())


# ============================================================
# SANITIZE FINAL ROW
# ============================================================

def uppercase_business_fields(row: Dict[str, Any]) -> Dict[str, Any]:
    for k in ["name", "address", "city", "dirigeant", "interlocuteur"]:
        row[k] = normalize_upper(row.get(k, ""))
    return row


def sanitize_final_row(row: Dict[str, Any]) -> Dict[str, Any]:
    row = dict(row)

    row["date"] = safe_strip(row.get("date", ""))
    row["agency"] = upper_or_empty(row.get("agency", ""))
    row["initials"] = upper_or_empty(row.get("initials", ""))

    row["name"] = safe_strip(row.get("name", ""))
    row["address"] = safe_strip(row.get("address", ""))
    row["postal_code"] = normalize_cp(row.get("postal_code", ""))
    row["city"] = safe_strip(row.get("city", ""))

    row["siret"] = validate_siret(row.get("siret", ""))
    row["naf"] = safe_strip(row.get("naf", "")).replace(".", "").replace(" ", "").upper()

    row["activity_summary"] = safe_strip(row.get("activity_summary", ""))[:60]

    row["dirigeant"] = safe_strip(row.get("dirigeant", ""))
    row["interlocuteur"] = safe_strip(row.get("interlocuteur", ""))
    row["contact_firstname"] = safe_strip(row.get("contact_firstname", ""))
    row["contact_lastname"] = safe_strip(row.get("contact_lastname", ""))

    row["phone"] = normalize_phone(row.get("phone", ""))
    row["phone2"] = normalize_phone(row.get("phone2", ""))

    if row["phone"] and row["phone2"] and row["phone"] == row["phone2"]:
        row["phone2"] = ""

    row["email"] = clean_email(row.get("email", "")) if is_valid_email(row.get("email", "")) else ""
    row["website"] = safe_strip(row.get("website", ""))
    if not row["website"] and row["email"]:
        dom = domain_from_email(row["email"])
        if dom and dom not in FREE_MAIL_DOMAINS:
            row["website"] = f"https://{dom}"

    row["resume"] = safe_strip(row.get("resume", ""))
    row["commande"] = safe_strip(row.get("commande", ""))

    row = ensure_contact_fallback(row)
    row = normalize_row_address(row)
    row = uppercase_business_fields(row)

    return row


# ============================================================
# ANTI-RÉGRESSION / GARDE-FOUS
# ============================================================

def guard_minimum_business_identity(row: Dict[str, Any]) -> Dict[str, Any]:
    row = dict(row)
    if not safe_strip(row.get("name", "")):
        row["name"] = ""
    if not safe_strip(row.get("city", "")) and safe_strip(row.get("address", "")):
        _, cp, city = extract_local_address_from_text(row.get("address", ""))
        if city:
            row["city"] = city
        if cp and not safe_strip(row.get("postal_code", "")):
            row["postal_code"] = cp
    return row


def ensure_required_headers(row: Dict[str, Any]) -> Dict[str, Any]:
    out = {}
    for h in HEADERS:
        out[h] = row.get(h, "")
    return out


def finalize_row_for_export(row: Dict[str, Any]) -> Dict[str, Any]:
    row = sanitize_final_row(row)
    row = guard_minimum_business_identity(row)
    row = ensure_required_headers(row)
    return row
    # ============================================================
# BLOC 3 / 6 — OCR HYBRIDE / GEMINI / DOUBLE PASSE CARTES
# VERSION OPTIMISÉE SANS PERTE DE FONCTIONNALITÉS
# ============================================================

USE_EASYOCR = env_str("USE_EASYOCR", "1").strip() == "1"


def pil_to_cv(img: Image.Image) -> np.ndarray:
    arr = np.array(img)
    if len(arr.shape) == 2:
        return arr
    return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)


def cv_to_pil(arr: np.ndarray) -> Image.Image:
    if arr is None:
        raise ValueError("cv array is None")
    if len(arr.shape) == 2:
        return Image.fromarray(arr)
    rgb = cv2.cvtColor(arr, cv2.COLOR_BGR2RGB)
    return Image.fromarray(rgb)


def open_image_from_bytes(img_bytes: bytes) -> Optional[Image.Image]:
    try:
        return Image.open(io.BytesIO(img_bytes)).convert("RGB")
    except Exception:
        return None


def preprocess_image_variants(img_bytes: bytes) -> List[Image.Image]:
    variants: List[Image.Image] = []
    img = open_image_from_bytes(img_bytes)
    if img is None:
        return variants

    variants.append(img)

    try:
        big = img.resize((max(1, img.width * 2), max(1, img.height * 2)))
        variants.append(big)
    except Exception:
        pass

    try:
        gray = img.convert("L")
        arr = np.array(gray)
        arr = cv2.equalizeHist(arr)
        variants.append(Image.fromarray(arr))
    except Exception:
        pass

    try:
        gray = img.convert("L")
        arr = np.array(gray)
        _, th = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        variants.append(Image.fromarray(th))
    except Exception:
        pass

    try:
        cv_img = pil_to_cv(img)
        gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (3, 3), 0)
        sharp = cv2.addWeighted(gray, 1.5, blur, -0.5, 0)
        variants.append(Image.fromarray(sharp))
    except Exception:
        pass

    try:
        cv_img = pil_to_cv(img)
        gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
        th = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 11
        )
        variants.append(Image.fromarray(th))
    except Exception:
        pass

    return variants


def crop_image_zones(img: Image.Image) -> List[Image.Image]:
    w, h = img.size
    zones = [img]

    try:
        zones.append(img.crop((0, 0, w, h // 3)))
        zones.append(img.crop((0, h // 3, w, 2 * h // 3)))
        zones.append(img.crop((0, 2 * h // 3, w, h)))
        zones.append(img.crop((0, 0, w // 2, h)))
        zones.append(img.crop((w // 2, 0, w, h)))
        zones.append(img.crop((0, 0, w, h // 2)))
        zones.append(img.crop((0, h // 2, w, h)))
    except Exception:
        pass

    return zones


def rotate_image_variants(img: Image.Image) -> List[Image.Image]:
    out = [img]
    for angle in (-2, 2, -4, 4, -90, 90):
        try:
            out.append(img.rotate(angle, expand=True, fillcolor="white"))
        except Exception:
            continue
    return out


def get_easyocr_reader():
    global EASYOCR_READER

    if not USE_EASYOCR:
        return False

    if EASYOCR_READER is None:
        try:
            EASYOCR_READER = easyocr.Reader(["fr", "en"], gpu=False)
        except Exception as e:
            print(f"[WARN] EasyOCR unavailable: {e}")
            EASYOCR_READER = False
    return EASYOCR_READER


def ocr_easyocr_on_pil(img: Image.Image) -> str:
    try:
        reader = get_easyocr_reader()
        if not reader:
            return ""
        STATS["ocr_easyocr_calls"] += 1
        arr = pil_to_cv(img)
        results = reader.readtext(arr, detail=0, paragraph=False)
        return "\n".join([str(x).strip() for x in results if str(x).strip()])
    except Exception:
        return ""


def ocr_tesseract_on_pil(img: Image.Image) -> str:
    try:
        STATS["ocr_tesseract_calls"] += 1
        txt1 = pytesseract.image_to_string(img, lang="fra+eng", config="--psm 6") or ""
        txt2 = pytesseract.image_to_string(img, lang="fra+eng", config="--psm 11") or ""
        return "\n".join([t for t in [txt1, txt2] if safe_strip(t)])
    except Exception:
        return ""


def merge_text_blocks(texts: List[str]) -> str:
    lines = []
    seen = set()
    for txt in texts:
        for ln in (txt or "").splitlines():
            ln = normalize_spaces_inline(ln)
            if not ln:
                continue
            key = ln.lower()
            if key not in seen:
                seen.add(key)
                lines.append(ln)
    return "\n".join(lines)


def extract_structured_from_text_fast(text: str) -> Dict[str, Any]:
    lines = clean_ocr_lines(text)
    joined = "\n".join(lines)

    emails = extract_all_emails(joined)
    urls = extract_urls(joined)
    phones = extract_all_phones(joined)
    phone, phone2 = split_phones_by_priority(phones)
    addr, cp, city = extract_local_address_from_text(joined)
    company = detect_company_from_lines(
        lines,
        website=(urls[0] if urls else ""),
        email=(emails[0] if emails else "")
    )
    person = detect_person_name_from_lines(lines, company_line=company)
    fn, ln = split_human_name(person)

    return {
        "lines": lines,
        "raw_text": joined,
        "emails": emails,
        "urls": urls,
        "phones": phones,
        "phone": phone,
        "phone2": phone2,
        "postal_code": cp,
        "city": city,
        "address": clean_address_candidate(addr),
        "company": normalize_spaces_inline(company),
        "person_name": normalize_spaces_inline(person),
        "contact_firstname": normalize_spaces_inline(fn),
        "contact_lastname": normalize_spaces_inline(ln),
    }


def enough_business_card_info(d: Dict[str, Any]) -> bool:
    filled = sum(
        1 for v in [
            safe_strip(d.get("name", "")),
            safe_strip(d.get("company", "")),
            clean_email(d.get("email", "")),
            normalize_phone(d.get("phone", "")),
            safe_strip(d.get("website", "")),
            safe_strip(d.get("city", "")),
            safe_strip(d.get("address", "")),
        ] if v
    )
    return filled >= 3


def enough_contact_info_for_skip_easyocr(d: Dict[str, Any], ocr_struct: Dict[str, Any]) -> bool:
    score = 0
    if safe_strip(d.get("company", "")) or safe_strip(ocr_struct.get("company", "")):
        score += 1
    if clean_email(d.get("email", "")) or (ocr_struct.get("emails") or []):
        score += 1
    if normalize_phone(d.get("phone", "")) or normalize_phone(ocr_struct.get("phone", "")):
        score += 1
    if safe_strip(d.get("website", "")) or (ocr_struct.get("urls") or []):
        score += 1
    if safe_strip(d.get("city", "")) or safe_strip(ocr_struct.get("city", "")):
        score += 1
    return score >= 3


def ocr_image_bytes(img_bytes: bytes, light: bool = False, allow_easyocr: bool = True) -> str:
    if not img_bytes:
        return ""

    h = image_hash_bytes(img_bytes) + ("|light" if light else "|full") + ("|easy" if allow_easyocr else "|noeasy")
    if h in OCR_TEXT_CACHE:
        return OCR_TEXT_CACHE[h]

    variants = preprocess_image_variants(img_bytes)
    all_texts = []

    base_variants = variants[:2] if light else variants[:4]

    for base_img in base_variants:
        rotated = rotate_image_variants(base_img)
        rotated = rotated[:2] if light else rotated[:4]

        for rot in rotated:
            zones = crop_image_zones(rot)
            zones = zones[:3] if light else zones[:5]

            for zone in zones:
                # Tesseract d'abord : plus léger
                t2 = ocr_tesseract_on_pil(zone)
                if t2:
                    all_texts.append(t2)

                # EasyOCR seulement si autorisé
                if allow_easyocr:
                    t1 = ocr_easyocr_on_pil(zone)
                    if t1:
                        all_texts.append(t1)

    merged = merge_text_blocks(all_texts)
    OCR_TEXT_CACHE[h] = merged
    return merged


def _guess_mime(img_bytes: bytes) -> str:
    try:
        im = Image.open(io.BytesIO(img_bytes))
        fmt = (im.format or "").upper()
        if fmt == "PNG":
            return "image/png"
        if fmt in ("JPG", "JPEG"):
            return "image/jpeg"
        if fmt == "WEBP":
            return "image/webp"
    except Exception:
        pass
    return "image/jpeg"


def gemini_vision_json(img_bytes: bytes, prompt: str, max_tokens: int = 650) -> Dict[str, Any]:
    if not GEMINI_API_KEY or not img_bytes:
        return {}

    cache_key = f"{image_hash_bytes(img_bytes)}|{sha1_text(prompt)}|{max_tokens}"
    if cache_key in SEARCH_MATCH_CACHE:
        cached = SEARCH_MATCH_CACHE[cache_key]
        if isinstance(cached, dict):
            return dict(cached)

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
        STATS["gemini_vision_calls"] += 1
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
        out = json.loads(m.group(0))
        SEARCH_MATCH_CACHE[cache_key] = out if isinstance(out, dict) else {}
        return out if isinstance(out, dict) else {}
    except Exception:
        return {}


def gemini_extract_business_card(img_bytes: bytes) -> Dict[str, str]:
    if not img_bytes:
        return {}
    key = image_hash_bytes(img_bytes)
    if key in GEMINI_CARD_CACHE:
        return GEMINI_CARD_CACHE[key]

    prompt = (
        "Tu es un extracteur expert de carte de visite.\n"
        "Lis l'image elle-même, pas seulement un OCR brut.\n"
        "Retourne STRICTEMENT un JSON avec les clés exactes :\n"
        "name, company, title, email, phone, website, postal_code, city, address\n"
        "Règles :\n"
        "- si inconnu : \"\"\n"
        "- name = prénom + nom si visibles\n"
        "- company = raison sociale ou marque la plus probable\n"
        "- email en minuscules\n"
        "- postal_code = 5 chiffres si visibles\n"
        "- address = adresse propre, sans inventer\n"
    )

    d = gemini_vision_json(img_bytes, prompt, max_tokens=700) or {}
    out = {
        "name": normalize_spaces_inline(d.get("name", "")),
        "company": normalize_spaces_inline(d.get("company", "")),
        "title": normalize_spaces_inline(d.get("title", "")),
        "email": clean_email(str(d.get("email", ""))),
        "phone": normalize_phone(str(d.get("phone", ""))),
        "website": normalize_spaces_inline(d.get("website", "")),
        "postal_code": normalize_cp(str(d.get("postal_code", ""))),
        "city": normalize_spaces_inline(d.get("city", "")),
        "address": normalize_spaces_inline(d.get("address", "")),
    }
    GEMINI_CARD_CACHE[key] = out
    return out


def gemini_extract_business_card_second_pass(img_bytes: bytes) -> Dict[str, str]:
    if not img_bytes:
        return {}

    prompt = (
        "Deuxième passe d'analyse d'une carte de visite difficile.\n"
        "Retourne STRICTEMENT un JSON avec les clés exactes :\n"
        "name, company, title, email, phone, website, postal_code, city, address\n"
        "Règles :\n"
        "- privilégie les éléments visuellement les plus probables\n"
        "- ne pas inventer\n"
        "- si doute, laisser vide\n"
    )
    d = gemini_vision_json(img_bytes, prompt, max_tokens=700) or {}
    return {
        "name": normalize_spaces_inline(d.get("name", "")),
        "company": normalize_spaces_inline(d.get("company", "")),
        "title": normalize_spaces_inline(d.get("title", "")),
        "email": clean_email(str(d.get("email", ""))),
        "phone": normalize_phone(str(d.get("phone", ""))),
        "website": normalize_spaces_inline(d.get("website", "")),
        "postal_code": normalize_cp(str(d.get("postal_code", ""))),
        "city": normalize_spaces_inline(d.get("city", "")),
        "address": normalize_spaces_inline(d.get("address", "")),
    }


def gemini_extract_facade_logo(img_bytes: bytes) -> Dict[str, str]:
    if not img_bytes:
        return {}
    key = image_hash_bytes(img_bytes)
    if key in GEMINI_FACADE_CACHE:
        return GEMINI_FACADE_CACHE[key]

    prompt = (
        "Tu analyses une photo de prospection de façade, enseigne ou logo.\n"
        "Retourne STRICTEMENT un JSON avec les clés exactes :\n"
        "company, city\n"
        "Règles :\n"
        "- si inconnu : \"\"\n"
        "- company = nom d'entreprise / enseigne le plus probable\n"
    )

    d = gemini_vision_json(img_bytes, prompt, max_tokens=350) or {}
    out = {
        "company": normalize_spaces_inline(d.get("company", "")),
        "city": normalize_spaces_inline(d.get("city", "")),
    }
    GEMINI_FACADE_CACHE[key] = out
    return out


def extract_structured_from_ocr_text(ocr_text: str) -> Dict[str, Any]:
    lines = clean_ocr_lines(ocr_text)
    text = "\n".join(lines)

    emails = extract_all_emails(text)
    urls = extract_urls(text)
    phones = extract_all_phones(text)
    phone, phone2 = split_phones_by_priority(phones)

    local_addr_ocr, local_cp_ocr, local_city_ocr = extract_local_address_from_text(text)

    company = detect_company_from_lines(
        lines,
        website=(urls[0] if urls else ""),
        email=(emails[0] if emails else "")
    )
    person = detect_person_name_from_lines(lines, company_line=company)
    fn, ln = split_human_name(person)

    return {
        "lines": lines,
        "raw_text": text,
        "emails": emails,
        "urls": urls,
        "phones": phones,
        "phone": phone,
        "phone2": phone2,
        "postal_code": local_cp_ocr,
        "city": local_city_ocr,
        "address": clean_address_candidate(local_addr_ocr),
        "company": normalize_spaces_inline(company),
        "person_name": normalize_spaces_inline(person),
        "contact_firstname": normalize_spaces_inline(fn),
        "contact_lastname": normalize_spaces_inline(ln),
    }


def merge_card_passes(pass1: Dict[str, Any], pass2: Dict[str, Any], ocr_struct: Dict[str, Any]) -> Dict[str, Any]:
    out = {
        "name": "",
        "company": "",
        "title": "",
        "email": "",
        "phone": "",
        "phone2": "",
        "website": "",
        "postal_code": "",
        "city": "",
        "address": "",
        "ocr_text": ocr_struct.get("raw_text", ""),
        "ocr_lines": ocr_struct.get("lines", []),
    }

    def pick(*vals) -> str:
        for v in vals:
            if safe_strip(v):
                return safe_strip(v)
        return ""

    out["name"] = pick(pass1.get("name"), pass2.get("name"), ocr_struct.get("person_name"))
    out["company"] = pick(pass1.get("company"), pass2.get("company"), ocr_struct.get("company"))
    out["title"] = pick(pass1.get("title"), pass2.get("title"))
    out["email"] = pick(pass1.get("email"), pass2.get("email"), (ocr_struct.get("emails") or [""])[0])
    out["website"] = pick(pass1.get("website"), pass2.get("website"), (ocr_struct.get("urls") or [""])[0])

    phones_all = []
    for x in [
        pass1.get("phone", ""),
        pass2.get("phone", ""),
        ocr_struct.get("phone", ""),
        ocr_struct.get("phone2", "")
    ] + list(ocr_struct.get("phones", []) or []):
        p = normalize_phone(x)
        if p and p not in phones_all:
            phones_all.append(p)

    out["phone"], out["phone2"] = split_phones_by_priority(phones_all)

    out["postal_code"] = pick(pass1.get("postal_code"), pass2.get("postal_code"), ocr_struct.get("postal_code"))
    out["city"] = pick(pass1.get("city"), pass2.get("city"), ocr_struct.get("city"))
    out["address"] = pick(pass1.get("address"), pass2.get("address"), ocr_struct.get("address"))

    if not out["website"] and out["email"]:
        dom = domain_from_email(out["email"])
        if dom and dom not in FREE_MAIL_DOMAINS:
            out["website"] = f"https://{dom}"

    return out


def analyze_business_card_bytes(img_bytes: bytes) -> Dict[str, Any]:
    if not img_bytes:
        return {
            "name": "",
            "company": "",
            "title": "",
            "email": "",
            "phone": "",
            "phone2": "",
            "website": "",
            "postal_code": "",
            "city": "",
            "address": "",
            "ocr_text": "",
            "ocr_lines": [],
        }

    key = image_hash_bytes(img_bytes)
    if key in PROSPECT_CARD_ANALYSIS_CACHE:
        return dict(PROSPECT_CARD_ANALYSIS_CACHE[key])

    # 1) Gemini passe 1
    pass1 = gemini_extract_business_card(img_bytes) if GEMINI_API_KEY else {}

    # 2) OCR léger sans EasyOCR d’abord
    light_ocr = enough_business_card_info(pass1)
    ocr_txt_fast = ocr_image_bytes(img_bytes, light=light_ocr, allow_easyocr=False)
    ocr_struct_fast = extract_structured_from_ocr_text(ocr_txt_fast)

    # 3) décider si on a besoin d’une passe 2 Gemini
    need_second_pass = (
        not safe_strip(pass1.get("name", "")) or
        not safe_strip(pass1.get("company", "")) or
        (
            not safe_strip(pass1.get("email", "")) and
            not safe_strip((ocr_struct_fast.get("emails") or [""])[0])
        )
    )

    pass2 = gemini_extract_business_card_second_pass(img_bytes) if GEMINI_API_KEY and need_second_pass else {}

    # 4) si Gemini + Tesseract suffisent, on n’appelle pas EasyOCR
    if enough_contact_info_for_skip_easyocr(
        {
            "company": pass1.get("company", "") or pass2.get("company", ""),
            "email": pass1.get("email", "") or pass2.get("email", ""),
            "phone": pass1.get("phone", "") or pass2.get("phone", ""),
            "website": pass1.get("website", "") or pass2.get("website", ""),
            "city": pass1.get("city", "") or pass2.get("city", ""),
        },
        ocr_struct_fast,
    ):
        merged = merge_card_passes(pass1, pass2, ocr_struct_fast)
        PROSPECT_CARD_ANALYSIS_CACHE[key] = dict(merged)
        return merged

    # 5) seulement ici, OCR complet avec EasyOCR
    ocr_txt_full = ocr_image_bytes(img_bytes, light=False, allow_easyocr=True)
    ocr_struct_full = extract_structured_from_ocr_text(ocr_txt_full)

    merged = merge_card_passes(pass1, pass2, ocr_struct_full)
    PROSPECT_CARD_ANALYSIS_CACHE[key] = dict(merged)
    return merged


def analyze_facade_bytes(img_bytes: bytes) -> Dict[str, Any]:
    if not img_bytes:
        return {
            "company": "",
            "city": "",
            "ocr_text": "",
            "ocr_lines": [],
            "urls": [],
            "emails": [],
            "phones": [],
            "phone": "",
            "phone2": "",
            "address": "",
            "postal_code": "",
        }

    key = f"facade|{image_hash_bytes(img_bytes)}"
    if key in PROSPECT_CARD_ANALYSIS_CACHE:
        return dict(PROSPECT_CARD_ANALYSIS_CACHE[key])

    gem = gemini_extract_facade_logo(img_bytes) if GEMINI_API_KEY else {"company": "", "city": ""}

    # OCR léger sans EasyOCR d’abord
    ocr_txt_fast = ocr_image_bytes(img_bytes, light=not bool(gem.get("company")), allow_easyocr=False)
    ocr_struct_fast = extract_structured_from_ocr_text(ocr_txt_fast)

    enough = sum(
        1 for v in [
            safe_strip(gem.get("company", "") or ocr_struct_fast.get("company", "")),
            safe_strip(gem.get("city", "") or ocr_struct_fast.get("city", "")),
            (ocr_struct_fast.get("urls") or [""])[0] if ocr_struct_fast.get("urls") else "",
            (ocr_struct_fast.get("emails") or [""])[0] if ocr_struct_fast.get("emails") else "",
        ] if safe_strip(v)
    ) >= 2

    if enough:
        ocr_struct = ocr_struct_fast
    else:
        ocr_txt_full = ocr_image_bytes(img_bytes, light=False, allow_easyocr=True)
        ocr_struct = extract_structured_from_ocr_text(ocr_txt_full)

    company = choose_final_company_name(
        gouv_name="",
        gemini_company=gem.get("company", ""),
        detected_company_line=ocr_struct.get("company", ""),
        person_name=ocr_struct.get("person_name", ""),
        website=(ocr_struct.get("urls") or [""])[0] if ocr_struct.get("urls") else "",
        email=(ocr_struct.get("emails") or [""])[0] if ocr_struct.get("emails") else "",
    )

    out = {
        "company": normalize_spaces_inline(company),
        "city": normalize_spaces_inline(gem.get("city", "") or ocr_struct.get("city", "")),
        "ocr_text": ocr_struct.get("raw_text", ""),
        "ocr_lines": ocr_struct.get("lines", []),
        "urls": ocr_struct.get("urls", []),
        "emails": ocr_struct.get("emails", []),
        "phones": ocr_struct.get("phones", []),
        "phone": ocr_struct.get("phone", ""),
        "phone2": ocr_struct.get("phone2", ""),
        "address": ocr_struct.get("address", ""),
        "postal_code": ocr_struct.get("postal_code", ""),
    }

    PROSPECT_CARD_ANALYSIS_CACHE[key] = dict(out)
    return out


def gemini_activity_summary(name: str, naf: str, website: str, company_hint: str = "") -> str:
    if not GEMINI_API_KEY:
        return ""

    cache_key = f"{name}|{naf}|{website}|{company_hint}"
    if cache_key in GEMINI_ACTIVITY_CACHE:
        return GEMINI_ACTIVITY_CACHE[cache_key]

    if not (name or naf or website or company_hint):
        return ""

    endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
    prompt = (
        "Tu es un normaliseur d'activité d'entreprise.\n"
        "Retourne STRICTEMENT un JSON avec la clé exacte : activity_summary\n"
        "Règles :\n"
        "- activité très courte et factuelle\n"
        "- maximum 60 caractères\n"
        "- pas de marketing\n"
        "- si inconnu : \"\"\n"
        f"Nom entreprise : {name}\n"
        f"Code NAF : {naf}\n"
        f"Site web : {website}\n"
        f"Indice complémentaire : {company_hint}\n"
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 120}
    }

    try:
        STATS["gemini_activity_calls"] += 1
        r = requests.post(endpoint, params={"key": GEMINI_API_KEY}, json=payload, timeout=30)
        if r.status_code >= 300:
            return ""
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
            return ""
        d = json.loads(m.group(0))
        out = safe_strip(d.get("activity_summary", ""))[:60]
        GEMINI_ACTIVITY_CACHE[cache_key] = out
        return out
    except Exception:
        return ""


def infer_activity_fallback(name: str, website: str = "", naf: str = "") -> str:
    n = normalize_text(name)
    w = normalize_text(domain_from_url(website))
    z = f"{n} {w} {naf}".strip()

    rules = [
        (["atelier delaye", "menuiserie delaye", "menuiserie", "bois"], "MENUISERIE"),
        (["vert marine", "aqualone", "aquatique", "piscine"], "CENTRE AQUATIQUE"),
        (["cheval", "transport", "travaux publics", "benne"], "TRANSPORT / TP"),
        (["serrurerie", "boret"], "SERRURERIE / MÉTALLERIE"),
        (["resine", "revetement"], "RÉSINE / REVÊTEMENTS"),
        (["euromaster", "pneu", "garage"], "PNEUMATIQUES / AUTO"),
        (["environnement", "dechets", "recyclage"], "ENVIRONNEMENT / DÉCHETS"),
        (["keolis", "voyageurs"], "TRANSPORT VOYAGEURS"),
        (["compagnons", "devoir", "formation"], "FORMATION"),
    ]

    for keys, label in rules:
        if any(k in z for k in keys):
            return label

    naf_map = {
        "4332B": "MENUISERIE",
        "5610C": "RESTAURATION",
        "3811Z": "COLLECTE DE DÉCHETS",
        "4531Z": "PNEUMATIQUES / AUTO",
        "4333Z": "REVÊTEMENTS / RÉSINE",
        "4399D": "TRAVAUX SPÉCIALISÉS",
        "2512Z": "SERRURERIE / MÉTALLERIE",
        "8559A": "FORMATION",
    }
    return naf_map.get((naf or "").upper(), "")


def enrich_activity_summary(name: str, naf: str, website: str, company_hint: str = "") -> str:
    val = gemini_activity_summary(name=name, naf=naf, website=website, company_hint=company_hint)
    if val:
        return val
    return infer_activity_fallback(name=name, website=website, naf=naf)
    # ============================================================
# BLOC 4 / 6 — WORKER DUMP / TELEGRAM / API GOUV / PLACES / SCRAPE / RANKING
# ============================================================

def worker_dump(kind: str, date: str) -> List[Dict[str, Any]]:
    if not WORKER_BASE_URL or not EXPORT_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN")

    STATS["worker_dump_calls"] += 1
    url = f"{WORKER_BASE_URL.rstrip('/')}/dump?date={date}&kind={kind}"
    r = requests.get(url, headers={"X-Export-Token": EXPORT_TOKEN}, timeout=45)

    if r.status_code != 200:
        raise RuntimeError(f"/dump failed {kind} {r.status_code}: {r.text[:400]}")

    return parse_json_lines(r.text)


def tg_get_file_path(file_id: str) -> Optional[str]:
    if not file_id:
        return None
    if file_id in TG_FILE_PATH_CACHE:
        return TG_FILE_PATH_CACHE[file_id]
    if not TELEGRAM_TOKEN:
        raise RuntimeError("Missing TELEGRAM_TOKEN")

    STATS["telegram_getfile_calls"] += 1
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile"
    r = requests.post(url, json={"file_id": file_id}, timeout=25)
    r.raise_for_status()
    j = r.json()
    if not j.get("ok"):
        return None
    fp = j["result"]["file_path"]
    TG_FILE_PATH_CACHE[file_id] = fp
    return fp


def tg_download_file(file_id: str) -> bytes:
    if not file_id:
        return b""
    if file_id in TG_FILE_BYTES_CACHE:
        return TG_FILE_BYTES_CACHE[file_id]

    file_path = tg_get_file_path(file_id)
    if not file_path:
        return b""

    STATS["telegram_file_downloads"] += 1
    r = requests.get(f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_path}", timeout=60)
    r.raise_for_status()
    content = r.content
    TG_FILE_BYTES_CACHE[file_id] = content
    return content


# ============================================================
# API GOUV — HELPERS
# ============================================================

def extract_dirigeant_from_gouv_result(e: Dict[str, Any]) -> str:
    dirigeant = ""
    arr = e.get("dirigeants") or e.get("representants") or []
    if arr:
        first = arr[0]
        if isinstance(first, str):
            dirigeant = first
        elif isinstance(first, dict):
            p = first.get("personne") or first
            prenom = p.get("prenom") or p.get("prenoms") or ""
            nom = p.get("nom") or p.get("nom_usage") or p.get("nomNaissance") or ""
            dirigeant = f"{prenom} {nom}".strip()
    return normalize_spaces_inline(dirigeant)


def result_to_gouv_record(e: Dict[str, Any], fallback_name: str = "", fallback_city: str = "") -> Dict[str, str]:
    siege = e.get("siege") or {}

    naf = safe_strip(e.get("activite_principale") or e.get("naf") or "").replace(".", "").replace(" ", "").upper()

    return {
        "name": e.get("nom_raison_sociale") or e.get("denomination") or fallback_name,
        "siret": validate_siret(siege.get("siret") or ""),
        "naf": naf,
        "address": safe_strip(siege.get("adresse") or siege.get("libelle_voie") or ""),
        "postal_code": normalize_cp(siege.get("code_postal") or ""),
        "city": safe_strip(siege.get("libelle_commune") or fallback_city),
        "dirigeant": extract_dirigeant_from_gouv_result(e),
    }


def gouv_search_raw(query: str, per_page: int = 8) -> List[Dict[str, Any]]:
    query = normalize_spaces_inline(query)
    if not query:
        return []

    cache_key = f"gouv_raw|{query}|{per_page}"
    if cache_key in GOUV_CACHE:
        cached = GOUV_CACHE[cache_key]
        if isinstance(cached, dict):
            return list(cached.get("_results", []) or [])

    try:
        STATS["gouv_calls"] += 1
        url = f"https://recherche-entreprises.api.gouv.fr/search?q={requests.utils.quote(query)}&page=1&per_page={per_page}"
        r = requests.get(url, headers={"accept": "application/json"}, timeout=20)
        r.raise_for_status()
        j = r.json()
        results = j.get("results") or []
        GOUV_CACHE[cache_key] = {"_results": results}
        return list(results)
    except Exception:
        return []


def search_gouv_company(name: str, city: str = "") -> Dict[str, str]:
    name = normalize_spaces_inline(name)
    city = normalize_spaces_inline(city)
    if not name:
        return {}

    cache_key = f"search_gouv_company|{name}|{city}"
    if cache_key in GOUV_CACHE:
        return dict(GOUV_CACHE[cache_key])

    q = name if not city else f"{name} {city}".strip()
    results = gouv_search_raw(q, per_page=1)
    if not results:
        GOUV_CACHE[cache_key] = {}
        return {}

    out = result_to_gouv_record(results[0], fallback_name=name, fallback_city=city)
    GOUV_CACHE[cache_key] = dict(out)
    return dict(out)


def build_company_candidates(
    company_raw: str,
    email: str,
    website: str,
    city: str,
    ocr_text: str,
    activity_hint: str = "",
    postal_code_hint: str = "",
    phone_hint: str = "",
) -> List[str]:
    cands: List[str] = []

    raw = normalize_spaces_inline(company_raw)
    dom_email = brand_from_domain(domain_from_email(email))
    dom_web = brand_from_domain(domain_from_url(website))
    city = normalize_spaces_inline(city)
    activity_hint = normalize_spaces_inline(activity_hint)
    postal_code_hint = normalize_cp(postal_code_hint)
    phone_hint = normalize_phone(phone_hint)

    for x in [raw, dom_email, dom_web]:
        x = normalize_spaces_inline(x)
        if x and x not in cands:
            cands.append(x)

    base = cands[:]
    for x in base:
        for extra in [city, activity_hint, postal_code_hint]:
            y = normalize_spaces_inline(f"{x} {extra}")
            if y and y not in cands:
                cands.append(y)

    if city and activity_hint:
        q = normalize_spaces_inline(f"{city} {activity_hint}")
        if q and q not in cands:
            cands.append(q)

    if raw and city:
        q = normalize_spaces_inline(f"{raw} {city}")
        if q and q not in cands:
            cands.append(q)

    if raw and postal_code_hint:
        q = normalize_spaces_inline(f"{raw} {postal_code_hint}")
        if q and q not in cands:
            cands.append(q)

    if phone_hint and city:
        q = normalize_spaces_inline(f"{phone_hint} {city}")
        if q and q not in cands:
            cands.append(q)

    if phone_hint and activity_hint:
        q = normalize_spaces_inline(f"{phone_hint} {activity_hint}")
        if q and q not in cands:
            cands.append(q)

    if postal_code_hint and phone_hint:
        q = normalize_spaces_inline(f"{postal_code_hint} {phone_hint}")
        if q and q not in cands:
            cands.append(q)

    if ocr_text:
        lines = [normalize_spaces_inline(ln) for ln in ocr_text.splitlines() if normalize_spaces_inline(ln)]
        existing = {normalize_company_name(c) for c in cands}
        for ln in lines[:10]:
            if len(ln) < 4 or len(ln) > 60:
                continue
            if "@" in ln or "www" in ln.lower():
                continue
            if re.search(r"\d", ln):
                continue
            norm = normalize_company_name(ln)
            if len(norm) >= 4 and norm not in existing:
                cands.append(ln)
                break

    out = []
    seen = set()
    for c in cands:
        key = normalize_company_name(c)
        if key and key not in seen:
            seen.add(key)
            out.append(c)

    return out[:12]


def search_gouv_by_address_hint(address: str, postal_code: str, city: str, website: str, email: str) -> Dict[str, str]:
    parts = []

    if safe_strip(address):
        parts.append(safe_strip(address))
    if normalize_cp(postal_code):
        parts.append(normalize_cp(postal_code))
    if safe_strip(city):
        parts.append(safe_strip(city))

    dom_web = brand_from_domain(domain_from_url(website))
    dom_email = brand_from_domain(domain_from_email(email))

    if dom_web:
        parts.append(dom_web)
    elif dom_email:
        parts.append(dom_email)

    q = " ".join([p for p in parts if p]).strip()
    if not q:
        return {}

    cache_key = f"search_gouv_by_address_hint|{q}"
    if cache_key in GOUV_CACHE:
        return dict(GOUV_CACHE[cache_key])

    results = gouv_search_raw(q, per_page=6)
    if not results:
        GOUV_CACHE[cache_key] = {}
        return {}

    city_n = normalize_text(city)
    best_score = -999
    best = {}

    for e in results:
        rec = result_to_gouv_record(e)
        score = 0

        if city_n:
            score += score_city_match(city, rec.get("city", ""))

        for tok in [dom_web, dom_email]:
            tok_n = normalize_company_name(tok)
            if tok_n and tok_n in normalize_company_name(rec.get("name", "")):
                score += 25

        if normalize_cp(postal_code) and normalize_cp(postal_code) == normalize_cp(rec.get("postal_code", "")):
            score += 20

        if normalize_text(address) and normalize_text(rec.get("address", "")):
            common = 0
            for tok in normalize_text(address).split():
                if len(tok) >= 3 and tok in normalize_text(rec.get("address", "")):
                    common += 1
            score += min(common * 5, 20)

        if score > best_score:
            best_score = score
            best = rec

    GOUV_CACHE[cache_key] = dict(best)
    return dict(best)


def search_gouv_company_ranked(
    company_raw: str,
    email: str,
    website: str,
    city: str,
    ocr_text: str,
    address_hint: str = "",
    postal_code_hint: str = "",
    activity_hint: str = "",
    phone_hint: str = "",
) -> Dict[str, str]:
    candidates = build_company_candidates(
        company_raw=company_raw,
        email=email,
        website=website,
        city=city,
        ocr_text=ocr_text,
        activity_hint=activity_hint,
        postal_code_hint=postal_code_hint,
        phone_hint=phone_hint,
    )
    if not candidates:
        return {}

    best_score = -999
    best = {}

    for q in candidates:
        cache_key = f"search_gouv_company_ranked|{q}|{city}|{postal_code_hint}|{activity_hint}|{phone_hint}"
        if cache_key in GOUV_CACHE:
            results = GOUV_CACHE[cache_key].get("_results") or []
        else:
            results = gouv_search_raw(q, per_page=8)
            GOUV_CACHE[cache_key] = {"_results": results}

        for e in results:
            rec = result_to_gouv_record(e, fallback_name=company_raw, fallback_city=city)
            score = score_candidate_from_signals(
                {
                    "name_parts": tokenize_search_input(company_raw),
                    "city_hints": [city] if city else [],
                    "activity_hints": [activity_hint] if activity_hint else [],
                    "postal_codes": [postal_code_hint] if postal_code_hint else [],
                    "phones": [phone_hint] if phone_hint else [],
                    "emails": [email] if email else [],
                },
                rec,
            )

            if address_hint and rec.get("address"):
                common = 0
                for tok in normalize_text(address_hint).split():
                    if len(tok) >= 3 and tok in normalize_text(rec.get("address", "")):
                        common += 1
                score += min(common * 5, 20)

            if company_raw and normalize_company_name(company_raw) and normalize_company_name(company_raw) in normalize_company_name(rec.get("name", "")):
                score += 25

            if score > best_score:
                best_score = score
                best = rec

    return dict(best)


# ============================================================
# GOOGLE PLACES
# ============================================================

def places_enrich(name: str, city: str = "") -> Dict[str, str]:
    name = normalize_spaces_inline(name)
    city = normalize_spaces_inline(city)

    if not GOOGLE_PLACES_API_KEY or not name:
        return {}

    cache_key = f"{name}|{city}"
    if cache_key in PLACES_CACHE:
        return dict(PLACES_CACHE[cache_key])

    try:
        STATS["places_calls"] += 1
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
            timeout=20,
        )
        r1.raise_for_status()
        j1 = r1.json()
        p = (j1.get("places") or [None])[0]
        if not p or not p.get("id"):
            PLACES_CACHE[cache_key] = {}
            return {}

        pid = p["id"]
        website = p.get("websiteUri") or ""

        r2 = requests.get(
            f"https://places.googleapis.com/v1/places/{pid}",
            headers={
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "nationalPhoneNumber,internationalPhoneNumber,websiteUri,formattedAddress",
            },
            timeout=20,
        )
        r2.raise_for_status()
        j2 = r2.json()

        phone = j2.get("nationalPhoneNumber") or j2.get("internationalPhoneNumber") or ""
        website2 = j2.get("websiteUri") or website
        addr = safe_strip(j2.get("formattedAddress") or "")

        out = {
            "phone": normalize_phone(phone),
            "website": normalize_spaces_inline(website2),
            "address": addr,
        }
        PLACES_CACHE[cache_key] = dict(out)
        return dict(out)
    except Exception:
        return {}


# ============================================================
# SCRAPING EMAIL
# ============================================================

def fetch_html(url: str, max_bytes: int = 250000) -> str:
    try:
        STATS["scrape_calls"] += 1
        r = requests.get(
            url,
            headers={"user-agent": "Mozilla/5.0 (ProspectionBot)"},
            timeout=10,
        )
        if r.status_code >= 300:
            return ""
        return r.text[:max_bytes]
    except Exception:
        return ""


def scrape_email_from_site(url: str) -> str:
    if not url:
        return ""

    domain = domain_from_url(url)
    if domain in EMAIL_SCRAPE_CACHE:
        cached = EMAIL_SCRAPE_CACHE[domain]
        return "" if cached == "none" else cached

    try:
        base = url.rstrip("/")
        candidates = [
            url,
            base + "/contact",
            base + "/nous-contacter",
            base + "/mentions-legales",
            base + "/legal",
        ]

        for attempt_url in candidates:
            html = fetch_html(attempt_url)
            if not html:
                continue

            matches = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", html)
            if not matches:
                continue

            valid = []
            for m in matches:
                em = clean_email(m)
                if not is_valid_email(em):
                    continue
                if re.search(r"no-?reply", em, re.I):
                    continue
                valid.append(em)

            if not valid:
                continue

            if domain:
                same_domain = [em for em in valid if domain_from_email(em) == domain]
                if same_domain:
                    EMAIL_SCRAPE_CACHE[domain] = same_domain[0]
                    return same_domain[0]

            EMAIL_SCRAPE_CACHE[domain] = valid[0]
            return valid[0]
    except Exception:
        pass

    if domain:
        EMAIL_SCRAPE_CACHE[domain] = "none"
    return ""


# ============================================================
# RECHERCHE MULTI-CRITÈRES HAUT NIVEAU
# ============================================================

def search_candidates_from_free_input(raw_input: str) -> List[Dict[str, Any]]:
    raw_input = normalize_spaces_inline(raw_input)
    if not raw_input:
        return []

    cache_key = f"search_candidates_from_free_input|{raw_input}"
    if cache_key in SEARCH_MATCH_CACHE:
        cached = SEARCH_MATCH_CACHE[cache_key]
        if isinstance(cached, dict):
            return list(cached.get("results", []) or [])

    parsed = classify_search_tokens(raw_input)
    if not has_minimum_search_signals(parsed):
        SEARCH_MATCH_CACHE[cache_key] = {"results": []}
        return []

    queries = build_search_queries(parsed)
    gathered: List[Dict[str, Any]] = []

    for q in queries:
        results = gouv_search_raw(q, per_page=8)
        for e in results:
            rec = result_to_gouv_record(e)
            rec["_query"] = q
            rec["_score"] = score_candidate_from_signals(parsed, rec)
            gathered.append(rec)

    # déduplication légère avant tri
    dedup_map: Dict[str, Dict[str, Any]] = {}
    for r in gathered:
        k = validate_siret(r.get("siret", "")) or (
            normalize_company_name(r.get("name", "")) + "|" +
            normalize_text(r.get("city", "")) + "|" +
            normalize_cp(r.get("postal_code", ""))
        )
        prev = dedup_map.get(k)
        if prev is None or int(r.get("_score", 0)) > int(prev.get("_score", 0)):
            dedup_map[k] = r

    out = sorted(dedup_map.values(), key=lambda x: int(x.get("_score", 0)), reverse=True)
    SEARCH_MATCH_CACHE[cache_key] = {"results": out[:12]}
    return out[:12]


def best_gouv_match_from_free_input(raw_input: str) -> Dict[str, Any]:
    results = search_candidates_from_free_input(raw_input)
    return dict(results[0]) if results else {}


# ============================================================
# ENRICHISSEMENT GLOBAL ENTREPRISE
# ============================================================

def enrich_company_from_identifiers(
    company_raw: str = "",
    city: str = "",
    email: str = "",
    website: str = "",
    phone: str = "",
    ocr_text: str = "",
    address_hint: str = "",
    postal_code_hint: str = "",
    activity_hint: str = "",
) -> Dict[str, str]:
    company_raw = normalize_spaces_inline(company_raw)
    city = normalize_spaces_inline(city)
    email = clean_email(email)
    website = normalize_spaces_inline(website)
    phone = normalize_phone(phone)
    address_hint = normalize_spaces_inline(address_hint)
    postal_code_hint = normalize_cp(postal_code_hint)
    activity_hint = normalize_spaces_inline(activity_hint)

    gouv = {}

    if company_raw or city or phone or activity_hint or postal_code_hint or email:
        gouv = search_gouv_company_ranked(
            company_raw=company_raw,
            email=email,
            website=website,
            city=city,
            ocr_text=ocr_text,
            address_hint=address_hint,
            postal_code_hint=postal_code_hint,
            activity_hint=activity_hint,
            phone_hint=phone,
        )

    if not gouv and address_hint:
        gouv = search_gouv_by_address_hint(
            address=address_hint,
            postal_code=postal_code_hint,
            city=city,
            website=website,
            email=email,
        )

    if not gouv and company_raw:
        gouv = search_gouv_company(company_raw, city)

    place = {}
    place_query_name = safe_strip(gouv.get("name", "") or company_raw or brand_from_domain(domain_from_email(email)) or brand_from_domain(domain_from_url(website)))
    if place_query_name:
        place = places_enrich(place_query_name, city) if city else places_enrich(place_query_name, "")

    final = {
        "name": safe_strip(gouv.get("name", "") or company_raw),
        "siret": validate_siret(gouv.get("siret", "")),
        "naf": safe_strip(gouv.get("naf", "")).replace(".", "").replace(" ", "").upper(),
        "address": safe_strip(
            address_hint or
            gouv.get("address", "") or
            place.get("address", "")
        ),
        "postal_code": normalize_cp(postal_code_hint or gouv.get("postal_code", "")),
        "city": safe_strip(city or gouv.get("city", "")),
        "dirigeant": safe_strip(gouv.get("dirigeant", "")),
        "website": normalize_spaces_inline(website or place.get("website", "")),
        "phone": normalize_phone(phone or place.get("phone", "")),
        "email": clean_email(email),
    }

    if not final["email"] and final["website"]:
        scraped = scrape_email_from_site(final["website"])
        if is_valid_email(scraped):
            final["email"] = clean_email(scraped)

    if not final["website"] and final["email"]:
        dom = domain_from_email(final["email"])
        if dom and dom not in FREE_MAIL_DOMAINS:
            final["website"] = f"https://{dom}"

    final["address"] = strip_postal_city_from_address(
        final.get("address", ""),
        final.get("postal_code", ""),
        final.get("city", ""),
    )

    return final


def enrich_company_from_worker_choice(choice: Dict[str, Any]) -> Dict[str, str]:
    choice = choice or {}

    company_name = safe_strip(choice.get("nom", "") or choice.get("name", ""))
    city = safe_strip(choice.get("ville", "") or choice.get("city", ""))
    address = safe_strip(choice.get("adresse", "") or choice.get("address", ""))
    postal_code = normalize_cp(choice.get("codePostal", "") or choice.get("postal_code", ""))
    siret = validate_siret(choice.get("siret", ""))
    naf = safe_strip(choice.get("naf", "")).replace(".", "").replace(" ", "").upper()

    enriched = enrich_company_from_identifiers(
        company_raw=company_name,
        city=city,
        address_hint=address,
        postal_code_hint=postal_code,
    )

    out = {
        "name": safe_strip(enriched.get("name", "") or company_name),
        "siret": validate_siret(enriched.get("siret", "") or siret),
        "naf": safe_strip(enriched.get("naf", "") or naf).replace(".", "").replace(" ", "").upper(),
        "address": safe_strip(enriched.get("address", "") or address),
        "postal_code": normalize_cp(enriched.get("postal_code", "") or postal_code),
        "city": safe_strip(enriched.get("city", "") or city),
        "dirigeant": safe_strip(enriched.get("dirigeant", "")),
        "website": safe_strip(enriched.get("website", "")),
        "phone": normalize_phone(enriched.get("phone", "")),
        "email": clean_email(enriched.get("email", "")) if is_valid_email(enriched.get("email", "")) else "",
    }

    out["address"] = strip_postal_city_from_address(out["address"], out["postal_code"], out["city"])
    return out
    # ============================================================
# BLOC 5 / 6 — FUSION MÉTIER / CARTES / FAÇADES / HISTORIQUES
# ============================================================

# ------------------------------------------------------------
# LECTURE DES CARTES KV
# ------------------------------------------------------------

def load_cards_for_day(date: str) -> List[Dict[str, Any]]:
    try:
        return worker_dump("cards", date)
    except Exception:
        return []


def group_cards_by_prospect(cards: List[Dict[str, Any]]) -> Dict[str, List[Dict]]:
    groups: Dict[str, List[Dict]] = {}

    for c in cards:
        linked = safe_strip(c.get("linked_prospect_name", ""))
        key = normalize_company_name(linked)

        if not key:
            key = f"__unlinked__|{c.get('chatId','')}|{c.get('ts','')}"

        groups.setdefault(key, []).append(c)

    return groups


# ------------------------------------------------------------
# ANALYSE CARTE KV
# ------------------------------------------------------------

def analyze_card_entry(card: Dict[str, Any]) -> Dict[str, Any]:

    file_id = card.get("file_id") or ""
    if not file_id:
        return {}

    img = tg_download_file(file_id)
    if not img:
        return {}

    analysis = analyze_business_card_bytes(img)

    if not analysis:
        return {}

    return {
        "company": safe_strip(analysis.get("company", "")),
        "name": safe_strip(analysis.get("name", "")),
        "email": clean_email(analysis.get("email", "")),
        "phone": normalize_phone(analysis.get("phone", "")),
        "phone2": normalize_phone(analysis.get("phone2", "")),
        "website": safe_strip(analysis.get("website", "")),
        "address": safe_strip(analysis.get("address", "")),
        "postal_code": normalize_cp(analysis.get("postal_code", "")),
        "city": safe_strip(analysis.get("city", "")),
        "ocr_text": analysis.get("ocr_text", ""),
    }


# ------------------------------------------------------------
# ANALYSE PHOTO FACADE
# ------------------------------------------------------------

def analyze_photo_entry(photo: Dict[str, Any]) -> Dict[str, Any]:

    file_id = photo.get("file_id") or ""
    if not file_id:
        return {}

    img = tg_download_file(file_id)
    if not img:
        return {}

    analysis = analyze_facade_bytes(img)

    return {
        "company": safe_strip(analysis.get("company", "")),
        "city": safe_strip(analysis.get("city", "")),
        "phone": normalize_phone(analysis.get("phone", "")),
        "phone2": normalize_phone(analysis.get("phone2", "")),
        "email": clean_email((analysis.get("emails") or [""])[0] if analysis.get("emails") else ""),
        "website": (analysis.get("urls") or [""])[0] if analysis.get("urls") else "",
        "address": safe_strip(analysis.get("address", "")),
        "postal_code": normalize_cp(analysis.get("postal_code", "")),
        "ocr_text": analysis.get("ocr_text", ""),
    }


# ------------------------------------------------------------
# FUSION PROSPECT + CARTE
# ------------------------------------------------------------

def merge_card_into_prospect(prospect: Dict[str, Any], card_data: Dict[str, Any]) -> Dict[str, Any]:

    p = dict(prospect)

    if not safe_strip(p.get("name")) and safe_strip(card_data.get("company")):
        p["name"] = card_data["company"]

    if not safe_strip(p.get("interlocuteur")) and safe_strip(card_data.get("name")):
        p["interlocuteur"] = card_data["name"]

    if not safe_strip(p.get("contact_firstname")) or not safe_strip(p.get("contact_lastname")):
        fn, ln = split_human_name(card_data.get("name", ""))
        if fn and not safe_strip(p.get("contact_firstname")):
            p["contact_firstname"] = fn
        if ln and not safe_strip(p.get("contact_lastname")):
            p["contact_lastname"] = ln

    if not safe_strip(p.get("email")) and is_valid_email(card_data.get("email")):
        p["email"] = card_data["email"]

    if not safe_strip(p.get("website")) and safe_strip(card_data.get("website")):
        p["website"] = card_data["website"]

    if not safe_strip(p.get("phone")) and safe_strip(card_data.get("phone")):
        p["phone"] = card_data["phone"]

    if not safe_strip(p.get("phone2")) and safe_strip(card_data.get("phone2")):
        p["phone2"] = card_data["phone2"]

    if not safe_strip(p.get("address")) and safe_strip(card_data.get("address")):
        p["address"] = card_data["address"]

    if not safe_strip(p.get("postal_code")) and safe_strip(card_data.get("postal_code")):
        p["postal_code"] = card_data["postal_code"]

    if not safe_strip(p.get("city")) and safe_strip(card_data.get("city")):
        p["city"] = card_data["city"]

    return p


# ------------------------------------------------------------
# PROSPECT FINAL
# ------------------------------------------------------------

def unify_from_prospect(prospect: Dict[str, Any], cards: List[Dict[str, Any]]) -> Dict[str, Any]:

    base = dict(prospect)

    card_data = {}

    if cards:
        for c in cards:
            analysis = analyze_card_entry(c)
            if analysis:
                card_data = merge_card_into_prospect(card_data, analysis)

    merged = merge_card_into_prospect(base, card_data)

    enriched = enrich_company_from_identifiers(
        company_raw=merged.get("name", ""),
        city=merged.get("city", ""),
        email=merged.get("email", ""),
        website=merged.get("website", ""),
        phone=merged.get("phone", ""),
        address_hint=merged.get("address", ""),
        postal_code_hint=merged.get("postal_code", ""),
    )

    merged = merge_rows_keep_best(merged, enriched)

    merged["activity_summary"] = enrich_activity_summary(
        merged.get("name", ""),
        merged.get("naf", ""),
        merged.get("website", ""),
    )

    merged["resume"] = safe_strip(prospect.get("resume", ""))
    merged["commande"] = safe_strip(prospect.get("commande", ""))

    return merged


# ------------------------------------------------------------
# LOT CARTES → PROSPECTS
# ------------------------------------------------------------

def promote_cards_to_prospects(cards: List[Dict[str, Any]]) -> List[Dict[str, Any]]:

    prospects = []

    for c in cards:

        analysis = analyze_card_entry(c)

        if not analysis:
            continue

        if not analysis.get("company"):
            continue

        p = {
            "name": analysis.get("company"),
            "interlocuteur": analysis.get("name"),
            "email": analysis.get("email"),
            "phone": analysis.get("phone"),
            "phone2": analysis.get("phone2"),
            "website": analysis.get("website"),
            "address": analysis.get("address"),
            "postal_code": analysis.get("postal_code"),
            "city": analysis.get("city"),
            "resume": "",
            "commande": "",
        }

        enriched = enrich_company_from_identifiers(
            company_raw=p["name"],
            city=p["city"],
            email=p["email"],
            website=p["website"],
        )

        p = merge_rows_keep_best(p, enriched)

        p["activity_summary"] = enrich_activity_summary(
            p.get("name", ""),
            p.get("naf", ""),
            p.get("website", ""),
        )

        prospects.append(p)

        STATS["lot_cards_promoted"] += 1

    return prospects


# ------------------------------------------------------------
# LOT FACADE → PROSPECTS
# ------------------------------------------------------------

def promote_photos_to_prospects(photos: List[Dict[str, Any]]) -> List[Dict[str, Any]]:

    prospects = []

    for ph in photos:

        analysis = analyze_photo_entry(ph)

        if not analysis:
            continue

        if not analysis.get("company"):
            continue

        p = {
            "name": analysis.get("company"),
            "city": analysis.get("city"),
            "phone": analysis.get("phone"),
            "phone2": analysis.get("phone2"),
            "email": analysis.get("email"),
            "website": analysis.get("website"),
            "address": analysis.get("address"),
            "postal_code": analysis.get("postal_code"),
        }

        enriched = enrich_company_from_identifiers(
            company_raw=p["name"],
            city=p["city"],
            email=p["email"],
            website=p["website"],
        )

        p = merge_rows_keep_best(p, enriched)

        p["activity_summary"] = enrich_activity_summary(
            p.get("name", ""),
            p.get("naf", ""),
            p.get("website", ""),
        )

        prospects.append(p)

        STATS["lot_photos_promoted"] += 1

    return prospects


# ------------------------------------------------------------
# PIPELINE GLOBAL CONSOLIDATION
# ------------------------------------------------------------

def consolidate_all_data(date: str) -> List[Dict[str, Any]]:

    prospects = worker_dump("prospects", date)
    photos = worker_dump("photos", date)

    cards = load_cards_for_day(date)

    card_groups = group_cards_by_prospect(cards)

    rows = []

    for p in prospects:

        key = normalize_company_name(p.get("name", ""))

        related_cards = card_groups.get(key, [])

        row = unify_from_prospect(p, related_cards)

        rows.append(row)

        if related_cards:
            STATS["prospect_card_links_found"] += 1
        else:
            STATS["prospect_card_links_missing"] += 1

    promoted_cards = promote_cards_to_prospects(cards)
    promoted_photos = promote_photos_to_prospects(photos)

    rows.extend(promoted_cards)
    rows.extend(promoted_photos)

    rows = dedupe_rows(rows)

    final_rows = []

    for r in rows:

        r["date"] = date
        r["agency"] = r.get("agency") or AGENCY
        r["initials"] = r.get("initials") or INITIALS

        final_rows.append(finalize_row_for_export(r))

    return final_rows
    # ============================================================
# BLOC 6 / 6 — EXCEL / ZIP / EMAIL / SEND MODES / MAIN
# ============================================================

def to_row(d: Dict[str, Any]) -> List[Any]:
    return [d.get(h, "") for h in HEADERS]


def build_excel_one_sheet(date: str, rows: List[Dict[str, Any]], suffix: str) -> str:
    rows = [finalize_row_for_export(r) for r in rows if any(safe_strip(v) for v in r.values())]
    rows = dedupe_rows(rows)
    rows_sorted = sorted(rows, key=lambda x: record_score(x), reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"
    ws.append(HEADERS)

    for r in rows_sorted:
        ws.append(to_row(r))

    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    autosize(ws)

    filename = os.path.join(OUT_DIR, f"PROSPECTION_{date}_{suffix}.xlsx")
    wb.save(filename)
    return filename


# ============================================================
# ZIP MÉDIAS
# ============================================================

def build_media_zip(
    date: str,
    agency: str,
    initials: str,
    photos: List[Dict[str, Any]],
    cards: List[Dict[str, Any]],
) -> Optional[Tuple[str, bytes]]:

    agency = (agency or "").upper().strip()
    initials = (initials or "").upper().strip()

    photos_u = [
        p for p in photos
        if (p.get("agency") or "").upper() == agency
        and (p.get("user") or p.get("initials") or "").upper() == initials
    ]
    cards_u = [
        c for c in cards
        if (c.get("agency") or "").upper() == agency
        and (c.get("user") or c.get("initials") or "").upper() == initials
    ]

    if not photos_u and not cards_u:
        return None

    photos_u = photos_u[:MAX_PHOTO_IMAGES]
    cards_u = cards_u[:MAX_OCR_IMAGES]

    buf = io.BytesIO()

    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        lines = ["type;file;date;agency;initials;city;comment;geo_lat;geo_lon"]

        for i, p in enumerate(photos_u, start=1):
            fid = p.get("file_id") or ""
            try:
                img = tg_download_file(fid) if fid else b""
            except Exception:
                img = b""
            if not img:
                continue

            fname = f"photos/{date}_{agency}_{initials}_{i:02d}.jpg"
            z.writestr(fname, img)

            geo = p.get("geo") or {}
            comment = safe_strip(p.get("comment") or p.get("meeting") or "")
            city = safe_strip(p.get("city") or "")

            lines.append(
                f"photo;{fname};{date};{agency};{initials};{city};"
                f"{comment};{geo.get('lat') or ''};{geo.get('lon') or ''}"
            )

        for i, c in enumerate(cards_u, start=1):
            fid = c.get("file_id") or ""
            try:
                img = tg_download_file(fid) if fid else b""
            except Exception:
                img = b""
            if not img:
                continue

            fname = f"cards/{date}_{agency}_{initials}_{i:02d}.jpg"
            z.writestr(fname, img)

            comment = safe_strip(c.get("comment") or "")
            lines.append(f"card;{fname};{date};{agency};{initials};;{comment};;")

        z.writestr("index.csv", ("\n".join(lines)).encode("utf-8"))

    zip_name = f"MEDIA_{date}_{agency}_{initials}.zip"
    return zip_name, buf.getvalue()


# ============================================================
# BREVO
# ============================================================

def brevo_send_email(
    to_email: str,
    subject: str,
    html: str,
    attachments: Optional[List[Tuple[str, bytes]]] = None,
):
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
            "content-type": "application/json",
        },
        data=json.dumps(payload),
        timeout=60,
    )

    if r.status_code >= 300:
        raise RuntimeError(f"Brevo send failed {r.status_code}: {r.text[:400]}")

    STATS["emails_sent"] += 1


# ============================================================
# FILTRES DE LIGNES
# ============================================================

def filter_rows_for_individual(rows: List[Dict[str, Any]], agency: str, initials: str) -> List[Dict[str, Any]]:
    agency = (agency or "").upper().strip()
    initials = (initials or "").upper().strip()

    out = []
    for r in rows:
        if (r.get("agency") or "").upper() == agency and (r.get("initials") or "").upper() == initials:
            out.append(r)
    return out


def filter_rows_for_agency(rows: List[Dict[str, Any]], agency: str) -> List[Dict[str, Any]]:
    agency = (agency or "").upper().strip()
    return [r for r in rows if (r.get("agency") or "").upper() == agency]


# ============================================================
# PACKS D’ENVOI
# ============================================================

def send_individual_pack(
    date: str,
    agency: str,
    initials: str,
    rows: List[Dict[str, Any]],
    photos: List[Dict[str, Any]],
    cards: List[Dict[str, Any]],
):
    to_email = email_for_initials(initials)
    if not to_email:
        print(f"[WARN] No email for initials={initials}, skip individual pack.")
        return

    indiv_rows = filter_rows_for_individual(rows, agency, initials)
    indiv_rows = [finalize_row_for_export(r) for r in indiv_rows if any(safe_strip(v) for v in r.values())]

    if not indiv_rows:
        print(f"[INFO] No rows for {agency}/{initials}, skip.")
        return

    xlsx = build_excel_one_sheet(date, indiv_rows, f"INDIV_{agency}_{initials}")

    attachments: List[Tuple[str, bytes]] = []
    with open(xlsx, "rb") as f:
        attachments.append((os.path.basename(xlsx), f.read()))

    media = build_media_zip(date, agency, initials, photos, cards)
    if media:
        attachments.append(media)

    subject = f"Prospection {date} — {agency}/{initials}"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Voici ton export de prospection du <b>{date}</b> pour <b>{agency}/{initials}</b>.</p>"
        f"<ul>"
        f"<li><b>Excel</b> consolidé</li>"
        f"<li><b>ZIP médias</b> si disponible</li>"
        f"</ul>"
        f"<p>— Bot Prospection</p>"
    )

    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] individual pack sent to {to_email} ({agency}/{initials})")


def send_agency_manager_pack(
    date: str,
    agency: str,
    rows: List[Dict[str, Any]],
):
    agencies_cfg = ROUTING.get("agencies") or {}
    cfg = agencies_cfg.get(agency) or {}
    manager = cfg.get("manager") or {}
    to_email = clean_email(manager.get("email") or "")

    if not is_valid_email(to_email):
        print(f"[WARN] No valid manager email for agency={agency} ({to_email})")
        return

    ag_rows = filter_rows_for_agency(rows, agency)
    ag_rows = [finalize_row_for_export(r) for r in ag_rows if any(safe_strip(v) for v in r.values())]

    if not ag_rows:
        print(f"[INFO] No rows for agency={agency}, skip manager mail.")
        return

    xlsx = build_excel_one_sheet(date, ag_rows, f"AGENCE_{agency}")

    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    brevo_send_email(
        to_email,
        f"Prospection {date} — Agence {agency}",
        (
            f"<p>Bonjour,</p>"
            f"<p>Voici le consolidé de prospection du <b>{date}</b> pour l’agence <b>{agency}</b>.</p>"
            f"<p>— Bot Prospection</p>"
        ),
        attachments=attachments,
    )
    print(f"[OK] agency manager pack sent to {to_email} (agency={agency})")


def send_admin_pack(
    date: str,
    rows: List[Dict[str, Any]],
):
    admin = ROUTING.get("admin") or {}
    to_email = clean_email(admin.get("email") or "")

    if not is_valid_email(to_email):
        print(f"[WARN] No valid admin email configured ({to_email})")
        return

    rows = [finalize_row_for_export(r) for r in rows if any(safe_strip(v) for v in r.values())]

    if not rows:
        print(f"[INFO] No rows for admin date={date}, skip.")
        return

    xlsx = build_excel_one_sheet(date, rows, "ADMIN_ALL")

    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    brevo_send_email(
        to_email,
        f"Prospection {date} — ADMIN",
        (
            f"<p>Bonjour,</p>"
            f"<p>Voici le consolidé global de prospection du <b>{date}</b>.</p>"
            f"<p>— Bot Prospection</p>"
        ),
        attachments=attachments,
    )
    print(f"[OK] admin pack sent to {to_email}")


# ============================================================
# MAIN
# ============================================================

def main():
    if not WORKER_BASE_URL or not EXPORT_TOKEN or not TELEGRAM_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN / TELEGRAM_TOKEN")

    run_date = RUN_DATE if RUN_DATE else paris_ymd_fallback()

    print(f"🚀 export_and_mail.py — mode={SEND_MODE} date={run_date} agency={AGENCY} initials={INITIALS}")
    print(f"📦 OUT_DIR={OUT_DIR}")
    print(f"🔑 Gemini={'ON' if GEMINI_API_KEY else 'OFF'} | Places={'ON' if GOOGLE_PLACES_API_KEY else 'OFF'}")

    # Données brutes Worker
    photos = worker_dump("photos", run_date)
    cards = worker_dump("cards", run_date)

    # Consolidation avancée
    rows = consolidate_all_data(run_date)

    if SEND_MODE == "individual":
        if AGENCY not in VALID_AGENCIES:
            raise RuntimeError("AGENCY required (GR|VR|GRS|SLS) for mode=individual")
        if not INITIALS:
            raise RuntimeError("INITIALS required for mode=individual")

        send_individual_pack(run_date, AGENCY, INITIALS, rows, photos, cards)
        print(json.dumps(STATS, ensure_ascii=False))
        return

    if SEND_MODE == "agency_manager":
        for ag in sorted(VALID_AGENCIES):
            send_agency_manager_pack(run_date, ag, rows)
        print(json.dumps(STATS, ensure_ascii=False))
        return

    if SEND_MODE == "admin":
        send_admin_pack(run_date, rows)
        print(json.dumps(STATS, ensure_ascii=False))
        return

    raise RuntimeError(f"Unknown SEND_MODE={SEND_MODE}")


if __name__ == "__main__":
    main()