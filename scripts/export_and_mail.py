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

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EMAIL_IN_TEXT_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:(?:\+33|0)[\s\.-]*[1-9](?:[\s\.-]*\d{2}){4})")
URL_RE = re.compile(r"(https?://[^\s)]+|www\.[^\s)]+)", re.I)
CP_RE = re.compile(r"\b((?:0[1-9]|[1-8]\d|9[0-5])\s?\d{3}|97\d{3}|98\d{3})\b")

DISPOSABLE_DOMAINS = {
    "tempmail.com",
    "guerrillamail.com",
    "yopmail.com",
    "10minutemail.com",
}

HEADERS = [
    "date", "agency", "initials", "name", "address", "postal_code", "city",
    "siret", "naf", "activity_summary", "dirigeant", "interlocuteur",
    "contact_firstname", "contact_lastname", "phone", "phone2", "email",
    "website", "resume", "commande"
]


# ============================================================
# ROUTING
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
# RUNTIME STATS / IN-MEMORY CACHE
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
}

GOUV_CACHE: Dict[str, Dict[str, Any]] = {}
PLACES_CACHE: Dict[str, Dict[str, str]] = {}
EMAIL_SCRAPE_CACHE: Dict[str, str] = {}
GEMINI_ACTIVITY_CACHE: Dict[str, str] = {}
GEMINI_CARD_CACHE: Dict[str, Dict[str, str]] = {}
GEMINI_FACADE_CACHE: Dict[str, Dict[str, str]] = {}
TG_FILE_PATH_CACHE: Dict[str, str] = {}
TG_FILE_BYTES_CACHE: Dict[str, bytes] = {}
CARD_PIPELINE_CACHE: Dict[str, Dict[str, str]] = {}


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
# NORMALISATION / VALIDATION
# ============================================================

def validate_siret(raw: str) -> str:
    siret = re.sub(r"\D", "", (raw or "").strip())
    if len(siret) != 14 or not siret.isdigit():
        if siret:
            print(f"[WARN] SIRET invalid length/non-digit: {raw}")
        return ""

    total = 0
    for i in range(14):
        digit = int(siret[13 - i])
        if i % 2 == 1:
            digit *= 2
        if digit > 9:
            digit -= 9
        total += digit

    if total % 10 == 0:
        return siret

    print(f"[WARN] SIRET invalid Luhn: {raw}")
    return ""


def normalize_phone(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    s = re.sub(r"[^\d+]", "", s)
    if s.startswith("+33"):
        s = "0" + s[3:]
    if not re.fullmatch(r"0\d{9}", s or ""):
        return ""
    return s


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
        if pn.startswith(("01", "02", "03", "04", "05")):
            if pn not in fixed:
                fixed.append(pn)
        elif pn.startswith(("06", "07")):
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
    return re.sub(r"\s+", "", (cp or "").strip())


def extract_postal_code(text: str) -> str:
    if not text:
        return ""
    cps = CP_RE.findall(text)
    return normalize_cp(cps[0]) if cps else ""


def domain_from_email(email: str) -> str:
    email = (email or "").strip().lower()
    if "@" not in email:
        return ""
    dom = email.split("@", 1)[1].strip()
    dom = re.sub(r"[^a-z0-9.\-]", "", dom)
    if dom in {
        "gmail.com", "outlook.com", "hotmail.com", "yahoo.com", "yahoo.fr",
        "icloud.com", "free.fr", "orange.fr", "laposte.net"
    }:
        return ""
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
    dom = (dom or "").strip().lower()
    if not dom:
        return ""
    dom = dom.split(".", 1)[0]
    dom = dom.replace("-", " ").replace("_", " ")
    dom = re.sub(r"\s+", " ", dom).strip()
    return dom


def normalize_text(s: str) -> str:
    s = (s or "").lower().strip()
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
    stop = {
        "sas", "sasu", "sa", "sarl", "eurl", "scop", "scp", "snc", "sci",
        "societe", "compagnie", "groupe", "generale", "france", "holding", "services"
    }
    parts = [p for p in s.split() if p not in stop]
    return " ".join(parts).strip()


def deduce_company(company_raw: str, email: str, website: str) -> str:
    company_raw = (company_raw or "").strip()
    dom_email = brand_from_domain(domain_from_email(email))
    dom_web = brand_from_domain(domain_from_url(website))

    for cand in [company_raw, dom_web, dom_email]:
        cand = re.sub(r"\s+", " ", (cand or "").strip())
        if len(cand) >= 4:
            return cand

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


def is_probable_person_name(s: str) -> bool:
    s = re.sub(r"\s+", " ", (s or "").strip())
    if not s:
        return False
    low = normalize_text(s)
    if "@" in s or "www" in low:
        return False
    if re.search(r"\d", s):
        return False

    bad_exact = {"alt gr", "ait gr", "alt qr", "ait qr", "cs", "sas", "sarl", "sa"}
    if low in bad_exact:
        return False

    bad_tokens = {
        "restaurant", "responsable", "commerce", "collectivite", "territoire",
        "recyclage", "valorisation", "dechets", "drh", "developpement", "rh",
        "onyx", "auvergne", "rhone", "alpes", "sas", "sarl", "eurl", "sa",
        "service", "environnement", "gestion", "cedex", "centr", "alp", "cs",
        "experience", "luxe", "scannez", "scanner", "savoir", "plus", "aqualone"
    }
    words = low.split()
    if len(words) < 2 or len(words) > 4:
        return False
    if any(w in bad_tokens for w in words):
        return False

    return sum(1 for w in words if len(w) >= 2) >= 2


def strip_postal_city_from_address(addr: str, postal_code: str, city: str) -> str:
    addr = (addr or "").strip()
    if not addr:
        return ""

    pc = (postal_code or "").strip()
    ct = (city or "").strip()

    addr = re.sub(r"\s+", " ", addr)

    if pc and ct:
        tail = f"{pc} {ct}".strip()
        addr = re.sub(rf"(?:\s|-)?{re.escape(tail)}\s*$", "", addr, flags=re.IGNORECASE).strip()

    if pc:
        addr = re.sub(rf"(?:\s|-)?{re.escape(pc)}\s*$", "", addr).strip()

    addr = re.sub(r"\s+", " ", addr).strip()
    return addr


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

    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
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

        city = after[:60].strip()

        addr_line = re.sub(r"\s+", " ", addr_line).strip()
        city = re.sub(r"\s+", " ", city).strip()

        if cp:
            best_addr, best_cp, best_city = addr_line, cp, city
            break

    return best_addr, best_cp, best_city


def ensure_contact_fallback(d: Dict[str, Any]) -> Dict[str, Any]:
    fn = (d.get("contact_firstname") or "").strip()
    ln = (d.get("contact_lastname") or "").strip()
    interlocuteur = (d.get("interlocuteur") or "").strip()
    dirigeant = (d.get("dirigeant") or "").strip()

    if interlocuteur:
        f, l = split_human_name(interlocuteur)
        if not fn and f:
            d["contact_firstname"] = f
        if not ln and l:
            d["contact_lastname"] = l

    fn = (d.get("contact_firstname") or "").strip()
    ln = (d.get("contact_lastname") or "").strip()

    if (not fn and not ln) and dirigeant:
        f, l = split_human_name(dirigeant)
        d["contact_firstname"] = f or ""
        d["contact_lastname"] = l or (dirigeant or "")

    if not (d.get("interlocuteur") or "").strip():
        combo = f"{(d.get('contact_firstname') or '').strip()} {(d.get('contact_lastname') or '').strip()}".strip()
        if combo:
            d["interlocuteur"] = combo
        elif dirigeant:
            d["interlocuteur"] = dirigeant

    return d


def normalize_upper(v: str) -> str:
    return re.sub(r"\s+", " ", (v or "").strip()).upper()


def clean_ocr_line(line: str) -> str:
    s = re.sub(r"\s+", " ", (line or "").strip())
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

        key = low
        if key in seen:
            continue
        seen.add(key)
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
    if len(ln) > 40:
        return False

    return True


def detect_company_from_lines(lines: List[str], gemini_company: str = "", website: str = "", email: str = "") -> str:
    if gemini_company and len(gemini_company.strip()) >= 3:
        return re.sub(r"\s+", " ", gemini_company).strip()

    dom_web = brand_from_domain(domain_from_url(website))
    dom_email = brand_from_domain(domain_from_email(email))

    candidates = []
    for i, ln in enumerate(lines[:8]):
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
        if len(ln) <= 20:
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
        return candidates[0][1].strip()

    for cand in [dom_web, dom_email]:
        cand = re.sub(r"\s+", " ", (cand or "").strip())
        if len(cand) >= 3:
            return cand

    return ""


def detect_person_name_from_lines(lines: List[str], company_line: str = "") -> str:
    company_n = normalize_company_name(company_line)

    bad_person_phrases = {
        "le bois essence",
        "de l experience du luxe",
        "scannez moi pour en savoir plus",
        "centre aquatique aqualone",
    }

    for ln in lines[:12]:
        low = normalize_text(ln)

        if low in bad_person_phrases:
            continue

        if any(x in low for x in ["experience", "luxe", "scannez", "savoir plus"]):
            continue

        if not is_probable_person_name(ln):
            continue

        if company_n and normalize_company_name(ln) == company_n:
            continue

        return ln.strip()

    return ""


def looks_like_person_line(line: str) -> bool:
    ln = clean_ocr_line(line)
    if not ln:
        return False
    return is_probable_person_name(ln)


def normalize_company_candidate(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def clean_address_candidate(addr: str) -> str:
    a = re.sub(r"\s+", " ", (addr or "").strip())
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


def extract_best_company_hint(lines: List[str], gemini_company: str, website: str, email: str) -> str:
    company = detect_company_from_lines(lines, gemini_company=gemini_company, website=website, email=email)
    return deduce_company(company, email, website)


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
        normalize_company_candidate(gouv_name),
        normalize_company_candidate(gemini_company),
        normalize_company_candidate(detected_company_line),
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
        if looks_like_person_line(cand):
            continue

        score = 0

        if gouv_name and cand.strip() == gouv_name.strip():
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
        return scored[0][1].strip()

    for cand in [dom_web, dom_mail]:
        cand = normalize_company_candidate(cand)
        if len(cand) >= 3:
            return cand

    return ""


def infer_activity_fallback(name: str, website: str = "", naf: str = "") -> str:
    n = normalize_text(name)
    w = normalize_text(domain_from_url(website))
    z = f"{n} {w} {naf}".strip()

    rules = [
        (["atelier delaye", "menuiserie delaye", "menuiserie", "bois"], "MENUISERIE"),
        (["vert marine", "aqualone", "aquatique", "piscine"], "CENTRE AQUATIQUE"),
        (["groupe cheval", "cheval molina", "transport", "travaux publics", "benne"], "TRANSPORT / TRAVAUX PUBLICS"),
        (["serrurerie boret", "boret", "serrurerie"], "SERRURERIE / MÉTALLERIE"),
        (["ad resine", "resine", "revetement"], "RÉSINE / REVÊTEMENTS"),
        (["euromaster", "pneu", "garage"], "PNEUMATIQUES / ENTRETIEN AUTO"),
        (["lely", "environnement", "dechets"], "ENVIRONNEMENT / DECHETS"),
        (["keolis", "transport voyageurs"], "TRANSPORT DE VOYAGEURS"),
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
    }
    if naf in naf_map:
        return naf_map[naf]

    return ""


def enrich_activity_summary(name: str, naf: str, website: str, company_hint: str = "") -> str:
    val = gemini_activity_summary(name=name, naf=naf, website=website, company_hint=company_hint)
    if val:
        return val
    return infer_activity_fallback(name=name, website=website, naf=naf)


def uppercase_business_fields(row: Dict[str, Any]) -> Dict[str, Any]:
    for k in ["name", "address", "city", "dirigeant", "interlocuteur"]:
        row[k] = normalize_upper(row.get(k, ""))
    return row


# ============================================================
# HTTP HELPERS
# ============================================================

def worker_dump(kind: str, date: str) -> List[Dict[str, Any]]:
    if not WORKER_BASE_URL or not EXPORT_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN")
    STATS["worker_dump_calls"] += 1
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
# OCR HYBRIDE
# ============================================================

def image_hash_bytes(img_bytes: bytes) -> str:
    return hashlib.sha1(img_bytes or b"").hexdigest()


def pil_to_cv(img: Image.Image) -> np.ndarray:
    arr = np.array(img)
    if len(arr.shape) == 2:
        return arr
    return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)


def preprocess_image_variants(img_bytes: bytes) -> List[Image.Image]:
    variants: List[Image.Image] = []
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    except Exception:
        return variants

    variants.append(img)

    try:
        big = img.resize((img.width * 2, img.height * 2))
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

    return variants


def crop_image_zones(img: Image.Image) -> List[Image.Image]:
    w, h = img.size
    zones = [img]
    zones.append(img.crop((0, 0, w, h // 3)))
    zones.append(img.crop((0, h // 3, w, 2 * h // 3)))
    zones.append(img.crop((0, 2 * h // 3, w, h)))
    zones.append(img.crop((0, 0, w // 2, h)))
    zones.append(img.crop((w // 2, 0, w, h)))
    return zones


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
        return pytesseract.image_to_string(img, lang="fra+eng") or ""
    except Exception:
        return ""


def merge_text_blocks(texts: List[str]) -> str:
    lines = []
    seen = set()
    for txt in texts:
        for ln in (txt or "").splitlines():
            ln = re.sub(r"\s+", " ", ln).strip()
            if not ln:
                continue
            key = ln.lower()
            if key not in seen:
                seen.add(key)
                lines.append(ln)
    return "\n".join(lines)


def ocr_image_bytes(img_bytes: bytes, light: bool = False) -> str:
    if not img_bytes:
        return ""

    variants = preprocess_image_variants(img_bytes)
    all_texts = []

    for base_img in variants[: (2 if light else len(variants))]:
        zones = crop_image_zones(base_img)
        for zone in zones[: (2 if light else len(zones))]:
            t1 = ocr_easyocr_on_pil(zone)
            if t1:
                all_texts.append(t1)
            t2 = ocr_tesseract_on_pil(zone)
            if t2:
                all_texts.append(t2)

    return merge_text_blocks(all_texts)


# ============================================================
# GEMINI
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
        return json.loads(m.group(0))
    except Exception:
        return {}


def gemini_extract_business_card(img_bytes: bytes) -> Dict[str, str]:
    if not img_bytes:
        return {}
    key = image_hash_bytes(img_bytes)
    if key in GEMINI_CARD_CACHE:
        return GEMINI_CARD_CACHE[key]

    prompt = (
        "Tu es un extracteur de carte de visite.\n"
        "Lis l'IMAGE (ne te base PAS uniquement sur un OCR).\n"
        "Retourne STRICTEMENT un JSON avec les clés EXACTES:\n"
        "name, company, title, email, phone, website, postal_code, city, address\n"
        "Règles:\n"
        "- Si inconnu: \"\"\n"
        "- email en minuscules\n"
        "- postal_code: 5 chiffres\n"
        "- name = prénom + nom si visibles\n"
        "- address = adresse la plus propre possible\n"
    )
    d = gemini_vision_json(img_bytes, prompt, max_tokens=700) or {}
    out = {
        "name": str(d.get("name") or "").strip(),
        "company": str(d.get("company") or "").strip(),
        "title": str(d.get("title") or "").strip(),
        "email": clean_email(str(d.get("email") or "")),
        "phone": normalize_phone(str(d.get("phone") or "")),
        "website": str(d.get("website") or "").strip(),
        "postal_code": normalize_cp(str(d.get("postal_code") or "").strip()),
        "city": str(d.get("city") or "").strip(),
        "address": str(d.get("address") or "").strip(),
    }
    GEMINI_CARD_CACHE[key] = out
    return out


def gemini_extract_facade_logo(img_bytes: bytes) -> Dict[str, str]:
    if not img_bytes:
        return {}
    key = image_hash_bytes(img_bytes)
    if key in GEMINI_FACADE_CACHE:
        return GEMINI_FACADE_CACHE[key]

    prompt = (
        "Tu analyses une photo de prospection (façade, enseigne, logo).\n"
        "Lis l'IMAGE.\n"
        "Retourne STRICTEMENT un JSON avec les clés:\n"
        "company, city\n"
        "Règles: si inconnu -> \"\".\n"
    )
    d = gemini_vision_json(img_bytes, prompt, max_tokens=350) or {}
    out = {
        "company": str(d.get("company") or "").strip(),
        "city": str(d.get("city") or "").strip()
    }
    GEMINI_FACADE_CACHE[key] = out
    return out


def gemini_activity_summary(name: str, naf: str, website: str, company_hint: str = "") -> str:
    if not GEMINI_API_KEY:
        return ""

    cache_key = f"{name}|{naf}|{website}|{company_hint}"
    if cache_key in GEMINI_ACTIVITY_CACHE:
        return GEMINI_ACTIVITY_CACHE[cache_key]

    if not (name or naf or website):
        return ""

    endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
    prompt = (
        "Tu es un normaliseur d'activité d'entreprise.\n"
        "Retourne STRICTEMENT un JSON avec la clé EXACTE: activity_summary\n"
        "Règles:\n"
        "- activité très courte, factuelle\n"
        "- maximum 60 caractères\n"
        "- pas de marketing\n"
        "- si inconnu: \"\"\n"
        f"Nom entreprise: {name}\n"
        f"Code NAF: {naf}\n"
        f"Site web: {website}\n"
        f"Indice complémentaire: {company_hint}\n"
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
        out = str(d.get("activity_summary") or "").strip()[:60]
        GEMINI_ACTIVITY_CACHE[cache_key] = out
        return out
    except Exception:
        return ""
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

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EMAIL_IN_TEXT_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:(?:\+33|0)[\s\.-]*[1-9](?:[\s\.-]*\d{2}){4})")
URL_RE = re.compile(r"(https?://[^\s)]+|www\.[^\s)]+)", re.I)
CP_RE = re.compile(r"\b((?:0[1-9]|[1-8]\d|9[0-5])\s?\d{3}|97\d{3}|98\d{3})\b")

DISPOSABLE_DOMAINS = {
    "tempmail.com",
    "guerrillamail.com",
    "yopmail.com",
    "10minutemail.com",
}

HEADERS = [
    "date", "agency", "initials", "name", "address", "postal_code", "city",
    "siret", "naf", "activity_summary", "dirigeant", "interlocuteur",
    "contact_firstname", "contact_lastname", "phone", "phone2", "email",
    "website", "resume", "commande"
]


# ============================================================
# ROUTING
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
# RUNTIME STATS / IN-MEMORY CACHE
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
}

GOUV_CACHE: Dict[str, Dict[str, Any]] = {}
PLACES_CACHE: Dict[str, Dict[str, str]] = {}
EMAIL_SCRAPE_CACHE: Dict[str, str] = {}
GEMINI_ACTIVITY_CACHE: Dict[str, str] = {}
GEMINI_CARD_CACHE: Dict[str, Dict[str, str]] = {}
GEMINI_FACADE_CACHE: Dict[str, Dict[str, str]] = {}
TG_FILE_PATH_CACHE: Dict[str, str] = {}
TG_FILE_BYTES_CACHE: Dict[str, bytes] = {}
CARD_PIPELINE_CACHE: Dict[str, Dict[str, str]] = {}


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
# NORMALISATION / VALIDATION
# ============================================================

def validate_siret(raw: str) -> str:
    siret = re.sub(r"\D", "", (raw or "").strip())
    if len(siret) != 14 or not siret.isdigit():
        if siret:
            print(f"[WARN] SIRET invalid length/non-digit: {raw}")
        return ""

    total = 0
    for i in range(14):
        digit = int(siret[13 - i])
        if i % 2 == 1:
            digit *= 2
        if digit > 9:
            digit -= 9
        total += digit

    if total % 10 == 0:
        return siret

    print(f"[WARN] SIRET invalid Luhn: {raw}")
    return ""


def normalize_phone(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    s = re.sub(r"[^\d+]", "", s)
    if s.startswith("+33"):
        s = "0" + s[3:]
    if not re.fullmatch(r"0\d{9}", s or ""):
        return ""
    return s


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
        if pn.startswith(("01", "02", "03", "04", "05")):
            if pn not in fixed:
                fixed.append(pn)
        elif pn.startswith(("06", "07")):
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
    return re.sub(r"\s+", "", (cp or "").strip())


def extract_postal_code(text: str) -> str:
    if not text:
        return ""
    cps = CP_RE.findall(text)
    return normalize_cp(cps[0]) if cps else ""


def domain_from_email(email: str) -> str:
    email = (email or "").strip().lower()
    if "@" not in email:
        return ""
    dom = email.split("@", 1)[1].strip()
    dom = re.sub(r"[^a-z0-9.\-]", "", dom)
    if dom in {
        "gmail.com", "outlook.com", "hotmail.com", "yahoo.com", "yahoo.fr",
        "icloud.com", "free.fr", "orange.fr", "laposte.net"
    }:
        return ""
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
    dom = (dom or "").strip().lower()
    if not dom:
        return ""
    dom = dom.split(".", 1)[0]
    dom = dom.replace("-", " ").replace("_", " ")
    dom = re.sub(r"\s+", " ", dom).strip()
    return dom


def normalize_text(s: str) -> str:
    s = (s or "").lower().strip()
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
    stop = {
        "sas", "sasu", "sa", "sarl", "eurl", "scop", "scp", "snc", "sci",
        "societe", "compagnie", "groupe", "generale", "france", "holding", "services"
    }
    parts = [p for p in s.split() if p not in stop]
    return " ".join(parts).strip()


def deduce_company(company_raw: str, email: str, website: str) -> str:
    company_raw = (company_raw or "").strip()
    dom_email = brand_from_domain(domain_from_email(email))
    dom_web = brand_from_domain(domain_from_url(website))

    for cand in [company_raw, dom_web, dom_email]:
        cand = re.sub(r"\s+", " ", (cand or "").strip())
        if len(cand) >= 4:
            return cand

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


def is_probable_person_name(s: str) -> bool:
    s = re.sub(r"\s+", " ", (s or "").strip())
    if not s:
        return False
    low = normalize_text(s)
    if "@" in s or "www" in low:
        return False
    if re.search(r"\d", s):
        return False

    bad_exact = {"alt gr", "ait gr", "alt qr", "ait qr", "cs", "sas", "sarl", "sa"}
    if low in bad_exact:
        return False

    bad_tokens = {
        "restaurant", "responsable", "commerce", "collectivite", "territoire",
        "recyclage", "valorisation", "dechets", "drh", "developpement", "rh",
        "onyx", "auvergne", "rhone", "alpes", "sas", "sarl", "eurl", "sa",
        "service", "environnement", "gestion", "cedex", "centr", "alp", "cs",
        "experience", "luxe", "scannez", "scanner", "savoir", "plus", "aqualone"
    }
    words = low.split()
    if len(words) < 2 or len(words) > 4:
        return False
    if any(w in bad_tokens for w in words):
        return False

    return sum(1 for w in words if len(w) >= 2) >= 2


def strip_postal_city_from_address(addr: str, postal_code: str, city: str) -> str:
    addr = (addr or "").strip()
    if not addr:
        return ""

    pc = (postal_code or "").strip()
    ct = (city or "").strip()

    addr = re.sub(r"\s+", " ", addr)

    if pc and ct:
        tail = f"{pc} {ct}".strip()
        addr = re.sub(rf"(?:\s|-)?{re.escape(tail)}\s*$", "", addr, flags=re.IGNORECASE).strip()

    if pc:
        addr = re.sub(rf"(?:\s|-)?{re.escape(pc)}\s*$", "", addr).strip()

    addr = re.sub(r"\s+", " ", addr).strip()
    return addr


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

    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
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

        city = after[:60].strip()

        addr_line = re.sub(r"\s+", " ", addr_line).strip()
        city = re.sub(r"\s+", " ", city).strip()

        if cp:
            best_addr, best_cp, best_city = addr_line, cp, city
            break

    return best_addr, best_cp, best_city


def ensure_contact_fallback(d: Dict[str, Any]) -> Dict[str, Any]:
    fn = (d.get("contact_firstname") or "").strip()
    ln = (d.get("contact_lastname") or "").strip()
    interlocuteur = (d.get("interlocuteur") or "").strip()
    dirigeant = (d.get("dirigeant") or "").strip()

    if interlocuteur:
        f, l = split_human_name(interlocuteur)
        if not fn and f:
            d["contact_firstname"] = f
        if not ln and l:
            d["contact_lastname"] = l

    fn = (d.get("contact_firstname") or "").strip()
    ln = (d.get("contact_lastname") or "").strip()

    if (not fn and not ln) and dirigeant:
        f, l = split_human_name(dirigeant)
        d["contact_firstname"] = f or ""
        d["contact_lastname"] = l or (dirigeant or "")

    if not (d.get("interlocuteur") or "").strip():
        combo = f"{(d.get('contact_firstname') or '').strip()} {(d.get('contact_lastname') or '').strip()}".strip()
        if combo:
            d["interlocuteur"] = combo
        elif dirigeant:
            d["interlocuteur"] = dirigeant

    return d


def normalize_upper(v: str) -> str:
    return re.sub(r"\s+", " ", (v or "").strip()).upper()


def clean_ocr_line(line: str) -> str:
    s = re.sub(r"\s+", " ", (line or "").strip())
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

        key = low
        if key in seen:
            continue
        seen.add(key)
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
    if len(ln) > 40:
        return False

    return True


def detect_company_from_lines(lines: List[str], gemini_company: str = "", website: str = "", email: str = "") -> str:
    if gemini_company and len(gemini_company.strip()) >= 3:
        return re.sub(r"\s+", " ", gemini_company).strip()

    dom_web = brand_from_domain(domain_from_url(website))
    dom_email = brand_from_domain(domain_from_email(email))

    candidates = []
    for i, ln in enumerate(lines[:8]):
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
        if len(ln) <= 20:
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
        return candidates[0][1].strip()

    for cand in [dom_web, dom_email]:
        cand = re.sub(r"\s+", " ", (cand or "").strip())
        if len(cand) >= 3:
            return cand

    return ""


def detect_person_name_from_lines(lines: List[str], company_line: str = "") -> str:
    company_n = normalize_company_name(company_line)

    bad_person_phrases = {
        "le bois essence",
        "de l experience du luxe",
        "scannez moi pour en savoir plus",
        "centre aquatique aqualone",
    }

    for ln in lines[:12]:
        low = normalize_text(ln)

        if low in bad_person_phrases:
            continue

        if any(x in low for x in ["experience", "luxe", "scannez", "savoir plus"]):
            continue

        if not is_probable_person_name(ln):
            continue

        if company_n and normalize_company_name(ln) == company_n:
            continue

        return ln.strip()

    return ""


def looks_like_person_line(line: str) -> bool:
    ln = clean_ocr_line(line)
    if not ln:
        return False
    return is_probable_person_name(ln)


def normalize_company_candidate(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def clean_address_candidate(addr: str) -> str:
    a = re.sub(r"\s+", " ", (addr or "").strip())
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


def extract_best_company_hint(lines: List[str], gemini_company: str, website: str, email: str) -> str:
    company = detect_company_from_lines(lines, gemini_company=gemini_company, website=website, email=email)
    return deduce_company(company, email, website)


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
        normalize_company_candidate(gouv_name),
        normalize_company_candidate(gemini_company),
        normalize_company_candidate(detected_company_line),
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
        if looks_like_person_line(cand):
            continue

        score = 0

        if gouv_name and cand.strip() == gouv_name.strip():
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
        return scored[0][1].strip()

    for cand in [dom_web, dom_mail]:
        cand = normalize_company_candidate(cand)
        if len(cand) >= 3:
            return cand

    return ""


def infer_activity_fallback(name: str, website: str = "", naf: str = "") -> str:
    n = normalize_text(name)
    w = normalize_text(domain_from_url(website))
    z = f"{n} {w} {naf}".strip()

    rules = [
        (["atelier delaye", "menuiserie delaye", "menuiserie", "bois"], "MENUISERIE"),
        (["vert marine", "aqualone", "aquatique", "piscine"], "CENTRE AQUATIQUE"),
        (["groupe cheval", "cheval molina", "transport", "travaux publics", "benne"], "TRANSPORT / TRAVAUX PUBLICS"),
        (["serrurerie boret", "boret", "serrurerie"], "SERRURERIE / MÉTALLERIE"),
        (["ad resine", "resine", "revetement"], "RÉSINE / REVÊTEMENTS"),
        (["euromaster", "pneu", "garage"], "PNEUMATIQUES / ENTRETIEN AUTO"),
        (["lely", "environnement", "dechets"], "ENVIRONNEMENT / DECHETS"),
        (["keolis", "transport voyageurs"], "TRANSPORT DE VOYAGEURS"),
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
    }
    if naf in naf_map:
        return naf_map[naf]

    return ""


def enrich_activity_summary(name: str, naf: str, website: str, company_hint: str = "") -> str:
    val = gemini_activity_summary(name=name, naf=naf, website=website, company_hint=company_hint)
    if val:
        return val
    return infer_activity_fallback(name=name, website=website, naf=naf)


def uppercase_business_fields(row: Dict[str, Any]) -> Dict[str, Any]:
    for k in ["name", "address", "city", "dirigeant", "interlocuteur"]:
        row[k] = normalize_upper(row.get(k, ""))
    return row


# ============================================================
# HTTP HELPERS
# ============================================================

def worker_dump(kind: str, date: str) -> List[Dict[str, Any]]:
    if not WORKER_BASE_URL or not EXPORT_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN")
    STATS["worker_dump_calls"] += 1
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
# OCR HYBRIDE
# ============================================================

def image_hash_bytes(img_bytes: bytes) -> str:
    return hashlib.sha1(img_bytes or b"").hexdigest()


def pil_to_cv(img: Image.Image) -> np.ndarray:
    arr = np.array(img)
    if len(arr.shape) == 2:
        return arr
    return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)


def preprocess_image_variants(img_bytes: bytes) -> List[Image.Image]:
    variants: List[Image.Image] = []
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    except Exception:
        return variants

    variants.append(img)

    try:
        big = img.resize((img.width * 2, img.height * 2))
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

    return variants


def crop_image_zones(img: Image.Image) -> List[Image.Image]:
    w, h = img.size
    zones = [img]
    zones.append(img.crop((0, 0, w, h // 3)))
    zones.append(img.crop((0, h // 3, w, 2 * h // 3)))
    zones.append(img.crop((0, 2 * h // 3, w, h)))
    zones.append(img.crop((0, 0, w // 2, h)))
    zones.append(img.crop((w // 2, 0, w, h)))
    return zones


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
        return pytesseract.image_to_string(img, lang="fra+eng") or ""
    except Exception:
        return ""


def merge_text_blocks(texts: List[str]) -> str:
    lines = []
    seen = set()
    for txt in texts:
        for ln in (txt or "").splitlines():
            ln = re.sub(r"\s+", " ", ln).strip()
            if not ln:
                continue
            key = ln.lower()
            if key not in seen:
                seen.add(key)
                lines.append(ln)
    return "\n".join(lines)


def ocr_image_bytes(img_bytes: bytes, light: bool = False) -> str:
    if not img_bytes:
        return ""

    variants = preprocess_image_variants(img_bytes)
    all_texts = []

    for base_img in variants[: (2 if light else len(variants))]:
        zones = crop_image_zones(base_img)
        for zone in zones[: (2 if light else len(zones))]:
            t1 = ocr_easyocr_on_pil(zone)
            if t1:
                all_texts.append(t1)
            t2 = ocr_tesseract_on_pil(zone)
            if t2:
                all_texts.append(t2)

    return merge_text_blocks(all_texts)


# ============================================================
# GEMINI
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
        return json.loads(m.group(0))
    except Exception:
        return {}


def gemini_extract_business_card(img_bytes: bytes) -> Dict[str, str]:
    if not img_bytes:
        return {}
    key = image_hash_bytes(img_bytes)
    if key in GEMINI_CARD_CACHE:
        return GEMINI_CARD_CACHE[key]

    prompt = (
        "Tu es un extracteur de carte de visite.\n"
        "Lis l'IMAGE (ne te base PAS uniquement sur un OCR).\n"
        "Retourne STRICTEMENT un JSON avec les clés EXACTES:\n"
        "name, company, title, email, phone, website, postal_code, city, address\n"
        "Règles:\n"
        "- Si inconnu: \"\"\n"
        "- email en minuscules\n"
        "- postal_code: 5 chiffres\n"
        "- name = prénom + nom si visibles\n"
        "- address = adresse la plus propre possible\n"
    )
    d = gemini_vision_json(img_bytes, prompt, max_tokens=700) or {}
    out = {
        "name": str(d.get("name") or "").strip(),
        "company": str(d.get("company") or "").strip(),
        "title": str(d.get("title") or "").strip(),
        "email": clean_email(str(d.get("email") or "")),
        "phone": normalize_phone(str(d.get("phone") or "")),
        "website": str(d.get("website") or "").strip(),
        "postal_code": normalize_cp(str(d.get("postal_code") or "").strip()),
        "city": str(d.get("city") or "").strip(),
        "address": str(d.get("address") or "").strip(),
    }
    GEMINI_CARD_CACHE[key] = out
    return out


def gemini_extract_facade_logo(img_bytes: bytes) -> Dict[str, str]:
    if not img_bytes:
        return {}
    key = image_hash_bytes(img_bytes)
    if key in GEMINI_FACADE_CACHE:
        return GEMINI_FACADE_CACHE[key]

    prompt = (
        "Tu analyses une photo de prospection (façade, enseigne, logo).\n"
        "Lis l'IMAGE.\n"
        "Retourne STRICTEMENT un JSON avec les clés:\n"
        "company, city\n"
        "Règles: si inconnu -> \"\".\n"
    )
    d = gemini_vision_json(img_bytes, prompt, max_tokens=350) or {}
    out = {
        "company": str(d.get("company") or "").strip(),
        "city": str(d.get("city") or "").strip()
    }
    GEMINI_FACADE_CACHE[key] = out
    return out


def gemini_activity_summary(name: str, naf: str, website: str, company_hint: str = "") -> str:
    if not GEMINI_API_KEY:
        return ""

    cache_key = f"{name}|{naf}|{website}|{company_hint}"
    if cache_key in GEMINI_ACTIVITY_CACHE:
        return GEMINI_ACTIVITY_CACHE[cache_key]

    if not (name or naf or website):
        return ""

    endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
    prompt = (
        "Tu es un normaliseur d'activité d'entreprise.\n"
        "Retourne STRICTEMENT un JSON avec la clé EXACTE: activity_summary\n"
        "Règles:\n"
        "- activité très courte, factuelle\n"
        "- maximum 60 caractères\n"
        "- pas de marketing\n"
        "- si inconnu: \"\"\n"
        f"Nom entreprise: {name}\n"
        f"Code NAF: {naf}\n"
        f"Site web: {website}\n"
        f"Indice complémentaire: {company_hint}\n"
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
        out = str(d.get("activity_summary") or "").strip()[:60]
        GEMINI_ACTIVITY_CACHE[cache_key] = out
        return out
    except Exception:
        return ""
        # ============================================================
# MEDIA ZIP
# ============================================================

def build_media_zip(date: str, agency: str, initials: str,
                    photos: List[Dict[str, Any]], cards: List[Dict[str, Any]]) -> Optional[Tuple[str, bytes]]:
    photos_u = [p for p in photos if (p.get("agency") == agency and (p.get("user") or p.get("initials") or "").upper() == initials)]
    cards_u = [c for c in cards if (c.get("agency") == agency and (c.get("user") or c.get("initials") or "").upper() == initials)]
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
            comment = p.get("comment") or p.get("meeting") or ""
            lines.append(
                f"photo;{fname};{date};{agency};{initials};{(p.get('city') or '')};"
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
    STATS["emails_sent"] += 1


# ============================================================
# SEND MODES
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

    rows = build_rows_with_linked_cards(
        date=date,
        prospects=prospects,
        photos=photos,
        cards=cards,
        agency_filter=agency,
        initials_filter=initials,
    )

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

    rows = build_rows_with_linked_cards(
        date=date,
        prospects=prospects,
        photos=photos,
        cards=cards,
        agency_filter=agency,
        initials_filter=None,
    )

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

    rows = build_rows_with_linked_cards(
        date=date,
        prospects=prospects,
        photos=photos,
        cards=cards,
        agency_filter=None,
        initials_filter=None,
    )

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

def attach_kv_keys(records: List[Dict[str, Any]], kind_prefix: str) -> List[Dict[str, Any]]:
    out = []
    for r in records:
        x = dict(r)
        if kind_prefix == "prospect":
            x["_kv_key"] = f"prospect:{x.get('date','')}:{x.get('agency','')}:{x.get('created_at','')}:{x.get('chatId','')}"
        elif kind_prefix == "card":
            x["_kv_key"] = x.get("card_kv_key") or ""
        out.append(x)
    return out


def main():
    if not WORKER_BASE_URL or not EXPORT_TOKEN or not TELEGRAM_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN / TELEGRAM_TOKEN")

    run_date = RUN_DATE if RUN_DATE else paris_ymd_fallback()

    print(f"🚀 export_and_mail.py — mode={SEND_MODE} date={run_date} agency={AGENCY} initials={INITIALS}")
    print(f"📦 OUT_DIR={OUT_DIR}")
    print(f"🔑 Gemini={'ON' if GEMINI_API_KEY else 'OFF'} | Places={'ON' if GOOGLE_PLACES_API_KEY else 'OFF'}")

    prospects = worker_dump("prospects", run_date)
    photos = worker_dump("photos", run_date)
    cards = worker_dump("cards", run_date)

    for p in prospects:
        if "_kv_key" not in p:
            p["_kv_key"] = ""
    for c in cards:
        if "_kv_key" not in c:
            c["_kv_key"] = c.get("card_kv_key", "")

    if SEND_MODE == "individual":
        if AGENCY not in VALID_AGENCIES:
            raise RuntimeError("AGENCY required (GR|VR|GRS|SLS) for mode=individual")
        if not INITIALS:
            raise RuntimeError("INITIALS required for mode=individual")
        send_individual_pack(run_date, AGENCY, INITIALS, prospects, photos, cards)
        print(json.dumps(STATS, ensure_ascii=False))
        return

    if SEND_MODE == "agency_manager":
        for ag in sorted(VALID_AGENCIES):
            send_agency_manager_pack(run_date, ag, prospects, photos, cards)
        print(json.dumps(STATS, ensure_ascii=False))
        return

    if SEND_MODE == "admin":
        send_admin_pack(run_date, prospects, photos, cards)
        print(json.dumps(STATS, ensure_ascii=False))
        return

    raise RuntimeError(f"Unknown SEND_MODE={SEND_MODE}")


if __name__ == "__main__":
    main()
