import os
import re
import json
import time
import base64
import shutil
import tempfile
import datetime
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse, urljoin

import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


# =====================================================
# ENV
# =====================================================
WORKER_BASE_URL = os.environ.get("WORKER_BASE_URL", "").rstrip("/")
EXPORT_TOKEN = os.environ.get("EXPORT_TOKEN", "")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")

SEND_MODE = os.environ.get("SEND_MODE", "individual")  # individual | agency_manager | admin
RUN_DATE = os.environ.get("RUN_DATE", "")              # YYYY-MM-DD Europe/Paris
AGENCY = os.environ.get("AGENCY", "")                  # GR | VR | GRS | SLS
INITIALS = os.environ.get("INITIALS", "")              # JL, CZ...
MAX_OCR_IMAGES = int(os.environ.get("MAX_OCR_IMAGES", "50"))        # legacy (unused here)
MAX_PHOTO_IMAGES = int(os.environ.get("MAX_PHOTO_IMAGES", "15"))    # lot photos V2

# V2 Vision + Places
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GOOGLE_PLACES_API_KEY = os.environ.get("GOOGLE_PLACES_API_KEY", "")

# Email (Brevo)
BREVO_API_KEY = os.environ.get("BREVO_API_KEY", "")
BREVO_SENDER_EMAIL = os.environ.get("BREVO_SENDER_EMAIL", "")
BREVO_SENDER_NAME = os.environ.get("BREVO_SENDER_NAME", "Prospection Bot")
MAIL_ROUTING_JSON = os.environ.get("MAIL_ROUTING_JSON", "")  # mapping recipients

# Admin Telegram (optional)
TELEGRAM_ADMIN_CHAT_ID = os.environ.get("TELEGRAM_ADMIN_CHAT_ID", "")

# Networking
HTTP_TIMEOUT = 20
UA = "prospection-export/2.1"
SCRAPE_MAX_BYTES = 250_000
SCRAPE_MAX_PAGES = 5  # homepage + up to 4 extra pages
SCRAPE_DELAY_S = 0.0  # keep 0 for speed; set 0.2 if needed


# =====================================================
# HELPERS
# =====================================================
def die(msg: str) -> None:
    raise SystemExit(msg)


def paris_today_ymd() -> str:
    if RUN_DATE and re.match(r"^\d{4}-\d{2}-\d{2}$", RUN_DATE):
        return RUN_DATE
    return datetime.date.today().isoformat()


def http_get(url: str, headers: Optional[Dict[str, str]] = None, stream: bool = False) -> requests.Response:
    h = {"user-agent": UA}
    if headers:
        h.update(headers)
    return requests.get(url, headers=h, timeout=HTTP_TIMEOUT, stream=stream, allow_redirects=True)


def http_post(url: str, json_body: Dict[str, Any], headers: Optional[Dict[str, str]] = None) -> requests.Response:
    h = {"user-agent": UA, "content-type": "application/json"}
    if headers:
        h.update(headers)
    return requests.post(url, headers=h, json=json_body, timeout=HTTP_TIMEOUT)


def safe_json_loads(s: str) -> Optional[Dict[str, Any]]:
    try:
        return json.loads(s)
    except Exception:
        m = re.search(r"\{.*\}", s, re.S)
        if not m:
            return None
        try:
            return json.loads(m.group(0))
        except Exception:
            return None


def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def normalize_naf(naf: str) -> str:
    return (naf or "").upper().replace(".", "").replace(" ", "").strip()


def only_digits(s: str) -> str:
    return re.sub(r"\D+", "", s or "")


def normalize_fr_phone(phone: str) -> str:
    p = (phone or "").strip()
    if not p:
        return ""
    p = p.replace("(", " ").replace(")", " ").replace(".", " ").replace("-", " ")
    digits = only_digits(p)

    # +33 / 0033
    if digits.startswith("33") and len(digits) >= 11:
        digits = "0" + digits[2:]
    if digits.startswith("0033") and len(digits) >= 13:
        digits = "0" + digits[4:]

    if len(digits) == 10 and digits.startswith("0"):
        return " ".join([digits[i:i+2] for i in range(0, 10, 2)])
    return (phone or "").strip()


def extract_emails(text: str) -> List[str]:
    if not text:
        return []
    emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    out = []
    for e in emails:
        if re.search(r"no-?reply", e, re.I):
            continue
        out.append(e)
    return list(dict.fromkeys(out))


def extract_fr_phones(text: str) -> List[str]:
    if not text:
        return []
    # broad candidates: +33, 0X..., with spaces allowed
    candidates = re.findall(r"(?:\+33\s?|\b0)(?:[\s\.\-]?\d){9,12}", text)
    norm = []
    for c in candidates:
        n = normalize_fr_phone(c)
        if n:
            norm.append(n)
    # keep uniques
    return list(dict.fromkeys(norm))


def best_phone_candidate(phones: List[str]) -> str:
    """
    Prefer fixed-line (01-05) then others; deprioritize 08/09; mobile 06/07 goes to phone2 generally.
    For main phone, prefer 01-05.
    """
    if not phones:
        return ""
    def score(p: str) -> int:
        d = only_digits(p)
        if len(d) != 10 or not d.startswith("0"):
            return 0
        if d.startswith(("01", "02", "03", "04", "05")):
            return 100
        if d.startswith(("09",)):
            return 60
        if d.startswith(("08",)):
            return 40
        if d.startswith(("06", "07")):
            return 20
        return 10

    best = sorted(phones, key=lambda x: score(x), reverse=True)[0]
    return best


def tokenize_name(s: str) -> List[str]:
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9àâäéèêëîïôöùûüç\s]", " ", s)
    s = normalize_spaces(s)
    toks = [t for t in s.split(" ") if t and t not in ("sas", "sarl", "eurl", "sa", "sasu", "snc", "groupe", "france")]
    return toks


def name_similarity(a: str, b: str) -> float:
    """
    Simple token overlap score 0..1
    """
    ta = set(tokenize_name(a))
    tb = set(tokenize_name(b))
    if not ta or not tb:
        return 0.0
    inter = len(ta & tb)
    union = len(ta | tb)
    return inter / union if union else 0.0


# =====================================================
# WORKER DUMP
# =====================================================
def worker_dump(kind: str, date_ymd: str) -> List[Dict[str, Any]]:
    if not WORKER_BASE_URL or not EXPORT_TOKEN:
        die("Missing WORKER_BASE_URL or EXPORT_TOKEN")

    url = f"{WORKER_BASE_URL}/dump?date={date_ymd}&kind={kind}"
    r = http_get(url, headers={"X-Export-Token": EXPORT_TOKEN})
    if r.status_code != 200:
        die(f"Worker /dump failed kind={kind} status={r.status_code} body={r.text[:500]}")

    lines = [ln.strip() for ln in r.text.splitlines() if ln.strip()]
    out: List[Dict[str, Any]] = []
    for ln in lines:
        try:
            out.append(json.loads(ln))
        except Exception:
            continue
    return out


# =====================================================
# TELEGRAM DOWNLOAD
# =====================================================
def tg_api(method: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    if not TELEGRAM_TOKEN:
        die("Missing TELEGRAM_TOKEN")
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/{method}"
    r = http_post(url, payload)
    j = r.json()
    if not j.get("ok"):
        raise RuntimeError(f"Telegram {method} failed: {j}")
    return j["result"]


def tg_get_file_path(file_id: str) -> str:
    res = tg_api("getFile", {"file_id": file_id})
    file_path = res.get("file_path", "")
    if not file_path:
        raise RuntimeError("Telegram getFile missing file_path")
    return file_path


def tg_download_file(file_path: str, dst_path: str) -> None:
    url = f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_path}"
    r = http_get(url, stream=True)
    if r.status_code != 200:
        raise RuntimeError(f"Telegram file download failed {r.status_code}")
    with open(dst_path, "wb") as f:
        for chunk in r.iter_content(chunk_size=65536):
            if chunk:
                f.write(chunk)


# =====================================================
# REVERSE GEOCODE (Option B)
# =====================================================
def reverse_geocode_city(lat: float, lon: float) -> Optional[str]:
    try:
        url = "https://nominatim.openstreetmap.org/reverse"
        params = {
            "format": "jsonv2",
            "lat": str(lat),
            "lon": str(lon),
            "zoom": "12",
            "addressdetails": "1",
        }
        r = requests.get(url, params=params, headers={"user-agent": "prospection-export/2.1 (contact: admin)"}, timeout=HTTP_TIMEOUT)
        if r.status_code != 200:
            return None
        j = r.json()
        addr = j.get("address", {}) or {}
        city = addr.get("city") or addr.get("town") or addr.get("village") or addr.get("municipality")
        if city:
            return normalize_spaces(city)
        return None
    except Exception:
        return None


# =====================================================
# GEMINI VISION (REST)
# =====================================================
GEMINI_MODEL = "gemini-1.5-flash"


def gemini_vision_extract(image_bytes: bytes) -> Dict[str, Any]:
    if not GEMINI_API_KEY:
        return {
            "kind": "unknown",
            "company_name_guess": None,
            "text_extracted": None,
            "contact": {
                "civility": None,
                "first_name": None,
                "last_name": None,
                "job_title": None,
                "email": None,
                "mobile": None,
                "phone": None,
            },
            "confidence": 0,
            "_note": "GEMINI_API_KEY missing",
        }

    b64 = base64.b64encode(image_bytes).decode("ascii")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"

    prompt = (
        "Tu es un extracteur strict JSON.\n"
        "Analyse l'image (carte de visite / logo / enseigne).\n"
        "Retourne UNIQUEMENT un JSON valide (sans markdown, sans texte autour), au format EXACT :\n"
        "{\n"
        "  \"kind\": \"business_card\" | \"logo\" | \"signage\" | \"unknown\",\n"
        "  \"company_name_guess\": null | string,\n"
        "  \"text_extracted\": null | string,\n"
        "  \"contact\": {\n"
        "    \"civility\": null | string,\n"
        "    \"first_name\": null | string,\n"
        "    \"last_name\": null | string,\n"
        "    \"job_title\": null | string,\n"
        "    \"email\": null | string,\n"
        "    \"mobile\": null | string,\n"
        "    \"phone\": null | string\n"
        "  },\n"
        "  \"confidence\": number\n"
        "}\n"
        "Contraintes:\n"
        "- N'invente jamais email ou téléphone.\n"
        "- confidence entre 0 et 100.\n"
        "- Si tu n'es pas sûr, mets null et baisse confidence.\n"
    )

    body = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": "image/jpeg", "data": b64}},
                ],
            }
        ],
        "generationConfig": {
            "temperature": 0.2,
            "topP": 0.9,
            "maxOutputTokens": 512,
        },
    }

    # retry simple
    last_err = None
    for attempt in range(1, 4):
        try:
            r = http_post(url, body)
            if r.status_code in (429, 500, 502, 503, 504):
                last_err = f"Gemini HTTP {r.status_code}"
                time.sleep(0.6 * attempt)
                continue
            if r.status_code != 200:
                last_err = f"Gemini HTTP {r.status_code}: {r.text[:200]}"
                break
            j = r.json()
            txt = ""
            try:
                txt = j["candidates"][0]["content"]["parts"][0].get("text", "")
            except Exception:
                txt = ""
            data = safe_json_loads(txt) or {}
            kind = data.get("kind") if data.get("kind") in ("business_card", "logo", "signage", "unknown") else "unknown"
            conf = data.get("confidence", 0)
            try:
                conf = int(float(conf))
            except Exception:
                conf = 0
            conf = max(0, min(100, conf))
            contact = data.get("contact") if isinstance(data.get("contact"), dict) else {}
            return {
                "kind": kind,
                "company_name_guess": data.get("company_name_guess"),
                "text_extracted": data.get("text_extracted"),
                "contact": {
                    "civility": contact.get("civility"),
                    "first_name": contact.get("first_name"),
                    "last_name": contact.get("last_name"),
                    "job_title": contact.get("job_title"),
                    "email": contact.get("email"),
                    "mobile": contact.get("mobile"),
                    "phone": contact.get("phone"),
                },
                "confidence": conf,
            }
        except Exception as e:
            last_err = str(e)
            time.sleep(0.6 * attempt)

    return {
        "kind": "unknown",
        "company_name_guess": None,
        "text_extracted": None,
        "contact": {
            "civility": None,
            "first_name": None,
            "last_name": None,
            "job_title": None,
            "email": None,
            "mobile": None,
            "phone": None,
        },
        "confidence": 0,
        "_note": f"Gemini failed: {last_err}",
    }


# =====================================================
# API GOUV ENRICH (deterministic)
# =====================================================
def api_gouv_search(company_name: str, city: str, topn: int = 3) -> List[Dict[str, Any]]:
    q = normalize_spaces(f"{company_name} {city}".strip())
    if not q:
        return []
    url = "https://recherche-entreprises.api.gouv.fr/search"
    params = {"q": q, "page": 1, "per_page": topn}
    r = requests.get(url, params=params, headers={"accept": "application/json", "user-agent": UA}, timeout=HTTP_TIMEOUT)
    if r.status_code != 200:
        return []
    j = r.json()
    results = j.get("results") or []
    out = []
    for e in results:
        siege = e.get("siege") or {}
        out.append({
            "name": e.get("nom_raison_sociale") or e.get("denomination") or "",
            "siret": siege.get("siret") or "",
            "siren": e.get("siren") or (str(siege.get("siret") or "")[:9] if siege.get("siret") else ""),
            "naf": e.get("activite_principale") or e.get("naf") or "",
            "address": siege.get("adresse") or siege.get("libelle_voie") or "",
            "postal_code": siege.get("code_postal") or "",
            "city": siege.get("libelle_commune") or "",
            "score": e.get("score"),
        })
    return out


# =====================================================
# GOOGLE PLACES — V2+ PHONE OPTIMIZATION
# =====================================================
def places_search_text(query: str, max_results: int = 5) -> List[Dict[str, Any]]:
    if not GOOGLE_PLACES_API_KEY:
        return []
    url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
        "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.websiteUri,places.nationalPhoneNumber,places.internationalPhoneNumber",
    }
    body = {
        "textQuery": normalize_spaces(query),
        "maxResultCount": max_results,
        "languageCode": "fr",
        "regionCode": "FR",
    }
    r = requests.post(url, headers=headers, json=body, timeout=HTTP_TIMEOUT)
    if r.status_code != 200:
        return []
    j = r.json()
    return j.get("places") or []


def place_details(place_id: str) -> Dict[str, Any]:
    if not GOOGLE_PLACES_API_KEY or not place_id:
        return {}
    url = f"https://places.googleapis.com/v1/places/{place_id}"
    headers = {
        "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
        "X-Goog-FieldMask": "id,displayName,nationalPhoneNumber,internationalPhoneNumber,websiteUri,formattedAddress,googleMapsUri",
    }
    r = requests.get(url, headers=headers, timeout=HTTP_TIMEOUT)
    if r.status_code != 200:
        return {}
    return r.json() or {}


def best_place_pick(query_name: str, query_city: str, places: List[Dict[str, Any]]) -> Tuple[str, str, List[str]]:
    """
    Score by name similarity + city presence in address. Then enrich using details if needed.
    Return (phone, website, suggestions_top3)
    """
    suggestions = []
    for p in places[:3]:
        dn = (p.get("displayName") or {}).get("text") if isinstance(p.get("displayName"), dict) else p.get("displayName")
        if dn:
            suggestions.append(normalize_spaces(dn))

    scored: List[Tuple[float, Dict[str, Any]]] = []
    for p in places:
        dn = (p.get("displayName") or {}).get("text") if isinstance(p.get("displayName"), dict) else p.get("displayName")
        dn = normalize_spaces(dn or "")
        addr = normalize_spaces(p.get("formattedAddress") or "")
        sim = name_similarity(query_name, dn)
        city_bonus = 0.0
        if query_city and addr and query_city.lower() in addr.lower():
            city_bonus = 0.15
        scored.append((sim + city_bonus, p))

    scored.sort(key=lambda x: x[0], reverse=True)
    best = scored[0][1] if scored else None

    # First try quick fields from searchText response
    phone = ""
    website = ""
    if best:
        phone = best.get("nationalPhoneNumber") or best.get("internationalPhoneNumber") or ""
        website = best.get("websiteUri") or ""

    # If missing phone/website, call details on top 2 candidates to improve coverage
    for _, cand in scored[:2]:
        if phone and website:
            break
        pid = cand.get("id") or ""
        if not pid:
            continue
        det = place_details(pid)
        if not phone:
            phone = det.get("nationalPhoneNumber") or det.get("internationalPhoneNumber") or phone
        if not website:
            website = det.get("websiteUri") or website

    return normalize_fr_phone(phone), (website or "").strip(), suggestions[:3]


# =====================================================
# SCRAPING — V2+ (multi pages)
# =====================================================
def is_same_domain(base: str, candidate: str) -> bool:
    try:
        b = urlparse(base)
        c = urlparse(candidate)
        return (b.netloc and c.netloc and b.netloc.lower() == c.netloc.lower())
    except Exception:
        return False


def extract_candidate_links(html: str, base_url: str) -> List[str]:
    """
    Extract a few useful internal links likely containing contact info.
    """
    if not html:
        return []
    links = re.findall(r'href=["\']([^"\']+)["\']', html, re.I)
    out = []
    for href in links:
        href = href.strip()
        if not href or href.startswith("#"):
            continue
        if href.startswith("mailto:") or href.startswith("tel:"):
            continue
        full = urljoin(base_url, href)
        # keep same domain only
        if not is_same_domain(base_url, full):
            continue
        low = full.lower()
        if any(k in low for k in ["contact", "mentions", "legal", "nous-contacter", "a-propos", "societe", "agence"]):
            out.append(full)
    # de-dup preserve order
    seen = set()
    dedup = []
    for u in out:
        if u in seen:
            continue
        seen.add(u)
        dedup.append(u)
    return dedup[: (SCRAPE_MAX_PAGES - 1)]


def fetch_html_limited(url: str) -> str:
    r = http_get(url, headers={"accept": "text/html,application/xhtml+xml"})
    if r.status_code != 200:
        return ""
    # limit bytes
    txt = r.text
    if len(txt) > SCRAPE_MAX_BYTES:
        txt = txt[:SCRAPE_MAX_BYTES]
    return txt


def scrape_site_for_email_phone_v2plus(url: str) -> Tuple[str, str]:
    """
    Homepage + contact/legal pages fallback.
    Return (email, phone)
    """
    if not url:
        return "", ""
    try:
        base = url.strip()
        html0 = fetch_html_limited(base)
        if not html0:
            return "", ""

        emails = extract_emails(html0)
        phones = extract_fr_phones(html0)

        # add common paths even if not linked
        common_paths = [
            "/contact", "/contacts", "/nous-contacter",
            "/mentions-legales", "/mentions-legale",
            "/legal", "/legals",
        ]
        candidates = extract_candidate_links(html0, base)
        for p in common_paths:
            candidates.append(urljoin(base, p))

        # de-dup
        seen = set()
        pages = [base]
        for u in candidates:
            if u in seen:
                continue
            seen.add(u)
            pages.append(u)
        pages = pages[:SCRAPE_MAX_PAGES]

        # crawl a few pages
        for u in pages[1:]:
            if SCRAPE_DELAY_S:
                time.sleep(SCRAPE_DELAY_S)
            h = fetch_html_limited(u)
            if not h:
                continue
            emails.extend(extract_emails(h))
            phones.extend(extract_fr_phones(h))

        emails = list(dict.fromkeys([e for e in emails if e]))
        phones = list(dict.fromkeys([p for p in phones if p]))

        email = emails[0] if emails else ""
        phone = best_phone_candidate(phones)

        return email, phone
    except Exception:
        return "", ""


# =====================================================
# EXCEL
# =====================================================
COLUMNS = [
    "NOM", "ADRESSE", "CODE POSTAL", "VILLE", "TELEPHONE", "TELEPHONE 2", "MAIL",
    "SIRET", "NAF", "SITE WEB",
    "Contact: civilité", "Contact : prénom", "Contact : nom",
    "RESUME ENTRETIEN", "COMMANDE",
    "SOURCE", "CONFIDENCE", "STATUS", "SUGGESTIONS",
]

def autosize_columns(ws) -> None:
    for col in range(1, len(COLUMNS) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 18


def build_workbook(rows: List[Dict[str, Any]]) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Prospection"
    ws.append(COLUMNS)

    for r in rows:
        ws.append([
            r.get("name", ""),
            r.get("address", ""),
            r.get("postal_code", ""),
            r.get("city", ""),
            r.get("phone", ""),
            r.get("phone2", ""),
            r.get("email", ""),
            r.get("siret", ""),
            r.get("naf", ""),
            r.get("website", ""),
            r.get("contact_civility", ""),
            r.get("contact_firstname", ""),
            r.get("contact_lastname", ""),
            r.get("resume", ""),
            r.get("commande", ""),
            r.get("source", ""),
            r.get("confidence", ""),
            r.get("status", ""),
            r.get("suggestions", ""),
        ])

    autosize_columns(ws)
    return wb


# =====================================================
# EMAIL (Brevo)
# =====================================================
def parse_routing() -> Dict[str, Any]:
    if not MAIL_ROUTING_JSON:
        return {}
    try:
        return json.loads(MAIL_ROUTING_JSON)
    except Exception:
        return {}


def recipients_for(mode: str, agency: str, initials: str) -> List[str]:
    routing = parse_routing()
    out: List[str] = []

    if isinstance(routing.get("global"), list):
        out += routing["global"]

    if mode == "admin" and isinstance(routing.get("admin"), list):
        out += routing["admin"]

    agencies = routing.get("agencies") if isinstance(routing.get("agencies"), dict) else {}
    ag = agencies.get(agency) if isinstance(agencies, dict) else None

    if isinstance(ag, dict):
        if mode == "agency_manager":
            mgr = ag.get("__agency_manager")
            if isinstance(mgr, list):
                out += mgr
        if mode == "individual":
            lst = ag.get(initials.upper())
            if isinstance(lst, list):
                out += lst

    clean = []
    for e in out:
        e = (e or "").strip()
        if e and "@" in e:
            clean.append(e)
    return list(dict.fromkeys(clean))


def brevo_send_email(to_emails: List[str], subject: str, html: str, attachments: List[Tuple[str, bytes]]) -> None:
    if not BREVO_API_KEY or not BREVO_SENDER_EMAIL:
        print("Brevo not configured -> skip send")
        return
    if not to_emails:
        print("No recipients -> skip send")
        return

    url = "https://api.brevo.com/v3/smtp/email"
    headers = {"api-key": BREVO_API_KEY, "content-type": "application/json", "accept": "application/json"}

    att = [{"name": fn, "content": base64.b64encode(b).decode("ascii")} for fn, b in attachments]

    body = {
        "sender": {"name": BREVO_SENDER_NAME, "email": BREVO_SENDER_EMAIL},
        "to": [{"email": e} for e in to_emails],
        "subject": subject,
        "htmlContent": html,
        "attachment": att,
    }

    r = requests.post(url, headers=headers, json=body, timeout=HTTP_TIMEOUT)
    if r.status_code not in (200, 201, 202):
        raise RuntimeError(f"Brevo send failed: {r.status_code} {r.text[:500]}")


# =====================================================
# OPTIONAL: notify admin via Telegram
# =====================================================
def notify_admin(text: str) -> None:
    if not TELEGRAM_ADMIN_CHAT_ID or not TELEGRAM_TOKEN:
        return
    try:
        tg_api("sendMessage", {"chat_id": TELEGRAM_ADMIN_CHAT_ID, "text": text})
    except Exception:
        pass


# =====================================================
# PHOTO PIPELINE
# =====================================================
def load_image_bytes(path: str, max_bytes: int = 1_800_000) -> bytes:
    b = open(path, "rb").read()
    if len(b) <= max_bytes:
        return b
    return b[:max_bytes]


def enrich_from_photo_record(photo: Dict[str, Any], tmpdir: str) -> Dict[str, Any]:
    file_id = photo.get("file_id") or ""
    if not file_id:
        return {"_skip": True, "_reason": "missing file_id"}

    # City determination: manual city > reverse geocode if geo present
    city = photo.get("city")
    if not city:
        geo = photo.get("geo") if isinstance(photo.get("geo"), dict) else None
        if geo and "lat" in geo and "lon" in geo:
            time.sleep(1.05)  # nominatim polite delay
            city = reverse_geocode_city(float(geo["lat"]), float(geo["lon"]))
    city = normalize_spaces(city or "")

    # Download image
    fp = tg_get_file_path(file_id)
    local_path = os.path.join(tmpdir, f"{file_id}.jpg")
    tg_download_file(fp, local_path)

    img_bytes = load_image_bytes(local_path)
    vision = gemini_vision_extract(img_bytes)

    # Company guess
    company_guess = normalize_spaces((vision.get("company_name_guess") or "")[:200])
    if not company_guess and vision.get("text_extracted"):
        first_line = normalize_spaces(str(vision["text_extracted"]).splitlines()[0])[:200]
        if first_line and "@" not in first_line and len(first_line) >= 2:
            company_guess = first_line

    # Contact from vision
    c = vision.get("contact") or {}
    v_email = normalize_spaces(c.get("email") or "")
    v_mobile = normalize_fr_phone(c.get("mobile") or "")
    v_phone = normalize_fr_phone(c.get("phone") or "")

    # Manual overrides (priority)
    manual_phone = normalize_fr_phone(photo.get("phone_manual") or "")
    manual_mail = normalize_spaces(photo.get("mail_manual") or "")
    manual_interloc = normalize_spaces(photo.get("interlocutor") or "")
    manual_meeting = normalize_spaces(photo.get("meeting") or "")

    # Enrich via API gouv
    status = "A_VERIFIER"
    suggestions = []
    best = None

    if company_guess and city:
        results = api_gouv_search(company_guess, city, topn=3)
        if results:
            best = results[0]
            suggestions = [r["name"] for r in results[:3] if r.get("name")]
            status = "OK" if len(results) == 1 else "A_VERIFIER"
        else:
            status = "A_VERIFIER"
    else:
        status = "A_VERIFIER"

    # V2+ Places: pick best with scoring + details
    places_phone = ""
    places_website = ""
    places_sugs: List[str] = []

    places_query_name = (best["name"] if best and best.get("name") else company_guess)
    places_city = (best["city"] if best and best.get("city") else city)
    if places_query_name and places_city and GOOGLE_PLACES_API_KEY:
        places = places_search_text(f"{places_query_name} {places_city}", max_results=5)
        if places:
            places_phone, places_website, places_sugs = best_place_pick(places_query_name, places_city, places)
            if places_sugs and not suggestions:
                suggestions = places_sugs

    # V2+ scrape multi pages if still missing info
    scraped_email = ""
    scraped_phone = ""
    if places_website:
        scraped_email, scraped_phone = scrape_site_for_email_phone_v2plus(places_website)
    scraped_phone = normalize_fr_phone(scraped_phone)

    # Priority strict: manual > places > vision > scrape
    email_final = manual_mail or v_email or scraped_email or ""
    phone_final = manual_phone or places_phone or v_phone or scraped_phone or ""
    phone_final = normalize_fr_phone(phone_final)

    # Mobile into TELEPHONE 2 if present
    phone2_final = v_mobile

    # Identity fields
    civ = normalize_spaces(c.get("civility") or "")
    first = normalize_spaces(c.get("first_name") or "")
    last = normalize_spaces(c.get("last_name") or "")

    if manual_interloc and (not first or not last):
        parts = manual_interloc.split(" ", 1)
        if len(parts) == 2:
            first = first or parts[0]
            last = last or parts[1]
        else:
            last = last or manual_interloc

    row = {
        "name": best["name"] if best and best.get("name") else (company_guess or ""),
        "address": best["address"] if best and best.get("address") else "",
        "postal_code": best["postal_code"] if best and best.get("postal_code") else "",
        "city": best["city"] if best and best.get("city") else (city or ""),
        "phone": phone_final,
        "phone2": phone2_final,
        "email": email_final,
        "siret": best["siret"] if best and best.get("siret") else "",
        "naf": normalize_naf(best["naf"]) if best and best.get("naf") else "",
        "website": places_website or "",
        "contact_civility": civ,
        "contact_firstname": first,
        "contact_lastname": last,
        "resume": manual_meeting,
        "commande": "",
        "source": "photo",
        "confidence": vision.get("confidence", 0),
        "status": status,
        "suggestions": " | ".join([s for s in suggestions[:3] if s]) if suggestions else "",
    }

    if not row["name"]:
        row["status"] = "A_VERIFIER"

    return row


# =====================================================
# CLASSIC PROSPECTS (company mode) normalization
# =====================================================
def map_prospect_record(p: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "name": p.get("name", "") or "",
        "address": p.get("address", "") or "",
        "postal_code": p.get("postal_code", "") or "",
        "city": p.get("city", "") or "",
        "phone": normalize_fr_phone(p.get("phone", "") or ""),
        "phone2": normalize_fr_phone(p.get("phone2", "") or ""),
        "email": (p.get("email", "") or "").strip(),
        "siret": (p.get("siret", "") or "").strip(),
        "naf": normalize_naf(p.get("naf", "") or ""),
        "website": (p.get("website", "") or "").strip(),
        "contact_civility": (p.get("contact_civility", "") or "").strip(),
        "contact_firstname": (p.get("contact_firstname", "") or "").strip(),
        "contact_lastname": (p.get("contact_lastname", "") or "").strip(),
        "resume": (p.get("resume", "") or "").strip(),
        "commande": (p.get("commande", "") or "").strip(),
        "source": "company",
        "confidence": "",
        "status": "OK",
        "suggestions": "",
    }


# =====================================================
# MAIN
# =====================================================
def main() -> None:
    date_ymd = paris_today_ymd()

    prospects = worker_dump("prospects", date_ymd)
    photos = worker_dump("photos", date_ymd)
    closes = worker_dump("closes", date_ymd)

    def match_scope(agency_val: str, initials_val: str) -> bool:
        ag = (agency_val or "").upper()
        ini = (initials_val or "").upper()
        if SEND_MODE == "individual":
            return ag == AGENCY.upper() and ini == INITIALS.upper()
        if SEND_MODE == "agency_manager":
            return ag == AGENCY.upper()
        return True

    prospects_scoped = [p for p in prospects if match_scope(p.get("agency", ""), p.get("initials", ""))]
    photos_scoped = [p for p in photos if match_scope(p.get("agency", ""), p.get("user", ""))]
    closes_scoped = [c for c in closes if match_scope(c.get("agency", ""), c.get("initials", ""))]

    rows: List[Dict[str, Any]] = [map_prospect_record(p) for p in prospects_scoped]

    tmpdir = tempfile.mkdtemp(prefix="prospection_photos_")
    try:
        processed = 0
        for ph in photos_scoped:
            if processed >= MAX_PHOTO_IMAGES:
                break
            row = enrich_from_photo_record(ph, tmpdir)
            if row.get("_skip"):
                continue
            rows.append(row)
            processed += 1
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    safe_mode = SEND_MODE
    safe_ag = (AGENCY.upper() if AGENCY else "ALL")
    safe_ini = (INITIALS.upper() if INITIALS else "ALL")
    filename = f"prospection_{date_ymd}_{safe_mode}_{safe_ag}_{safe_ini}.xlsx"

    wb = build_workbook(rows)
    wb.save(filename)
    print(f"✅ Excel generated: {filename} rows={len(rows)}")

    to_emails = recipients_for(SEND_MODE, AGENCY.upper(), INITIALS.upper())
    subject = f"Prospection {date_ymd} — {SEND_MODE} {safe_ag} {safe_ini}"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Veuillez trouver en pièce jointe l'export prospection du <b>{date_ymd}</b>.</p>"
        f"<ul>"
        f"<li>Mode : <b>{SEND_MODE}</b></li>"
        f"<li>Agence : <b>{safe_ag}</b></li>"
        f"<li>Initiales : <b>{safe_ini}</b></li>"
        f"<li>Prospects (Entreprise & Ville) : <b>{len(prospects_scoped)}</b></li>"
        f"<li>Photos traitées : <b>{min(len(photos_scoped), MAX_PHOTO_IMAGES)}</b></li>"
        f"</ul>"
        f"<p>— Bot Prospection</p>"
    )

    with open(filename, "rb") as f:
        content = f.read()

    try:
        brevo_send_email(to_emails, subject, html, [(filename, content)])
        print(f"✅ Email sent to: {to_emails}")
    except Exception as e:
        print(f"⚠️ Email send failed: {e}")
        notify_admin(f"⚠️ Prospection export email failed ({date_ymd}) — {e}")

    if closes_scoped:
        print("=== CLOSES (log) ===")
        for c in closes_scoped[:50]:
            print(json.dumps(c, ensure_ascii=False))

    print("DONE")


if __name__ == "__main__":
    if not WORKER_BASE_URL:
        die("WORKER_BASE_URL missing")
    if not EXPORT_TOKEN:
        die("EXPORT_TOKEN missing")
    if not TELEGRAM_TOKEN:
        die("TELEGRAM_TOKEN missing")

    if SEND_MODE not in ("individual", "agency_manager", "admin"):
        die(f"SEND_MODE invalid: {SEND_MODE}")

    if SEND_MODE in ("individual", "agency_manager") and not AGENCY:
        die("AGENCY required for individual/agency_manager")
    if SEND_MODE == "individual" and not INITIALS:
        die("INITIALS required for individual")

    main()