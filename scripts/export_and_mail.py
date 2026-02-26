#!/usr/bin/env python3
"""
export_and_mail.py — VERSION PRO V2 (V1 stable + Gemini 1.5 Flash “safe add-on”)

✅ V1 conservée :
- /dump prospects + closes (Cloudflare Worker inchangé)
- OCR Tesseract (cartes de visite)
- Export Excel depuis template Table.2.xlsx (openpyxl)
- Envoi emails Brevo (multi-agences / multi-modes)

✅ V2 ajout (SANS casser l’existant) :
- Gemini 1.5 Flash (Google AI Studio) OPTIONNEL
- GEMINI_API_KEY absent => fallback silencieux (aucun plantage)
- JSON STRICT, retries, timeout, température <= 0.2, maxOutputTokens <= 512
- Complète Excel uniquement si vide (n’écrase jamais)
- Ajoute colonnes AI si nécessaires (en fin de feuille)

MODES :
- SEND_MODE=individual       -> mail immédiat collaborateur (filtré AGENCY+INITIALS)
- SEND_MODE=agency_manager   -> mail manager par agence (consolidé)
- SEND_MODE=admin            -> mail SL global + tableaux HTML + classement

ENV (Secrets / env) :
- WORKER_BASE_URL
- EXPORT_TOKEN
- TELEGRAM_TOKEN
- BREVO_API_KEY
- BREVO_SENDER_EMAIL
- BREVO_SENDER_NAME
- MAIL_ROUTING_JSON   (json recipients)
OPTIONNEL :
- GEMINI_API_KEY

ENV inputs (workflow/dispatch) :
- SEND_MODE: individual | agency_manager | admin
- RUN_DATE: YYYY-MM-DD (optionnel)
- AGENCY: GR|VR|GRS|SLS (requis si individual; optionnel si agency_manager)
- INITIALS: JL|CZ|JB|LB|PV|ST|AC|SL (requis si individual)
- MAX_OCR_IMAGES: (optionnel, défaut 50)
"""

import os
import re
import io
import json
import base64
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

import requests
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image
import pytesseract


# =========================
# REGEX
# =========================
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:(?:\+33|0)\s*[1-9](?:[\s.\-]*\d{2}){4})")


# =========================
# ENV
# =========================
def must_env(name: str) -> str:
    v = os.getenv(name, "").strip()
    if not v:
        raise RuntimeError(f"Missing required env var: {name}")
    return v


def opt_env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def opt_int(name: str, default: int) -> int:
    s = os.getenv(name, "").strip()
    if not s:
        return default
    try:
        return int(s)
    except Exception:
        return default


# =========================
# GLOBAL CONFIG
# =========================
TEMPLATE_PATH = "Table.2.xlsx"  # doit être committé à la racine du repo

# Colonnes attendues (template utilisateur)
TEMPLATE_HEADERS = [
    "NOM",
    "RUE",
    "CODE POSTAL",
    "VILLE",
    "Téléphone",
    "Téléphone (Portable)",
    "Mail générique",
    "SIRET",
    "NAF",
    "SITE WEB",
    "INTERLOCUTEUR",
    "DIRIGEANT",
    "RESUME ENTRETIEN",
    "COMMANDE",
    "CARTE DE VISITE",
]

# Colonnes ajoutées fin (V1)
EXTRA_HEADERS = ["AGENCE", "INITIALS"]  # INITIALS en dernière colonne

# Colonnes AI ajoutées fin (V2)
AI_HEADERS_BUSINESS_CARD = [
    "AI_CIVILITY",
    "AI_FIRST_NAME",
    "AI_LAST_NAME",
    "AI_JOB_TITLE",
    "AI_EMAIL",
    "AI_MOBILE",
    "AI_PHONE",
]

AI_HEADERS_MEETING = [
    "AI_NEED",
    "AI_POSITIONS",
    "AI_VOLUME",
    "AI_CONSTRAINTS",
    "AI_DECISION_MAKER",
    "AI_NEXT_STEP",
    "AI_URGENCY",
    "AI_NOTES",
]

AI_HEADERS_SCORE = [
    "AI_SCORE",
    "AI_SCORE_JUSTIFICATION",
]

# Gemini config
GEMINI_MODEL = "gemini-1.5-flash"
GEMINI_TEMPERATURE = 0.2
GEMINI_MAX_TOKENS = 512
GEMINI_TIMEOUT_S = 20
GEMINI_MAX_RETRIES = 3


# =========================
# WORKER DUMP
# =========================
def build_dump_url(base: str, run_date: str, kind: str) -> str:
    base = base.rstrip("/")
    return f"{base}/dump?date={run_date}&kind={kind}"


def fetch_dump_jsonl(run_date: str, kind: str) -> List[dict]:
    base = must_env("WORKER_BASE_URL")
    token = must_env("EXPORT_TOKEN")

    url = build_dump_url(base, run_date, kind)
    r = requests.get(url, headers={"X-Export-Token": token}, timeout=60)
    r.raise_for_status()

    out: List[dict] = []
    for ln in r.text.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        try:
            out.append(json.loads(ln))
        except Exception:
            continue
    return out


# =========================
# TELEGRAM DOWNLOAD (for OCR)
# =========================
def telegram_get_file_path(file_id: str) -> Optional[str]:
    token = must_env("TELEGRAM_TOKEN")
    url = f"https://api.telegram.org/bot{token}/getFile"
    r = requests.post(url, json={"file_id": file_id}, timeout=30)
    r.raise_for_status()
    j = r.json()
    if not j.get("ok"):
        return None
    return j["result"].get("file_path")


def download_card_image_bytes(record: dict) -> Optional[bytes]:
    """
    Worker stocke:
      - card_photo_url (direct)
      - card_photo_file_id
    """
    url = (record.get("card_photo_url") or "").strip()
    if url:
        r = requests.get(url, timeout=60)
        if r.ok and r.content:
            return r.content

    file_id = (record.get("card_photo_file_id") or "").strip()
    if not file_id:
        return None

    fp = telegram_get_file_path(file_id)
    if not fp:
        return None

    token = must_env("TELEGRAM_TOKEN")
    url2 = f"https://api.telegram.org/file/bot{token}/{fp}"
    r2 = requests.get(url2, timeout=60)
    if r2.ok and r2.content:
        return r2.content

    return None


# =========================
# OCR + NORMALISATION
# =========================
def normalize_phone_fr(raw: str) -> str:
    p = re.sub(r"[^\d+]", "", raw or "")
    if p.startswith("+33"):
        p = "0" + p[3:]
    p = re.sub(r"\D", "", p)
    if len(p) == 10 and p.startswith("0"):
        return p
    return ""


def split_name_simple(full: str) -> Tuple[str, str]:
    """
    "Jean Dupont" -> ("Jean","Dupont")
    "Dupont" -> ("","Dupont")
    """
    s = (full or "").strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return ("", "")
    parts = s.split(" ")
    if len(parts) == 1:
        return ("", parts[0])
    return (parts[0], " ".join(parts[1:]))


def ocr_extract(image_bytes: bytes) -> Dict[str, str]:
    """
    Extrait:
      - ocr_text (brut)
      - email
      - mobile (06/07)
      - phone (fixe prioritaire 04 sinon autre)
      - contact_full (heuristique)
    """
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    except Exception:
        return {"ocr_text": "", "email": "", "mobile": "", "phone": "", "contact_full": ""}

    text = pytesseract.image_to_string(img, lang="fra+eng")
    text = re.sub(r"[ \t]+", " ", text)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    email = ""
    m = EMAIL_RE.search(text)
    if m:
        email = m.group(0).strip()

    phones = PHONE_RE.findall(text)
    normalized = []
    for p in phones:
        np = normalize_phone_fr(p)
        if np:
            normalized.append(np)

    mobile = next((x for x in normalized if x.startswith(("06", "07"))), "")
    phone = next((x for x in normalized if x.startswith("04")), "")
    if not phone:
        phone = next((x for x in normalized if x.startswith(("01", "02", "03", "05", "08", "09"))), "")

    # contact heuristic
    contact_full = ""
    candidates = []
    for ln in lines[:25]:
        if len(ln) < 4 or len(ln) > 60:
            continue
        upper = ln.upper()
        if any(k in upper for k in ["SAS", "SARL", "EURL", "FRANCE", "GROUPE", "WWW", "HTTP", "@"]):
            continue
        if len(ln.split()) < 2:
            continue
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", ln):
            continue
        candidates.append(ln)

    # prefer line with uppercase token (often lastname)
    for ln in candidates:
        if re.search(r"\b[A-ZÀ-ÖØ-Ý]{3,}\b", ln):
            contact_full = ln
            break
    if not contact_full and candidates:
        contact_full = candidates[0]

    return {"ocr_text": text, "email": email, "mobile": mobile, "phone": phone, "contact_full": contact_full}


def ensure_contact_fields(record: dict) -> None:
    """
    Priorité (V1 conservée) :
      1) saisie manuelle (interlocuteur) -> split -> contact_firstname/contact_lastname
      2) si contact fields déjà remplis -> rien
      3) dirigeant -> split -> contact fields

    Important: ne jamais écraser une valeur existante.
    """
    interloc = (record.get("interlocuteur") or "").strip()
    if interloc:
        if not (record.get("contact_firstname") or "").strip() and not (record.get("contact_lastname") or "").strip():
            fn, ln = split_name_simple(interloc)
            record["contact_firstname"] = fn
            record["contact_lastname"] = ln or interloc
        return

    if (record.get("contact_firstname") or "").strip() or (record.get("contact_lastname") or "").strip():
        return

    dirg = (record.get("dirigeant") or "").strip()
    if dirg:
        fn, ln = split_name_simple(dirg)
        record["contact_firstname"] = fn
        record["contact_lastname"] = ln or dirg


# =========================
# GEMINI 1.5 FLASH (OPTIONNEL)
# =========================
def _extract_json_object(text: str) -> Optional[str]:
    """
    Essaye d’extraire un objet JSON {...} depuis une réponse qui peut contenir
    du markdown, ```json, ou du texte autour.
    """
    if not text:
        return None
    s = text.strip()

    # Enlever fences markdown si présents
    s = re.sub(r"^```(?:json)?\s*", "", s, flags=re.IGNORECASE).strip()
    s = re.sub(r"\s*```$", "", s).strip()

    # Chercher premier '{' et parse "balanced braces" simple
    start = s.find("{")
    if start < 0:
        return None

    depth = 0
    in_str = False
    esc = False
    for i in range(start, len(s)):
        ch = s[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        else:
            if ch == '"':
                in_str = True
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return s[start : i + 1]
    return None


def _gemini_enabled() -> bool:
    return bool(opt_env("GEMINI_API_KEY"))


def gemini_generate_json(system_instruction: str, user_text: str, schema_hint: str) -> Optional[dict]:
    """
    Appel Gemini (Google AI Studio API) et retourne un dict.
    - Si GEMINI_API_KEY absent => return None (fallback silencieux)
    - Retry simple (max 3)
    - Timeout défini
    - JSON strict (on extrait/clean si nécessaire)
    """
    api_key = opt_env("GEMINI_API_KEY")
    if not api_key:
        return None

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    headers = {"Content-Type": "application/json"}
    params = {"key": api_key}

    # Prompt strict JSON
    # On force un objet JSON unique, sans texte, sans markdown.
    sys = (system_instruction or "").strip()
    usr = (user_text or "").strip()
    sch = (schema_hint or "").strip()

    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": (
                            "SYSTEM INSTRUCTION:\n"
                            f"{sys}\n\n"
                            "OUTPUT RULES (STRICT):\n"
                            "- Return ONLY one JSON object.\n"
                            "- No markdown.\n"
                            "- No explanations.\n"
                            "- Use null for unknown.\n\n"
                            "SCHEMA (HINT):\n"
                            f"{sch}\n\n"
                            "INPUT:\n"
                            f"{usr}\n"
                        )
                    }
                ],
            }
        ],
        "generationConfig": {
            "temperature": GEMINI_TEMPERATURE,
            "maxOutputTokens": GEMINI_MAX_TOKENS,
        },
    }

    last_err = None
    for attempt in range(1, GEMINI_MAX_RETRIES + 1):
        try:
            r = requests.post(url, headers=headers, params=params, json=payload, timeout=GEMINI_TIMEOUT_S)
            # Rate-limit / transient
            if r.status_code in (429, 500, 502, 503, 504):
                last_err = f"HTTP {r.status_code}"
                time.sleep(0.7 * attempt)
                continue
            r.raise_for_status()

            j = r.json()
            # Extraction texte
            text = ""
            try:
                candidates = j.get("candidates") or []
                if candidates:
                    parts = candidates[0].get("content", {}).get("parts") or []
                    if parts:
                        text = parts[0].get("text", "") or ""
            except Exception:
                text = ""

            raw = text.strip()
            if not raw:
                return None

            obj_str = _extract_json_object(raw) or raw
            try:
                parsed = json.loads(obj_str)
            except Exception:
                # tentative : extraire à nouveau si le raw contient du bruit
                obj_str2 = _extract_json_object(raw)
                if not obj_str2:
                    return None
                parsed = json.loads(obj_str2)

            if isinstance(parsed, dict):
                return parsed
            return None

        except Exception as e:
            last_err = str(e)
            time.sleep(0.7 * attempt)
            continue

    # Fallback silencieux : on ne plante pas le job
    _ = last_err
    return None


# =========================
# AI FUNCTIONS (SPEC)
# =========================
def ai_parse_business_card(ocr_text: str) -> Optional[dict]:
    """
    Retour EXACT :
    {
      "civility": null,
      "first_name": null,
      "last_name": null,
      "job_title": null,
      "email": null,
      "mobile": null,
      "phone": null
    }
    """
    system_instruction = (
        "You extract business card fields from noisy OCR text. "
        "Be conservative and only fill fields you are confident about."
    )
    schema_hint = json.dumps(
        {
            "civility": None,
            "first_name": None,
            "last_name": None,
            "job_title": None,
            "email": None,
            "mobile": None,
            "phone": None,
        },
        ensure_ascii=False,
    )
    res = gemini_generate_json(system_instruction, ocr_text, schema_hint)
    if not res:
        return None

    # Normalize exact keys + nulls
    out = {
        "civility": res.get("civility", None),
        "first_name": res.get("first_name", None),
        "last_name": res.get("last_name", None),
        "job_title": res.get("job_title", None),
        "email": res.get("email", None),
        "mobile": res.get("mobile", None),
        "phone": res.get("phone", None),
    }

    # clean strings
    for k, v in list(out.items()):
        if isinstance(v, str):
            vv = v.strip()
            out[k] = vv if vv else None
    return out


def ai_structure_meeting(meeting_text: str) -> Optional[dict]:
    """
    Retour EXACT :
    {
      "need": null,
      "positions": null,
      "volume": null,
      "constraints": null,
      "decision_maker": null,
      "next_step": null,
      "urgency": null,
      "notes": null
    }
    """
    system_instruction = (
        "You structure a free-form commercial meeting note into a strict JSON object. "
        "Do not invent facts; use null if missing."
    )
    schema_hint = json.dumps(
        {
            "need": None,
            "positions": None,
            "volume": None,
            "constraints": None,
            "decision_maker": None,
            "next_step": None,
            "urgency": None,
            "notes": None,
        },
        ensure_ascii=False,
    )
    res = gemini_generate_json(system_instruction, meeting_text, schema_hint)
    if not res:
        return None

    out = {
        "need": res.get("need", None),
        "positions": res.get("positions", None),
        "volume": res.get("volume", None),
        "constraints": res.get("constraints", None),
        "decision_maker": res.get("decision_maker", None),
        "next_step": res.get("next_step", None),
        "urgency": res.get("urgency", None),
        "notes": res.get("notes", None),
    }
    for k, v in list(out.items()):
        if isinstance(v, str):
            vv = v.strip()
            out[k] = vv if vv else None
    return out


def ai_score_prospect(data: dict) -> Optional[dict]:
    """
    BONUS (optionnel) :
    {
      "score": 0-100,
      "justification": ""
    }
    """
    system_instruction = (
        "You estimate prospect potential for a staffing/recruitment agency based on provided structured data. "
        "Return a score 0-100 and a short justification. Do not invent missing facts."
    )
    schema_hint = json.dumps({"score": 0, "justification": ""}, ensure_ascii=False)
    user_text = json.dumps(data, ensure_ascii=False)
    res = gemini_generate_json(system_instruction, user_text, schema_hint)
    if not res:
        return None
    score = res.get("score", None)
    justif = res.get("justification", "")
    try:
        score_int = int(score)
    except Exception:
        score_int = None
    if score_int is None or score_int < 0 or score_int > 100:
        return None
    if isinstance(justif, str):
        justif = justif.strip()
    return {"score": score_int, "justification": justif or ""}


# =========================
# EXCEL (Template)
# =========================
def build_header_map(ws) -> Dict[str, int]:
    max_col = ws.max_column
    return {str(ws.cell(row=1, column=c).value or "").strip(): c for c in range(1, max_col + 1)}


def ensure_headers(ws, headers_to_ensure: List[str]) -> Dict[str, int]:
    header_map = build_header_map(ws)
    col = ws.max_column
    for h in headers_to_ensure:
        if h not in header_map:
            col += 1
            ws.cell(row=1, column=col).value = h
            header_map[h] = col
    return header_map


def autosize(ws, width: int = 22) -> None:
    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def build_excel_from_template(records: List[dict], out_path: str) -> None:
    if not os.path.exists(TEMPLATE_PATH):
        raise RuntimeError(f"Missing Excel template at repo root: {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Ensure base extra headers
    header_map = ensure_headers(ws, EXTRA_HEADERS)

    # Ensure AI headers (safe append)
    header_map = ensure_headers(ws, AI_HEADERS_BUSINESS_CARD)
    header_map = ensure_headers(ws, AI_HEADERS_MEETING)
    header_map = ensure_headers(ws, AI_HEADERS_SCORE)

    start_row = ws.max_row + 1

    for idx, r in enumerate(records):
        row = start_row + idx

        # ensure contact fallback (V1)
        ensure_contact_fields(r)

        # Interlocuteur displayed:
        interloc_excel = (r.get("interlocuteur") or "").strip()
        if not interloc_excel:
            fn = (r.get("contact_firstname") or "").strip()
            ln = (r.get("contact_lastname") or "").strip()
            interloc_excel = f"{fn} {ln}".strip() or (r.get("dirigeant") or "")

        # Base template values
        values = {
            "NOM": r.get("name", "") or "",
            "RUE": r.get("address", "") or "",
            "CODE POSTAL": r.get("postal_code", "") or "",
            "VILLE": r.get("city", "") or "",
            "Téléphone": r.get("phone", "") or "",
            "Téléphone (Portable)": r.get("phone2", "") or "",
            "Mail générique": r.get("email", "") or "",
            "SIRET": r.get("siret", "") or "",
            "NAF": (r.get("naf", "") or "").replace(".", "").replace(" ", ""),
            "SITE WEB": r.get("website", "") or "",
            "INTERLOCUTEUR": interloc_excel,
            "DIRIGEANT": r.get("dirigeant", "") or "",
            "RESUME ENTRETIEN": r.get("resume", "") or "",
            "COMMANDE": r.get("commande", "") or "",
            "CARTE DE VISITE": r.get("card_photo_url", "") or r.get("card_photo_file_id", "") or "",
            "AGENCE": r.get("agency", "") or "",
            "INITIALS": r.get("initials", "") or "",
        }

        # AI values (only if present, never overwrite existing excel cells because row is new)
        ai_bc = r.get("_ai_business_card") or {}
        ai_meet = r.get("_ai_meeting") or {}
        ai_score = r.get("_ai_score") or {}

        values.update(
            {
                "AI_CIVILITY": ai_bc.get("civility") or "",
                "AI_FIRST_NAME": ai_bc.get("first_name") or "",
                "AI_LAST_NAME": ai_bc.get("last_name") or "",
                "AI_JOB_TITLE": ai_bc.get("job_title") or "",
                "AI_EMAIL": ai_bc.get("email") or "",
                "AI_MOBILE": ai_bc.get("mobile") or "",
                "AI_PHONE": ai_bc.get("phone") or "",
                "AI_NEED": ai_meet.get("need") or "",
                "AI_POSITIONS": ai_meet.get("positions") or "",
                "AI_VOLUME": ai_meet.get("volume") or "",
                "AI_CONSTRAINTS": ai_meet.get("constraints") or "",
                "AI_DECISION_MAKER": ai_meet.get("decision_maker") or "",
                "AI_NEXT_STEP": ai_meet.get("next_step") or "",
                "AI_URGENCY": ai_meet.get("urgency") or "",
                "AI_NOTES": ai_meet.get("notes") or "",
                "AI_SCORE": ai_score.get("score") if isinstance(ai_score.get("score"), int) else "",
                "AI_SCORE_JUSTIFICATION": ai_score.get("justification") or "",
            }
        )

        for h, v in values.items():
            col = header_map.get(h)
            if col:
                ws.cell(row=row, column=col).value = v

    autosize(ws, 22)
    wb.save(out_path)


# =========================
# BREVO
# =========================
def send_mail_brevo(to_list: List[str], subject: str, html: str, attachments: List[Tuple[str, bytes]]) -> None:
    api_key = must_env("BREVO_API_KEY")
    sender_email = must_env("BREVO_SENDER_EMAIL")
    sender_name = must_env("BREVO_SENDER_NAME")

    payload = {
        "sender": {"email": sender_email, "name": sender_name},
        "to": [{"email": x} for x in to_list],
        "subject": subject,
        "htmlContent": html,
        "attachment": [{"name": name, "content": base64.b64encode(b).decode("utf-8")} for name, b in attachments],
    }

    r = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={"api-key": api_key, "content-type": "application/json"},
        json=payload,
        timeout=60,
    )
    r.raise_for_status()


# =========================
# STATS (closes)
# =========================
def to_int(x) -> int:
    try:
        s = str(x).strip()
        return int(s) if s else 0
    except Exception:
        return 0


def compute_stats(prospects: List[dict], closes: List[dict]) -> dict:
    """
    actions = nb prospects enregistrés (records)
    clients/prospects/commandes = venant des closes (/dump?kind=closes)
    """
    by_agency: Dict[str, dict] = {}
    by_initials: Dict[str, dict] = {}

    # actions by prospect records
    for p in prospects:
        ag = (p.get("agency") or "").upper()
        ini = (p.get("initials") or "").upper()
        if not ag:
            continue

        by_agency.setdefault(ag, {"actions": 0, "clients": 0, "prospects": 0, "commandes": 0, "by_initials": {}})
        by_agency[ag]["actions"] += 1
        by_agency[ag]["by_initials"].setdefault(ini, {"actions": 0, "clients": 0, "prospects": 0, "commandes": 0})
        by_agency[ag]["by_initials"][ini]["actions"] += 1

        by_initials.setdefault(ini, {"agencies": set(), "actions": 0, "clients": 0, "prospects": 0, "commandes": 0})
        by_initials[ini]["actions"] += 1
        by_initials[ini]["agencies"].add(ag)

    # closes numbers
    for c in closes:
        ag = (c.get("agency") or "").upper()
        ini = (c.get("initials") or "").upper()

        clients = to_int(c.get("visits_clients"))
        prosps = to_int(c.get("visits_prospects"))
        comm = to_int(c.get("commandes"))

        if ag:
            by_agency.setdefault(ag, {"actions": 0, "clients": 0, "prospects": 0, "commandes": 0, "by_initials": {}})
            by_agency[ag]["clients"] += clients
            by_agency[ag]["prospects"] += prosps
            by_agency[ag]["commandes"] += comm

            by_agency[ag]["by_initials"].setdefault(ini, {"actions": 0, "clients": 0, "prospects": 0, "commandes": 0})
            by_agency[ag]["by_initials"][ini]["clients"] += clients
            by_agency[ag]["by_initials"][ini]["prospects"] += prosps
            by_agency[ag]["by_initials"][ini]["commandes"] += comm

        by_initials.setdefault(ini, {"agencies": set(), "actions": 0, "clients": 0, "prospects": 0, "commandes": 0})
        by_initials[ini]["clients"] += clients
        by_initials[ini]["prospects"] += prosps
        by_initials[ini]["commandes"] += comm
        if ag:
            by_initials[ini]["agencies"].add(ag)

    totals = {
        "actions": sum(v["actions"] for v in by_agency.values()),
        "clients": sum(v["clients"] for v in by_agency.values()),
        "prospects": sum(v["prospects"] for v in by_agency.values()),
        "commandes": sum(v["commandes"] for v in by_agency.values()),
    }

    ranking = sorted(
        [(ini, v["actions"]) for ini, v in by_initials.items() if ini],
        key=lambda x: x[1],
        reverse=True,
    )

    for ini in list(by_initials.keys()):
        by_initials[ini]["agencies"] = sorted(list(by_initials[ini]["agencies"]))

    return {"by_agency": by_agency, "by_initials": by_initials, "totals": totals, "ranking": ranking}


def html_table(headers: List[str], rows: List[List[Any]]) -> str:
    th = "".join([f"<th style='border:1px solid #ddd;padding:6px;text-align:left'>{h}</th>" for h in headers])
    trs = []
    for r in rows:
        tds = "".join([f"<td style='border:1px solid #ddd;padding:6px'>{str(x)}</td>" for x in r])
        trs.append(f"<tr>{tds}</tr>")
    return f"<table style='border-collapse:collapse;border:1px solid #ddd;margin:10px 0'><tr>{th}</tr>{''.join(trs)}</table>"


# =========================
# FILTERING
# =========================
def filter_by_date(items: List[dict], run_date: str) -> List[dict]:
    return [x for x in items if str(x.get("date", "")).strip() == run_date]


def filter_prospects(items: List[dict], run_date: str, mode: str, agency: str, initials: str) -> List[dict]:
    items = filter_by_date(items, run_date)
    if mode == "individual":
        return [x for x in items if (x.get("agency") or "").upper() == agency and (x.get("initials") or "").upper() == initials]
    if mode == "agency_manager":
        return [x for x in items if (x.get("agency") or "").upper() == agency] if agency else items
    return items  # admin


def filter_closes(items: List[dict], run_date: str, mode: str, agency: str, initials: str) -> List[dict]:
    items = filter_by_date(items, run_date)
    if mode == "individual":
        return [x for x in items if (x.get("agency") or "").upper() == agency and (x.get("initials") or "").upper() == initials]
    if mode == "agency_manager":
        return [x for x in items if (x.get("agency") or "").upper() == agency] if agency else items
    return items  # admin


# =========================
# V2 INTEGRATION HELPERS
# =========================
def fill_if_empty(record: dict, key: str, value: Any) -> None:
    """Ne jamais écraser une donnée existante."""
    if value is None:
        return
    cur = record.get(key, None)
    if cur is None:
        record[key] = value
        return
    if isinstance(cur, str) and not cur.strip():
        record[key] = value
        return


def maybe_ai_enrich_business_card(record: dict, ocr_text: str) -> None:
    """
    Après OCR -> ai_parse_business_card
    Complète UNIQUEMENT si vides :
      - email -> record["email"]
      - mobile -> record["phone2"]
      - phone -> record["phone"]
      - interlocuteur/contact -> record["contact_firstname/contact_lastname"] (si vides)
    Stocke aussi le bloc AI dans record["_ai_business_card"] pour Excel (colonnes AI_*)
    """
    if not ocr_text or not _gemini_enabled():
        return

    ai = ai_parse_business_card(ocr_text)
    if not ai:
        return

    record["_ai_business_card"] = ai

    # Fill missing
    fill_if_empty(record, "email", ai.get("email"))
    fill_if_empty(record, "phone2", ai.get("mobile"))
    fill_if_empty(record, "phone", ai.get("phone"))

    # Contact fields: only if missing
    if not (record.get("contact_firstname") or "").strip() and not (record.get("contact_lastname") or "").strip():
        fn = ai.get("first_name")
        ln = ai.get("last_name")
        if fn or ln:
            fill_if_empty(record, "contact_firstname", fn or "")
            fill_if_empty(record, "contact_lastname", ln or "")
            # If interlocuteur empty, reconstruct (but do not overwrite manual)
            if not (record.get("interlocuteur") or "").strip():
                combo = f"{fn or ''} {ln or ''}".strip()
                if combo:
                    fill_if_empty(record, "interlocuteur", combo)


def maybe_ai_structure_meeting(record: dict) -> None:
    """
    Après récupération résumé entretien -> ai_structure_meeting
    - Ne modifie pas record["resume"] (V1 inchangé)
    - Stocke le JSON structuré dans record["_ai_meeting"] pour Excel (colonnes AI_*)
    """
    txt = (record.get("resume") or "").strip()
    if not txt or not _gemini_enabled():
        return
    ai = ai_structure_meeting(txt)
    if not ai:
        return
    record["_ai_meeting"] = ai


def maybe_ai_score(record: dict) -> None:
    """
    BONUS : score potentiel prospect (optionnel)
    - Fait uniquement si Gemini OK
    - Ne bloque jamais le flux
    """
    if not _gemini_enabled():
        return

    # Données minimales (sans complexifier)
    payload = {
        "company": record.get("name") or "",
        "city": record.get("city") or "",
        "naf": record.get("naf") or "",
        "website": record.get("website") or "",
        "meeting": (record.get("_ai_meeting") or {}) if isinstance(record.get("_ai_meeting"), dict) else {},
        "order": (record.get("commande") or ""),
        "notes": (record.get("resume") or ""),
    }
    ai = ai_score_prospect(payload)
    if not ai:
        return
    record["_ai_score"] = ai


# =========================
# MAIN
# =========================
def main() -> None:
    routing = json.loads(must_env("MAIL_ROUTING_JSON"))

    send_mode = opt_env("SEND_MODE", "individual").strip()
    run_date = opt_env("RUN_DATE")
    if not run_date:
        run_date = datetime.utcnow().strftime("%Y-%m-%d")

    agency = opt_env("AGENCY", "").upper()
    initials = opt_env("INITIALS", "").upper()
    max_ocr = opt_int("MAX_OCR_IMAGES", 50)

    prospects_all = fetch_dump_jsonl(run_date, "prospects")
    closes_all = fetch_dump_jsonl(run_date, "closes")

    # -------------------------
    # INDIVIDUAL (mail immédiat)
    # -------------------------
    if send_mode == "individual":
        if not agency or not initials:
            raise RuntimeError("SEND_MODE=individual requires AGENCY and INITIALS")

        prospects = filter_prospects(prospects_all, run_date, "individual", agency, initials)
        closes = filter_closes(closes_all, run_date, "individual", agency, initials)

        card_attachments: List[Tuple[str, bytes]] = []
        ocr_done = 0

        for rec in prospects:
            # V1 fallback contact (manual > dirigeant)
            ensure_contact_fields(rec)

            need_email = not (rec.get("email") or "").strip()
            need_mobile = not (rec.get("phone2") or "").strip()
            need_contact = not ((rec.get("contact_firstname") or "").strip() or (rec.get("contact_lastname") or "").strip())
            need_phone = not (rec.get("phone") or "").strip()

            has_card = (rec.get("card_photo_url") or rec.get("card_photo_file_id") or "").strip() != ""
            if has_card and (need_email or need_mobile or need_contact or need_phone) and ocr_done < max_ocr:
                img = download_card_image_bytes(rec)
                if img:
                    o = ocr_extract(img)
                    ocr_text = o.get("ocr_text", "")

                    # V1: regex fill if missing
                    if need_email and o.get("email"):
                        fill_if_empty(rec, "email", o["email"])
                    if need_mobile and o.get("mobile"):
                        fill_if_empty(rec, "phone2", o["mobile"])
                    if need_phone and o.get("phone"):
                        fill_if_empty(rec, "phone", o["phone"])

                    if need_contact and o.get("contact_full"):
                        fn, ln = split_name_simple(o["contact_full"])
                        if fn and not (rec.get("contact_firstname") or "").strip():
                            fill_if_empty(rec, "contact_firstname", fn)
                        if ln and not (rec.get("contact_lastname") or "").strip():
                            fill_if_empty(rec, "contact_lastname", ln)

                    # V2: Gemini parse business card (only fills empty)
                    maybe_ai_enrich_business_card(rec, ocr_text)

                    ocr_done += 1
                    card_attachments.append((f"carte_{agency}_{initials}_{ocr_done}.jpg", img))

            # V2: structure meeting (non destructif) + optional score
            maybe_ai_structure_meeting(rec)
            maybe_ai_score(rec)

            # final V1 fallback
            ensure_contact_fields(rec)

        # Excel
        xlsx_name = f"prospection_{agency}_{initials}_{run_date}.xlsx"
        build_excel_from_template(prospects, xlsx_name)
        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        attachments = [(xlsx_name, excel_bytes)] + card_attachments

        # recipients
        to_list = routing.get(agency, {}).get(initials)
        if not to_list:
            raise RuntimeError(f"No routing found for {agency}/{initials}")

        # stats from closes (requested)
        clients = sum(to_int(x.get("visits_clients")) for x in closes)
        prosps = sum(to_int(x.get("visits_prospects")) for x in closes)
        comm = sum(to_int(x.get("commandes")) for x in closes)

        ai_note = ""
        if _gemini_enabled():
            ai_note = "<p><i>AI activée (Gemini 1.5 Flash) : structuration entretien + parsing carte (si utile).</i></p>"

        html = f"""
        <p>Bonjour,</p>
        <p>Voici ton export de prospection <b>{agency}/{initials}</b> du <b>{run_date}</b>.</p>
        <ul>
          <li><b>Actions (prospects saisis)</b> : {len(prospects)}</li>
          <li><b>Clients</b> : {clients}</li>
          <li><b>Prospects</b> : {prosps}</li>
          <li><b>Commandes</b> : {comm}</li>
        </ul>
        {ai_note}
        <p>Pièces jointes : Excel{" + cartes de visite" if card_attachments else ""}.</p>
        """
        subject = f"[PROSPECTION] {agency}/{initials} - {run_date}"
        send_mail_brevo(to_list, subject, html, attachments)

        print(f"OK individual: {agency}/{initials} prospects={len(prospects)} cards={len(card_attachments)} closes={len(closes)} ai={'on' if _gemini_enabled() else 'off'}")
        return

    # ---------------------------------------
    # AGENCY MANAGER (17:45) — 1 mail/agence
    # ---------------------------------------
    if send_mode == "agency_manager":
        agencies = [agency] if agency else [k for k in routing.keys() if k != "_admin"]

        for ag in agencies:
            if not ag:
                continue

            prospects = filter_prospects(prospects_all, run_date, "agency_manager", ag, "")
            closes = filter_closes(closes_all, run_date, "agency_manager", ag, "")

            for rec in prospects:
                ensure_contact_fields(rec)
                maybe_ai_structure_meeting(rec)
                maybe_ai_score(rec)

            xlsx_name = f"prospection_{ag}_CONSOLIDE_{run_date}.xlsx"
            build_excel_from_template(prospects, xlsx_name)
            with open(xlsx_name, "rb") as f:
                excel_bytes = f.read()

            to_list = routing.get(ag, {}).get("_manager")
            if not to_list:
                print(f"SKIP agency {ag}: missing _manager")
                continue

            clients = sum(to_int(x.get("visits_clients")) for x in closes)
            prosps = sum(to_int(x.get("visits_prospects")) for x in closes)
            comm = sum(to_int(x.get("commandes")) for x in closes)

            ai_note = ""
            if _gemini_enabled():
                ai_note = "<p><i>AI activée (Gemini 1.5 Flash) : structuration entretien + score (si possible).</i></p>"

            html = f"""
            <p>Bonjour,</p>
            <p>Voici le récapitulatif <b>agence {ag}</b> du <b>{run_date}</b>.</p>
            <ul>
              <li><b>Actions (prospects saisis)</b> : {len(prospects)}</li>
              <li><b>Clients</b> : {clients}</li>
              <li><b>Prospects</b> : {prosps}</li>
              <li><b>Commandes</b> : {comm}</li>
            </ul>
            {ai_note}
            <p>Pièce jointe : Excel consolidé agence.</p>
            """
            subject = f"[PROSPECTION] Récap agence {ag} - {run_date}"
            send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
            print(f"OK agency_manager: {ag} prospects={len(prospects)} closes={len(closes)} ai={'on' if _gemini_enabled() else 'off'}")

        return

    # ------------------------------------------------
    # ADMIN (17:47) — 1 mail global SL + tableaux HTML
    # ------------------------------------------------
    if send_mode == "admin":
        prospects = filter_by_date(prospects_all, run_date)
        closes = filter_by_date(closes_all, run_date)

        for rec in prospects:
            ensure_contact_fields(rec)
            maybe_ai_structure_meeting(rec)
            maybe_ai_score(rec)

        xlsx_name = f"prospection_GLOBAL_{run_date}.xlsx"
        build_excel_from_template(prospects, xlsx_name)
        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        stats = compute_stats(prospects, closes)
        totals = stats["totals"]

        agency_rows = []
        for ag, v in sorted(stats["by_agency"].items(), key=lambda x: x[0]):
            agency_rows.append([ag, v["actions"], v["clients"], v["prospects"], v["commandes"]])

        initials_rows = []
        for ini, v in sorted(stats["by_initials"].items(), key=lambda x: x[0]):
            if not ini:
                continue
            initials_rows.append(
                [
                    ini,
                    ", ".join(v["agencies"]) if v["agencies"] else "",
                    v["actions"],
                    v["clients"],
                    v["prospects"],
                    v["commandes"],
                ]
            )

        ranking_rows = [[ini, actions] for (ini, actions) in stats["ranking"]]

        ai_note = ""
        if _gemini_enabled():
            ai_note = "<p><i>AI activée (Gemini 1.5 Flash) : structuration entretien + score (si possible).</i></p>"

        html = f"""
        <p>Bonsoir SL,</p>
        <p>Voici le <b>récapitulatif GLOBAL</b> du <b>{run_date}</b>.</p>

        <h3 style="margin:16px 0 6px 0">Totaux jour</h3>
        <ul>
          <li><b>Actions (prospects saisis)</b> : {totals["actions"]}</li>
          <li><b>Clients</b> : {totals["clients"]}</li>
          <li><b>Prospects</b> : {totals["prospects"]}</li>
          <li><b>Commandes</b> : {totals["commandes"]}</li>
        </ul>

        <h3 style="margin:16px 0 6px 0">Tableau par agence</h3>
        {html_table(["Agence","Actions","Clients","Prospects","Commandes"], agency_rows)}

        <h3 style="margin:16px 0 6px 0">Tableau par commercial</h3>
        {html_table(["Initiales","Agences","Actions","Clients","Prospects","Commandes"], initials_rows)}

        <h3 style="margin:16px 0 6px 0">Classement (par nombre d’actions)</h3>
        {html_table(["Initiales","Actions"], ranking_rows)}

        {ai_note}
        <p>Pièce jointe : Excel GLOBAL.</p>
        """

        to_list = routing.get("_admin")
        if not to_list:
            raise RuntimeError("MAIL_ROUTING_JSON missing _admin recipients")

        subject = f"[PROSPECTION] Récap GLOBAL - {run_date}"
        send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
        print(f"OK admin: prospects={len(prospects)} closes={len(closes)} ai={'on' if _gemini_enabled() else 'off'}")
        return

    raise RuntimeError("Invalid SEND_MODE. Use individual|agency_manager|admin")


if __name__ == "__main__":
    main()