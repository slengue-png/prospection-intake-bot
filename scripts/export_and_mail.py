#!/usr/bin/env python3
"""
export_and_mail.py — VERSION PRO (V1 FIGÉE)

✅ OCR cartes de visite (email/portable + nom contact heuristique)
✅ Reconstruit contact si besoin (interlocuteur > OCR > dirigeant)
✅ Export Excel basé sur template Table.2.xlsx (colonnes exactes) + AGENCE + INITIALS (INITIALS en dernière colonne)
✅ Consolidation "closes" (clients/prospects/commandes) via /dump?kind=closes
✅ Emails Brevo:
   - SEND_MODE=individual       -> mail immédiat du collaborateur
   - SEND_MODE=agency_manager   -> mail manager par agence (consolidé)
   - SEND_MODE=admin            -> mail SL (global) avec:
        - tableau par agence
        - tableau par commercial
        - totaux clients/prospects/commandes
        - classement par nombre d’actions

ENV attendues (GitHub Secrets / env):
- WORKER_BASE_URL
- EXPORT_TOKEN
- TELEGRAM_TOKEN
- BREVO_API_KEY
- BREVO_SENDER_EMAIL
- BREVO_SENDER_NAME
- MAIL_ROUTING_JSON

ENV inputs (workflow):
- SEND_MODE: individual | agency_manager | admin
- RUN_DATE: YYYY-MM-DD (optionnel -> UTC fallback)
- AGENCY: GR|VR|GRS|SLS (requis si individual; optionnel si agency_manager)
- INITIALS: JL|CZ|... (requis si individual)
- MAX_OCR_IMAGES: (optionnel, défaut 50)
"""

import os
import re
import io
import json
import base64
from datetime import datetime
from typing import Dict, List, Optional, Tuple

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


# =========================
# GLOBAL CONFIG
# =========================
TEMPLATE_PATH = "Table.2.xlsx"  # doit être committé à la racine du repo

# Colonnes attendues dans ton template (1ère ligne)
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

# Colonnes ajoutées en fin
EXTRA_HEADERS = ["AGENCE", "INITIALS"]  # INITIALS en dernière colonne


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
    Worker V4.2 stocke:
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
      - email
      - mobile (06/07)
      - phone (fixe prioritaire 04 sinon autre)
      - contact_full (heuristique)
    """
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    except Exception:
        return {"email": "", "mobile": "", "phone": "", "contact_full": ""}

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

    return {"email": email, "mobile": mobile, "phone": phone, "contact_full": contact_full}


def ensure_contact_fields(record: dict) -> None:
    """
    Priorité:
      1) saisie manuelle (interlocuteur) -> split -> contact_firstname/contact_lastname
      2) OCR (si contact fields vides) -> split -> contact fields
      3) dirigeant -> split -> contact fields

    Et si contact fields existent mais interlocuteur vide -> on reconstruit interlocuteur (pour Excel)
    """
    # 1) manual interlocuteur
    interloc = (record.get("interlocuteur") or "").strip()
    if interloc:
        fn, ln = split_name_simple(interloc)
        record["contact_firstname"] = fn
        record["contact_lastname"] = ln or interloc
        return

    # 2) if already filled (by worker) do nothing
    if (record.get("contact_firstname") or "").strip() or (record.get("contact_lastname") or "").strip():
        return

    # 3) fallback dirigeant
    dirg = (record.get("dirigeant") or "").strip()
    if dirg:
        fn, ln = split_name_simple(dirg)
        record["contact_firstname"] = fn
        record["contact_lastname"] = ln or dirg


# =========================
# EXCEL (Template)
# =========================
def build_header_map(ws) -> Dict[str, int]:
    max_col = ws.max_column
    return {str(ws.cell(row=1, column=c).value or "").strip(): c for c in range(1, max_col + 1)}


def ensure_extra_headers(ws) -> Dict[str, int]:
    header_map = build_header_map(ws)
    col = ws.max_column

    if "AGENCE" not in header_map:
        col += 1
        ws.cell(row=1, column=col).value = "AGENCE"
    if "INITIALS" not in header_map:
        col += 1
        ws.cell(row=1, column=col).value = "INITIALS"

    return build_header_map(ws)


def autosize(ws, width: int = 22) -> None:
    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def build_excel_from_template(records: List[dict], out_path: str) -> None:
    if not os.path.exists(TEMPLATE_PATH):
        raise RuntimeError(f"Missing Excel template at repo root: {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Soft check headers
    current_headers = [ws.cell(row=1, column=i).value for i in range(1, len(TEMPLATE_HEADERS) + 1)]
    if [str(x or "").strip() for x in current_headers] != TEMPLATE_HEADERS:
        # On ne bloque pas, mais ton template doit rester stable
        pass

    header_map = ensure_extra_headers(ws)

    start_row = ws.max_row + 1

    for idx, r in enumerate(records):
        row = start_row + idx

        # ensure contact fallback
        ensure_contact_fields(r)

        # Interlocuteur displayed:
        interloc_excel = (r.get("interlocuteur") or "").strip()
        if not interloc_excel:
            fn = (r.get("contact_firstname") or "").strip()
            ln = (r.get("contact_lastname") or "").strip()
            interloc_excel = f"{fn} {ln}".strip() or (r.get("dirigeant") or "")

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
    actions = nb prospects enregistrés
    clients/prospects/commandes = venant des closes (mini menu Clore session)
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

    # cast agencies sets -> list
    for ini in list(by_initials.keys()):
        by_initials[ini]["agencies"] = sorted(list(by_initials[ini]["agencies"]))

    return {"by_agency": by_agency, "by_initials": by_initials, "totals": totals, "ranking": ranking}


def html_table(headers: List[str], rows: List[List[str]]) -> str:
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
# MAIN
# =========================
def main() -> None:
    routing = json.loads(must_env("MAIL_ROUTING_JSON"))

    send_mode = opt_env("SEND_MODE", "individual").strip()
    run_date = opt_env("RUN_DATE")
    if not run_date:
        # fallback (workflow schedule te passera RUN_DATE si tu veux; sinon UTC)
        run_date = datetime.utcnow().strftime("%Y-%m-%d")

    agency = opt_env("AGENCY", "").upper()
    initials = opt_env("INITIALS", "").upper()
    max_ocr = int(opt_env("MAX_OCR_IMAGES", "50") or "50")

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

        # OCR only if missing and card exists
        card_attachments: List[Tuple[str, bytes]] = []
        ocr_done = 0

        for rec in prospects:
            ensure_contact_fields(rec)

            need_email = not (rec.get("email") or "").strip()
            need_mobile = not (rec.get("phone2") or "").strip()
            need_contact = not ((rec.get("contact_firstname") or "").strip() or (rec.get("contact_lastname") or "").strip())

            has_card = (rec.get("card_photo_url") or rec.get("card_photo_file_id") or "").strip() != ""
            if has_card and (need_email or need_mobile or need_contact) and ocr_done < max_ocr:
                img = download_card_image_bytes(rec)
                if img:
                    o = ocr_extract(img)

                    if need_email and o.get("email"):
                        rec["email"] = o["email"]
                    if need_mobile and o.get("mobile"):
                        rec["phone2"] = o["mobile"]
                    if not (rec.get("phone") or "").strip() and o.get("phone"):
                        rec["phone"] = o["phone"]

                    if need_contact and o.get("contact_full"):
                        fn, ln = split_name_simple(o["contact_full"])
                        if fn and not (rec.get("contact_firstname") or "").strip():
                            rec["contact_firstname"] = fn
                        if ln and not (rec.get("contact_lastname") or "").strip():
                            rec["contact_lastname"] = ln

                    ocr_done += 1
                    card_attachments.append((f"carte_{agency}_{initials}_{ocr_done}.jpg", img))

            # final fallback
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

        # stats from closes (as requested)
        clients = sum(to_int(x.get("visits_clients")) for x in closes)
        prosps = sum(to_int(x.get("visits_prospects")) for x in closes)
        comm = sum(to_int(x.get("commandes")) for x in closes)

        html = f"""
        <p>Bonjour,</p>
        <p>Voici ton export de prospection <b>{agency}/{initials}</b> du <b>{run_date}</b>.</p>
        <ul>
          <li><b>Actions (prospects saisis)</b> : {len(prospects)}</li>
          <li><b>Clients</b> : {clients}</li>
          <li><b>Prospects</b> : {prosps}</li>
          <li><b>Commandes</b> : {comm}</li>
        </ul>
        <p>Pièces jointes : Excel{ " + cartes de visite" if card_attachments else "" }.</p>
        """
        subject = f"[PROSPECTION] {agency}/{initials} - {run_date}"
        send_mail_brevo(to_list, subject, html, attachments)

        print(f"OK individual: {agency}/{initials} prospects={len(prospects)} cards={len(card_attachments)} closes={len(closes)}")
        return

    # ---------------------------------------
    # AGENCY MANAGER (17:45) — 1 mail/agence
    # ---------------------------------------
    if send_mode == "agency_manager":
        # si AGENCY fourni -> une agence, sinon toutes les agences du routing (sauf _admin)
        agencies = [agency] if agency else [k for k in routing.keys() if k != "_admin"]

        for ag in agencies:
            if not ag:
                continue

            prospects = filter_prospects(prospects_all, run_date, "agency_manager", ag, "")
            closes = filter_closes(closes_all, run_date, "agency_manager", ag, "")

            for rec in prospects:
                ensure_contact_fields(rec)

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

            html = f"""
            <p>Bonjour,</p>
            <p>Voici le récapitulatif <b>agence {ag}</b> du <b>{run_date}</b>.</p>
            <ul>
              <li><b>Actions (prospects saisis)</b> : {len(prospects)}</li>
              <li><b>Clients</b> : {clients}</li>
              <li><b>Prospects</b> : {prosps}</li>
              <li><b>Commandes</b> : {comm}</li>
            </ul>
            <p>Pièce jointe : Excel consolidé agence.</p>
            """
            subject = f"[PROSPECTION] Récap agence {ag} - {run_date}"
            send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
            print(f"OK agency_manager: {ag} prospects={len(prospects)} closes={len(closes)}")

        return

    # ------------------------------------------------
    # ADMIN (17:47) — 1 mail global SL + tableaux HTML
    # ------------------------------------------------
    if send_mode == "admin":
        prospects = filter_by_date(prospects_all, run_date)
        closes = filter_by_date(closes_all, run_date)

        for rec in prospects:
            ensure_contact_fields(rec)

        xlsx_name = f"prospection_GLOBAL_{run_date}.xlsx"
        build_excel_from_template(prospects, xlsx_name)
        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        stats = compute_stats(prospects, closes)
        totals = stats["totals"]

        # table by agency
        agency_rows = []
        for ag, v in sorted(stats["by_agency"].items(), key=lambda x: x[0]):
            agency_rows.append([ag, v["actions"], v["clients"], v["prospects"], v["commandes"]])

        # table by commercial
        initials_rows = []
        for ini, v in sorted(stats["by_initials"].items(), key=lambda x: x[0]):
            if not ini:
                continue
            initials_rows.append([
                ini,
                ", ".join(v["agencies"]) if v["agencies"] else "",
                v["actions"],
                v["clients"],
                v["prospects"],
                v["commandes"],
            ])

        # ranking by actions
        ranking_rows = [[ini, actions] for (ini, actions) in stats["ranking"]]

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

        <p>Pièce jointe : Excel GLOBAL.</p>
        """

        to_list = routing.get("_admin")
        if not to_list:
            raise RuntimeError("MAIL_ROUTING_JSON missing _admin recipients")

        subject = f"[PROSPECTION] Récap GLOBAL - {run_date}"
        send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
        print(f"OK admin: prospects={len(prospects)} closes={len(closes)}")
        return

    raise RuntimeError("Invalid SEND_MODE. Use individual|agency_manager|admin")


if __name__ == "__main__":
    main()