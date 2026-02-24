import os
import re
import io
import json
import base64
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any

import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from PIL import Image
import pytesseract

# =========================
# CONFIG / CONSTANTS
# =========================

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:(?:\+33|0)\s*[1-9](?:[\s.\-]*\d{2}){4})")

# Colonnes Excel (stable)
# NOTE: tu as demandé "INITIALS en dernière colonne" + "AGENCY" aussi (utile manager/admin)
EXCEL_COLUMNS = [
    "NOM",
    "ADRESSE",
    "CODE POSTAL",
    "VILLE",
    "TELEPHONE",
    "TELEPHONE 2",
    "MAIL",
    "SIRET",
    "NAF",
    "SITE WEB",
    "Contact: civilité",
    "Contact : prénom",
    "Contact : nom",
    "DIRIGEANT",
    "RESUME ENTRETIEN",
    "COMMANDE",
    "INFOS_COMMERCIALES",
    "CARTE_VISITE_FILE_ID",
    "CARTE_VISITE_URL",
    "AGENCY",
    "INITIALS",
]

# =========================
# ENV HELPERS
# =========================

def must_env(name: str) -> str:
    v = os.getenv(name, "").strip()
    if not v:
        raise RuntimeError(f"Missing required env var: {name}")
    return v

def opt_env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()

def safe_filename(s: str) -> str:
    s = re.sub(r"[^\w\-\. ]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:120] if len(s) > 120 else s

def normalize_naf(naf: str) -> str:
    if not naf:
        return ""
    naf = str(naf).strip().upper()
    return naf.replace(".", "").replace(" ", "")

def join_address(address: str, postal: str, city: str) -> str:
    parts = [str(x).strip() for x in [address, postal, city] if str(x).strip()]
    return " ".join(parts).strip()

# =========================
# WORKER DUMP
# =========================

def build_dump_url(base: str, date: str, kind: str) -> str:
    base = base.strip()
    if base.endswith("/"):
        base = base[:-1]
    return f"{base}/dump?date={date}&kind={kind}"

def fetch_dump_jsonl(date: str, kind: str) -> List[dict]:
    base = must_env("WORKER_BASE_URL")
    token = must_env("EXPORT_TOKEN")
    url = build_dump_url(base, date, kind)
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
            # ignore bad line
            continue
    return out

# =========================
# TELEGRAM DOWNLOAD (fallback if only file_id)
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

def telegram_download_file_by_path(file_path: str) -> bytes:
    token = must_env("TELEGRAM_TOKEN")
    url = f"https://api.telegram.org/file/bot{token}/{file_path}"
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content

def download_card_image(record: dict) -> Optional[bytes]:
    """
    Prefer card_photo_url (direct) else card_photo_file_id (via getFile).
    """
    url = (record.get("card_photo_url") or "").strip()
    if url:
        try:
            r = requests.get(url, timeout=60)
            r.raise_for_status()
            return r.content
        except Exception:
            pass

    file_id = (record.get("card_photo_file_id") or "").strip()
    if not file_id:
        return None
    try:
        fp = telegram_get_file_path(file_id)
        if not fp:
            return None
        return telegram_download_file_by_path(fp)
    except Exception:
        return None

# =========================
# OCR
# =========================

def ocr_extract_contact_fields(image_bytes: bytes) -> Dict[str, str]:
    """
    Retourne: {contact_full, email, mobile, phone}
    Heuristiques FR.
    """
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    except Exception:
        return {"contact_full": "", "email": "", "mobile": "", "phone": ""}

    text = pytesseract.image_to_string(img, lang="fra+eng")
    text = re.sub(r"[ \t]+", " ", text)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    email = ""
    m = EMAIL_RE.search(text)
    if m:
        email = m.group(0).strip()

    phones = PHONE_RE.findall(text)
    normalized: List[str] = []
    for p in phones:
        p2 = re.sub(r"[^\d+]", "", p)
        if p2.startswith("+33"):
            p2 = "0" + p2[3:]
        p2 = re.sub(r"\D", "", p2)
        if len(p2) == 10 and p2.startswith("0"):
            normalized.append(p2)

    mobile = next((x for x in normalized if x.startswith("06") or x.startswith("07")), "")
    phone = next((x for x in normalized if x.startswith("04")), "")
    if not phone:
        phone = next((x for x in normalized if x.startswith(("01", "02", "03", "05", "09", "08"))), "")

    # Candidate name line
    contact_full = ""
    candidates = []
    for ln in lines[:25]:
        if len(ln) < 4 or len(ln) > 60:
            continue
        up = ln.upper()
        if any(k in up for k in ["SAS", "SARL", "EURL", "FRANCE", "GROUPE", "WWW", "HTTP", "@", "TEL", "MOBILE"]):
            continue
        if len(ln.split()) < 2:
            continue
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", ln):
            continue
        candidates.append(ln)

    for ln in candidates:
        if re.search(r"\b[A-ZÀ-ÖØ-Ý]{3,}\b", ln):
            contact_full = ln
            break
    if not contact_full and candidates:
        contact_full = candidates[0]

    return {"contact_full": contact_full, "email": email, "mobile": mobile, "phone": phone}

def split_name_simple(full: str) -> Tuple[str, str]:
    """
    "Jean Dupont" -> ("Jean", "Dupont")
    "Dupont" -> ("", "Dupont")
    """
    s = (full or "").strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return ("", "")
    parts = s.split(" ")
    if len(parts) == 1:
        return ("", parts[0])
    return (parts[0], " ".join(parts[1:]))

def normalize_civility(s: str) -> Tuple[str, str]:
    """
    Extract civility if present at start.
    Returns (civility, rest)
    """
    if not s:
        return ("", "")
    s = s.strip()
    m = re.match(r"^(M\.?|MME|Mme|Monsieur|Madame)\s+(.*)$", s, flags=re.IGNORECASE)
    if not m:
        return ("", s)
    civ_raw = m.group(1)
    rest = m.group(2).strip()
    civ = civ_raw
    civ = civ.replace("Monsieur", "M.").replace("Madame", "Mme")
    civ = civ.replace("MME", "Mme").replace("Mme.", "Mme").strip()
    return (civ, rest)

def ensure_contact_fields(record: dict) -> None:
    """
    Ta règle:
    1) Saisie manuelle interlocuteur => CONTACT prénom/nom (split)
    2) Sinon OCR (si présent via contact_full / contact_first/last)
    3) Sinon dirigeant => Dirigeant + Contact nom (et prénom si split)
    """
    interloc = (record.get("interlocuteur") or "").strip()
    dirg = (record.get("dirigeant") or "").strip()

    # 1) manuel
    if interloc:
        civ, rest = normalize_civility(interloc)
        if civ and not (record.get("contact_civility") or "").strip():
            record["contact_civility"] = civ
        fn, ln = split_name_simple(rest)
        record["contact_firstname"] = fn
        record["contact_lastname"] = ln or rest
        return

    # 2) si déjà rempli par OCR ou worker (contact_first/last) -> ok
    has_contact = (record.get("contact_firstname") or "").strip() or (record.get("contact_lastname") or "").strip()
    if has_contact:
        return

    # 3) fallback dirigeant
    if dirg:
        fn, ln = split_name_simple(dirg)
        record["contact_firstname"] = fn
        record["contact_lastname"] = ln or dirg

# =========================
# OPTIONAL: light email scraping if missing
# =========================

def scrape_email_from_website(website: str) -> str:
    if not website:
        return ""
    try:
        r = requests.get(website, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        r.raise_for_status()
        matches = EMAIL_RE.findall(r.text)
        if not matches:
            return ""
        # remove noreply
        matches = [m for m in matches if not re.search(r"no-?reply", m, re.I)]
        return matches[0] if matches else ""
    except Exception:
        return ""

# =========================
# EXCEL
# =========================

def build_excel(records: List[dict], xlsx_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "PROSPECTION"

    ws.append(EXCEL_COLUMNS)

    for r in records:
        naf = normalize_naf(r.get("naf", ""))
        row = [
            r.get("name", "") or "",
            join_address(r.get("address", ""), r.get("postal_code", ""), r.get("city", "")),
            r.get("postal_code", "") or "",
            r.get("city", "") or "",
            r.get("phone", "") or "",
            r.get("phone2", "") or "",
            r.get("email", "") or "",
            r.get("siret", "") or "",
            naf,
            r.get("website", "") or "",
            r.get("contact_civility", "") or "",
            r.get("contact_firstname", "") or "",
            r.get("contact_lastname", "") or "",
            r.get("dirigeant", "") or "",
            r.get("resume", "") or "",
            r.get("commande", "") or "",
            r.get("infos_commerciales", "") or "",
            r.get("card_photo_file_id", "") or "",
            r.get("card_photo_url", "") or "",
            r.get("agency", "") or "",
            r.get("initials", "") or "",
        ]
        ws.append(row)

    # autosize simple
    for col in range(1, len(EXCEL_COLUMNS) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 20

    wb.save(xlsx_path)

# =========================
# BREVO EMAIL
# =========================

def send_mail_brevo(to_list: List[str], subject: str, html: str, attachments: List[Tuple[str, bytes]]) -> None:
    api_key = must_env("BREVO_API_KEY")
    sender_email = must_env("BREVO_SENDER_EMAIL")
    sender_name = must_env("BREVO_SENDER_NAME")

    payload: Dict[str, Any] = {
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
# STATS / REPORTING
# =========================

def html_table(headers: List[str], rows: List[List[str]]) -> str:
    th = "".join([f"<th style='border:1px solid #ddd;padding:6px;text-align:left'>{h}</th>" for h in headers])
    trs = []
    for row in rows:
        tds = "".join([f"<td style='border:1px solid #ddd;padding:6px'>{(c if c is not None else '')}</td>" for c in row])
        trs.append(f"<tr>{tds}</tr>")
    return f"""
    <table style="border-collapse:collapse;border:1px solid #ddd;width:100%;margin:10px 0">
      <thead><tr style="background:#f6f6f6">{th}</tr></thead>
      <tbody>
        {''.join(trs)}
      </tbody>
    </table>
    """

def compute_action_counts(prospects: List[dict]) -> Dict[Tuple[str, str], int]:
    # (agency, initials) -> number of prospect records
    counts: Dict[Tuple[str, str], int] = {}
    for r in prospects:
        ag = (r.get("agency") or "").upper()
        ini = (r.get("initials") or "").upper()
        if not ag or not ini:
            continue
        counts[(ag, ini)] = counts.get((ag, ini), 0) + 1
    return counts

def compute_close_stats(closes: List[dict]) -> Dict[Tuple[str, str], Dict[str, int]]:
    # (agency, initials) -> totals of clients/prospects/commandes
    out: Dict[Tuple[str, str], Dict[str, int]] = {}
    for c in closes:
        ag = (c.get("agency") or "").upper()
        ini = (c.get("initials") or "").upper()
        if not ag or not ini:
            continue
        def to_int(x) -> int:
            try:
                x = str(x).strip()
                return int(x) if x else 0
            except Exception:
                return 0
        vc = to_int(c.get("visits_clients", "0"))
        vp = to_int(c.get("visits_prospects", "0"))
        co = to_int(c.get("commandes", "0"))
        key = (ag, ini)
        if key not in out:
            out[key] = {"clients": 0, "prospects": 0, "commandes": 0}
        out[key]["clients"] += vc
        out[key]["prospects"] += vp
        out[key]["commandes"] += co
    return out

def sum_close_stats_by_agency(close_stats: Dict[Tuple[str, str], Dict[str, int]]) -> Dict[str, Dict[str, int]]:
    out: Dict[str, Dict[str, int]] = {}
    for (ag, _ini), v in close_stats.items():
        if ag not in out:
            out[ag] = {"clients": 0, "prospects": 0, "commandes": 0}
        out[ag]["clients"] += v["clients"]
        out[ag]["prospects"] += v["prospects"]
        out[ag]["commandes"] += v["commandes"]
    return out

# =========================
# ROUTING / MODES
# =========================

def load_routing() -> dict:
    routing_raw = must_env("MAIL_ROUTING_JSON")
    return json.loads(routing_raw)

def get_agencies_from_routing(routing: dict) -> List[str]:
    return [k for k in routing.keys() if k not in ("_admin",)]

def recipients_individual(routing: dict, agency: str, initials: str) -> List[str]:
    return routing.get(agency, {}).get(initials, [])

def recipients_manager(routing: dict, agency: str) -> List[str]:
    return routing.get(agency, {}).get("_manager", [])

def recipients_admin(routing: dict) -> List[str]:
    return routing.get("_admin", [])

# =========================
# MAIN LOGIC
# =========================

def enrich_records_with_ocr_and_fallback(records: List[dict]) -> List[Tuple[str, bytes]]:
    """
    OCR:
    - Only if card exists.
    - Fill missing email/phone2/contact fields.
    Returns list of attachment images (name, bytes).
    """
    attachments: List[Tuple[str, bytes]] = []

    for r in records:
        img_bytes = download_card_image(r)
        if not img_bytes:
            # fallback contact via dirigeant if needed
            ensure_contact_fields(r)
            continue

        # OCR extract
        o = ocr_extract_contact_fields(img_bytes)

        # If no manual interlocuteur, OCR may provide contact
        if not (r.get("interlocuteur") or "").strip() and o.get("contact_full"):
            # store as interlocuteur for traceability (but contact fields filled below)
            r["interlocuteur"] = o["contact_full"]

        # civility + split
        if o.get("contact_full"):
            civ, rest = normalize_civility(o["contact_full"])
            if civ and not (r.get("contact_civility") or "").strip():
                r["contact_civility"] = civ
            # Only fill name parts if empty (manual should win; worker sets manual already)
            if not (r.get("contact_firstname") or "").strip() and not (r.get("contact_lastname") or "").strip():
                fn, ln = split_name_simple(rest)
                if fn or ln:
                    r["contact_firstname"] = fn
                    r["contact_lastname"] = ln or rest

        # email
        if not (r.get("email") or "").strip() and o.get("email"):
            r["email"] = o["email"]

        # phones
        if not (r.get("phone2") or "").strip() and o.get("mobile"):
            r["phone2"] = o["mobile"]
        if not (r.get("phone") or "").strip() and o.get("phone"):
            r["phone"] = o["phone"]

        # Website email scraping if still missing
        if not (r.get("email") or "").strip():
            site = (r.get("website") or "").strip()
            if site:
                scraped = scrape_email_from_website(site)
                if scraped:
                    r["email"] = scraped

        # Apply final fallback rule (manuel > OCR > dirigeant)
        ensure_contact_fields(r)

        # Add image attachment
        file_id = (r.get("card_photo_file_id") or "card").strip()
        company = safe_filename(r.get("name", "societe") or "societe")
        attachments.append((safe_filename(f"carte_{company}_{file_id}.jpg"), img_bytes))

    # ensure fallback for all records
    for r in records:
        ensure_contact_fields(r)

    return attachments

def filter_prospects(prospects: List[dict], date: str, agency: Optional[str], initials: Optional[str], mode: str) -> List[dict]:
    out = []
    for r in prospects:
        if str(r.get("date", "")).strip() != date:
            continue
        if mode in ("individual", "agency_manager") and agency:
            if str(r.get("agency", "")).strip().upper() != agency.upper():
                continue
        if mode == "individual" and initials:
            if str(r.get("initials", "")).strip().upper() != initials.upper():
                continue
        out.append(r)
    return out

def build_admin_html(run_date: str, prospects_all: List[dict], closes_all: List[dict]) -> str:
    action_counts = compute_action_counts(prospects_all)
    close_stats = compute_close_stats(closes_all)
    agency_totals = sum_close_stats_by_agency(close_stats)

    # Table by agency
    agency_rows = []
    for ag in sorted(agency_totals.keys()):
        v = agency_totals[ag]
        agency_rows.append([ag, str(v["clients"]), str(v["prospects"]), str(v["commandes"]), str(sum(action_counts.get((ag, ini), 0) for (_, ini) in action_counts.keys() if _ == ag))])
    agency_tbl = html_table(
        ["Agence", "Clients", "Prospects", "Commandes", "Nb actions (fiches)"],
        agency_rows or [["—", "0", "0", "0", "0"]],
    )

    # Table by commercial (agency+initials)
    commercial_rows = []
    all_keys = set(action_counts.keys()) | set(close_stats.keys())
    for (ag, ini) in sorted(all_keys):
        actions = action_counts.get((ag, ini), 0)
        st = close_stats.get((ag, ini), {"clients": 0, "prospects": 0, "commandes": 0})
        commercial_rows.append([ag, ini, str(actions), str(st["clients"]), str(st["prospects"]), str(st["commandes"])])
    commercial_tbl = html_table(
        ["Agence", "Initiales", "Nb actions (fiches)", "Clients", "Prospects", "Commandes"],
        commercial_rows or [["—", "—", "0", "0", "0", "0"]],
    )

    # Ranking by actions
    ranking = sorted([(ag, ini, cnt) for (ag, ini), cnt in action_counts.items()], key=lambda x: x[2], reverse=True)
    rank_rows = [[str(i+1), ag, ini, str(cnt)] for i, (ag, ini, cnt) in enumerate(ranking)]
    rank_tbl = html_table(
        ["Rang", "Agence", "Initiales", "Nb actions (fiches)"],
        rank_rows or [["—", "—", "—", "0"]],
    )

    # Totals
    tot_clients = sum(v["clients"] for v in close_stats.values()) if close_stats else 0
    tot_prospects = sum(v["prospects"] for v in close_stats.values()) if close_stats else 0
    tot_commandes = sum(v["commandes"] for v in close_stats.values()) if close_stats else 0
    tot_actions = sum(action_counts.values()) if action_counts else 0

    return f"""
    <p>Bonsoir,</p>
    <p><b>Récap global</b> — <b>{run_date}</b></p>

    <p>
      <b>Totaux jour</b><br/>
      Actions (fiches): <b>{tot_actions}</b><br/>
      Clients: <b>{tot_clients}</b> — Prospects: <b>{tot_prospects}</b> — Commandes: <b>{tot_commandes}</b>
    </p>

    <h3>📊 Synthèse par agence</h3>
    {agency_tbl}

    <h3>👥 Synthèse par commercial</h3>
    {commercial_tbl}

    <h3>🏆 Classement commerciaux (nb d’actions)</h3>
    {rank_tbl}

    <p>Pièces jointes : Excel global.</p>
    """

def build_manager_html(run_date: str, agency: str, prospects_agency: List[dict], closes_all: List[dict]) -> str:
    action_counts = compute_action_counts(prospects_agency)
    close_stats = compute_close_stats(closes_all)

    keys = sorted(set([(agency.upper(), (r.get("initials") or "").upper()) for r in prospects_agency if (r.get("initials") or "").strip()] +
                      [(ag, ini) for (ag, ini) in close_stats.keys() if ag.upper() == agency.upper()]))

    rows = []
    for (ag, ini) in keys:
        actions = action_counts.get((ag, ini), 0)
        st = close_stats.get((ag, ini), {"clients": 0, "prospects": 0, "commandes": 0})
        rows.append([ini, str(actions), str(st["clients"]), str(st["prospects"]), str(st["commandes"])])

    tot_actions = sum(action_counts.values()) if action_counts else 0
    tot_clients = sum(v["clients"] for (ag, _), v in close_stats.items() if ag.upper() == agency.upper())
    tot_prospects = sum(v["prospects"] for (ag, _), v in close_stats.items() if ag.upper() == agency.upper())
    tot_commandes = sum(v["commandes"] for (ag, _), v in close_stats.items() if ag.upper() == agency.upper())

    tbl = html_table(
        ["Initiales", "Nb actions (fiches)", "Clients", "Prospects", "Commandes"],
        rows or [["—", "0", "0", "0", "0"]],
    )

    return f"""
    <p>Bonjour,</p>
    <p><b>Récap agence {agency}</b> — <b>{run_date}</b></p>
    <p>
      Totaux agence — Actions: <b>{tot_actions}</b>,
      Clients: <b>{tot_clients}</b>,
      Prospects: <b>{tot_prospects}</b>,
      Commandes: <b>{tot_commandes}</b>
    </p>
    {tbl}
    <p>Pièce jointe : Excel consolidé agence.</p>
    """

def build_individual_html(run_date: str, agency: str, initials: str, closes_all: List[dict]) -> str:
    close_stats = compute_close_stats(closes_all).get((agency.upper(), initials.upper()), {"clients": 0, "prospects": 0, "commandes": 0})
    return f"""
    <p>Bonjour,</p>
    <p>Voici ton export de prospection <b>{agency}/{initials}</b> du <b>{run_date}</b>.</p>
    <p>
      <b>Clôture</b> — Clients: <b>{close_stats["clients"]}</b>,
      Prospects: <b>{close_stats["prospects"]}</b>,
      Commandes: <b>{close_stats["commandes"]}</b>
    </p>
    <p>Pièces jointes : Excel + cartes de visite (si disponibles).</p>
    """

def main():
    send_mode = opt_env("SEND_MODE", "individual").strip()
    run_date = opt_env("RUN_DATE", "")
    if not run_date:
        run_date = datetime.utcnow().strftime("%Y-%m-%d")

    routing = load_routing()
    agencies = get_agencies_from_routing(routing)

    agency = opt_env("AGENCY", "").upper()
    initials = opt_env("INITIALS", "").upper()

    # 1) dumps
    prospects_all = fetch_dump_jsonl(run_date, "prospects")
    closes_all = fetch_dump_jsonl(run_date, "closes")

    # 2) Mode individual => 1 mail pour AGENCY/INITIALS
    if send_mode == "individual":
        if not agency or not initials:
            raise RuntimeError("SEND_MODE=individual requires AGENCY and INITIALS")

        prospects = filter_prospects(prospects_all, run_date, agency, initials, "individual")

        # OCR + fallback + attachments
        attachments = enrich_records_with_ocr_and_fallback(prospects)

        # excel
        xlsx_name = safe_filename(f"prospection_{agency}_{initials}_{run_date}.xlsx")
        build_excel(prospects, xlsx_name)

        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        attach_all = [(xlsx_name, excel_bytes)] + attachments

        to_list = recipients_individual(routing, agency, initials)
        if not to_list:
            raise RuntimeError(f"No recipient mapping for {agency}/{initials}")

        subject = f"[PROSPECTION] {agency}/{initials} - {run_date}"
        html = build_individual_html(run_date, agency, initials, closes_all)
        send_mail_brevo(to_list, subject, html, attach_all)

        print(f"OK individual: {agency}/{initials} records={len(prospects)} attachments={len(attach_all)}")
        return

    # 3) Mode agency_manager => 1 mail par agence (manager) + excel agence
    if send_mode == "agency_manager":
        # si AGENCY vide => loop toutes agences
        target_agencies = [agency] if agency else agencies

        for ag in target_agencies:
            prospects_agency = filter_prospects(prospects_all, run_date, ag, None, "agency_manager")

            # Pas d'OCR sur recap agence (plus rapide). Mais on applique fallback contact.
            for r in prospects_agency:
                ensure_contact_fields(r)

            xlsx_name = safe_filename(f"prospection_{ag}_CONSOLIDE_{run_date}.xlsx")
            build_excel(prospects_agency, xlsx_name)
            with open(xlsx_name, "rb") as f:
                excel_bytes = f.read()

            to_list = recipients_manager(routing, ag)
            if not to_list:
                # on skip si pas de manager
                print(f"Skip agency {ag}: no _manager recipient")
                continue

            subject = f"[PROSPECTION] Récap agence {ag} - {run_date}"
            html = build_manager_html(run_date, ag, prospects_agency, closes_all)
            send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])

            print(f"OK agency_manager: {ag} records={len(prospects_agency)}")

        return

    # 4) Mode admin => 1 mail global SL + excel global + tableaux stats
    if send_mode == "admin":
        # fallback contact fields
        for r in prospects_all:
            ensure_contact_fields(r)

        xlsx_name = safe_filename(f"prospection_GLOBAL_{run_date}.xlsx")
        build_excel(prospects_all, xlsx_name)
        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        to_list = recipients_admin(routing)
        if not to_list:
            raise RuntimeError("No _admin recipients configured in MAIL_ROUTING_JSON")

        subject = f"[PROSPECTION] Récap GLOBAL - {run_date}"
        html = build_admin_html(run_date, prospects_all, closes_all)
        send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])

        print(f"OK admin: records={len(prospects_all)}")
        return

    raise RuntimeError("Invalid SEND_MODE. Use individual|agency_manager|admin")

if __name__ == "__main__":
    main()
