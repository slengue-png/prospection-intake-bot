"""
export_and_mail.py — Script principal export prospection v8.4
=============================================================
Règles d'envoi :

INDIVIDUAL (clôture d'un commercial) :
  - Commercial → son travail (.xls + .xlsx)
  - Manager agence → travail du commercial (.xls + .xlsx)
  - SL admin → travail du commercial si SL a travaillé pour cette agence (.xls + .xlsx)

AGENCY_MANAGER (cron 17h47) :
  - Chaque manager → tout ce que son agence a produit (.xls uniquement)

ADMIN (cron 17h49) :
  - SL → recap toutes agences (.xlsx uniquement)
"""

import os, sys, json, datetime, base64, requests
from typing import Any, Dict, List, Optional

from excel_export import (
    export_commercial, export_agency_manager, export_admin,
    build_email_attachments, get_xlsx_bytes,
)

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

WORKER_BASE_URL    = os.environ.get("WORKER_BASE_URL", "").rstrip("/")
EXPORT_TOKEN       = os.environ.get("EXPORT_TOKEN", "")
TELEGRAM_TOKEN     = os.environ.get("TELEGRAM_TOKEN", "")
BREVO_API_KEY      = os.environ.get("BREVO_API_KEY", "")
BREVO_SENDER_EMAIL = os.environ.get("BREVO_SENDER_EMAIL", "")
BREVO_SENDER_NAME  = os.environ.get("BREVO_SENDER_NAME", "Prospection Bot")
MAIL_ROUTING_JSON  = os.environ.get("MAIL_ROUTING_JSON", "{}")
SEND_MODE          = os.environ.get("SEND_MODE", "individual")
RUN_DATE           = os.environ.get("RUN_DATE", datetime.date.today().strftime("%Y-%m-%d"))
AGENCY             = os.environ.get("AGENCY", "")
INITIALS           = os.environ.get("INITIALS", "")
OUT_DIR            = "/tmp/prospection_out"

try:
    MAIL_ROUTING = json.loads(MAIL_ROUTING_JSON)
except Exception:
    MAIL_ROUTING = {}

# Fallback hardcodé si le JSON est vide ou mal parsé
if not MAIL_ROUTING.get("agencies"):
    MAIL_ROUTING = {
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
        "admin": {"initials": "SL", "email": "samuel.lengue@ras-interim.fr"},
    }
    print("  ⚠️  MAIL_ROUTING_JSON invalide — routing hardcodé utilisé")

ADMIN_EMAIL    = MAIL_ROUTING.get("admin", {}).get("email", "")
ADMIN_INITIALS = MAIL_ROUTING.get("admin", {}).get("initials", "SL")


# ─────────────────────────────────────────────
# RÉCUPÉRATION DONNÉES
# ─────────────────────────────────────────────

def fetch_prospects(agency="", initials="", date="") -> List[Dict]:
    params = {}
    if date:     params["date"]     = date
    if agency:   params["agency"]   = agency
    if initials: params["initials"] = initials
    try:
        r = requests.get(
            f"{WORKER_BASE_URL}/dump",
            params=params,
            headers={"X-Export-Token": EXPORT_TOKEN},
            timeout=60,
        )
        r.raise_for_status()
        prospects = []
        for line in r.text.strip().split("\n"):
            line = line.strip()
            if not line: continue
            try:
                obj = json.loads(line)
                if isinstance(obj, dict): prospects.append(obj)
            except: pass
        print(f"  → {len(prospects)} prospect(s) | headers: count={r.headers.get('x-record-count','?')}")
        return prospects
    except Exception as e:
        print(f"  ✗ fetch: {e}")
        return []


def normalize_prospect(p: Dict) -> Dict:
    d = p.get("draft", p)
    return {
        "name":               d.get("name", ""),
        "nom_commercial":     d.get("nom_commercial", ""),
        "address":            d.get("address", ""),
        "postal_code":        d.get("postal_code", ""),
        "city":               d.get("city", ""),
        "phone":              d.get("phone", ""),
        "phone2":             d.get("phone2", ""),
        "email":              d.get("email", ""),
        "siret":              d.get("siret", ""),
        "naf":                d.get("naf", ""),
        "activity_summary":   d.get("secteur_activite", ""),
        "effectif":           d.get("effectif", ""),
        "date_creation":      d.get("date_creation", ""),
        "capital":            d.get("capital", ""),
        "website":            d.get("website", ""),
        "contact_civility":   d.get("contact_civility", ""),
        "interlocuteur":      d.get("interlocuteur", ""),
        "dirigeant":          d.get("dirigeant", ""),
        "resume":             d.get("resume", ""),
        "commande":           d.get("commande", ""),
        "qualification":      d.get("qualification", ""),
        "card_photo_url":     d.get("card_photo_url", ""),
        "card_photo_file_id": d.get("card_photo_file_id", ""),
        "facade_photo_url":   d.get("facade_photo_url", ""),
        "facade_photo_file_id": d.get("facade_photo_file_id", ""),
        "agency":             p.get("agency", d.get("agency", "")),
        "initials":           p.get("initials", d.get("initials", "")),
        "_confidence":        p.get("_confidence", d.get("_confidence", 0)),
    }


# ─────────────────────────────────────────────
# EMAIL BREVO
# ─────────────────────────────────────────────

def send_brevo(to_email: str, to_name: str, subject: str, html: str,
               attachments: List = None) -> bool:
    if not BREVO_API_KEY or not to_email:
        print(f"  ✗ Brevo non configuré ou email vide")
        return False
    payload = {
        "sender":      {"email": BREVO_SENDER_EMAIL, "name": BREVO_SENDER_NAME},
        "to":          [{"email": to_email, "name": to_name}],
        "subject":     subject,
        "htmlContent": html,
    }
    if attachments:
        payload["attachment"] = [
            {"name": name, "content": base64.b64encode(content).decode()}
            for name, content in attachments
        ]
    try:
        r = requests.post(
            "https://api.brevo.com/v3/smtp/email",
            headers={"api-key": BREVO_API_KEY, "Content-Type": "application/json"},
            json=payload, timeout=60,
        )
        if r.status_code in (200, 201, 202):
            print(f"  ✓ → {to_email}")
            return True
        print(f"  ✗ Brevo {r.status_code}: {r.text[:200]}")
        return False
    except Exception as e:
        print(f"  ✗ Brevo: {e}")
        return False


def html_body(rows, agency, initials, date, mode):
    count = len(rows)
    if mode == "individual":
        title = f"Prospection {date} — {agency}/{initials}"
        intro = f"<b>{count} fiche(s)</b> saisie(s) le {date} par <b>{initials}</b> ({agency})."
    elif mode == "agency_manager":
        title = f"Récap journalier {date} — Agence {agency}"
        intro = f"<b>{count} fiche(s)</b> au total pour l'agence <b>{agency}</b> le {date}."
    else:
        title = f"Récap admin {date} — Toutes agences"
        intro = f"<b>{count} fiche(s)</b> toutes agences confondues le {date}."
    return f"""<html><body style="font-family:Arial,sans-serif;color:#333">
    <h2 style="color:#2C3E50">{title}</h2><p>{intro}</p>
    <p>Fichiers joints :<br>• <b>.xlsx</b> : consultation<br>• <b>.xls</b> : import CRM</p>
    <hr style="border:none;border-top:1px solid #eee">
    <p style="color:#999;font-size:11px">Généré par Prospection Bot v8.4</p>
    </body></html>"""


def attachments_xls_only(path_xlsx, path_xls):
    """Retourne uniquement le .xls pour les envois CRM."""
    result = []
    if os.path.exists(path_xls):
        result.append((os.path.basename(path_xls), get_xlsx_bytes(path_xls)))
    return result


def attachments_xlsx_only(path_xlsx, path_xls):
    """Retourne uniquement le .xlsx pour l'admin."""
    result = []
    if os.path.exists(path_xlsx):
        result.append((os.path.basename(path_xlsx), get_xlsx_bytes(path_xlsx)))
    return result


# ─────────────────────────────────────────────
# MODE INDIVIDUAL
# ─────────────────────────────────────────────

def run_individual():
    print(f"\n=== INDIVIDUAL — {AGENCY}/{INITIALS} — {RUN_DATE} ===")
    if not AGENCY or not INITIALS:
        print("✗ AGENCY et INITIALS requis"); sys.exit(1)

    prospects_raw = fetch_prospects(agency=AGENCY, initials=INITIALS, date=RUN_DATE)
    if not prospects_raw:
        print("  Aucun prospect"); return

    # Filtrer par agence et initiales
    rows = [normalize_prospect(p) for p in prospects_raw
            if (p.get("agency") or "").upper() == AGENCY.upper()
            and (p.get("initials") or "").upper() == INITIALS.upper()]
    print(f"  → {len(rows)} fiche(s) après filtre {AGENCY}/{INITIALS}")
    path_xlsx, path_xls = export_commercial(RUN_DATE, AGENCY, INITIALS, rows, OUT_DIR)

    ag_cfg     = MAIL_ROUTING.get("agencies", {}).get(AGENCY, {})
    commercial = ag_cfg.get("commercial", {})
    manager    = ag_cfg.get("manager", {})
    subject    = f"Prospection {RUN_DATE} — {AGENCY}/{INITIALS} ({len(rows)} fiche(s))"
    html       = html_body(rows, AGENCY, INITIALS, RUN_DATE, "individual")

    sent_to = set()

    # 1. Commercial → .xls + .xlsx
    if commercial.get("email"):
        att = build_email_attachments(path_xlsx, path_xls)
        send_brevo(commercial["email"], f"Commercial {INITIALS}", subject, html, att)
        sent_to.add(commercial["email"])

    # 2. Manager agence → .xls + .xlsx
    if manager.get("email") and manager["email"] not in sent_to:
        att = build_email_attachments(path_xlsx, path_xls)
        send_brevo(manager["email"], f"Manager {AGENCY}", subject, html, att)
        sent_to.add(manager["email"])

    # 3. Admin SL → .xls + .xlsx (si SL a travaillé pour cette agence)
    #    SL reçoit toujours l'export individuel pour les agences où il intervient
    if ADMIN_EMAIL and ADMIN_EMAIL not in sent_to:
        att = build_email_attachments(path_xlsx, path_xls)
        send_brevo(ADMIN_EMAIL, "Admin SL", subject, html, att)
        sent_to.add(ADMIN_EMAIL)

    print(f"✅ Individual terminé — {len(sent_to)} destinataire(s)")


# ─────────────────────────────────────────────
# MODE AGENCY_MANAGER
# ─────────────────────────────────────────────

def run_agency_manager():
    print(f"\n=== AGENCY_MANAGER — {RUN_DATE} ===")
    ag_config = MAIL_ROUTING.get("agencies", {})

    for agency in ["GR", "VR", "GRS", "SLS"]:
        config = ag_config.get(agency, {})
        print(f"\n  Agence {agency}...")
        prospects_raw = fetch_prospects(agency=agency, date=RUN_DATE)
        if not prospects_raw:
            print(f"    Aucun prospect"); continue

        rows = [normalize_prospect(p) for p in prospects_raw
                if (p.get("agency") or "").upper() == agency.upper()]        path_xlsx, path_xls = export_agency_manager(RUN_DATE, agency, rows, OUT_DIR)
        manager = config.get("manager", {})
        subject = f"Récap {RUN_DATE} — Agence {agency} ({len(rows)} fiche(s))"
        html    = html_body(rows, agency, "", RUN_DATE, "agency_manager")

        # Manager → .xls uniquement
        if manager.get("email"):
            att = attachments_xls_only(path_xlsx, path_xls)
            send_brevo(manager["email"], f"Manager {agency}", subject, html, att)

    print(f"\n✅ Agency_manager terminé")


# ─────────────────────────────────────────────
# MODE ADMIN
# ─────────────────────────────────────────────

def run_admin():
    print(f"\n=== ADMIN — {RUN_DATE} ===")
    prospects_raw = fetch_prospects(date=RUN_DATE)
    if not prospects_raw:
        print("  Aucun prospect"); return

    rows = [normalize_prospect(p) for p in prospects_raw]
    path_xlsx, path_xls = export_admin(RUN_DATE, rows, OUT_DIR)

    subject = f"Récap admin {RUN_DATE} — Toutes agences ({len(rows)} fiche(s))"
    html    = html_body(rows, "", "", RUN_DATE, "admin")

    # Admin SL → .xlsx uniquement
    if ADMIN_EMAIL:
        att = attachments_xlsx_only(path_xlsx, path_xls)
        send_brevo(ADMIN_EMAIL, "Admin SL", subject, html, att)

    print(f"✅ Admin terminé")


# ─────────────────────────────────────────────
# POINT D'ENTRÉE
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print(f"Prospection Export — mode={SEND_MODE} date={RUN_DATE}")
    if not WORKER_BASE_URL: print("✗ WORKER_BASE_URL manquant"); sys.exit(1)
    if not EXPORT_TOKEN:    print("✗ EXPORT_TOKEN manquant");    sys.exit(1)

    if   SEND_MODE == "individual":     run_individual()
    elif SEND_MODE == "agency_manager": run_agency_manager()
    elif SEND_MODE == "admin":          run_admin()
    else: print(f"✗ SEND_MODE inconnu: {SEND_MODE}"); sys.exit(1)
