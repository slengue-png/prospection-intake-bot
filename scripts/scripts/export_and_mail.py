"""
export_and_mail.py — Script principal export prospection
=========================================================
Appelé par GitHub Actions (prospection-close.yml et prospection-daily.yml)

Variables d'environnement requises :
  WORKER_BASE_URL   : URL de base du worker (ex: https://prospection-intake-worker.slengue.workers.dev)
  EXPORT_TOKEN      : Token d'authentification pour /dump
  TELEGRAM_TOKEN    : Token bot Telegram (pour notifications + images)
  BREVO_API_KEY     : Clé API Brevo
  BREVO_SENDER_EMAIL: Email expéditeur
  BREVO_SENDER_NAME : Nom expéditeur
  MAIL_ROUTING_JSON : JSON de routing email (voir config)
  SEND_MODE         : individual | agency_manager | admin
  RUN_DATE          : Date YYYY-MM-DD
  AGENCY            : GR|VR|GRS|SLS (mode individual)
  INITIALS          : Initiales (mode individual)
"""

import os
import sys
import json
import datetime
import base64
import requests
from typing import Any, Dict, List, Optional

from excel_export import (
    export_commercial,
    export_agency_manager,
    export_admin,
    build_email_attachments,
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
OUT_DIR            = "out"

try:
    MAIL_ROUTING = json.loads(MAIL_ROUTING_JSON)
except Exception:
    MAIL_ROUTING = {}


# ─────────────────────────────────────────────
# RÉCUPÉRATION DONNÉES WORKER
# ─────────────────────────────────────────────

def fetch_prospects(agency: str = "", initials: str = "", date: str = "") -> List[Dict]:
    """Récupère les prospects depuis le Worker via /dump."""
    params = {"token": EXPORT_TOKEN}
    if date:     params["date"]     = date
    if agency:   params["agency"]   = agency
    if initials: params["initials"] = initials

    url = f"{WORKER_BASE_URL}/dump"
    try:
        r = requests.get(url, params=params, timeout=60)
        r.raise_for_status()
        lines = r.text.strip().split("\n")
        prospects = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
                if isinstance(obj, dict):
                    prospects.append(obj)
            except Exception:
                pass
        print(f"  → {len(prospects)} prospect(s) récupéré(s)")
        return prospects
    except Exception as e:
        print(f"  ✗ Erreur fetch prospects: {e}")
        return []


def normalize_prospect(p: Dict) -> Dict:
    """Normalise un prospect brut depuis le KV vers les champs Excel."""
    d = p.get("draft", p)
    return {
        "name":             d.get("name", ""),
        "nom_commercial":   d.get("nom_commercial", ""),
        "address":          d.get("address", ""),
        "postal_code":      d.get("postal_code", ""),
        "city":             d.get("city", ""),
        "phone":            d.get("phone", ""),
        "phone2":           d.get("phone2", ""),
        "email":            d.get("email", ""),
        "siret":            d.get("siret", ""),
        "naf":              d.get("naf", ""),
        "activity_summary": d.get("secteur_activite", ""),
        "effectif":         d.get("effectif", ""),
        "date_creation":    d.get("date_creation", ""),
        "capital":          d.get("capital", ""),
        "website":          d.get("website", ""),
        "contact_civility": d.get("contact_civility", ""),
        "interlocuteur":    d.get("interlocuteur", ""),
        "dirigeant":        d.get("dirigeant", ""),
        "resume":           d.get("resume", ""),
        "commande":         d.get("commande", ""),
        "qualification":    d.get("qualification", ""),
        "card_photo_url":   d.get("card_photo_url", ""),
        "card_photo_file_id": d.get("card_photo_file_id", ""),
        "facade_photo_url":   d.get("facade_photo_url", ""),
        "facade_photo_file_id": d.get("facade_photo_file_id", ""),
        "agency":           p.get("agency", d.get("agency", "")),
        "initials":         p.get("initials", d.get("initials", "")),
        "_confidence":      p.get("_confidence", d.get("_confidence", 0)),
        "_score":           p.get("_score", d.get("_score", 0)),
    }


# ─────────────────────────────────────────────
# ENVOI EMAIL BREVO
# ─────────────────────────────────────────────

def send_email_brevo(
    to_email: str,
    to_name: str,
    subject: str,
    html_body: str,
    attachments: List = None,
) -> bool:
    """Envoie un email via l'API Brevo."""
    if not BREVO_API_KEY or not BREVO_SENDER_EMAIL:
        print(f"  ✗ Brevo non configuré")
        return False

    payload = {
        "sender":  {"email": BREVO_SENDER_EMAIL, "name": BREVO_SENDER_NAME},
        "to":      [{"email": to_email, "name": to_name}],
        "subject": subject,
        "htmlContent": html_body,
    }

    if attachments:
        payload["attachment"] = [
            {
                "name":    name,
                "content": base64.b64encode(content).decode("utf-8"),
            }
            for name, content in attachments
        ]

    try:
        r = requests.post(
            "https://api.brevo.com/v3/smtp/email",
            headers={
                "api-key": BREVO_API_KEY,
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=60,
        )
        if r.status_code in (200, 201):
            print(f"  ✓ Email envoyé à {to_email}")
            return True
        else:
            print(f"  ✗ Brevo erreur {r.status_code}: {r.text[:200]}")
            return False
    except Exception as e:
        print(f"  ✗ Brevo exception: {e}")
        return False


def build_html_body(rows: List[Dict], agency: str, initials: str, date: str, mode: str) -> str:
    """Construit le corps HTML de l'email."""
    count = len(rows)
    if mode == "individual":
        title = f"Prospection {date} — {agency} / {initials}"
        intro = f"Bonjour,<br><br>Veuillez trouver en pièce jointe le rapport de prospection du <b>{date}</b> pour <b>{initials}</b> ({agency}).<br><br><b>{count} fiche(s)</b> saisie(s) ce jour."
    elif mode == "agency_manager":
        title = f"Récapitulatif journalier {date} — Agence {agency}"
        intro = f"Bonjour,<br><br>Voici le récapitulatif de prospection du <b>{date}</b> pour l'agence <b>{agency}</b>.<br><br><b>{count} fiche(s)</b> au total."
    else:
        title = f"Récapitulatif admin {date}"
        intro = f"Bonjour,<br><br>Récapitulatif complet de prospection du <b>{date}</b>.<br><br><b>{count} fiche(s)</b> toutes agences confondues."

    return f"""
    <html><body style="font-family: Arial, sans-serif; color: #333;">
    <h2 style="color: #2C3E50;">{title}</h2>
    <p>{intro}</p>
    <p>Les fichiers joints :<br>
    • <b>.xlsx</b> : consultation avec miniatures photos<br>
    • <b>.xls</b> : import CRM</p>
    <hr style="border: none; border-top: 1px solid #eee;">
    <p style="color: #999; font-size: 11px;">Généré automatiquement par Prospection Bot v8.4</p>
    </body></html>
    """


# ─────────────────────────────────────────────
# NOTIFICATION TELEGRAM
# ─────────────────────────────────────────────

def get_chat_id_for_initials(initials: str) -> Optional[str]:
    """Récupère le chat_id Telegram depuis le routing."""
    users = MAIL_ROUTING.get("users", {})
    # Pas de chat_id dans le routing — on skip la notification Telegram
    return None


def notify_telegram(chat_id: str, message: str) -> bool:
    """Envoie une notification Telegram."""
    if not TELEGRAM_TOKEN or not chat_id:
        return False
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            json={"chat_id": chat_id, "text": message, "parse_mode": "HTML"},
            timeout=15,
        )
        return r.status_code == 200
    except Exception:
        return False


# ─────────────────────────────────────────────
# MODES D'EXÉCUTION
# ─────────────────────────────────────────────

def run_individual():
    """Export individuel pour un commercial (déclenché par le Worker)."""
    print(f"\n=== MODE INDIVIDUAL — {AGENCY}/{INITIALS} — {RUN_DATE} ===")

    if not AGENCY or not INITIALS:
        print("✗ AGENCY et INITIALS requis en mode individual")
        sys.exit(1)

    # Récupérer les prospects
    prospects_raw = fetch_prospects(agency=AGENCY, initials=INITIALS, date=RUN_DATE)
    if not prospects_raw:
        print(f"  Aucun prospect pour {AGENCY}/{INITIALS} le {RUN_DATE}")
        return

    rows = [normalize_prospect(p) for p in prospects_raw]
    print(f"  {len(rows)} fiche(s) normalisée(s)")

    # Générer les fichiers Excel
    path_xlsx, path_xls = export_commercial(RUN_DATE, AGENCY, INITIALS, rows, OUT_DIR)
    print(f"  Excel: {path_xlsx}, {path_xls}")

    attachments = build_email_attachments(path_xlsx, path_xls)

    # Routing email
    ag_config   = MAIL_ROUTING.get("agencies", {}).get(AGENCY, {})
    commercial  = ag_config.get("commercial", {})
    manager     = ag_config.get("manager", {})
    admin_cfg   = MAIL_ROUTING.get("admin", {})

    subject = f"Prospection {RUN_DATE} — {AGENCY}/{INITIALS} ({len(rows)} fiche(s))"
    html    = build_html_body(rows, AGENCY, INITIALS, RUN_DATE, "individual")

    # Envoyer au commercial
    if commercial.get("email"):
        send_email_brevo(commercial["email"], f"Commercial {INITIALS}", subject, html, attachments)

    # Envoyer au manager d'agence
    if manager.get("email") and manager.get("email") != commercial.get("email"):
        send_email_brevo(manager["email"], f"Manager {AGENCY}", subject, html, attachments)

    # Envoyer à l'admin
    if admin_cfg.get("email"):
        send_email_brevo(admin_cfg["email"], "Admin", subject, html, attachments)

    print(f"✅ Export individual terminé")


def run_agency_manager():
    """Récapitulatif journalier par agence pour les managers."""
    print(f"\n=== MODE AGENCY_MANAGER — {RUN_DATE} ===")

    ag_config = MAIL_ROUTING.get("agencies", {})

    for agency, config in ag_config.items():
        print(f"\n  Agence {agency}...")
        prospects_raw = fetch_prospects(agency=agency, date=RUN_DATE)
        if not prospects_raw:
            print(f"    Aucun prospect")
            continue

        rows = [normalize_prospect(p) for p in prospects_raw]
        path_xlsx, path_xls = export_agency_manager(RUN_DATE, agency, rows, OUT_DIR)
        attachments = build_email_attachments(path_xlsx, path_xls)

        manager = config.get("manager", {})
        subject = f"Récap prospection {RUN_DATE} — Agence {agency} ({len(rows)} fiche(s))"
        html    = build_html_body(rows, agency, "", RUN_DATE, "agency_manager")

        if manager.get("email"):
            send_email_brevo(manager["email"], f"Manager {agency}", subject, html, attachments)

    print(f"\n✅ Export agency_manager terminé")


def run_admin():
    """Récapitulatif complet pour l'admin."""
    print(f"\n=== MODE ADMIN — {RUN_DATE} ===")

    prospects_raw = fetch_prospects(date=RUN_DATE)
    if not prospects_raw:
        print("  Aucun prospect")
        return

    rows = [normalize_prospect(p) for p in prospects_raw]
    path_xlsx, path_xls = export_admin(RUN_DATE, rows, OUT_DIR)
    attachments = build_email_attachments(path_xlsx, path_xls)

    admin_cfg = MAIL_ROUTING.get("admin", {})
    subject   = f"Récap admin prospection {RUN_DATE} ({len(rows)} fiche(s))"
    html      = build_html_body(rows, "", "", RUN_DATE, "admin")

    if admin_cfg.get("email"):
        send_email_brevo(admin_cfg["email"], "Admin", subject, html, attachments)

    print(f"✅ Export admin terminé")


# ─────────────────────────────────────────────
# POINT D'ENTRÉE
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print(f"Prospection Export — mode={SEND_MODE} date={RUN_DATE}")

    if not WORKER_BASE_URL:
        print("✗ WORKER_BASE_URL manquant")
        sys.exit(1)
    if not EXPORT_TOKEN:
        print("✗ EXPORT_TOKEN manquant")
        sys.exit(1)

    if SEND_MODE == "individual":
        run_individual()
    elif SEND_MODE == "agency_manager":
        run_agency_manager()
    elif SEND_MODE == "admin":
        run_admin()
    else:
        print(f"✗ SEND_MODE inconnu: {SEND_MODE}")
        sys.exit(1)
