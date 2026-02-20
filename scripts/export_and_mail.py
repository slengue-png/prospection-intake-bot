import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime, date
from openpyxl import Workbook
import yaml
import re
from pathlib import Path

# =========================
# EXCEL IMPORT (AGENCES)
# =========================
IMPORT_COLUMNS = [
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
]

BROWSER_HEADERS = {
"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
"Accept": "*/*",
"Accept-Language": "fr-FR,fr;q=0.9",
"Connection": "close",
}

# =========================
# CONFIG
# =========================
def load_config(path="config.yml"):
with open(path, "r", encoding="utf-8") as f:
return yaml.safe_load(f)

def normalize_emails(x):
if x is None:
return []
if isinstance(x, list):
raw = x
elif isinstance(x, str):
raw = [p.strip() for p in x.split(",")]
else:
raw = []
out = []
for e in raw:
e = (e or "").strip()
if e:
out.append(e)
return out

# =========================
# FETCH /dump (Worker)
# =========================
def fetch_jsonl(url, export_token):
print(f"[FETCH] {url}")
if not export_token:
raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

req = urllib.request.Request(url)
for k, v in BROWSER_HEADERS.items():
req.add_header(k, v)
req.add_header("X-Export-Token", export_token)

try:
with urllib.request.urlopen(req, timeout=45) as resp:
raw = resp.read().decode("utf-8", errors="replace").strip()
print("[HTTP]", resp.status)
except urllib.error.HTTPError as e:
body = e.read().decode("utf-8", errors="replace")
raise RuntimeError(f"HTTPError {e.code} on {url} -> {body[:300]}")
except urllib.error.URLError as e:
raise RuntimeError(f"URLError on {url} -> {e.reason}")

if not raw:
return []
rows = []
for line in raw.splitlines():
line = line.strip()
if not line:
continue
try:
rows.append(json.loads(line))
except:
pass
return rows

# =========================
# DATE PARSING (ROBUSTE)
# =========================
def parse_date_flexible(s: str) -> date:
"""
Accepte:
- YYYY-MM-DD
- DD/MM/YYYY
- DD-MM-YYYY
Rejette le reste avec une erreur claire.
"""
s = (s or "").strip()
if not s:
raise ValueError("Empty date")

# 1) YYYY-MM-DD
try:
return datetime.strptime(s, "%Y-%m-%d").date()
except:
pass

# 2) DD/MM/YYYY
try:
return datetime.strptime(s, "%d/%m/%Y").date()
except:
pass

# 3) DD-MM-YYYY
try:
return datetime.strptime(s, "%d-%m-%Y").date()
except:
pass

raise ValueError(
f"RUN_DATE invalide: '{s}'. Formats acceptés: YYYY-MM-DD (ex: 2026-02-19) ou DD/MM/YYYY (ex: 19/02/2026)."
)

def last_day_of_month(d: date) -> date:
if d.month == 12:
return date(d.year, 12, 31)
nxt = date(d.year, d.month + 1, 1)
return nxt.fromordinal(nxt.toordinal() - 1)

def month_key(d: date) -> str:
return f"{d.year:04d}-{d.month:02d}"

# =========================
# COMMANDES -> comptage intelligent
# =========================
PARASITE_WORDS = {
"besoin","poste","postes","profil","profils","recherche","recherchons",
"urgent","urgente","asap","svp","stp","merci","cdi","cdd","interim",
"intérim","mission","missions","pour","de","des","du","la","le","les",
"a","à","au","aux","et","ou","sur","en","d","l","un","une",
}

def _stem_fr(word: str) -> str:
w = word.strip().lower()
w = re.sub(r"[^a-zàâäéèêëîïôöùûüç0-9\-]", "", w)
if not w:
return ""
for suf in ["es", "s", "x"]:
if w.endswith(suf) and len(w) > 3:
w = w[: -len(suf)]
break
return w

def count_commandes(commande_text: str) -> int:
if not commande_text:
return 0
s = str(commande_text).strip()
if not s:
return 0

if re.fullmatch(r"\d+", s):
return 1

s = s.replace("\n", " ")
s = re.sub(r"[;,/|]+", " ", s)
s = re.sub(r"\s+", " ", s).strip()

groups = re.findall(r"\b\d+\s+([A-Za-zÀ-ÖØ-öø-ÿ\-]+)", s)
if groups:
uniq = set()
for g in groups:
w = _stem_fr(g)
if w and w not in PARASITE_WORDS:
uniq.add(w)
return max(1, len(uniq)) if uniq else 1

words = re.split(r"\s+", s)
uniq = set()
for w in words:
ww = _stem_fr(w)
if not ww:
continue
if ww in PARASITE_WORDS:
continue
if ww.isdigit():
continue
uniq.add(ww)

return len(uniq) if uniq else 1

# =========================
# EXCEL
# =========================
def build_excel(records):
wb = Workbook()
ws = wb.active
ws.title = "IMPORT"
ws.append(IMPORT_COLUMNS)

for r in records:
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
r.get("contact_lastname", "") or r.get("interlocuteur", ""),
r.get("dirigeant", ""),
r.get("resume", ""),
r.get("commande", ""),
])

bio = io.BytesIO()
wb.save(bio)
return bio.getvalue()

# =========================
# EMAIL (Brevo API)
# =========================
def send_email_brevo(subject, body, to_list, attachments):
api_key = os.environ.get("BREVO_API_KEY", "").strip()
if not api_key:
raise RuntimeError("BREVO_API_KEY missing in GitHub Secrets")

to_list = normalize_emails(to_list)
if not to_list:
print(f"[BREVO][SKIP] no recipients for subject: {subject}")
return

sender_email = os.environ.get("BREVO_SENDER_EMAIL", "").strip()
if not sender_email:
raise RuntimeError("BREVO_SENDER_EMAIL missing in GitHub Secrets")

payload = {
"sender": {"email": sender_email, "name": "Prospection Bot"},
"to": [{"email": x} for x in to_list],
"subject": subject,
"textContent": body,
}

if attachments:
payload["attachment"] = []
for filename, content_bytes in attachments:
payload["attachment"].append({
"name": filename,
"content": base64.b64encode(content_bytes).decode("utf-8"),
})

data = json.dumps(payload).encode("utf-8")
req = urllib.request.Request("https://api.brevo.com/v3/smtp/email", data=data, method="POST")
req.add_header("Content-Type", "application/json")
req.add_header("api-key", api_key)

try:
with urllib.request.urlopen(req, timeout=30) as resp:
print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)
_ = resp.read()
except urllib.error.HTTPError as e:
err = e.read().decode("utf-8", errors="replace")
print("[BREVO][HTTPERROR]", e.code, err[:400].replace("\n", " "))
raise

# =========================
# CUMUL MENSUEL (repo)
# =========================
def cumul_path_for(d: date) -> Path:
Path("data").mkdir(parents=True, exist_ok=True)
return Path("data") / f"cumul_{month_key(d)}.json"

def load_cumul(path: Path) -> dict:
if not path.exists():
return {"month": path.stem.replace("cumul_", ""), "days": {}, "updated_at_utc": ""}
try:
return json.loads(path.read_text(encoding="utf-8"))
except:
return {"month": path.stem.replace("cumul_", ""), "days": {}, "updated_at_utc": ""}

def save_cumul(path: Path, data: dict):
data["updated_at_utc"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def aggregate_day(prospects: list[dict], agencies: dict) -> dict:
out = {
"agencies": {},
"by_user": {},
"totals": {"prospects": 0, "clients": 0, "commandes": 0},
}

for ag in agencies.keys():
out["agencies"][ag] = {"prospects": 0, "clients": 0, "commandes": 0}

for r in prospects:
ag = (r.get("agency") or "").strip()
ini = (r.get("initials") or "NA").strip().upper()

cmds = count_commandes(r.get("commande", ""))

if ag in out["agencies"]:
out["agencies"][ag]["prospects"] += 1
out["agencies"][ag]["commandes"] += cmds

key = f"{ag}|{ini}"
out["by_user"].setdefault(key, {"agency": ag, "initials": ini, "prospects": 0, "clients": 0, "commandes": 0})
out["by_user"][key]["prospects"] += 1
out["by_user"][key]["commandes"] += cmds

out["totals"]["prospects"] += 1
out["totals"]["commandes"] += cmds

return out

def aggregate_period(cumul: dict, start_day: int, end_day: int) -> dict:
res = {
"agencies": {},
"by_user": {},
"totals": {"prospects": 0, "clients": 0, "commandes": 0},
}

for day_str, day_data in cumul.get("days", {}).items():
try:
dd = datetime.strptime(day_str, "%Y-%m-%d").date()
except:
continue
if not (start_day <= dd.day <= end_day):
continue

t = day_data.get("totals", {})
res["totals"]["prospects"] += int(t.get("prospects", 0))
res["totals"]["clients"] += int(t.get("clients", 0))
res["totals"]["commandes"] += int(t.get("commandes", 0))

for ag, a in (day_data.get("agencies", {}) or {}).items():
res["agencies"].setdefault(ag, {"prospects": 0, "clients": 0, "commandes": 0})
res["agencies"][ag]["prospects"] += int(a.get("prospects", 0))
res["agencies"][ag]["clients"] += int(a.get("clients", 0))
res["agencies"][ag]["commandes"] += int(a.get("commandes", 0))

for k, u in (day_data.get("by_user", {}) or {}).items():
res["by_user"].setdefault(k, {"agency": u.get("agency",""), "initials": u.get("initials","NA"),
"prospects": 0, "clients": 0, "commandes": 0})
res["by_user"][k]["prospects"] += int(u.get("prospects", 0))
res["by_user"][k]["clients"] += int(u.get("clients", 0))
res["by_user"][k]["commandes"] += int(u.get("commandes", 0))

return res

def make_ranking(by_user: dict) -> list[dict]:
rows = []
for _, u in by_user.items():
p = int(u.get("prospects", 0))
c = int(u.get("clients", 0))
cmd = int(u.get("commandes", 0))
tx = (cmd / p * 100.0) if p else 0.0
rows.append({
"agency": u.get("agency", ""),
"initials": u.get("initials", "NA"),
"prospects": p,
"clients": c,
"commandes": cmd,
"tx": tx,
})
rows.sort(key=lambda r: (r["commandes"], r["tx"], r["prospects"]), reverse=True)
return rows

# =========================
# MAIN
# =========================
def main():
cfg = load_config()
worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
if not worker.startswith("https://"):
raise RuntimeError("config.yml: worker_base_url must start with https://")

export_token = os.environ.get("EXPORT_TOKEN", "").strip()
if not export_token:
raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

run_date_str = (os.environ.get("RUN_DATE") or "").strip()

if run_date_str:
d = parse_date_flexible(run_date_str)
else:
d = datetime.utcnow().date()

date_str = d.strftime("%Y-%m-%d")

agencies_cfg = cfg.get("agencies", {})

# ===== dump prospects du jour =====
url = f"{worker}/dump?date={date_str}&kind=prospects"
prospects = fetch_jsonl(url, export_token)
print("[OK] prospects rows=", len(prospects))

agency_records = {ag: [] for ag in agencies_cfg.keys()}
for p in prospects:
ag = p.get("agency", "")
if ag in agency_records:
agency_records[ag].append(p)

# ===== emails agences (excel import) =====
for ag, recs in agency_records.items():
to_list = agencies_cfg.get(ag, {}).get("daily_to", [])
excel = build_excel(recs)

nb_prospects = len(recs)
nb_clients = 0
nb_commandes = sum(count_commandes(r.get("commande", "")) for r in recs)

send_email_brevo(
f"[PROSPECTION] {ag} — {date_str}",
f"Export agence {ag}\n"
f"Prospects: {nb_prospects}\n"
f"Clients: {nb_clients}\n"
f"Commandes: {nb_commandes}\n",
to_list,
[(f"{date_str}_{ag}_IMPORT.xlsx", excel)],
)

# ===== email global quotidien =====
lines = [f"Résumé global prospection — {date_str}", ""]
total_p = 0
total_c = 0
total_cmd = 0

for ag in agency_records.keys():
recs = agency_records[ag]
nb_p = len(recs)
nb_c = 0
nb_cmd = sum(count_commandes(r.get("commande", "")) for r in recs)

total_p += nb_p
total_c += nb_c
total_cmd += nb_cmd

lines.append(f"{ag} — Prospects: {nb_p} | Clients: {nb_c} | Commandes: {nb_cmd}")

by_init = {}
for r in recs:
ini = (r.get("initials") or "NA").strip().upper()
by_init.setdefault(ini, []).append(r)

for ini, rr in sorted(by_init.items()):
pcount = len(rr)
ccount = 0
cmdcount = sum(count_commandes(x.get("commande", "")) for x in rr)
tx = (cmdcount / pcount * 100.0) if pcount else 0.0
lines.append(f" - {ini}: {pcount} prospects | {ccount} clients | {cmdcount} commandes | Tx: {tx:.1f}%")

lines.append("")

lines.append(f"TOTAL GLOBAL — Prospects: {total_p} | Clients: {total_c} | Commandes: {total_cmd}")

global_to = cfg.get("global_to", [])
send_email_brevo(
f"[PROSPECTION] GLOBAL — {date_str}",
"\n".join(lines),
global_to,
[],
)

# ===== cumul mensuel =====
cumul_file = cumul_path_for(d)
cumul = load_cumul(cumul_file)

day_agg = aggregate_day(prospects, agencies_cfg)
cumul["days"][date_str] = day_agg
save_cumul(cumul_file, cumul)
print("[CUMUL] updated:", str(cumul_file))

# ===== envois 15/15 et fin de mois =====
ld = last_day_of_month(d)
send_mid = (d.day == 15)
send_end = (d == ld)

def send_agency_period(ag: str, start_day: int, end_day: int, label: str):
period = aggregate_period(cumul, start_day, end_day)
a = (period.get("agencies", {}) or {}).get(ag, {"prospects":0,"clients":0,"commandes":0})

lines_ag = [f"Récap {label} — {month_key(d)} — agence {ag}", ""]
lines_ag.append(f"Prospects: {a.get('prospects',0)} | Clients: {a.get('clients',0)} | Commandes: {a.get('commandes',0)}")
lines_ag.append("")

users = {}
for k, u in (period.get("by_user", {}) or {}).items():
if (u.get("agency") or "") == ag:
users[k] = u
rank = make_ranking(users)
if rank:
lines_ag.append("Classement commerciaux (agence)")
for i, r in enumerate(rank, 1):
lines_ag.append(f"{i}. {r['initials']} — {r['prospects']} prospects | {r['commandes']} commandes | Tx: {r['tx']:.1f}%")
else:
lines_ag.append("Aucune donnée sur la période.")

to_list = agencies_cfg.get(ag, {}).get("daily_to", [])
send_email_brevo(
f"[PROSPECTION] {ag} — {label} — {month_key(d)}",
"\n".join(lines_ag),
to_list,
[],
)

def send_global_period(start_day: int, end_day: int, label: str):
period = aggregate_period(cumul, start_day, end_day)
t = period.get("totals", {"prospects":0,"clients":0,"commandes":0})

lines_g = [f"Récap GLOBAL {label} — {month_key(d)}", ""]
lines_g.append("Par agence")
for ag in agencies_cfg.keys():
a = (period.get("agencies", {}) or {}).get(ag, {"prospects":0,"clients":0,"commandes":0})
lines_g.append(f"- {ag}: Prospects {a.get('prospects',0)} | Clients {a.get('clients',0)} | Commandes {a.get('commandes',0)}")
lines_g.append("")
lines_g.append(f"TOTAL: Prospects {t.get('prospects',0)} | Clients {t.get('clients',0)} | Commandes {t.get('commandes',0)}")
lines_g.append("")

rank = make_ranking(period.get("by_user", {}) or {})
lines_g.append("Classement commerciaux (global)")
if rank:
for i, r in enumerate(rank, 1):
lines_g.append(f"{i}. {r['agency']} / {r['initials']} — {r['prospects']} prospects | {r['commandes']} commandes | Tx: {r['tx']:.1f}%")
else:
lines_g.append("Aucune donnée sur la période.")

send_email_brevo(
f"[PROSPECTION] GLOBAL — {label} — {month_key(d)}",
"\n".join(lines_g),
global_to,
[],
)

if send_mid:
for ag in agencies_cfg.keys():
send_agency_period(ag, 1, 15, "01-15")
send_global_period(1, 15, "01-15")

if send_end:
for ag in agencies_cfg.keys():
send_agency_period(ag, 16, ld.day, "16-fin")
send_global_period(16, ld.day, "16-fin")


if __name__ == "__main__":
main()
