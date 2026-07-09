// =====================================================
// PROSPECTION INTAKE WORKER — V8.4
//
// Hérite de v8.3 — Cherry-pick des meilleures idées v10
//
// Nouveautés v8.4 :
// SÉCURITÉ
// - timingSafeEqual : comparaison constante des secrets (anti timing-attack)
// - checkAdminToken Bearer pour /dump, /remind, /update
//
// ROBUSTESSE FETCH
// - safeFetchJson / safeFetchText : retry intégré + timeout sur tous les appels HTTP
// - Tous les appels API (Gouv, Places, Gemini, Vision, BAN) passent par safeFetch
//
// PERFORMANCE KV
// - Index KV dédié idx:day:{date}:{agency} pour la déduplication
//   → plus de list scan sur tous les prospects, O(1) au lieu de O(n)
// - Index idx:dup:{date}:{agency} pour fingerprint rapide
// - addDayIndex + addDupIndex non bloquants (waitUntil)
//
// UX TELEGRAM
// - Messages spinner ⏳ pendant enrichissements lourds (supprimés après)
// - pinFormMessage / unpinFormMessage : fiche active épinglée dans le chat
// - Commandes rapides : /n (nouvelle visite), /s (save), /l (liste), /g (GPS)
// - force_reply sur les champs texte : réponse liée au message
// - buildPersistentBottomKeyboard contextuel (présence fiche active)
//
// BACKGROUND TASKS
// - scheduleBackgroundTask via event.waitUntil : GitHub dispatch non bloquant
// - Enrichissement OCR GitHub Actions en arrière-plan
//
// Conservé de v8.3 (inchangé) :
// - Qualification 1-tap, saisie vocale, lot photo/carte
// - OCR hybride Gemini + Google Vision
// - Validation BAN, auto-save, dédoublonnage 30j
// - Champs effectif, secteur_activite, date_creation, capital
//
// Endpoints :
// - GET  /health
// - GET  /dump?date=YYYY-MM-DD&kind=prospects|closes|photos|cards
// - POST /remind
// - POST /telegram
// - POST /update   (enrichissement OCR GitHub Actions)
// =====================================================

const JSON_HEADERS = { "content-type": "application/json; charset=utf-8" };

const SCRAPE_MAX_BYTES  = 200_000;
const SCRAPE_TIMEOUT_MS = 4_000;

const TTL_SESSION  = 60 * 60 * 24 * 7;   // 7j  — sessions actives
const TTL_PREFS    = 60 * 60 * 24 * 30;  // 30j — préférences utilisateur
const TTL_INDEX    = 60 * 60 * 24 * 35;  // 35j — index KV (> DEDUP_WINDOW_DAYS)
const TTL_PROSPECT = 60 * 60 * 24 * 90;  // 90j — fiches prospects (durabilité CRM)

const CACHE_TTL_GOUV   = 60 * 60 * 24;
const CACHE_TTL_PLACES = 60 * 60 * 24 * 7;
const CACHE_TTL_EMAIL  = 60 * 60 * 24 * 7;
const CACHE_TTL_GEMINI = 60 * 60 * 24 * 7;

const GITHUB_QUEUE_TTL = 60 * 60 * 24 * 7;

const PHOTO_LIMIT = 15;
const CARD_LIMIT  = 15;

const MAX_FIELD_LEN = 500;
const GEO_MAX_AGE_MS = 10_000;

// Dédoublonnage élargi à 30 jours
const DEDUP_WINDOW_DAYS = 30;

// Timeouts fetch
const DEFAULT_FETCH_TIMEOUT_MS = 8_000;
const GEMINI_FETCH_TIMEOUT_MS  = 12_000;
const VISION_FETCH_TIMEOUT_MS  = 10_000;
const PLACES_FETCH_TIMEOUT_MS  = 8_000;
const GOUV_FETCH_TIMEOUT_MS    = 8_000;
const TG_FILE_TIMEOUT_MS       = 12_000;
const GITHUB_FETCH_TIMEOUT_MS  = 12_000;
const API_RETRIES = 1;
const MAX_RETRY_SLEEP_MS = 1_500;

const VALID_AGENCIES = ["GR", "VR", "GRS", "SLS"];

// Qualification
const QUALIF_HOT  = "chaud";
const QUALIF_WARM = "tiede";
const QUALIF_COLD = "froid";

// NAF → libellé secteur (extrait partiel INSEE)
const NAF_LABELS = {
  "4332A": "Menuiserie bois et matières plastiques",
  "4332B": "Menuiserie métallique et serrurerie",
  "4333Z": "Travaux de revêtement des sols et des murs",
  "4399D": "Autres travaux spécialisés de construction",
  "4120A": "Construction de maisons individuelles",
  "4120B": "Construction d'autres bâtiments",
  "4941A": "Transports routiers de fret interurbains",
  "4941B": "Transports routiers de fret de proximité",
  "4942Z": "Services de déménagement",
  "5610A": "Restauration traditionnelle",
  "5610C": "Restauration rapide",
  "3811Z": "Collecte des déchets non dangereux",
  "3812Z": "Collecte des déchets dangereux",
  "4531Z": "Commerce de véhicules automobiles",
  "4532Z": "Commerce de détail de carburants",
  "4711A": "Commerce alimentaire",
  "4711B": "Commerce alimentaire de proximité",
  "8610Z": "Activités hospitalières",
  "8621Z": "Activité des médecins généralistes",
  "8690A": "Ambulances",
  "4910Z": "Transport ferroviaire interurbain",
  "5110Z": "Transport aérien de passagers",
  "6201Z": "Programmation informatique",
  "6202A": "Conseil en systèmes informatiques",
  "7022Z": "Conseil pour les affaires et la gestion",
  "6920Z": "Activités comptables",
  "6910Z": "Activités juridiques",
  "7112B": "Ingénierie et études techniques",
  "4321A": "Travaux d'installation électrique",
  "4322A": "Travaux d'installation d'eau et de gaz",
  "4322B": "Travaux d'installation d'équipements thermiques",
  "4531Z": "Pneumatiques / Entretien auto",
  // Commerce de gros
  "4669C": "Commerce de gros d'instruments de mesure et de contrôle",
  "4669A": "Commerce de gros de matériel électrique",
  "4669B": "Commerce de gros de fournitures industrielles",
  "4641Z": "Commerce de gros de textiles",
  "4642Z": "Commerce de gros d'habillement",
  "4611Z": "Agents en matières premières agricoles",
  "4621Z": "Commerce de gros de céréales",
  "4631Z": "Commerce de gros de fruits et légumes",
  "4632Z": "Commerce de gros de viandes",
  "4633Z": "Commerce de gros de produits laitiers",
  "4634Z": "Commerce de gros de boissons",
  "4635Z": "Commerce de gros de tabac",
  "4636Z": "Commerce de gros de sucre",
  "4637Z": "Commerce de gros de café",
  "4638A": "Commerce de gros de poissons",
  "4639A": "Commerce de gros alimentaire non spécialisé",
  "4651Z": "Commerce de gros d'ordinateurs",
  "4652Z": "Commerce de gros de composants électroniques",
  "4661Z": "Commerce de gros de matériel agricole",
  "4662Z": "Commerce de gros de machines-outils",
  "4663Z": "Commerce de gros de machines pour l'extraction",
  "4664Z": "Commerce de gros de machines pour l'industrie textile",
  "4665Z": "Commerce de gros de mobilier de bureau",
  // Commerce de détail spécialisé
  "4752A": "Commerce de détail de quincaillerie",
  "4753Z": "Commerce de détail de moquettes et revêtements",
  "4759A": "Commerce de détail de meubles",
  "4761Z": "Commerce de détail de livres",
  "4762Z": "Commerce de détail de journaux",
  "4763Z": "Commerce de détail d'enregistrements musicaux",
  "4764Z": "Commerce de détail d'articles de sport",
  "4765Z": "Commerce de détail de jeux et jouets",
  "4771Z": "Commerce de détail d'habillement",
  "4772A": "Commerce de détail de chaussures",
  "4773Z": "Commerce de détail de produits pharmaceutiques",
  "4774Z": "Commerce de détail d'articles médicaux",
  "4775Z": "Commerce de détail de parfumerie",
  "4776Z": "Commerce de détail de fleurs",
  "4777Z": "Commerce de détail d'articles d'horlogerie",
  "4778A": "Commerces de détail d'optique",
  "4778B": "Commerce de détail d'appareils photographiques",
  "4779Z": "Commerce de détail de biens d'occasion",
  // Services
  "7711A": "Location de courte durée de véhicules",
  "7712Z": "Location de camions",
  "7721Z": "Location d'articles de sport",
  "7729Z": "Location d'autres biens",
  "8110Z": "Services intégrés aux bâtiments",
  "8121Z": "Nettoyage courant des bâtiments",
  "8122Z": "Autres activités de nettoyage",
  "8129A": "Désinfection / dératisation",
};

// Effectif INSEE
const EFFECTIF_LABELS = {
  "NN": "Non employeur",
  "00": "0 salarié",
  "01": "1 ou 2 salariés",
  "02": "3 à 5 salariés",
  "03": "6 à 9 salariés",
  "11": "10 à 19 salariés",
  "12": "20 à 49 salariés",
  "21": "50 à 99 salariés",
  "22": "100 à 199 salariés",
  "31": "200 à 249 salariés",
  "32": "250 à 499 salariés",
  "41": "500 à 999 salariés",
  "42": "1 000 à 1 999 salariés",
  "51": "2 000 à 4 999 salariés",
  "52": "5 000 à 9 999 salariés",
  "53": "10 000 salariés et plus",
};

// =====================================================
// POINT D'ENTRÉE — Module Worker (env injection)
// =====================================================
// Cloudflare Workers modules syntax : les secrets sont dans `env`,
// pas dans globalThis. On les injecte une fois au démarrage.

export default {
  async fetch(req, env, ctx) {
    // Injection des secrets dans globalThis pour compatibilité du code existant
    if (env) Object.assign(globalThis, env);
    return handle(req, { waitUntil: ctx.waitUntil.bind(ctx) });
  },
  async scheduled(event, env, ctx) {
    if (env) Object.assign(globalThis, env);
    // Point d'entrée cron si besoin futur
  },
};


// =====================================================
// ROUTER
// =====================================================

async function handle(req, event = null) {
  const url = new URL(req.url);

  if (url.pathname === "/health") {
    // Test CSE — essayer GOOGLE_CSE_KEY puis GEMINI_API_KEY
    let cseStatus = "not_configured";
    const cseId = String(globalThis.GOOGLE_CSE_ID || "").trim();
    const cseKeysToTest = [
      ["GOOGLE_CSE_KEY",  String(globalThis.GOOGLE_CSE_KEY || "").trim()],
      ["GEMINI_API_KEY",  String(globalThis.GEMINI_API_KEY || "").trim()],
      ["PLACES_API_KEY",  String(globalThis.GOOGLE_PLACES_API_KEY || "").trim()],
    ].filter(([,v]) => v);
    if (cseId && cseKeysToTest.length) {
      for (const [name, cseKey] of cseKeysToTest) {
        try {
          const r = await fetch(
            `https://www.googleapis.com/customsearch/v1?key=${encodeURIComponent(cseKey)}&cx=${encodeURIComponent(cseId)}&q=test&num=1`,
            { signal: AbortSignal.timeout(5000) }
          );
          const j = await r.json();
          if (r.ok) { cseStatus = `ok (${name})`; break; }
          cseStatus = `${name}: error_${r.status}: ${j?.error?.message?.slice(0,60) || ""}`;
        } catch (e) {
          cseStatus = `${name}: exception: ${e?.message?.slice(0,40)}`;
        }
      }
    }

    // Test Places
    let placesStatus = "not_configured";
    const placesKey = String(globalThis.GOOGLE_PLACES_API_KEY || "").trim();
    if (placesKey) {
      try {
        const r = await fetch(
          "https://places.googleapis.com/v1/places:searchText",
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              "X-Goog-Api-Key": placesKey,
              "X-Goog-FieldMask": "places.displayName",
            },
            body: JSON.stringify({ textQuery: "Paris", maxResultCount: 1 }),
            signal: AbortSignal.timeout(5000),
          }
        );
        const j = await r.json();
        if (r.ok && j.places?.length) placesStatus = "ok";
        else placesStatus = `error_${r.status}: ${j?.error?.message?.slice(0,80) || ""}`;
      } catch (e) {
        placesStatus = `exception: ${e?.message?.slice(0,50)}`;
      }
    }

    return json({
      ok: true,
      service: "prospection-worker",
      version: "8.4",
      ts: new Date().toISOString(),
      config: {
        gemini:   !!String(globalThis.GEMINI_API_KEY        || "").trim(),
        vision:   !!String(globalThis.GOOGLE_VISION_API_KEY || "").trim(),
        places:   placesStatus,
        telegram: !!String(globalThis.TELEGRAM_TOKEN        || "").trim(),
        brevo:    !!String(globalThis.BREVO_API_KEY         || "").trim(),
        github:   !!String(globalThis.GITHUB_TOKEN          || "").trim(),
        kv:       !!globalThis.PROSPECTION_KV,
        cse:      cseStatus,
        brave:    !!String(globalThis.BRAVE_SEARCH_KEY || "").trim(),
      },
    });
  }

  if (url.pathname === "/status") {
    const tok = url.searchParams.get("token") || req.headers.get("X-Export-Token") || "";
    const configured = String(globalThis.EXPORT_TOKEN || "").trim();
    if (!configured || tok !== configured) return new Response("Unauthorized", { status: 401 });

    const date = url.searchParams.get("date") || ymdFR();
    const agencies = ["GR", "VR", "GRS", "SLS"];

    // Récupérer toutes les clés du jour en parallèle
    const keyArrays = await Promise.all(
      agencies.map(ag => listAllKeys(`prospect:${date}:${ag}:`))
    );

    // Récupérer toutes les fiches en parallèle
    const recordArrays = await Promise.all(
      keyArrays.map(keys => batchGet(keys))
    );

    // Récupérer les clôtures
    const closeKeys = await listAllKeys(`close:${date}:`);
    const closedAgencies = new Set(closeKeys.map(k => {
      const parts = (typeof k === "string" ? k : k.name).split(":");
      return parts[2] || "";
    }));

    let totalFiches = 0;
    let lastActivity = "";
    const parAgence = {};

    for (let i = 0; i < agencies.length; i++) {
      const ag = agencies[i];
      const records = recordArrays[i];
      const keys = keyArrays[i];

      const commerciaux = new Set();
      for (const r of records) {
        if (!r) continue;
        if (r.initials) commerciaux.add(r.initials);
        if (r.created_at && r.created_at > lastActivity) lastActivity = r.created_at;
      }

      const count = records.filter(Boolean).length;
      totalFiches += count;

      parAgence[ag] = {
        fiches:      count,
        commerciaux: [...commerciaux],
        cloture:     closedAgencies.has(ag),
      };
    }

    return json({
      ok:              true,
      date,
      total:           totalFiches,
      par_agence:      parAgence,
      derniere_activite: lastActivity || null,
      generated_at:    new Date().toISOString(),
    });
  }

  if (url.pathname === "/test-brave") {
    const q = url.searchParams.get("q") || "Intermarché Varces SIRET société";
    const key = String(globalThis.BRAVE_SEARCH_KEY || "").trim();
    if (!key) return json({ error: "BRAVE_SEARCH_KEY manquant" });
    try {
      const r = await fetch(
        "https://api.search.brave.com/res/v1/web/search?" + new URLSearchParams({ q, count: "5", country: "fr", search_lang: "fr" }),
        { headers: { "Accept": "application/json", "X-Subscription-Token": key }, signal: AbortSignal.timeout(8000) }
      );
      const j = await r.json();
      const items = (j?.web?.results || []).slice(0,5).map(i => ({
        title: i.title,
        url: i.url,
        description: (i.description||"").slice(0,120),
      }));
      const result = await searchViaBrave(q);
      return json({ q, result, items });
    } catch(e) { return json({ error: e.message }); }
  }

  if (url.pathname === "/test-scrape") {
    const testUrl = url.searchParams.get("url") || "";
    const search  = url.searchParams.get("search") || "";
    if (!testUrl) return json({ error: "param url manquant" });
    try {
      const r = await fetch(testUrl, {
        headers: {
          "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
          "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
          "Accept-Language": "fr-FR,fr;q=0.9",
        },
        signal: AbortSignal.timeout(8000),
      });
      const html = await r.text();
      const siretRe = /\b(\d{3}[\s.\-]?\d{3}[\s.\-]?\d{3}[\s.\-]?\d{5})\b/g;
      const sirenRe = /\b(\d{3}[\s.\-]?\d{3}[\s.\-]?\d{3})\b/g;
      const sirets = [...html.matchAll(siretRe)].map(m => m[1].replace(/[\s.\-]/g,"")).filter(s => s.length === 14 && !s.startsWith("2026") && !s.startsWith("2025"));
      const sirens = [...html.matchAll(sirenRe)].map(m => m[1].replace(/[\s.\-]/g,"")).filter(s => s.length === 9);
      const text = html.replace(/<[^>]+>/g," ").replace(/\s+/g," ");
      
      // Chercher terme spécifique dans le HTML brut
      let context = "";
      if (search) {
        const idx = html.toLowerCase().indexOf(search.toLowerCase());
        if (idx >= 0) context = html.slice(Math.max(0, idx-100), idx+300);
      }

      // Chercher "siret" ou "siren" dans le HTML brut
      const siretCtx = [];
      const siretIdx = html.toLowerCase().indexOf("siret");
      if (siretIdx >= 0) siretCtx.push(html.slice(Math.max(0,siretIdx-20), siretIdx+100));
      const sirenIdx = html.toLowerCase().indexOf("siren");
      if (sirenIdx >= 0) siretCtx.push(html.slice(Math.max(0,sirenIdx-20), sirenIdx+100));

      return json({ 
        status: r.status, 
        sirets: [...new Set(sirets)].slice(0,5), 
        sirens: [...new Set(sirens)].slice(0,10),
        siretContexts: siretCtx.slice(0,3),
        searchContext: context,
        preview: text.slice(0,1000),
      });
    } catch(e) {
      return json({ error: e.message });
    }
  }

  if (url.pathname === "/test-gouv") {
    // Test direct de l'API Gouv pour debug
    const q = url.searchParams.get("q") || "4S Menuiseries 38430";
    const results = [];
    try {
      const r = await safeFetchJson(
        "https://recherche-entreprises.api.gouv.fr/search?q=" + encodeURIComponent(q) + "&limite=5",
        {}, { timeoutMs: 8000, retries: 0, label: "TEST_GOUV" }
      );
      const raw = r.data?.results || [];
      for (const res of raw.slice(0, 5)) {
        const siege = res.siege || {};
        results.push({
          nom: res.nom_raison_sociale,
          nom_complet: res.nom_complet,
          siren: res.siren,
          siret: siege.siret,
          adresse: siege.adresse,
          cp: siege.code_postal,
          ville: siege.libelle_commune,
          naf: res.activite_principale,
          enseigne1: res.enseigne1 || "",
          enseigne2: res.enseigne2 || "",
          nom_commercial: res.nom_commercial || "",
        });
      }
    } catch (e) {
      return json({ error: e.message });
    }
    return json({ q, count: results.length, results });
  }

  if (url.pathname === "/dump") {
    return handleDump(req, url);
  }

  if (url.pathname === "/remind" && req.method === "POST") {
    return handleRemind(req);
  }

  if (url.pathname === "/update" && req.method === "POST") {
    return handleOCRUpdate(req);
  }

  if (url.pathname === "/telegram" && req.method === "POST") {
    const secretCheck = await checkTelegramSecretAsync(req);
    if (!secretCheck.ok) {
      log("warn", null, "WEBHOOK_SECRET_REJECTED", { reason: secretCheck.reason });
      return new Response("Unauthorized", { status: 401 });
    }

    let update;
    try {
      update = await req.json();
    } catch (_) {
      return json({ ok: false, error: "invalid JSON" }, 400);
    }

    try {
      await handleTelegram(update, event);
    } catch (e) {
      log("error", null, "TG_UNHANDLED", {
        err: e?.message || String(e),
        stack: (e?.stack || "").slice(0, 500),
        name: e?.name || "",
      });
    }
    return json({ ok: true });
  }

  return new Response("Not found", { status: 404 });
}


// =====================================================
// BASICS
// =====================================================

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: JSON_HEADERS });
}

function kv() {
  const store = globalThis.PROSPECTION_KV;
  if (!store) {
    // Log critique — le binding KV n'est pas accessible
    // En modules syntax, il vient de env via Object.assign
    // Vérifier que PROSPECTION_KV est bien déclaré dans wrangler.toml [[kv_namespaces]]
    console.error("[KV_MISSING] PROSPECTION_KV binding not found in globalThis");
  }
  return store;
}

function ymdFR() {
  const parts = new Intl.DateTimeFormat("fr-FR", {
    timeZone: "Europe/Paris",
    year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(new Date());
  const y = parts.find((p) => p.type === "year")?.value;
  const m = parts.find((p) => p.type === "month")?.value;
  const d = parts.find((p) => p.type === "day")?.value;
  return `${y}-${m}-${d}`;
}

function ymdFROffset(offsetDays) {
  const d = new Date();
  d.setDate(d.getDate() - offsetDays);
  const parts = new Intl.DateTimeFormat("fr-FR", {
    timeZone: "Europe/Paris",
    year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(d);
  const y = parts.find((p) => p.type === "year")?.value;
  const mo = parts.find((p) => p.type === "month")?.value;
  const da = parts.find((p) => p.type === "day")?.value;
  return `${y}-${mo}-${da}`;
}

function escapeHtml(s) {
  return String(s || "").replace(/[<>&]/g, (m) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" }[m]));
}

function trunc(s, max = MAX_FIELD_LEN) {
  return String(s || "").slice(0, max);
}

function parseComment(text) {
  const t = String(text || "").trim();
  return t === "." || t === "" ? null : trunc(t, 300);
}

function log(level, chatId, event, extra = {}) {
  console[level === "error" ? "error" : "log"](
    JSON.stringify({ level, chatId: chatId ? String(chatId) : null, event, ts: Date.now(), ...extra })
  );
}

function isValidFileId(id) {
  return typeof id === "string" && id.length >= 10 && id.length <= 300;
}

function normalizeSpaces(s) {
  return String(s || "").replace(/\s+/g, " ").trim();
}

function normalizeNAF(naf) {
  return String(naf || "").trim().toUpperCase().replaceAll(".", "").replaceAll(" ", "");
}

function normalizePhone(input) {
  let s = String(input || "").trim();
  if (!s) return "";
  s = s.replace(/[^\d+]/g, "");
  if (s.startsWith("+33")) s = "0" + s.slice(3);
  if (!/^0\d{9}$/.test(s)) return "";
  return s;
}

function formatPhone(phone) {
  // "0476401234" → "04 76 40 12 34"
  const s = String(phone || "").replace(/\D/g, "");
  if (s.length !== 10) return phone;
  return s.match(/.{2}/g).join(" ");
}

function validateSIRET(siret) {
  if (!siret) return "";
  const s = String(siret).replace(/\D/g, "");
  if (s.length !== 14) return "";
  let sum = 0;
  for (let i = 0; i < 14; i++) {
    let digit = parseInt(s[13 - i], 10);
    if (i % 2 === 1) digit *= 2;
    if (digit > 9) digit -= 9;
    sum += digit;
  }
  if (sum % 10 === 0) return s;
  log("warn", null, "SIRET_INVALID", { siret: String(siret) });
  return "";
}

function isValidEmail(email) {
  const e = String(email || "").trim().toLowerCase();
  if (!e || e.length > 254) return false;
  const re = /^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$/;
  if (!re.test(e)) return false;
  const disposable = new Set(["tempmail.com", "guerrillamail.com", "yopmail.com", "10minutemail.com"]);
  const domain = (e.split("@")[1] || "").toLowerCase();
  return !disposable.has(domain);
}

function normalizeEmail(email) {
  const e = String(email || "").trim().toLowerCase();
  return isValidEmail(e) ? e : "";
}

function splitHumanName(full) {
  const s = normalizeSpaces(full);
  if (!s) return { first: "", last: "" };
  const parts = s.split(" ");
  if (parts.length === 1) return { first: "", last: parts[0] };
  return { first: parts[0], last: parts.slice(1).join(" ") };
}


// Nettoyer le format "NOM (NOM)" ou "Prénom NOM (NOM)" → "Prénom NOM"
function cleanPersonName(name) {
  if (!name) return "";
  // Supprimer les parties entre parenthèses ex: "BOUJEAT (BOUJEAT)" → "BOUJEAT"
  return normalizeSpaces(name.replace(/\s*\([^)]*\)\s*/g, "").trim());
}

function ensureContactFallback(d) {
  const hasContact =
    (d.contact_firstname && String(d.contact_firstname).trim()) ||
    (d.contact_lastname && String(d.contact_lastname).trim());
  const dirigeant = normalizeSpaces(d.dirigeant || "");
  if (!hasContact && dirigeant) {
    const sp = splitHumanName(dirigeant);
    d.contact_firstname = d.contact_firstname || sp.first;
    d.contact_lastname = d.contact_lastname || sp.last || dirigeant;
  }
  if (!d.interlocuteur) {
    const combo = `${(d.contact_firstname || "").trim()} ${(d.contact_lastname || "").trim()}`.trim();
    if (combo) d.interlocuteur = combo;
    else if (dirigeant) d.interlocuteur = dirigeant;
  }
  return d;
}

function normalizeForMatch(str) {
  return String(str || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\b(sas|sasu|sarl|eurl|sa|sci|snc|societe|groupe|compagnie|cie)\b/g, " ")
    .replace(/[^a-z0-9]/g, "")
    .trim();
}

// ── Sécurité webhook ──────────────────────────────
async function checkTelegramSecretAsync(req) {
  const expected = String(globalThis.TELEGRAM_WEBHOOK_SECRET || "").trim();
  if (!expected) {
    log("error", null, "WEBHOOK_SECRET_MISSING", {});
    return { ok: false, reason: "secret_not_configured" };
  }
  const got = req.headers.get("X-Telegram-Bot-Api-Secret-Token") || "";
  if (!got) return { ok: false, reason: "secret_missing" };
  const equal = await timingSafeEqual(got, expected);
  if (!equal) return { ok: false, reason: "secret_mismatch" };
  return { ok: true };
}

// Alias synchrone conservé pour compatibilité interne
function checkTelegramSecret(req) {
  const expected = String(globalThis.TELEGRAM_WEBHOOK_SECRET || "").trim();
  if (!expected) return { ok: false, reason: "secret_not_configured" };
  const got = req.headers.get("X-Telegram-Bot-Api-Secret-Token");
  if (!got || got !== expected) return { ok: false, reason: "secret_mismatch" };
  return { ok: true };
}

// Version async timing-safe (pour /dump et /remind via X-Export-Token)
async function checkExportToken(req) {
  const configured = String(globalThis.EXPORT_TOKEN || "").trim();
  if (!configured) return false;
  const got = String(req.headers.get("X-Export-Token") || "").trim();
  if (!got) return false;
  return timingSafeEqual(got, configured);
}

// ── SSRF protection ───────────────────────────────
function isSafeUrl(url) {
  try {
    const u = new URL(url);
    if (u.protocol !== "https:") return false;
    const h = u.hostname.toLowerCase();
    if (h === "localhost") return false;
    if (/^127\./.test(h)) return false;
    if (/^10\./.test(h)) return false;
    if (/^192\.168\./.test(h)) return false;
    if (/^172\.(1[6-9]|2\d|3[01])\./.test(h)) return false;
    if (/^169\.254\./.test(h)) return false;
    if (/^::1$/.test(h)) return false;
    if (/^fd/.test(h)) return false;
    return true;
  } catch (_) {
    return false;
  }
}

function getFreshGeoOrNull(s, maxAgeMs = GEO_MAX_AGE_MS) {
  const g = s?.current_geo;
  if (!g || !g.at) return null;
  if ((Date.now() - g.at) > maxAgeMs) return null;
  return { lat: g.lat, lon: g.lon, at: g.at };
}

function escapeRegExp(s) {
  return String(s || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function stripPostalCityFromAddress(addr, postalCode, city) {
  let a = normalizeSpaces(addr);
  if (!a) return "";
  const cp = normalizeSpaces(postalCode);
  const ct = normalizeSpaces(city);
  if (cp && ct) {
    const tail = `${cp} ${ct}`.trim();
    a = a.replace(new RegExp(`(?:\\s|-)?${escapeRegExp(tail)}\\s*$`, "i"), "").trim();
  }
  if (cp) {
    a = a.replace(new RegExp(`(?:\\s|-)?${escapeRegExp(cp)}\\s*$`, "i"), "").trim();
  }
  return normalizeSpaces(a);
}

function nafLabel(naf) {
  return NAF_LABELS[String(naf || "").toUpperCase()] || "";
}

function effectifLabel(code) {
  return EFFECTIF_LABELS[String(code || "")] || String(code || "");
}

function parsePositiveInt(text) {
  const n = parseInt(String(text || "").replace(/\D/g, ""), 10);
  return Number.isFinite(n) && n >= 0 ? n : null;
}


// =====================================================
// TIMING-SAFE COMPARE (anti timing-attack) — v10
// =====================================================

async function timingSafeEqual(a, b) {
  const enc = new TextEncoder();
  const [ha, hb] = await Promise.all([
    crypto.subtle.digest("SHA-256", enc.encode(String(a || ""))),
    crypto.subtle.digest("SHA-256", enc.encode(String(b || ""))),
  ]);
  const x = new Uint8Array(ha);
  const y = new Uint8Array(hb);
  if (x.length !== y.length) return false;
  let diff = 0;
  for (let i = 0; i < x.length; i++) diff |= x[i] ^ y[i];
  return diff === 0;
}


// =====================================================
// SLEEP / RETRY HELPERS — v10
// =====================================================

function sleep(ms) {
  const capped = Math.min(Number(ms) || 0, MAX_RETRY_SLEEP_MS);
  return new Promise((r) => setTimeout(r, capped));
}

async function safeFetchJson(url, options = {}, cfg = {}) {
  const timeoutMs   = cfg.timeoutMs   || DEFAULT_FETCH_TIMEOUT_MS;
  const retries     = Number.isFinite(cfg.retries) ? cfg.retries : API_RETRIES;
  const retryDelay  = cfg.retryDelayMs || 400;
  const label       = cfg.label || "SAFE_FETCH_JSON";
  const acceptCodes = cfg.acceptStatuses || [200];

  let lastErr = null;
  for (let i = 0; i <= retries; i++) {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      let res;
      try {
        res = await fetch(url, { ...options, signal: controller.signal });
      } finally {
        clearTimeout(timer);
      }
      if (!acceptCodes.includes(res.status)) {
        const txt = await res.text().catch(() => "");
        throw new Error(`${label} bad status ${res.status}: ${txt.slice(0, 200)}`);
      }
      return { ok: true, status: res.status, data: await res.json() };
    } catch (e) {
      lastErr = e;
      if (i < retries) await sleep(retryDelay * (i + 1));
    }
  }
  return { ok: false, status: 0, error: lastErr?.message || String(lastErr), data: null };
}

async function safeFetchText(url, options = {}, cfg = {}) {
  const timeoutMs   = cfg.timeoutMs   || DEFAULT_FETCH_TIMEOUT_MS;
  const retries     = Number.isFinite(cfg.retries) ? cfg.retries : API_RETRIES;
  const retryDelay  = cfg.retryDelayMs || 400;
  const label       = cfg.label || "SAFE_FETCH_TEXT";
  const acceptCodes = cfg.acceptStatuses || [200];

  let lastErr = null;
  for (let i = 0; i <= retries; i++) {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      let res;
      try {
        res = await fetch(url, { ...options, signal: controller.signal });
      } finally {
        clearTimeout(timer);
      }
      const text = await res.text().catch(() => "");
      if (!acceptCodes.includes(res.status)) {
        throw new Error(`${label} bad status ${res.status}: ${text.slice(0, 200)}`);
      }
      return { ok: true, status: res.status, text };
    } catch (e) {
      lastErr = e;
      if (i < retries) await sleep(retryDelay * (i + 1));
    }
  }
  return { ok: false, status: 0, error: lastErr?.message || String(lastErr), text: "" };
}


// =====================================================
// BACKGROUND TASKS (waitUntil) — v10
// =====================================================

function scheduleBackgroundTask(event, promise, label = "BG_TASK") {
  const task = Promise.resolve(promise).catch((e) => {
    log("warn", null, label, { err: e?.message || String(e) });
  });
  if (event && typeof event.waitUntil === "function") {
    event.waitUntil(task);
    return true;
  }
  task.catch(() => {});
  return false;
}


// =====================================================
// KV INDEX HELPERS — v10
// (remplace le list-scan O(n) par un index O(1))
// =====================================================

async function readIndexArray(key) {
  const raw = await kv().get(key, { type: "json" });
  return Array.isArray(raw) ? raw.filter(Boolean) : [];
}

async function writeIndexArray(key, values, ttl = TTL_INDEX) {
  const arr = [...new Set(values.map(String).filter(Boolean))];
  await kv().put(key, JSON.stringify(arr), { expirationTtl: ttl });
  return arr;
}

async function appendIndexKey(indexKey, value, ttl = TTL_INDEX) {
  if (!indexKey || !value) return readIndexArray(indexKey);
  for (let attempt = 0; attempt < 3; attempt++) {
    const arr = await readIndexArray(indexKey);
    if (arr.includes(String(value))) return arr;
    const next = [...arr, String(value)];
    await writeIndexArray(indexKey, next, ttl);
    const verify = await readIndexArray(indexKey);
    if (verify.includes(String(value))) return verify;
    if (attempt < 2) await sleep(50 * Math.pow(2, attempt) + Math.floor(Math.random() * 50));
  }
  return readIndexArray(indexKey);
}

async function addDayIndex(date, agency, recordKey) {
  return appendIndexKey(`idx:day:${date}:${agency}`, recordKey, TTL_INDEX);
}

async function addDupIndex(date, agency, fingerprint) {
  return appendIndexKey(`idx:dup:${date}:${agency}`, fingerprint, TTL_INDEX);
}

async function getDayIndexedKeys(date, agency) {
  return readIndexArray(`idx:day:${date}:${agency}`);
}

async function getDupFingerprints(date, agency) {
  return readIndexArray(`idx:dup:${date}:${agency}`);
}


// =====================================================
// GEMINI FAILURE BLACKLIST
// =====================================================

const GEMINI_FAILURE_STRINGS = [
  "je ne peux pas",
  "impossible",
  "illisible",
  "inconnu",
  "n/a",
  "not available",
  "cannot",
  "unable",
  "sorry",
  "désolé",
  "je suis désolé",
  "aucun texte",
  "pas de texte",
  "image vide",
];

function isGeminiFailureValue(val) {
  if (!val) return false;
  const v = String(val).toLowerCase().trim();
  if (!v || v.length > 120) return false;
  return GEMINI_FAILURE_STRINGS.some((f) => v.includes(f));
}

function sanitizeGeminiOutput(d) {
  if (!d || typeof d !== "object") return {};
  const out = {};
  for (const [k, v] of Object.entries(d)) {
    out[k] = isGeminiFailureValue(v) ? "" : v;
  }
  return out;
}


// =====================================================
// CRYPTO / HASH
// =====================================================

async function sha1HexFromString(str) {
  const buf = await crypto.subtle.digest("SHA-1", new TextEncoder().encode(String(str || "").toLowerCase()));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function sha1HexFromBytes(bytes) {
  const buf = await crypto.subtle.digest("SHA-1", bytes);
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}


// =====================================================
// KV LIST HELPERS + PARALLÉLISATION
// =====================================================

async function listAllKeys(prefix) {
  let cursor = undefined;
  const keys = [];
  while (true) {
    const res = await kv().list({ prefix, cursor });
    keys.push(...(res.keys || []));
    if (res.list_complete) break;
    cursor = res.cursor;
  }
  return keys;
}

// Lecture parallèle de N clés KV
async function batchGet(keys, type = "json") {
  return Promise.all(
    keys.map((k) =>
      kv().get(typeof k === "string" ? k : k.name, { type }).catch(() => null)
    )
  );
}


// =====================================================
// CACHE HELPERS
// =====================================================

async function withCache(cacheKey, ttl, fetcher) {
  const fullKey = `cache:${cacheKey}`;
  const cached = await kv().get(fullKey, { type: "json" });
  if (cached !== null) {
    // Ne pas utiliser le cache si le résultat était vide (erreur précédente)
    const isEmpty = !cached || (typeof cached === "object" && Object.keys(cached).length === 0);
    if (!isEmpty) {
      log("info", null, "CACHE_HIT", { key: fullKey });
      return { value: cached, fromCache: true };
    }
    // Cache vide → on relance Gemini
    log("info", null, "CACHE_EMPTY_BYPASS", { key: fullKey });
  }
  const data = await fetcher();
  // Mettre en cache uniquement si le résultat contient au moins un champ utile
  const hasContent = data && typeof data === "object" && Object.keys(data).some(k =>
    !["confidence", "_confidence"].includes(k) && data[k] && String(data[k]).trim()
  );
  if (hasContent) {
    await kv().put(fullKey, JSON.stringify(data), { expirationTtl: ttl });
  }
  return { value: data, fromCache: false };
}


// =====================================================
// USER PREFS
// =====================================================

async function getPrefs(chatId) {
  return kv().get(`user_prefs:${chatId}`, { type: "json" });
}

async function savePrefs(chatId, agency, initials) {
  await kv().put(
    `user_prefs:${chatId}`,
    JSON.stringify({ agency, initials }),
    { expirationTtl: TTL_PREFS }
  );
}


// =====================================================
// SESSION / VISITS
// =====================================================

function emptyDraft() {
  return {
    _source: "",
    _edit_key: "",
    _visit_id: "",
    _created_from: "",

    name: "",
    nom_commercial: "",   // Nom enseigne/commercial (ex: "Intermarché SUPER Varces")
    address: "",
    postal_code: "",
    city: "",
    gps_lat: null,
    gps_lon: null,
    siret: "",
    naf: "",
    secteur_activite: "",
    effectif: "",
    date_creation: "",
    capital: "",
    dirigeant: "",

    interlocuteur: "",
    contact_civility: "",
    contact_firstname: "",
    contact_lastname: "",
    titre: "",           // fonction / poste (extrait OCR carte)

    phone: "",
    phone2: "",
    email: "",
    email_card: "",    // Email personnel de la carte de visite
    website: "",

    resume: "",
    commande: "",
    qualification: "",

    card_photo_file_id: "",
    card_photo_url: "",
    card_added_at: null,

    facade_photo_file_id: "",
    facade_photo_url: "",
    facade_added_at: null,

    card_kv_key: "",
    linked_card_keys: [],
    linked_photo_keys: [],

    source_mode: "",
    _ocr_confidence: "",
  };
}

function sanitizeVisitNo(v) {
  const n = Number(v || 0);
  if (!Number.isFinite(n) || n < 1) return 0;
  return Math.floor(n);
}

function sanitizeProspectDraft(d) {
  const out = { ...(d || {}) };

  out.name = normalizeSpaces(out.name || "");
  out.address = stripPostalCityFromAddress(out.address || "", out.postal_code || "", out.city || "");
  out.postal_code = normalizeSpaces(out.postal_code || "").replace(/\s+/g, "");
  out.city = normalizeSpaces(out.city || "");
  out.siret = validateSIRET(String(out.siret || "").replace(/\D/g, ""));
  out.naf = normalizeNAF(out.naf || "");
  out.secteur_activite = normalizeSpaces(out.secteur_activite || "") || nafLabel(out.naf);
  out.effectif = normalizeSpaces(out.effectif || "");
  out.date_creation = normalizeSpaces(out.date_creation || "");
  out.capital = normalizeSpaces(out.capital || "");
  out.dirigeant = normalizeSpaces(out.dirigeant || "");
  out.interlocuteur = normalizeSpaces(out.interlocuteur || "");
  out.contact_civility = normalizeSpaces(out.contact_civility || "");
  out.contact_firstname = normalizeSpaces(out.contact_firstname || "");
  out.contact_lastname = normalizeSpaces(out.contact_lastname || "");
  out.phone = normalizePhone(out.phone || "");
  out.phone2 = normalizePhone(out.phone2 || "");
  if (out.phone && out.phone2 && out.phone === out.phone2) out.phone2 = "";
  out.email = normalizeEmail(out.email || "");
  out.website = normalizeSpaces(out.website || "");
  out.resume = trunc(out.resume || "");
  out.commande = trunc(out.commande || "");
  out.qualification = [QUALIF_HOT, QUALIF_WARM, QUALIF_COLD].includes(out.qualification)
    ? out.qualification
    : "";

  out.card_photo_file_id = String(out.card_photo_file_id || "");
  out.card_photo_url = String(out.card_photo_url || "");
  out.facade_photo_file_id = String(out.facade_photo_file_id || "");
  out.facade_photo_url = String(out.facade_photo_url || "");
  out.card_kv_key = String(out.card_kv_key || "");

  out.linked_card_keys = Array.isArray(out.linked_card_keys) ? out.linked_card_keys : [];
  out.linked_photo_keys = Array.isArray(out.linked_photo_keys) ? out.linked_photo_keys : [];

  return ensureContactFallback(out);
}

function normalizeDraft(d) {
  return sanitizeProspectDraft({ ...emptyDraft(), ...(d || {}) });
}

function initSession(agency = "", initials = "", prev = null) {
  const p = prev || {};
  return {
    step: p.step || (agency ? "MENU" : "AGENCY"),
    agency: p.agency || agency || "",
    initials: p.initials || initials || "",
    mode: p.mode || "company",
    draft: normalizeDraft(p.draft || {}),
    companyChoices: Array.isArray(p.companyChoices) ? p.companyChoices : [],
    last_search_query: p.last_search_query || "",
    form_message_id: p.form_message_id || null,
    awaitingField: p.awaitingField || null,
    close_stats: p.close_stats || null,
    photo_batch: Array.isArray(p.photo_batch) ? p.photo_batch : [],
    photo_ts0: p.photo_ts0 || null,
    photo_city: p.photo_city || "",
    _last_photo_city: p._last_photo_city || "",
    current_geo: p.current_geo || null,
    card_batch: Array.isArray(p.card_batch) ? p.card_batch : [],
    card_ts0: p.card_ts0 || null,
    card_batch_prospect: Array.isArray(p.card_batch_prospect) ? p.card_batch_prospect : [],
    _manual_city: p._manual_city || "",
    _pending_duplicate_record: p._pending_duplicate_record || null,
    active_visit_id: p.active_visit_id || "",
    active_visit_no: sanitizeVisitNo(p.active_visit_no || 0),
    visits_index: Array.isArray(p.visits_index) ? p.visits_index : [],
    visit_seq: sanitizeVisitNo(p.visit_seq || 0),
    pending_delete_visit_id: p.pending_delete_visit_id || "",
    pending_open_visit_id: p.pending_open_visit_id || "",
    // Qualification en attente avant save
    _pending_qualif_save: p._pending_qualif_save || false,
  };
}

function resetPhotoState(s) {
  s.photo_batch = [];
  s.photo_ts0 = null;
  s.photo_city = "";
  return s;
}

function resetCardState(s) {
  s.card_batch = [];
  s.card_ts0 = null;
  s.card_batch_prospect = [];
  return s;
}

async function getSession(chatId) {
  const raw = await kv().get(`session:${chatId}`, { type: "json" });
  if (!raw) return null;
  return initSession(raw.agency || "", raw.initials || "", raw);
}

async function setSession(chatId, data) {
  const s = initSession((data && data.agency) || "", (data && data.initials) || "", data || {});
  await kv().put(`session:${chatId}`, JSON.stringify(s), { expirationTtl: TTL_SESSION });
}

async function resetSession(chatId) {
  await kv().delete(`session:${chatId}`);
}

function makeVisitId(chatId) {
  return `${chatId}:${Date.now()}:${Math.floor(Math.random() * 1000)}`;
}

function buildVisitStoreKey(chatId, visitId) {
  return `session_visit:${chatId}:${visitId}`;
}

function currentVisitLabel(s) {
  const n = sanitizeVisitNo((s && s.active_visit_no) || 0);
  return n ? String(n) : "—";
}

function compactVisitTitle(draft) {
  const d = normalizeDraft(draft || {});
  return `${d.name || "SOCIÉTÉ"} — ${d.city || "VILLE"}`;
}

function getVisitIndexEntry(s, visitId) {
  const arr = Array.isArray((s && s.visits_index) || null) ? s.visits_index : [];
  return arr.find((v) => String((v && v.visit_id) || "") === String(visitId || ""));
}

function getVisitNoFromId(s, visitId) {
  const it = getVisitIndexEntry(s, visitId);
  return sanitizeVisitNo((it && it.no) || 0);
}

function getVisitIdFromNo(s, no) {
  const target = sanitizeVisitNo(no);
  const arr = Array.isArray((s && s.visits_index) || null) ? s.visits_index : [];
  const it = arr.find((v) => sanitizeVisitNo((v && v.no) || 0) === target);
  return (it && it.visit_id) || "";
}

function upsertVisitIndexEntry(s, visitId, draft) {
  const title = compactVisitTitle(draft);
  const arr = Array.isArray(s.visits_index) ? [...s.visits_index] : [];
  const idx = arr.findIndex((v) => String((v && v.visit_id) || "") === String(visitId || ""));
  if (idx >= 0) {
    arr[idx] = { ...arr[idx], title, updated_at: Date.now() };
  } else {
    const nextNo = sanitizeVisitNo((s.visit_seq || 0) + 1);
    s.visit_seq = nextNo;
    arr.push({
      visit_id: visitId,
      no: nextNo,
      title,
      created_at: Date.now(),
      updated_at: Date.now(),
    });
    s.active_visit_no = nextNo;
  }
  s.visits_index = arr;
  return s;
}

async function getStoredVisit(chatId, visitId) {
  if (!visitId) return null;
  return kv().get(buildVisitStoreKey(chatId, visitId), { type: "json" });
}

async function setStoredVisit(chatId, visitId, payload) {
  await kv().put(buildVisitStoreKey(chatId, visitId), JSON.stringify(payload), {
    expirationTtl: TTL_SESSION,
  });
}

async function deleteStoredVisit(chatId, visitId) {
  await kv().delete(buildVisitStoreKey(chatId, visitId));
}

async function persistActiveVisit(chatId, s) {
  if (!s || !s.active_visit_id) return;
  const draft = normalizeDraft(s.draft || {});
  s.draft = draft;
  upsertVisitIndexEntry(s, s.active_visit_id, draft);
  const meta = getVisitIndexEntry(s, s.active_visit_id);
  await setStoredVisit(chatId, s.active_visit_id, {
    visit_id: s.active_visit_id,
    no: sanitizeVisitNo(((meta && meta.no) || s.active_visit_no || 0)),
    agency: s.agency || "",
    initials: s.initials || "",
    updated_at: Date.now(),
    title: compactVisitTitle(draft),
    draft,
  });
  s.active_visit_no = sanitizeVisitNo(((meta && meta.no) || s.active_visit_no || 0));
}

async function createNewActiveVisit(chatId, s, sourceMode = "company") {
  const visitId = makeVisitId(chatId);
  const d = normalizeDraft({
    ...emptyDraft(),
    _visit_id: visitId,
    _created_from: sourceMode,
    source_mode: sourceMode,
  });
  s.active_visit_id = visitId;
  s.active_visit_no = 0;
  s.draft = d;
  s.awaitingField = null;
  s.pending_delete_visit_id = "";
  s.pending_open_visit_id = "";
  s.step = "VISIT_ACTIVE";
  s.mode = sourceMode === "photo" ? "photo" : sourceMode === "card" ? "card" : "company";
  resetPhotoState(s);
  resetCardState(s);
  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);
  return visitId;
}

async function openVisitById(chatId, s, visitId) {
  const rec = await getStoredVisit(chatId, visitId);
  if (!rec || !rec.draft) return false;
  s.active_visit_id = visitId;
  s.active_visit_no = sanitizeVisitNo(rec.no || getVisitNoFromId(s, visitId));
  s.draft = normalizeDraft(rec.draft || {});
  s.awaitingField = null;
  s.pending_delete_visit_id = "";
  s.pending_open_visit_id = "";
  s.step = "VISIT_ACTIVE";
  // Réinitialiser form_message_id pour forcer l'affichage en bas du fil
  s.form_message_id = null;
  await setSession(chatId, s);
  return true;
}

async function deleteVisitById(chatId, s, visitId) {
  if (!visitId) return false;
  await deleteStoredVisit(chatId, visitId);
  const arr = Array.isArray(s.visits_index) ? [...s.visits_index] : [];
  s.visits_index = arr.filter((v) => String((v && v.visit_id) || "") !== String(visitId || ""));
  if (String(s.active_visit_id || "") === String(visitId || "")) {
    s.active_visit_id = "";
    s.active_visit_no = 0;
    s.draft = normalizeDraft({});
    s.awaitingField = null;
    s.step = "MENU";
  }
  s.pending_delete_visit_id = "";
  await setSession(chatId, s);
  return true;
}

async function buildVisitsList(chatId, s) {
  const arr = Array.isArray(s.visits_index) ? [...s.visits_index] : [];
  arr.sort((a, b) => sanitizeVisitNo((a && a.no) || 0) - sanitizeVisitNo((b && b.no) || 0));

  // Lecture parallèle
  const recs = await Promise.all(
    arr.map((it) => getStoredVisit(chatId, it.visit_id))
  );

  return arr.map((it, i) => {
    const draft = normalizeDraft((recs[i] && recs[i].draft) || {});
    return {
      visit_id: it.visit_id,
      no: sanitizeVisitNo((it && it.no) || 0),
      title: compactVisitTitle(draft),
      draft,
      updated_at: (recs[i] && recs[i].updated_at) || it.updated_at || 0,
    };
  });
}


// =====================================================
// VALIDATION MINIMUM AVANT SAVE
// =====================================================

function validateDraftMinimum(d) {
  const errors = [];
  if (!d.name || !d.name.trim()) {
    errors.push("le nom de l'entreprise est obligatoire");
  }
  if ((!d.city || !d.city.trim()) && (!d.postal_code || !d.postal_code.trim())) {
    errors.push("la ville ou le code postal est obligatoire");
  }
  return errors;
}

function draftHasContact(d) {
  return !!(
    normalizePhone(d.phone) ||
    normalizePhone(d.phone2) ||
    normalizeEmail(d.email) ||
    (d.website && d.website.trim())
  );
}


// =====================================================
// DAY COUNTS
// =====================================================

async function getDayCount(date, agency) {
  const [pKeys, phKeys, cKeys] = await Promise.all([
    listAllKeys(`prospect:${date}:${agency}:`),
    listAllKeys(`photo:${date}:${agency}:`),
    listAllKeys(`card:${date}:${agency}:`),
  ]);
  return { fiches: pKeys.length, photos: phKeys.length, cartes: cKeys.length };
}


// =====================================================
// LAST PROSPECT
// =====================================================

async function saveLastProspectKey(chatId, key) {
  await kv().put(`prospect:last:${chatId}`, key, { expirationTtl: TTL_PROSPECT });
}

async function getLastProspectKey(chatId) {
  return kv().get(`prospect:last:${chatId}`);
}


// =====================================================
// KEYBOARDS
// =====================================================

function agenciesKeyboard() {
  return {
    inline_keyboard: [
      [{ text: "GR", callback_data: "agency:GR" }, { text: "VR", callback_data: "agency:VR" }],
      [{ text: "GRS", callback_data: "agency:GRS" }, { text: "SLS", callback_data: "agency:SLS" }],
    ],
  };
}

function confirmProfileKeyboard(agency, initials) {
  return {
    inline_keyboard: [
      [{ text: `✅ Continuer ${agency} / ${initials}`, callback_data: "profile:keep" }],
      [{ text: "🔄 Changer de profil", callback_data: "profile:change" }],
    ],
  };
}

function mainMenuKeyboard(hasLastProspect = false, hasVisits = false, isAdmin = false) {
  const kb = [
    [{ text: "➕ Nouvelle visite", callback_data: "visit:new" }],
    [
      { text: "📸 Lot façades", callback_data: "mode:photo" },
      { text: "💳 Lot cartes",  callback_data: "mode:card"  },
    ],
    [{ text: "📋 Liste du jour", callback_data: "visits:list" }],
    [{ text: "✅ Clore la journée", callback_data: "close" }],
  ];
  if (hasLastProspect) kb.splice(3, 0, [{ text: "✏️ Modifier dernier prospect", callback_data: "edit:last" }]);
  if (hasVisits) kb.splice(3, 0, [{ text: "📌 Reprendre fiche active", callback_data: "visit:resume" }]);
  if (isAdmin) kb.push([{ text: "🔄 Nouvelle journée (reset cache)", callback_data: "admin:reset_cache" }]);
  kb.push([{ text: "ℹ️ Aide", callback_data: "help:inline" }]);
  return { inline_keyboard: kb };
}

function failSafeKeyboard(extraRows = [], useWarnMenu = false) {
  return {
    inline_keyboard: [
      ...extraRows,
      [{ text: "🔁 Refaire une recherche", callback_data: "retry_search" }],
      [{ text: "✏️ Saisie manuelle", callback_data: "manual_entry" }],
      [{ text: "📋 Résumé session", callback_data: "summary" }],
      [{ text: "⬅️ Menu", callback_data: useWarnMenu ? "menu_with_warn" : "menu" }],
    ],
  };
}

function companyPickKeyboard(list) {
  const rows = (list || []).map((e, i) => [
    { text: `${e.nom}${e.codePostal ? " (" + e.codePostal + ")" : ""}`, callback_data: `company:${i}` },
  ]);
  rows.push([{ text: "❌ Aucun ne convient", callback_data: "company:none" }]);
  rows.push([{ text: "🔁 Modifier recherche", callback_data: "retry_search" }]);
  rows.push([{ text: "📋 Résumé session", callback_data: "summary" }]);
  rows.push([{ text: "⬅️ Menu", callback_data: "menu" }]);
  return { inline_keyboard: rows };
}

// Auto-save 1-clic après enrichissement complet
function autoSaveKeyboard(summary) {
  return {
    inline_keyboard: [
      [{ text: "💾 Enregistrer direct", callback_data: "save_direct" }],
      [{ text: "✏️ Modifier d'abord", callback_data: "visit:show_form" }],
      [{ text: "⬅️ Menu", callback_data: "menu" }],
    ],
  };
}

function duplicateConfirmKeyboard() {
  return {
    inline_keyboard: [
      [{ text: "✅ Confirmer quand même", callback_data: "save_force" }],
      [{ text: "❌ Annuler", callback_data: "save_cancel_duplicate" }],
    ],
  };
}

function locationRequestKeyboard() {
  return {
    keyboard: [
      [{ text: "📍 Envoyer position", request_location: true }],
      [{ text: "⬅️ Menu" }],
    ],
    resize_keyboard: true,
    one_time_keyboard: true,
  };
}

// Clavier bas d'écran pour les lots photo/carte
// Note : Telegram n'a pas de request_photo. On affiche un bouton texte clair.
function photoReplyKeyboard(label = "📸 Envoyer une photo", showFinish = false) {
  const rows = [
    [{ text: label }],   // bouton texte — Telegram suggère la galerie/caméra via l'icône 📎
  ];
  if (showFinish) rows.push([{ text: "✅ Terminer" }]);
  rows.push([{ text: "⬅️ Menu" }]);
  return {
    keyboard: rows,
    resize_keyboard: true,
    one_time_keyboard: false,
    input_field_placeholder: "📎 Trombone = qualité originale",
  };
}

// Supprime le clavier bas d'écran (après lot terminé)
function removeReplyKeyboard() {
  return { remove_keyboard: true };
}

function startVisitKeyboard() {
  return {
    inline_keyboard: [
      [{ text: "🔎 Recherche souple", callback_data: "visit:start_search" }],
      [{ text: "💳 Carte de visite", callback_data: "visit:start_card" }],
      [{ text: "📷 Photo façade", callback_data: "visit:start_facade" }],
      [{ text: "⬅️ Menu", callback_data: "menu" }],
    ],
  };
}

function visitsListKeyboard(items) {
  const rows = [];
  for (const it of items) {
    rows.push([{
      text: `${it.no}. ${it.draft.name || "SOCIÉTÉ"} — ${it.draft.city || "VILLE"}`,
      callback_data: `visit:open_no:${it.no}`,
    }]);
  }
  rows.push([{ text: "⬅️ Retour", callback_data: "visit:back_to_active_or_menu" }]);
  return { inline_keyboard: rows };
}

// Qualification 1-tap
function qualificationKeyboard() {
  return {
    inline_keyboard: [
      [
        { text: "🔥 Chaud", callback_data: "qualif:chaud" },
        { text: "📅 Tiède", callback_data: "qualif:tiede" },
        { text: "❄️ Froid", callback_data: "qualif:froid" },
      ],
      [{ text: "⏭ Passer", callback_data: "qualif:skip" }],
    ],
  };
}

function buildActiveVisitKeyboard(s) {
  const d = normalizeDraft((s && s.draft) || {});
  const qualifEmoji = d.qualification === QUALIF_HOT ? "🔥"
    : d.qualification === QUALIF_WARM ? "📅"
    : d.qualification === QUALIF_COLD ? "❄️"
    : "⭐";

  const rows = [];

  rows.push([
    { text: d.name ? "✏️ Entreprise" : "🏢 Entreprise", callback_data: "field:name" },
    { text: d.interlocuteur ? "✏️ Interlocuteur" : "👤 Interlocuteur", callback_data: "field:interlocuteur" },
  ]);
  rows.push([
    { text: d.phone ? "✏️ Téléphone" : "☎️ Téléphone", callback_data: "field:phone" },
    { text: d.phone2 ? "✏️ Portable" : "📱 Portable", callback_data: "field:phone2" },
  ]);
  rows.push([
    { text: d.email ? "✏️ Email" : "✉️ Email", callback_data: "field:email" },
    { text: d.website ? "✏️ Site" : "🌐 Site", callback_data: "field:website" },
  ]);
  rows.push([
    { text: d.address ? "✏️ Adresse" : "📍 Adresse", callback_data: "field:address" },
    { text: d.city ? "✏️ Ville" : "🏙️ Ville", callback_data: "field:city" },
  ]);
  rows.push([
    { text: d.resume ? "✏️ Entretien" : "📝 Entretien", callback_data: "field:resume" },
    { text: d.commande ? "✏️ Commande" : "📦 Commande", callback_data: "field:commande" },
  ]);
  rows.push([
    { text: `${qualifEmoji} Qualification`, callback_data: "field:qualification" },
    { text: d.contact_civility ? "✏️ Civilité" : "🎩 Civilité", callback_data: "field:contact_civility" },
  ]);
  rows.push([
    { text: "💳 Carte", callback_data: "visit:add_card" },
    { text: "📷 Façade", callback_data: "visit:add_facade" },
  ]);
  rows.push([
    { text: "🏢 Enrichir", callback_data: "visit:enrich" },
    { text: "📋 Liste du jour", callback_data: "visits:list" },
  ]);
  rows.push([
    { text: "💾 Enregistrer", callback_data: "save" },
    { text: "🗑 Supprimer", callback_data: "visit:delete_active" },
  ]);
  rows.push([
    { text: "➕ Nouvelle visite", callback_data: "visit:new" },
    { text: "⬅️ Menu", callback_data: "menu" },
  ]);

  return { inline_keyboard: rows };
}


// =====================================================
// MAIN MENU / SUMMARY
// =====================================================

async function buildSessionSummary(chatId, s) {
  const date = ymdFR();
  const counts = await getDayCount(date, s.agency);
  let lastProspectName = "";
  let lastProspectCity = "";
  const lastKey = await getLastProspectKey(chatId);
  if (lastKey) {
    const rec = await kv().get(lastKey, { type: "json" });
    if (rec) { lastProspectName = rec.name || ""; lastProspectCity = rec.city || ""; }
  }
  if (!lastProspectName && s?.draft?.name) lastProspectName = s.draft.name;
  if (!lastProspectCity && s?.draft?.city) lastProspectCity = s.draft.city;
  const visitsCount = Array.isArray(s?.visits_index) ? s.visits_index.length : 0;

  return (
    `📋 <b>Résumé session</b>\n\n` +
    `🏷️ Agence : <b>${escapeHtml(s?.agency || "—")}</b>\n` +
    `👤 Initiales : <b>${escapeHtml(s?.initials || "—")}</b>\n` +
    `🧭 Étape : <b>${escapeHtml(s?.step || "—")}</b>\n\n` +
    `📊 <b>Aujourd'hui</b>\n` +
    `• Fiches KV : <b>${counts.fiches}</b>\n` +
    `• Photos lot : <b>${counts.photos}</b>\n` +
    `• Cartes lot : <b>${counts.cartes}</b>\n` +
    `• Visites session : <b>${visitsCount}</b>\n\n` +
    `🏢 Dernière entreprise : <b>${escapeHtml(lastProspectName || "—")}</b>\n` +
    `🏙️ Ville : <b>${escapeHtml(lastProspectCity || "—")}</b>\n` +
    `📍 GPS frais : <b>${getFreshGeoOrNull(s) ? "OUI" : "NON"}</b>\n` +
    `📌 Fiche active : <b>${escapeHtml(currentVisitLabel(s))}</b>`
  );
}

async function showMainMenu(chatId, s, note = "") {
  const date = ymdFR();
  const counts = await getDayCount(date, s.agency);
  const lastKey = await getLastProspectKey(chatId);
  const hasVisits = Array.isArray(s.visits_index) && s.visits_index.length > 0;
  const stats =
    `📋 <b>Activité agence</b> : ${counts.fiches} fiche(s) · ${counts.photos} photo(s) · ${counts.cartes} carte(s)\n` +
    `🏷️ <b>${escapeHtml(s.agency)}</b> | 👤 <b>${escapeHtml(s.initials)}</b>\n\n` +
    `Que veux-tu faire ?`;
  const header = note ? `${note}\n\n` : "";
  const isAdmin = (s.initials || "").toUpperCase() === "SL";
  return sendMessage(chatId, header + stats, mainMenuKeyboard(!!lastKey, hasVisits, isAdmin));
}

function renderVisitsListText(items) {
  if (!items.length) return "📋 <b>Liste du jour</b>\n\nAucune visite enregistrée dans la session.";
  const lines = ["📋 <b>Liste du jour</b>", ""];
  for (const it of items) {
    const qualif = it.draft?.qualification === QUALIF_HOT ? " 🔥"
      : it.draft?.qualification === QUALIF_WARM ? " 📅"
      : it.draft?.qualification === QUALIF_COLD ? " ❄️"
      : "";
    lines.push(`${it.no}. <b>${escapeHtml(it.draft?.name || "SOCIÉTÉ")}</b> — ${escapeHtml(it.draft?.city || "VILLE")}${qualif}`);
  }
  lines.push(""); lines.push("Sélectionne une fiche à rouvrir.");
  return lines.join("\n");
}


// =====================================================
// FORM RENDER
// =====================================================

function draftCompletionPercent(d) {
  const fields = [
    d.name, d.city || d.postal_code, d.phone || d.phone2,
    d.interlocuteur, d.email, d.website, d.address, d.resume, d.commande,
  ];
  const filled = fields.filter((v) => normalizeSpaces(String(v || ""))).length;
  return Math.round((filled / fields.length) * 100);
}

function renderActiveVisit(s) {
  const d = normalizeDraft((s && s.draft) || {});
  const no = currentVisitLabel(s);
  const qualifText = d.qualification === QUALIF_HOT ? "🔥 Chaud"
    : d.qualification === QUALIF_WARM ? "📅 Tiède"
    : d.qualification === QUALIF_COLD ? "❄️ Froid"
    : "—";

  const pct = draftCompletionPercent(d);
  const pctIcon = pct >= 75 ? "🟢" : pct >= 40 ? "🟡" : "🔴";

  // Champs manquants importants
  const missing = [];
  if (!d.name) missing.push("nom");
  if (!d.city && !d.postal_code) missing.push("ville");
  if (!d.phone && !d.phone2) missing.push("tél");
  if (!d.interlocuteur) missing.push("interlocuteur");
  if (!d.email) missing.push("email");

  const header = `📌 <b>Visite ${escapeHtml(no)}</b> ${pctIcon} <i>${pct}%</i>`;
  const subtitle = d._created_from ? `\n<i>Origine : ${escapeHtml(d._created_from)}</i>\n` : `\n`;
  const missingLine = missing.length
    ? `\n⚠️ <i>Manquant : ${missing.join(", ")}</i>`
    : `\n✅ <i>Fiche complète</i>`;

  return (
    `${header}${subtitle}\n` +
    `🏢 <b>${escapeHtml(d.name || "—")}</b>\n` +
    (d.nom_commercial ? `🏪 Enseigne : ${escapeHtml(d.nom_commercial)}\n` : "") +
    `📍 Adresse : ${escapeHtml(d.address || "—")}\n` +
    `🏙️ Ville : ${escapeHtml(`${d.postal_code || ""} ${d.city || ""}`.trim() || "—")}\n` +
    `🆔 SIRET : ${escapeHtml(d.siret || "—")} | NAF : ${escapeHtml(d.naf || "—")}\n` +
    `🏭 Activité : ${escapeHtml(d.secteur_activite || "—")}\n` +
    `👥 Effectif : ${escapeHtml(d.effectif || "—")}\n` +
    `📅 Création : ${escapeHtml(d.date_creation || "—")}\n\n` +
    `👔 Dirigeant : ${escapeHtml(d.dirigeant || "—")}\n` +
    `🎩 Civilité : ${escapeHtml(d.contact_civility || "—")}\n` +
    `👤 Interlocuteur : ${escapeHtml(d.interlocuteur || "—")}\n\n` +
    `☎️ Téléphone : ${escapeHtml(formatPhone(d.phone) || "—")}\n` +
    `📱 Portable : ${escapeHtml(formatPhone(d.phone2) || "—")}\n` +
    `✉️ Email : ${escapeHtml(d.email || "—")}\n` +
    `🌐 Site : ${escapeHtml(d.website || "—")}\n\n` +
    `📝 Entretien : ${escapeHtml(d.resume || "—")}\n` +
    `📦 Commande : ${escapeHtml(d.commande || "—")}\n` +
    `⭐ Qualification : ${qualifText}` +
    missingLine + `\n\n` +
    `💳 Carte liée : ${d.card_photo_file_id ? "✅" : "—"}\n` +
    `📷 Façade liée : ${d.facade_photo_file_id ? "✅" : "—"}`
  );
}


// =====================================================
// TELEGRAM HELPERS
// =====================================================

async function tg(method, payload) {
  const token = globalThis.TELEGRAM_TOKEN;
  if (!token) throw new Error("Missing TELEGRAM_TOKEN");
  const res = await safeFetchJson(
    `https://api.telegram.org/bot${token}/${method}`,
    {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    },
    { timeoutMs: DEFAULT_FETCH_TIMEOUT_MS, retries: 1, label: `TG_${String(method).toUpperCase()}` }
  );
  if (!res.ok || !res.data?.ok) {
    throw new Error(`Telegram ${method} failed: ${res.data?.description || res.error || ""}`);
  }
  return res.data.result;
}

async function sendMessage(chat_id, text, keyboard = null) {
  const payload = { chat_id, text, parse_mode: "HTML", disable_web_page_preview: true };
  if (keyboard) payload.reply_markup = keyboard;
  return tg("sendMessage", payload);
}

async function deleteMessage(chat_id, message_id) {
  try { return await tg("deleteMessage", { chat_id, message_id }); } catch (_) { return null; }
}

async function answerCb(id, text = "") {
  try { await tg("answerCallbackQuery", { callback_query_id: id, text }); } catch (_) {}
}

async function postOrReplaceForm(chatId, s) {
  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);
  const prev = s.form_message_id;

  // Envoyer la fiche comme nouveau message en bas de la conversation
  const sent = await sendMessage(chatId, renderActiveVisit(s), buildActiveVisitKeyboard(s));
  s.form_message_id = sent.message_id;
  await setSession(chatId, s);

  // Supprimer l'ancienne fiche si elle existe (pour éviter les doublons dans le fil)
  if (prev && prev !== sent.message_id) {
    await deleteMessage(chatId, prev).catch(() => {});
  }

  return sent;
}

async function tgGetFile(file_id) {
  if (!file_id) return null;
  return tg("getFile", { file_id });
}

async function tgDownloadFileBytes(file_path) {
  const token = globalThis.TELEGRAM_TOKEN;
  if (!token || !file_path) return new Uint8Array();
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), TG_FILE_TIMEOUT_MS);
    let r;
    try {
      r = await fetch(`https://api.telegram.org/file/bot${token}/${file_path}`, { signal: controller.signal });
    } finally {
      clearTimeout(timer);
    }
    if (!r.ok) return new Uint8Array();
    return new Uint8Array(await r.arrayBuffer());
  } catch (_) {
    return new Uint8Array();
  }
}

async function tgGetFileUrl(file_id) {
  if (!file_id) return "";
  try {
    const file = await tgGetFile(file_id);
    if (!file?.file_path) return "";
    return `https://api.telegram.org/file/bot${globalThis.TELEGRAM_TOKEN}/${file.file_path}`;
  } catch (_) {
    return "";
  }
}


// =====================================================
// GEMINI VISION
// =====================================================

const PROMPT_CARD =
  "Tu es un expert en extraction de cartes de visite. " +
  "IMPORTANT : la carte peut être pivotée (0°, 90°, 180° ou 270°) — lit dans TOUTES les orientations. " +
  "Lis l'IMAGE et retourne STRICTEMENT un JSON avec les clés exactes : " +
  "name (prénom + nom si visibles), company, title, " +
  "contact_civility (M. ou Mme si déductible sinon vide), " +
  "email, phone, phone2, website, postal_code, city, address, " +
  "confidence (low/medium/high selon la lisibilité). " +
  "Règles : si inconnu mettre chaîne vide ; email en minuscules ; " +
  "postal_code = 5 chiffres si visible. Ne pas inventer. " +
  "Si la carte est pivotée, lis quand même toutes les informations visibles.";

const PROMPT_FACADE =
  "Tu analyses une photo de prospection terrain (façade, enseigne, vitrine, logo). " +
  "Retourne STRICTEMENT un JSON avec les clés exactes : " +
  "company (nom exact sur l'enseigne), " +
  "activity (secteur d'activité déduit, max 60 caractères), " +
  "phone (numéro visible sur la façade sinon vide), " +
  "website (url visible sinon vide), " +
  "city (ville si visible sinon vide), " +
  "confidence (low/medium/high selon la lisibilité). " +
  "Si plusieurs enseignes visibles, prendre la plus grande ou la plus centrale. " +
  "Ne pas inventer. Si illisible : confidence low, autres champs vides.";

// Helper interne : parse réponse Gemini
function _parseGeminiJsonResponse(data) {
  const parts = (((data && data.candidates) || [])[0]?.content?.parts) || [];
  let raw = "";
  for (const p of parts) { if (p && typeof p.text === "string") raw += p.text; }
  raw = String(raw || "").trim().replace(/^```json/i, "").replace(/```$/i, "").trim();
  const m = raw.match(/\{[\s\S]*\}/);
  if (!m) return null;
  try { return JSON.parse(m[0]); } catch (_) { return null; }
}

const GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-lite-latest:generateContent";

async function getGeminiEndpoint() {
  return GEMINI_ENDPOINT;
}

async function geminiVisionJsonFromImageBytes(imgBytes, prompt, maxTokens = 700, temperature = 0.1) {
  const apiKey = String(globalThis.GEMINI_API_KEY || "").trim();
  if (!apiKey || !(imgBytes instanceof Uint8Array) || !imgBytes.length) return {};

  try {
    const mime = guessMimeFromBytes(imgBytes);
    const b64  = bytesToBase64(imgBytes);

    const endpoint = await getGeminiEndpoint(apiKey);
    log("info", null, "GEMINI_CALL", { imgLen: imgBytes.length, b64Len: b64.length, temperature, model: endpoint.split("/models/")[1]?.split(":")[0] });

    const r = await fetch(`${endpoint}?key=${encodeURIComponent(apiKey)}`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }, { inlineData: { mimeType: mime, data: b64 } }] }],
        generationConfig: { temperature, maxOutputTokens: maxTokens },
      }),
      signal: AbortSignal.timeout(20000),
    });

    log("info", null, "GEMINI_HTTP_STATUS", { status: r.status, ok: r.ok });

    if (!r.ok) {
      const errText = await r.text().catch(() => "");
      log("warn", null, "GEMINI_HTTP_ERROR", { status: r.status, body: errText.slice(0, 300) });
      return {};
    }

    const j = await r.json();
    const parts = (((j && j.candidates) || [])[0]?.content?.parts) || [];
    let raw = "";
    for (const p of parts) { if (p && typeof p.text === "string") raw += p.text; }
    raw = String(raw || "").trim().replace(/^```json/i, "").replace(/```$/i, "").trim();

    log("info", null, "GEMINI_RAW_TEXT", {
      rawPreview: raw.slice(0, 200),
      finishReason: j?.candidates?.[0]?.finishReason || "",
      hasCandidate: !!j?.candidates?.length,
    });

    const m = raw.match(/\{[\s\S]*\}/);
    if (!m) {
      log("warn", null, "GEMINI_NO_JSON", { raw: raw.slice(0, 200) });
      return {};
    }
    return sanitizeGeminiOutput(JSON.parse(m[0]));
  } catch (e) {
    log("error", null, "GEMINI_VISION_ERROR", { err: e?.message || String(e) });
    return {};
  }
}

// Structuration texte brut par Gemini (sans image = 10x moins cher)
async function geminiStructureFromText(rawText, type) {
  const apiKey = String(globalThis.GEMINI_API_KEY || "").trim();
  if (!apiKey || !rawText) return {};

  const prompt = type === "card"
    ? `Texte extrait d'une carte de visite par OCR:\n\n${rawText}\n\n` +
      `Retourne STRICTEMENT un JSON: name, company, title, contact_civility, ` +
      `email, phone, phone2, website, postal_code, city, address, confidence (low/medium/high). ` +
      `Vide si inconnu. Ne pas inventer.`
    : `Texte extrait d'une enseigne / façade par OCR:\n\n${rawText}\n\n` +
      `Retourne STRICTEMENT un JSON: company, activity, phone, website, city, confidence (low/medium/high). ` +
      `Vide si inconnu. Ne pas inventer.`;

  try {
    const endpoint = await getGeminiEndpoint(apiKey);
    const r = await fetch(`${endpoint}?key=${encodeURIComponent(apiKey)}`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 500 },
      }),
      signal: AbortSignal.timeout(10000),
    });
    if (!r.ok) return {};
    const j = await r.json();
    return sanitizeGeminiOutput(_parseGeminiJsonResponse(j) || {});
  } catch (_) {
    return {};
  }
}

// Transcription audio via Gemini
async function geminiTranscribeAudio(audioBytes, mimeType) {
  const apiKey = String(globalThis.GEMINI_API_KEY || "").trim();
  if (!apiKey || !audioBytes?.length) return "";

  try {
    const b64 = bytesToBase64(audioBytes);
    const endpoint = await getGeminiEndpoint(apiKey);
    const r = await fetch(`${endpoint}?key=${encodeURIComponent(apiKey)}`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        contents: [{
          parts: [
            { text: "Transcris ce message vocal en français. Retourne uniquement la transcription brute, sans commentaire." },
            { inlineData: { mimeType: mimeType || "audio/ogg", data: b64 } },
          ],
        }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 500 },
      }),
      signal: AbortSignal.timeout(15000),
    });
    if (!r.ok) return "";
    const j = await r.json();
    const parts = j?.candidates?.[0]?.content?.parts || [];
    return parts.map((p) => p.text || "").join("").trim();
  } catch (_) {
    return "";
  }
}


// =====================================================
// GOOGLE VISION OCR
// =====================================================

async function googleVisionOCR(imageBase64) {
  const apiKey = String(globalThis.GOOGLE_VISION_API_KEY || "").trim();
  if (!apiKey) return { fullText: "" };
  try {
    const res = await safeFetchJson(
      `https://vision.googleapis.com/v1/images:annotate?key=${encodeURIComponent(apiKey)}`,
      {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          requests: [{
            image: { content: imageBase64 },
            features: [{ type: "DOCUMENT_TEXT_DETECTION", maxResults: 1 }],
          }],
        }),
      },
      { timeoutMs: VISION_FETCH_TIMEOUT_MS, retries: 0, label: "GOOGLE_VISION" }
    );
    if (!res.ok || !res.data) return { fullText: "" };
    return { fullText: res.data.responses?.[0]?.fullTextAnnotation?.text || "" };
  } catch (_) {
    return { fullText: "" };
  }
}


// =====================================================
// OCR HYBRIDE FUSION
// =====================================================

function mergeOCRResults(gemini, visionStructured, fallbacks = {}) {
  const fields = [
    "name", "company", "title", "contact_civility",
    "email", "phone", "phone2", "website",
    "city", "address", "postal_code",
    "activity",
  ];
  const result = {};
  const confidence = {};

  for (const field of fields) {
    const g = String((gemini && gemini[field]) || "").trim();
    const v = String((visionStructured && visionStructured[field]) || "").trim();

    if (g && v && g.toLowerCase() === v.toLowerCase()) {
      result[field] = g;
      confidence[field] = "high";
    } else if (g && v) {
      result[field] = g.length >= v.length ? g : v;
      confidence[field] = "medium";
    } else {
      result[field] = g || v || "";
      confidence[field] = (g || v) ? "medium" : "low";
    }
  }

  if (!result.phone && fallbacks.phone_fallback) {
    result.phone = fallbacks.phone_fallback;
    confidence.phone = "medium";
  }
  if (!result.phone2 && fallbacks.phone2_fallback) {
    result.phone2 = fallbacks.phone2_fallback;
    confidence.phone2 = "medium";
  }
  if (!result.email && fallbacks.email_fallback) {
    result.email = fallbacks.email_fallback;
    confidence.email = "medium";
  }

  const highCount   = Object.values(confidence).filter((c) => c === "high").length;
  const filledCount = Object.values(result).filter(Boolean).length;

  result._confidence = highCount >= 3 ? "high"
    : filledCount >= 3 ? "medium"
    : "low";

  // Remettre la confidence globale Gemini si présente
  const geminiConf = (gemini && gemini.confidence) || "";
  if (geminiConf && result._confidence !== "high") {
    result._confidence = geminiConf;
  }

  return result;
}

// =====================================================
// PRÉPARATION IMAGE CARTE
// =====================================================
// Note : OffscreenCanvas + convertToBlob peut être lent sur CF Workers.
// On préfère envoyer l'image brute avec des prompts orientés.
// prepareCardImage ne fait plus de manipulation pixel — juste les bytes.

function prepareCardImage(imgBytes) {
  // Retourne directement les bytes — pas de Canvas pour éviter les timeouts
  return { variants: [imgBytes], rotated: false };
}

// Timeout helper pour les appels OCR individuels
function withTimeout(promise, ms, fallback) {
  return Promise.race([
    promise,
    new Promise((resolve) => setTimeout(() => resolve(fallback), ms)),
  ]);
}

async function hybridExtractCard(imgBytes) {
  const b64 = bytesToBase64(imgBytes);

  log("info", null, "CARD_OCR_START", { imgLen: imgBytes.length });

  // Gemini + Vision en parallèle — fetch direct sans withTimeout imbriqués
  const [geminiRaw, visionResult] = await Promise.all([
    geminiVisionJsonFromImageBytes(imgBytes, PROMPT_CARD, 800).catch((e) => {
      log("warn", null, "CARD_GEMINI_FAIL", { err: e?.message });
      return {};
    }),
    googleVisionOCR(b64).catch((e) => {
      log("warn", null, "CARD_VISION_FAIL", { err: e?.message });
      return { fullText: "" };
    }),
  ]);

  const gemini  = sanitizeGeminiOutput(geminiRaw || {});
  const rawText = visionResult?.fullText || "";

  log("info", null, "CARD_OCR_RAW", {
    geminiName:    gemini.name    || "(vide)",
    geminiCompany: gemini.company || "(vide)",
    geminiEmail:   gemini.email   || "(vide)",
    geminiConf:    gemini.confidence || "(vide)",
    visionLen:     rawText.length,
    visionPreview: rawText.slice(0, 100).replace(/\n/g, "|"),
  });

  // Fallbacks regex depuis Vision
  const phoneMatches = [...(rawText.matchAll(/(?:0|\+33)[1-9](?:[\s.\-]?\d{2}){4}/g))].map(m => normalizePhone(m[0])).filter(Boolean);
  const emailMatch = rawText.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/);
  // Premier téléphone → fixe, second → portable si différent
  const fallbacks = {
    phone_fallback:  phoneMatches[0] || "",
    phone2_fallback: phoneMatches[1] && phoneMatches[1] !== phoneMatches[0] ? phoneMatches[1] : "",
    email_fallback:  emailMatch ? normalizeEmail(emailMatch[0]) : "",
  };

  // Injection directe des fallbacks si Gemini a raté email/tel
  if (!gemini.email && fallbacks.email_fallback) gemini.email = fallbacks.email_fallback;
  if (!gemini.phone && fallbacks.phone_fallback) gemini.phone = fallbacks.phone_fallback;

  const merged = mergeOCRResults(gemini, {}, fallbacks);

  log("info", null, "CARD_OCR_DONE", {
    name: merged.name || "(vide)",
    company: merged.company || "(vide)",
    conf: merged._confidence,
  });

  return merged;
}

async function hybridExtractFacade(imgBytes, cityHint = "") {
  const b64 = bytesToBase64(imgBytes);
  const hasGemini = !!String(globalThis.GEMINI_API_KEY || "").trim();
  const hasVision = !!String(globalThis.GOOGLE_VISION_API_KEY || "").trim();

  log("info", null, "FACADE_OCR_START", { imgLen: imgBytes.length, cityHint, hasGemini, hasVision });

  // Appels en parallèle — chacun a son propre timeout interne
  const [geminiRaw, visionResult] = await Promise.all([
    geminiVisionJsonFromImageBytes(imgBytes, PROMPT_FACADE, 500, 0.3).catch((e) => {
      log("warn", null, "FACADE_GEMINI_FAIL", { err: e?.message });
      return {};
    }),
    googleVisionOCR(b64).catch((e) => {
      log("warn", null, "FACADE_VISION_FAIL", { err: e?.message });
      return { fullText: "" };
    }),
  ]);

  const gemini  = sanitizeGeminiOutput(geminiRaw || {});
  const rawText = visionResult?.fullText || "";

  log("info", null, "FACADE_OCR_RAW", {
    geminiCompany: gemini.company || "(vide)",
    geminiConf:    gemini.confidence || "(vide)",
    visionLen:     rawText.length,
    visionPreview: rawText.slice(0, 120).replace(/\n/g, "|"),
  });

  const phoneMatch = rawText.match(/(?:0|\+33)[1-9](?:[\s.\-]?\d{2}){4}/);
  const urlMatch   = rawText.match(/(?:www\.|https?:\/\/)[^\s]{4,}/i);

  let company  = gemini.company  || "";
  let activity = gemini.activity || "";

  // Fallback Vision si Gemini n'a pas trouvé le nom
  if (!company && rawText) {
    const lines = rawText.split("\n").map(l => l.trim()).filter(l => l.length >= 2 && l.length <= 50);
    const allCaps = lines.find(l => /^[A-Z0-9\s\-&'\.]{2,30}$/.test(l) && l.replace(/\s/g, "").length >= 2);
    if (allCaps) company = allCaps;
    if (!company) {
      const firstShort = lines.find(l => l.length <= 25 && l.length >= 3);
      if (firstShort) company = firstShort;
    }
  }

  const finalConf = company ? "medium" : "low";
  log("info", null, "FACADE_OCR_DONE", { company: company || "(vide)", activity, finalConf });

  return {
    company,
    activity,
    city:        gemini.city || cityHint || "",
    phone:       normalizePhone(gemini.phone || (phoneMatch ? phoneMatch[0] : "") || ""),
    website:     normalizeSpaces(gemini.website || (urlMatch ? urlMatch[0] : "") || ""),
    confidence:  finalConf,
    _confidence: finalConf,
  };
}

// =====================================================
// GÉOLOCALISATION INVERSE — API BAN (gratuit, France)
// =====================================================

async function reverseGeocode(lat, lon) {
  try {
    const res = await safeFetchJson(
      `https://api-adresse.data.gouv.fr/reverse/?lon=${lon}&lat=${lat}&limit=1`,
      {},
      { timeoutMs: 5000, retries: 0, label: "BAN_REVERSE" }
    );
    if (!res.ok || !res.data) return null;
    const feat = res.data?.features?.[0];
    if (!feat) return null;
    const props = feat.properties || {};
    return {
      city:        normalizeSpaces(props.city || props.municipality || ""),
      postal_code: normalizeSpaces(props.postcode || ""),
      address:     normalizeSpaces(props.name || ""),
      label:       normalizeSpaces(props.label || ""),
      score:       props.score || 0,
    };
  } catch (_) {
    return null;
  }
}

// Formater la durée depuis le dernier GPS
function geoAgeLabel(at) {
  if (!at) return "";
  const mins = Math.round((Date.now() - at) / 60000);
  if (mins < 1)  return "à l'instant";
  if (mins < 60) return `il y a ${mins} min`;
  return `il y a ${Math.round(mins / 60)}h`;
}
function countFilledFields(obj) {
  if (!obj || typeof obj !== "object") return 0;
  const skip = new Set(["confidence", "_confidence"]);
  return Object.entries(obj).filter(([k, v]) =>
    !skip.has(k) && v && String(v).trim() !== "" && String(v).toLowerCase() !== "low"
  ).length;
}

async function sendOCRFeedback(chatId, result, type) {
  const conf    = result._confidence || result.confidence || "low";
  const filled  = countFilledFields(result);

  if (conf === "low" || filled < 2) {
    return sendMessage(
      chatId,
      "⚠️ <b>Photo difficile à lire</b> — " + filled + " champ(s) extrait(s).\n\n" +
      "<b>Conseils pour une meilleure lecture :</b>\n" +
      "• 📐 Cadre la carte <b>à plat</b>, en remplissant l'image\n" +
      "• 💡 Évite les ombres et les reflets\n" +
      "• 📎 Utilise le <b>trombone</b> pour envoyer sans compression\n" +
      "• 🔄 Si la carte est verticale, essaie de la tourner avant de photographier\n\n" +
      "Tu peux renvoyer la photo ou compléter les champs manuellement.",
      {
        inline_keyboard: [
          [{ text: "📷 Renvoyer la photo", callback_data: "card_retry_photo" }],
          [{ text: "✏️ Saisir manuellement", callback_data: "edit:fields" }],
        ],
      }
    );
  }

  if (type === "card") {
    const icon = conf === "high"
      ? "✅ <b>Carte lue — confiance HAUTE</b>"
      : "🟡 <b>Carte lue — confiance MOYENNE</b>";
    const hi = () => conf === "high" ? " ✅✅" : " ✅";
    const lines = [icon, "━━━━━━━━━━━━━━━━━━━━━━━"];
    if (result.company || result.name) {
      lines.push(`🏢 ${escapeHtml(result.company || result.name)}${hi()}`);
    }
    if (result.name && result.name !== result.company) {
      lines.push(`👤 ${escapeHtml(result.name)}${hi()}`);
    }
    if (result.title)   lines.push(`💼 ${escapeHtml(result.title)}`);
    if (result.phone)   lines.push(`📞 ${escapeHtml(formatPhone(result.phone))}${hi()}`);
    if (result.email)   lines.push(`✉️ ${escapeHtml(result.email)}${hi()}`);
    if (result.website) lines.push(`🌐 ${escapeHtml(result.website)}`);
    if (result.city || result.postal_code) {
      lines.push(`📍 ${escapeHtml([result.postal_code, result.city].filter(Boolean).join(" "))}`);
    }
    return sendMessage(chatId, lines.join("\n"));
  }

  if (type === "facade") {
    if (result.company) {
      return sendMessage(
        chatId,
        `${conf === "high" ? "✅" : "🟡"} <b>${escapeHtml(result.company)}</b> détecté.\n` +
        (result.activity ? `🏭 ${escapeHtml(result.activity)}\n` : "") +
        (result.phone    ? `📞 ${escapeHtml(formatPhone(result.phone))}\n`    : "") +
        (result.website  ? `🌐 ${escapeHtml(result.website)}\n`  : "")
      );
    }
  }
}


// =====================================================
// BYTES / BASE64
// =====================================================

function bytesToBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const sub = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, Array.from(sub));
  }
  return btoa(binary);
}

function guessMimeFromBytes(bytes) {
  if (bytes.length >= 4) {
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
    if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
    if (bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46) return "image/webp";
  }
  return "image/jpeg";
}

async function getBytesFromFileId(file_id) {
  if (!file_id) return new Uint8Array();
  try {
    const file = await tgGetFile(file_id);
    if (!file?.file_path) return new Uint8Array();
    return tgDownloadFileBytes(file.file_path);
  } catch (_) {
    return new Uint8Array();
  }
}


// =====================================================
// VALIDATION ADRESSE BAN
// =====================================================

async function banNormalizeAddress(address, postalCode, city) {
  const q = [address, postalCode, city].filter(Boolean).join(" ").trim();
  if (!q || q.length < 5) return null;
  try {
    const res = await safeFetchJson(
      `https://api-adresse.data.gouv.fr/search/?q=${encodeURIComponent(q)}&limit=1`,
      {},
      { timeoutMs: 5000, retries: 0, label: "BAN_ADDRESS" }
    );
    if (!res.ok || !res.data) return null;
    const feat = res.data?.features?.[0];
    if (!feat) return null;
    const props = feat.properties || {};
    const coords = feat.geometry?.coordinates || [];
    return {
      address:     normalizeSpaces(props.name || ""),
      postal_code: normalizeSpaces(props.postcode || postalCode || ""),
      city:        normalizeSpaces(props.city || city || ""),
      score:       props.score || 0,
      lat:         coords[1] || null,
      lon:         coords[0] || null,
    };
  } catch (_) {
    return null;
  }
}


// =====================================================
// ENRICHISSEMENT ENTREPRISE
// =====================================================

async function searchCompany(query) {
  const url = `https://recherche-entreprises.api.gouv.fr/search?q=${encodeURIComponent(query)}&page=1&per_page=5`;
  const res = await safeFetchJson(
    url,
    { headers: { accept: "application/json" } },
    { timeoutMs: GOUV_FETCH_TIMEOUT_MS, retries: 1, label: "GOUV_SEARCH" }
  );
  if (!res.ok || !res.data) throw new Error(`API Gouv search failed: ${res.error || ""}`);

  return (res.data.results || []).map((e) => {
    const siege = e.siege || {};
    const trancheEff = e.tranche_effectif_salarie || siege.tranche_effectif_salarie || "";
    const dateCreation = e.date_creation || "";
    const capital = e.capital || "";

    return {
      nom: e.nom_raison_sociale || e.denomination || "",
      siret: String(siege.siret || ""),
      siren: e.siren || "",
      adresse: siege.adresse || siege.libelle_voie || "",
      codePostal: siege.code_postal || "",
      ville: siege.libelle_commune || "",
      naf: e.activite_principale || e.naf || "",
      secteur_activite: nafLabel(e.activite_principale || e.naf || ""),
      dirigeants: e.dirigeants || e.representants || [],
      effectif: effectifLabel(trancheEff),
      date_creation: dateCreation,
      capital: capital ? String(capital) : "",
    };
  });
}

async function searchCompanyCached(query) {
  const q = normalizeSpaces(query);
  if (!q) return [];
  const h = await sha1HexFromString(q);
  const { value } = await withCache(`gouv:${h}`, CACHE_TTL_GOUV, async () => {
    const results = await searchCompany(q);
    return { ts: Date.now(), results };
  });
  return Array.isArray(value?.results) ? value.results : [];
}

function extractDirigeantFromSearch(e) {
  const arr = Array.isArray(e.dirigeants) ? e.dirigeants : [];
  if (!arr.length) return "";
  const first = arr[0];
  if (typeof first === "string") return normalizeSpaces(first);
  const p = first.personne || first;
  // Garder uniquement le premier prénom (évite "LIONEL JACQUES JEAN LACRAMPE")
  const prenomFull = p.prenom || p.prenoms || "";
  const prenom = prenomFull.split(/\s+/)[0] || "";
  const nom = p.nom || p.nom_usage || "";
  // Capitaliser : "LACRAMPE" → "Lacrampe", "lionel" → "Lionel"
  const capitalize = (s) => s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : "";
  const result = [capitalize(prenom), nom.toUpperCase()].filter(Boolean).join(" ");
  // Ne pas utiliser denomination comme fallback — c'est le nom de l'entreprise, pas du dirigeant
  return cleanPersonName(result);
}

async function googlePlacesEnrich(name, city) {
  const apiKey = String(globalThis.GOOGLE_PLACES_API_KEY || "").trim();
  if (!apiKey) return { phone: "", website: "", address: "" };
  try {
    const r1 = await safeFetchJson(
      "https://places.googleapis.com/v1/places:searchText",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Goog-Api-Key": apiKey,
          "X-Goog-FieldMask": "places.id,places.websiteUri",
        },
        body: JSON.stringify({ textQuery: `${name} ${city}`.trim(), maxResultCount: 1, languageCode: "fr", regionCode: "FR" }),
      },
      { timeoutMs: PLACES_FETCH_TIMEOUT_MS, retries: 1, label: "PLACES_SEARCH" }
    );
    if (!r1.ok || !r1.data) return { phone: "", website: "", address: "" };
    const p = r1.data.places?.[0];
    if (!p) return { phone: "", website: "", address: "" };
    const placeId = p.id || "";
    let website = normalizeSpaces(p.websiteUri || "");
    let phone = "", address = "";
    if (placeId) {
      const r2 = await safeFetchJson(
        `https://places.googleapis.com/v1/places/${encodeURIComponent(placeId)}`,
        {
          method: "GET",
          headers: {
            "X-Goog-Api-Key": apiKey,
            "X-Goog-FieldMask": "nationalPhoneNumber,internationalPhoneNumber,websiteUri,formattedAddress",
          },
        },
        { timeoutMs: PLACES_FETCH_TIMEOUT_MS, retries: 1, label: "PLACES_DETAIL" }
      );
      if (r2.ok && r2.data) {
        phone   = normalizePhone(r2.data.nationalPhoneNumber || r2.data.internationalPhoneNumber || "");
        website = normalizeSpaces(r2.data.websiteUri || website);
        address = normalizeSpaces(r2.data.formattedAddress || "");
      }
    }
    return { phone, website, address };
  } catch (_) {
    return { phone: "", website: "", address: "" };
  }
}

async function googlePlacesEnrichCached(name, city) {
  const raw = `${normalizeSpaces(name)}|${normalizeSpaces(city)}`;
  const h   = await sha1HexFromString(raw);
  const { value } = await withCache(`places:${h}`, CACHE_TTL_PLACES, async () => {
    const d = await googlePlacesEnrich(name, city);
    return { ts: Date.now(), phone: d.phone || "", website: d.website || "", address: d.address || "" };
  });
  return {
    phone:   normalizePhone(value?.phone || ""),
    website: normalizeSpaces(value?.website || ""),
    address: normalizeSpaces(value?.address || ""),
  };
}

async function fetchHtml(url) {
  if (!isSafeUrl(url)) return "";
  try {
    const headers = globalThis.BROWSER_HEADERS
      ? JSON.parse(globalThis.BROWSER_HEADERS)
      : { "user-agent": "Mozilla/5.0 (compatible; ProspectionBot/8.4)" };
    const res = await safeFetchText(url, { headers }, {
      timeoutMs: SCRAPE_TIMEOUT_MS,
      retries: 0,
      label: "SCRAPE_HTML",
    });
    if (!res.ok || !res.text) return "";
    return res.text.slice(0, SCRAPE_MAX_BYTES);
  } catch (_) {
    return "";
  }
}

function pickEmailFromHtml(html, websiteUrl = "") {
  if (!html) return "";

  // Déobfuscation courante avant extraction
  let text = html
    .replace(/\[at\]/gi, "@").replace(/\(at\)/gi, "@").replace(/\s+at\s+/gi, "@")
    .replace(/\[dot\]/gi, ".").replace(/\(dot\)/gi, ".").replace(/\s+dot\s+/gi, ".")
    .replace(/&#64;/g, "@").replace(/&commat;/g, "@")
    .replace(/mailto:/gi, "");

  const matches = text.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || [];
  if (!matches.length) return "";

  const domain = (() => {
    try { return new URL(websiteUrl).hostname.replace(/^www\./i, "").toLowerCase(); }
    catch (_) { return ""; }
  })();

  const filtered = matches
    .map((m) => normalizeEmail(m))
    .filter(Boolean)
    .filter((m) => !/no-?reply|spam|robot|bot|example|test@/i.test(m))
    .filter((m) => !m.endsWith(".png") && !m.endsWith(".jpg"));

  if (!filtered.length) return "";
  const sameDomain = domain ? filtered.find((m) => m.endsWith("@" + domain)) : "";
  return sameDomain || filtered[0] || "";
}

async function scrapeEmailFromSite(url) {
  if (!url || !isSafeUrl(url)) return "";
  try {
    const base = new URL(url);
    const domain = base.hostname.replace(/^www\./i, "");

    // Pages à scraper dans l'ordre (les plus susceptibles d'avoir un email)
    const pagePaths = [
      "/contact", "/nous-contacter", "/contactez-nous", "/contact-us",
      "/mentions-legales", "/mentions-légales", "/legal",
      "/a-propos", "/qui-sommes-nous", "/equipe", "/team", "/about",
      "/", // page d'accueil en dernier
    ];

    const candidates = [...new Set(pagePaths.map(p => {
      try { return new URL(p, base).toString(); } catch (_) { return ""; }
    }))].filter(u => u && isSafeUrl(u));

    let allText = "";
    for (const candidate of candidates) {
      const html = await fetchHtml(candidate);
      if (!html) continue;
      const email = pickEmailFromHtml(html, url);
      if (email) return email;
      // Accumuler le texte brut pour Gemini si pas d'email trouvé
      if (allText.length < 3000) {
        // Extraire le texte visible uniquement (sans balises HTML)
        const text = html.replace(/<script[\s\S]*?<\/script>/gi, "")
          .replace(/<style[\s\S]*?<\/style>/gi, "")
          .replace(/<[^>]+>/g, " ")
          .replace(/\s+/g, " ")
          .slice(0, 1000);
        allText += text + " ";
      }
    }

    // Fallback Gemini : chercher l'email dans le texte brut accumulé
    if (allText.trim() && String(globalThis.GEMINI_API_KEY || "").trim()) {
      const prompt =
        `Voici le texte extrait du site web de l'entreprise "${domain}".\n\n` +
        `${allText.slice(0, 2500)}\n\n` +
        `Trouve l'adresse email de contact de cette entreprise. ` +
        `Réponds UNIQUEMENT avec l'email (ex: contact@entreprise.fr) ou "aucun" si introuvable.`;
      try {
        const r = await fetch(`${GEMINI_ENDPOINT}?key=${encodeURIComponent(String(globalThis.GEMINI_API_KEY).trim())}`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { temperature: 0.0, maxOutputTokens: 50 },
          }),
          signal: AbortSignal.timeout(8000),
        });
        if (r.ok) {
          const j = await r.json();
          const raw = (j?.candidates?.[0]?.content?.parts?.[0]?.text || "").trim().toLowerCase();
          const emailMatch = raw.match(/[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/);
          if (emailMatch) {
            const found = normalizeEmail(emailMatch[0]);
            if (found && !/no-?reply/i.test(found)) {
              log("info", null, "EMAIL_GEMINI_FOUND", { email: found, domain });
              return found;
            }
          }
        }
      } catch (_) {}
    }

    return "";
  } catch (_) {
    return "";
  }
}

// Deviner l'email à partir du domaine du site web
// Ex: pissard.fr → contact@pissard.fr (vérifié via DNS MX)
async function guessEmailFromDomain(website, markAsGuessed = true) {
  if (!website) return "";
  let domain = "";
  try {
    domain = new URL(website.startsWith("http") ? website : `https://${website}`)
      .hostname.replace(/^www\./, "");
  } catch (_) { return ""; }
  if (!domain) return "";

  // Vérifier via DNS-over-HTTPS Cloudflare si le domaine a un MX
  let hasMX = false;
  try {
    const r = await fetch(
      `https://cloudflare-dns.com/dns-query?name=${encodeURIComponent(domain)}&type=MX`,
      { headers: { "accept": "application/dns-json" }, signal: AbortSignal.timeout(3000) }
    );
    if (r.ok) {
      const j = await r.json();
      hasMX = !!(j?.Answer?.length);
    }
  } catch (_) {}

  if (!hasMX) return "";

  // Pattern le plus probable pour les PME françaises
  const prefixes = ["contact", "info", "commercial", "accueil"];
  return `${prefixes[0]}@${domain}`;
}

// Chercher l'email avec headers navigateur réalistes
async function fetchEmailFromPage(url) {
  if (!url || !isSafeUrl(url)) return "";
  try {
    const r = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
        "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "fr-FR,fr;q=0.9",
      },
      signal: AbortSignal.timeout(6000),
    });
    if (!r.ok) return "";
    const html = await r.text();
    return pickEmailFromHtml(html, url);
  } catch (_) { return ""; }
}

// Recherche Google CSE — email ET infos générales entreprise
// =====================================================
// BRAVE SEARCH — enseigne→société, email, site, tel
// =====================================================

async function searchViaBrave(query) {
  const key = String(globalThis.BRAVE_SEARCH_KEY || "").trim();
  if (!key || !query) return {};

  const EMAIL_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/gi;
  const PHONE_RE = /(?:0|\+33)[1-9](?:[\s.\-]?\d{2}){4}/g;
  const SIREN_RE = /\b(\d{3}[\s.\-]?\d{3}[\s.\-]?\d{3})\b/g;
  const SIRET_RE = /\b(\d{3}[\s.\-]?\d{3}[\s.\-]?\d{3}[\s.\-]?\d{5})\b/g;

  try {
    const url = "https://api.search.brave.com/res/v1/web/search?" +
      new URLSearchParams({ q: query, count: "5", country: "fr", search_lang: "fr" });

    const r = await fetch(url, {
      headers: {
        "Accept": "application/json",
        "Accept-Encoding": "gzip",
        "X-Subscription-Token": key,
      },
      signal: AbortSignal.timeout(8000),
    });

    if (!r.ok) {
      log("warn", null, "BRAVE_ERROR", { status: r.status, query: query.slice(0,50) });
      return {};
    }

    const j = await r.json();
    const results = j?.web?.results || [];

    // Trier : résultats avec CP en premier
    const loc = query.match(/\b\d{5}\b/)?.[0] || "";
    results.sort((a, b) => {
      const aLoc = loc && (a.title||"").includes(loc) ? 1 : 0;
      const bLoc = loc && (b.title||"").includes(loc) ? 1 : 0;
      return bLoc - aLoc;
    });

    let email = "", phone = "", phone2 = "", siren = "", siret = "", website = "";

    // Passe 1 : chercher SIRET 14 chiffres dans TOUS les résultats
    for (const item of results) {
      if (siret) break;
      const sources = [item.url || "", item.title || ""];
      for (const src of sources) {
        const m = src.match(/\b(\d{3}[\s.\-]?\d{3}[\s.\-]?\d{3}[\s.\-]?\d{5})\b/);
        if (m) {
          const s = m[1].replace(/[\s.\-]/g, "");
          if (s.length === 14 && !/^202[0-9]/.test(s)) {
            siret = s;
            siren = s.slice(0, 9);
            log("info", null, "BRAVE_SIRET_FOUND", { siret, from: src.slice(0,80) });
            break;
          }
        }
      }
    }

    // Passe 2 : email, téléphone, SIREN, site web
    for (const item of results) {
      const text = `${item.title || ""} ${item.description || ""}`.replace(/&amp;/g, "&");

      if (!email) {
        for (const m of (text.match(EMAIL_RE) || [])) {
          const e = normalizeEmail(m);
          if (e && !/no-?reply|google|example|@brave/i.test(e)) { email = e; break; }
        }
      }

      if (!phone || !phone2) {
        const phones = (text.match(PHONE_RE) || []).map(m => normalizePhone(m)).filter(Boolean);
        if (!phone && phones[0]) phone = phones[0];
        if (!phone2 && phones[1] && phones[1] !== phone) phone2 = phones[1];
      }

      // SIREN depuis URL si pas de SIRET
      if (!siren && item.url) {
        const sirenM = item.url.match(/[-\/](\d{9})(?:[-\/\.]|$)/);
        if (sirenM) siren = sirenM[1];
      }

      if (!website && item.url) {
        try {
          const u = new URL(item.url);
          if (!u.hostname.includes("google") && !u.hostname.includes("brave") &&
              !u.hostname.includes("facebook") && !u.hostname.includes("linkedin")) {
            website = `${u.protocol}//${u.hostname}`;
          }
        } catch (_) {}
      }
    }

    if (email || phone || siren || siret) {
      log("info", null, "BRAVE_FOUND", { email: email||"", phone: phone||"", siren: siren||"", siret: siret||"", query: query.slice(0,60) });
    }

    return { email, phone, phone2, siren, siret, website };
  } catch (e) {
    log("warn", null, "BRAVE_EXCEPTION", { err: e?.message, query: query.slice(0,50) });
    return {};
  }
}

async function searchEmailViaGoogle(companyName, city) {
  const result = await searchViaBrave(`"${companyName}" "${city}" "@"`);
  return result.email || "";
}




// Deviner l'email à partir du domaine du site web
// Ex: pissard.fr → contact@pissard.fr (vérifié via DNS MX)
async function scrapeEmailCached(url) {
  if (!url || !isSafeUrl(url)) return "";
  let host = "";
  try { host = new URL(url).hostname.replace(/^www\./i, "").toLowerCase(); }
  catch (_) { return ""; }
  if (!host) return "";
  const h = await sha1HexFromString(host);
  const cacheKey = `email:${h}`;

  // Vérifier le cache — ne pas utiliser si "none" (résultat vide précédent)
  const cached = await kv().get(cacheKey, { type: "text" }).catch(() => null);
  if (cached && cached !== "none") {
    const e = normalizeEmail(cached);
    if (e) return e;
  }

  // Scraping frais
  const email = await scrapeEmailFromSite(url);
  const toStore = email || "none";
  // Ne mettre en cache "none" que 1 jour (réessai demain)
  // Mettre en cache un email valide 7 jours
  const ttl = email ? CACHE_TTL_EMAIL : 60 * 60 * 24;
  await kv().put(cacheKey, toStore, { expirationTtl: ttl }).catch(() => {});
  return email;
}


// =====================================================
// BRAVE SEARCH — Recherche web pour enseignes/emails
// =====================================================

async function braveSearch(query, count = 5) {
  const key = String(globalThis.BRAVE_SEARCH_KEY || "").trim();
  if (!key) return [];
  try {
    const url = "https://api.search.brave.com/res/v1/web/search?q=" +
      encodeURIComponent(query) + "&count=" + count + "&country=fr&lang=fr&search_lang=fr";
    const r = await fetch(url, {
      headers: {
        "Accept": "application/json",
        "Accept-Encoding": "gzip",
        "X-Subscription-Token": key,
      },
      signal: AbortSignal.timeout(8000),
    });
    if (!r.ok) {
      log("warn", null, "BRAVE_ERROR", { status: r.status, query: query.slice(0,50) });
      return [];
    }
    const j = await r.json();
    return j?.web?.results || [];
  } catch (e) {
    log("warn", null, "BRAVE_EXCEPTION", { err: e?.message, query: query.slice(0,50) });
    return [];
  }
}

// Extraire email depuis les snippets Brave
function extractEmailFromBraveResults(results) {
  const EMAIL_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g;
  for (const r of results) {
    const text = `${r.title || ""} ${r.description || ""} ${r.url || ""}`;
    for (const m of (text.match(EMAIL_RE) || [])) {
      const e = normalizeEmail(m);
      if (e && !/no-?reply|google|example|brave|w3c/i.test(e)) return e;
    }
  }
  return "";
}

// Extraire SIREN depuis les snippets Brave
function extractSirenFromBraveResults(results) {
  const SIREN_RE = /\b(\d{3}[\s.\-]?\d{3}[\s.\-]?\d{3})\b/g;
  const SIRET_RE = /\b(\d{3}[\s.\-]?\d{3}[\s.\-]?\d{3}[\s.\-]?\d{5})\b/g;
  for (const r of results) {
    const text = `${r.title || ""} ${r.description || ""}`;
    // SIRET d'abord (plus précis)
    for (const m of (text.match(SIRET_RE) || [])) {
      const s = m.replace(/[\s.\-]/g, "");
      if (s.length === 14) return s.slice(0, 9); // retourner le SIREN
    }
    // SIREN
    for (const m of (text.match(SIREN_RE) || [])) {
      const s = m.replace(/[\s.\-]/g, "");
      if (s.length === 9 && !s.startsWith("000")) return s;
    }
  }
  return "";
}

// Trouver la société juridique d'une enseigne via Brave
async function findJuridicalFromEnseigne(enseigneName, ville, cp, phone = "") {
  const locFull = [ville, cp].filter(Boolean).join(" ");
  const queries = [
    `${enseigneName} ${locFull} SIRET`,
    phone ? `${enseigneName} ${locFull} ${phone}` : null,
    `${enseigneName} ${locFull} pappers`,
    `${enseigneName} ${locFull} siren`,
    `${enseigneName} ${locFull} societe.com`,
    `${enseigneName} ${locFull} verif.com`,
    `${enseigneName} ${locFull} data.gouv.fr`,
  ].filter(Boolean);

  let bestResult = {};
  for (const q of queries) {
    const result = await searchViaBrave(q);
    log("info", null, "BRAVE_JURIDIQUE", { q, siren: result.siren, siret: result.siret });
    if (result.siret) return result; // SIRET trouvé → stop
    if (result.siren && !bestResult.siren) bestResult = result; // Garder le SIREN en fallback
  }
  return bestResult; // Retourner le meilleur résultat (avec SIREN si trouvé)
}


// Chercher un établissement précis par son SIRET complet
async function searchEtablissementBySiret(siret) {
  if (!siret || siret.length !== 14) return null;
  try {
    // L'API Gouv permet de chercher par SIRET dans le champ q
    const url = "https://recherche-entreprises.api.gouv.fr/search?q=" + encodeURIComponent(siret) + "&limite=5";
    const r = await safeFetchJson(url, {}, { timeoutMs: 5000, retries: 0, label: "GOUV_SIRET" });
    if (!r.ok || !r.data?.results?.length) return null;

    // Chercher l'établissement avec ce SIRET exact dans les résultats
    for (const res of r.data.results) {
      // Vérifier dans les établissements de la société
      if (res.matching_etablissements) {
        for (const etab of res.matching_etablissements) {
          if ((etab.siret || "").replace(/\s/g,"") === siret) {
            return {
              nom:              res.nom_raison_sociale || res.nom_complet || "",
              siret:            siret,
              naf:              etab.activite_principale || res.activite_principale || "",
              adresse:          etab.adresse || "",
              codePostal:       etab.code_postal || "",
              ville:            etab.libelle_commune || "",
              effectif:         res.tranche_effectif_salarie || "",
              date_creation:    res.date_creation || "",
              dirigeants:       res.dirigeants || [],
              secteur_activite: res.libelle_activite_principale || "",
            };
          }
        }
      }
      // Sinon prendre le siège si SIRET correspond
      const siege = res.siege || {};
      if ((siege.siret || "").replace(/\s/g,"") === siret) {
        return {
          nom:              res.nom_raison_sociale || res.nom_complet || "",
          siret:            siret,
          naf:              siege.activite_principale || res.activite_principale || "",
          adresse:          siege.adresse || "",
          codePostal:       siege.code_postal || "",
          ville:            siege.libelle_commune || "",
          effectif:         res.tranche_effectif_salarie || "",
          date_creation:    res.date_creation || "",
          dirigeants:       res.dirigeants || [],
          secteur_activite: res.libelle_activite_principale || "",
        };
      }
    }
    return null;
  } catch (_) { return null; }
}

// Enseignes nationales connues — mode franchise
const KNOWN_BRANDS = new Set([
  "intermarche","intermarché","netto","carrefour","carrefourmarket","superU","hyperU",
  "leclerc","lidl","aldi","auchan","casino","monoprix","franprix","spar","vival",
  "bricomarche","bricodepot","leroymerlin","castorama","mr bricolage","gedimat","pointp",
  "mcdonalds","burgerking","kfc","subway","pizzahut","dominos",
  "roady","norauto","feuvert","speedy","midas","euromaster",
  "renault","peugeot","citroen","volkswagen","ford","toyota","bmw","mercedes",
  "ibis","mercure","novotel","campanile","kyriad","formule1",
]);

// Cache terrain enseigne → SIRET (persisté en KV)
async function getEnseigneCache(key) {
  try { return await kv().get(`enseigne_cache:${key}`, { type: "json" }); } catch (_) { return null; }
}
async function setEnseigneCache(key, data) {
  try { await kv().put(`enseigne_cache:${key}`, JSON.stringify(data), { expirationTtl: 60*60*24*90 }); } catch (_) {}
}

async function enrichCompany(e) {
  const naf     = normalizeNAF(e.naf || "");
  let dirigeant = extractDirigeantFromSearch(e);
  let gouvData  = e;

  // Si SIRET présent mais mauvais département → on cherche quand même
  const eSiretDept = (e.codePostal || "").slice(0, 2);
  const pDeptCheck = (e._places_phone || e.adresse) ? (e.codePostal || "").slice(0, 2) : "";
  const siretWrongDept = e.siret && pDeptCheck && eSiretDept && eSiretDept !== pDeptCheck;

  if ((!e.siret || siretWrongDept) && e.nom) {
    const cp      = e.codePostal || "";
    const ville   = e.ville || "";
    const phone   = normalizePhone(e._places_phone || "");
    const adresse = e.adresse || "";
    const pNomNorm = normalizeForMatch(e.nom);

    // Si pas de CP mais GPS disponible → reverse geocoding pour obtenir le CP
    let pDept = cp.slice(0, 2);
    if (!pDept && (e.gps_lat || e._places_lat)) {
      try {
        const lat = e.gps_lat || e._places_lat;
        const lon = e.gps_lon || e._places_lon;
        const geoRes = await reverseGeocode(lat, lon);
        if (geoRes?.postal_code) {
          pDept = geoRes.postal_code.slice(0, 2);
          log("info", null, "ENRICH_GPS_DEPT", { lat, lon, dept: pDept, cp: geoRes.postal_code });
        }
      } catch (_) {}
    }
    const isFranchise = [...KNOWN_BRANDS].some(b => pNomNorm.includes(normalizeForMatch(b)));

    // ── Cache terrain ─────────────────────────────────────────────
    const cacheKey = normalizeForMatch(`${e.nom} ${ville || cp}`);
    const cached = await getEnseigneCache(cacheKey);
    if (cached?.siret) {
      // Vérifier que le nom en cache correspond au nom recherché
      const cachedNomNorm = normalizeForMatch(cached.nom || "");
      const nameMatch = cachedNomNorm.includes(pNomNorm.slice(0,4)) ||
                        pNomNorm.includes(cachedNomNorm.slice(0,4));
      if (nameMatch) {
        log("info", null, "ENRICH_CACHE_HIT", { nom: cached.nom, siret: cached.siret });
        gouvData = { ...e, ...cached };
        dirigeant = extractDirigeantFromSearch(gouvData) || dirigeant;
      } else {
        log("info", null, "ENRICH_CACHE_REJECTED", { cachedNom: cached.nom, searched: e.nom });
        // Supprimer le cache corrompu
        try { await kv().delete(`enseigne_cache:${cacheKey}`); } catch (_) {}
      }
    }

    if (!gouvData.siret) {
      // ── Stratégie 1 : nom + CP/ville (uniquement même département) ──
      if (!isFranchise) {
        const queries = new Set();
        if (e.nom && cp)    queries.add(`${e.nom} ${cp}`);
        if (e.nom && ville) queries.add(`${e.nom} ${ville}`);
        const nomSansS = e.nom.replace(/ies$/i, "ie").replace(/es$/i, "e");
        if (nomSansS !== e.nom && cp)    queries.add(`${nomSansS} ${cp}`);
        if (nomSansS !== e.nom && ville) queries.add(`${nomSansS} ${ville}`);
        queries.add(e.nom);

        for (const q of queries) {
          try {
            const results = await searchCompanyCached(q.trim());
            const filtered = pDept
              ? (results || []).filter(r => (r.codePostal || "").slice(0, 2) === pDept)
              : (results || []);
            const best = filtered.find(r => {
              const enseignes = [r.nom_commercial, r.enseigne1, r.enseigne2, r.sigle]
                .filter(Boolean).map(normalizeForMatch);
              if (enseignes.some(en => en && (en.includes(pNomNorm) || pNomNorm.includes(en)))) return true;
              return normalizeForMatch(r.nom || r.nom_raison_sociale || "").includes(pNomNorm);
            });
            if (best) {
              // Vérifier cohérence NAF avant d'accepter
              const bestNafNorm = (best.naf || best.activite_principale || "").replace(".", "");
              const pNomLow2 = e.nom.toLowerCase();
              const isMenuiserie2 = pNomLow2.includes("menuis") || pNomLow2.includes("fenetre") || pNomLow2.includes("bois");
              const isAlim2 = pNomLow2.includes("march") || pNomLow2.includes("super") || pNomLow2.includes("hyper");
              const badNaf = (isMenuiserie2 && ["4673","3511","6810","9499"].some(n => bestNafNorm.startsWith(n))) ||
                             (isAlim2 && ["4332","3511","4673"].some(n => bestNafNorm.startsWith(n)));
              if (!badNaf) {
                log("info", null, "ENRICH_NAME_MATCH", { nom: best.nom || best.nom_raison_sociale, siret: best.siret, cp: best.codePostal });
                gouvData = { ...e, ...best };
                dirigeant = extractDirigeantFromSearch(best) || dirigeant;
                break;
              } else {
                log("info", null, "ENRICH_NAME_MATCH_REJECTED", { nom: best.nom || best.nom_raison_sociale, naf: best.naf, reason: "NAF incohérent" });
              }
            }
          } catch (_) {}
        }
      }

      // ── Stratégie 1b : Brave Search (franchises uniquement) ─────
      const gouvDept   = (gouvData.codePostal || "").slice(0, 2);
      const wrongDept  = pDept && gouvDept && gouvDept !== pDept;
      const needsBrave = isFranchise && (!gouvData.siret || wrongDept);
      log("info", null, "BRAVE_CHECK", { needsBrave, isFranchise, wrongDept, hasSiret: !!gouvData.siret, gouvDept, pDept });
      if (needsBrave) {
        try {
          const brave = await findJuridicalFromEnseigne(e.nom, ville, cp, phone);
          log("info", null, "BRAVE_JURIDIQUE_RAW", { siret: brave.siret||"", siren: brave.siren||"", ville, cp });

          // Priorité SIRET complet → recherche directe établissement
          if (brave.siret && brave.siret.length === 14) {
            // Chercher d'abord par SIRET dans matching_etablissements
            const etab = await searchEtablissementBySiret(brave.siret);
            if (etab && etab.nom) {
              log("info", null, "ENRICH_BRAVE_SIRET_MATCH", { nom: etab.nom, siret: brave.siret, cp: etab.codePostal });
              gouvData = { ...e, ...etab };
              dirigeant = extractDirigeantFromSearch(etab) || dirigeant;
            } else {
              // Fallback : chercher par SIREN et prendre établissement local
              const results = await searchCompanyCached(brave.siret.slice(0,9));
              const best = cp
                ? (results || []).find(r => (r.codePostal||"").slice(0,2) === pDept) || results?.[0]
                : results?.[0];
              if (best) {
                log("info", null, "ENRICH_BRAVE_SIREN_FALLBACK", { nom: best.nom, siret: brave.siret, cp: best.codePostal });
                gouvData = { ...e, ...best, siret: brave.siret };
                dirigeant = extractDirigeantFromSearch(best) || dirigeant;
              }
            }
          }
          // Sinon SIREN → valider par ville/CP
          else if (brave.siren && brave.siren.length === 9) {
            const results = await searchCompanyCached(brave.siren);
            const best = cp
              ? (results || []).find(r => (r.codePostal||"").slice(0,2) === pDept) || results?.[0]
              : results?.[0];
            if (best) {
              log("info", null, "ENRICH_BRAVE_SIREN_MATCH", { nom: best.nom, siren: brave.siren, cp: best.codePostal });
              gouvData = { ...e, ...best };
              dirigeant = extractDirigeantFromSearch(best) || dirigeant;
            }
          }
        } catch (_) {}
      }

      // ── Stratégie 2 : téléphone (preuve absolue) ─────────────────
      if (!gouvData.siret && phone) {
        try {
          const results = await searchCompanyCached(phone);
          const best = pDept
            ? (results || []).find(r => (r.codePostal || "").slice(0, 2) === pDept)
            : results?.[0];
          if (best) {
            log("info", null, "ENRICH_PHONE_MATCH", { nom: best.nom || best.nom_raison_sociale, siret: best.siret });
            gouvData = { ...e, ...best };
            dirigeant = extractDirigeantFromSearch(best) || dirigeant;
          }
        } catch (_) {}
      }

      // ── Stratégie 3 : recherche par rue + CP avec scoring NAF ────
      if (!gouvData.siret && cp) {
        try {
            const rueNom = adresse
              ? adresse.replace(/,.*$/, "").replace(/^\d+\s*/, "")
                .replace(/^(?:rue|avenue|boulevard|impasse|chemin|allée)\s+/i, "")
                .trim().split(/\s+/).slice(0, 3).join(" ")
              : ""; // Pas d'adresse pour les façades → chercher par nom + CP

            // Si pas de rue : utiliser le nom de la société comme query
            const searchQuery = rueNom.length >= 4 ? rueNom : normalizeSpaces(e.nom || "").slice(0, 30);

            const searchUrl = searchQuery.length >= 3
              ? "https://recherche-entreprises.api.gouv.fr/search?q=" + encodeURIComponent(searchQuery) + (adresse ? "&code_postal=" + encodeURIComponent(cp) : "&departement=" + encodeURIComponent(cp.slice(0,2))) + "&limite=50"
              : "https://recherche-entreprises.api.gouv.fr/search?code_postal=" + encodeURIComponent(cp) + "&limite=50";

            const nearR = await safeFetchJson(searchUrl, {}, { timeoutMs: 5000, retries: 0, label: "GOUV_ADDR_SEARCH" });

            log("info", null, "ENRICH_ADDR_SEARCH", {
              rueNom: searchQuery, cp, url: searchUrl.slice(0, 100),
              count: nearR.data?.results?.length || 0,
              hasBG2M: (nearR.data?.results || []).some(r => normalizeForMatch(r.nom_raison_sociale || "").includes("bg2m")),
            });

            if (nearR.ok && nearR.data?.results?.length) {
                const pNomLow = e.nom.toLowerCase();
                const isMenuiserie = pNomLow.includes("menuis") || pNomLow.includes("fenetre") || pNomLow.includes("bois") || pNomLow.includes("pvc");
                const isAlim = pNomLow.includes("march") || pNomLow.includes("super") || pNomLow.includes("hyper") || pNomLow.includes("casino") || pNomLow.includes("leclerc");
                const isAuto = pNomLow.includes("auto") || pNomLow.includes("garage") || pNomLow.includes("renault") || pNomLow.includes("peugeot");
                const isBrico = pNomLow.includes("brico") || pNomLow.includes("castorama") || pNomLow.includes("leroy");

                const scored = nearR.data.results.map(c => {
                  let score = 0;
                  const siege = c.siege || {};
                  const cCp   = (siege.code_postal || "").slice(0, 5);
                  const cDept = cCp.slice(0, 2);
                  const cAdresse = (siege.adresse || "").toUpperCase();
                  const cNaf = c.activite_principale || "";
                  const cEnseignes = [c.nom_commercial, c.enseigne1, c.enseigne2, c.sigle]
                    .filter(Boolean).map(normalizeForMatch);

                  // Coordonnées GPS de la recherche
                  const lat = e._places_lat || null;
                  const lon = e._places_lon || null;

                  // ── SCORING ENQUÊTEUR ──────────────────────────────
                  // PREUVES QUASI-ABSOLUES
                  // Téléphone identique → +120
                  const cPhone = normalizePhone(siege.telephone || "");
                  if (phone && cPhone && phone === cPhone) score += 120;

                  // Même domaine web → +90
                  const eDomain = (e._places_website || "").replace(/^https?:\/\/(www\.)?/, "").split("/")[0].toLowerCase();
                  const cWebsite = (c.site_internet || siege.site_internet || "").replace(/^https?:\/\/(www\.)?/, "").split("/")[0].toLowerCase();
                  if (eDomain && cWebsite && eDomain === cWebsite) score += 90;

                  // PREUVES GÉOGRAPHIQUES
                  // GPS distance
                  const cLat = siege.latitude  || null;
                  const cLon = siege.longitude || null;
                  if (lat && lon && cLat && cLon) {
                    const dLat = (lat - cLat) * 111000;
                    const dLon = (lon - cLon) * 111000 * Math.cos(lat * Math.PI / 180);
                    const dist = Math.sqrt(dLat*dLat + dLon*dLon);
                    if      (dist <=   5) score += 70;
                    else if (dist <=  20) score += 50;
                    else if (dist <=  50) score += 35;
                    else if (dist <= 200) score += 15;
                    else                  score -= 20;
                  }

                  // Département identique → +20 / différent → -80
                  if (cDept === pDept) score += 20;
                  else score -= 80;

                  // CP identique → +15, même dept → +10
                  if (cCp === cp) score += 15;
                  else if (cCp.slice(0,2) === cp.slice(0,2)) score += 10;

                  // Distance numéro de rue
                  const pNum = parseInt(adresse.match(/^\d+/)?.[0] || "0");
                  const cNum = parseInt(cAdresse.match(/^\d+/)?.[0] || "0");
                  if (pNum > 0 && cNum > 0) {
                    const diff = Math.abs(pNum - cNum);
                    if      (diff === 0)   score += 30;
                    else if (diff <= 10)   score += 20;
                    else if (diff <= 30)   score += 12;
                    else if (diff <= 100)  score += 5;
                    else                   score -= 10;
                  }
                  // Même nom de rue → +10 / différente → -5
                  const pRue = adresse.toUpperCase().replace(/^\d+\s*/, "").slice(0, 12);
                  if (pRue.length > 4 && cAdresse.includes(pRue)) score += 10;
                  else if (pRue.length > 4 && pNum > 0) score -= 5;

                  // PREUVES MÉTIER
                  // NAF cohérent → +40 / contradictoire → -120
                  const cNafNorm = cNaf.replace(".", "");
                  const menuNafs  = ["4332","4333","4391","4399"];
                  const alimentNafs = ["4711","4721","4722","4723","4724","4729"];
                  const autoNafs  = ["4511","4519","4520","4532","4540"];
                  const bricoNafs = ["4752","4753","4759"];
                  const dechetNafs = ["3811","3812","3821","3822","3832"];
                  const plastiNafs = ["4674","2221","2229","4669"];

                  // Bonus NAF cohérent
                  if (isMenuiserie && menuNafs.some(n => cNafNorm.startsWith(n)))   score += 40;
                  if (isAlim      && alimentNafs.some(n => cNafNorm.startsWith(n))) score += 40;
                  if (isAuto      && autoNafs.some(n => cNafNorm.startsWith(n)))    score += 40;
                  if (isBrico     && bricoNafs.some(n => cNafNorm.startsWith(n)))   score += 40;

                  // Malus NAF contradictoire → -120
                  const nomLow = (e.nom || "").toLowerCase();
                  const isDéchet  = nomLow.includes("déchet") || nomLow.includes("dechet") || nomLow.includes("environnement");
                  const isPlasti  = nomLow.includes("plastic") || nomLow.includes("plastique") || nomLow.includes("caoutchouc");
                  if (isMenuiserie && ["3511","4673","6810","6820","9329"].some(n => cNafNorm.startsWith(n))) score -= 120;
                  if (isAlim      && ["4332","3511","6810","9329"].some(n => cNafNorm.startsWith(n)))        score -= 120;
                  if (isDéchet    && dechetNafs.some(n => cNafNorm.startsWith(n)))  score += 40;
                  if (isPlasti    && plastiNafs.some(n => cNafNorm.startsWith(n)))  score += 40;
                  if (isDéchet    && ["5610","5630","9329"].some(n => cNafNorm.startsWith(n)))              score -= 120;
                  // Carrosserie/auto/pneumatiques ne peut pas être plastiques/déchets
                  if (isPlasti    && ["4520","4511","4519","4531","4532","4540","9329","5610"].some(n => cNafNorm.startsWith(n))) score -= 120;
                  if (isDéchet    && ["4520","4511","4531","5610","9329"].some(n => cNafNorm.startsWith(n))) score -= 120;

                  // PREUVES ENSEIGNE/LEXICALES (faibles)
                  // Enseigne déclarée → +30
                  if (cEnseignes.some(en => en && (en.includes(pNomNorm) || pNomNorm.includes(en)))) score += 30;
                  // Nom proche → +15 max
                  else if (normalizeForMatch(c.nom_raison_sociale || "").includes(pNomNorm)) score += 15;

                  // MULTIPLICATEUR CONVERGENCE
                  // Si téléphone + département + NAF cohérent → ×1.5
                  const hasPhone = phone && cPhone && phone === cPhone;
                  const hasDept  = cDept === pDept;
                  const hasNaf   = (isMenuiserie && menuNafs.some(n => cNafNorm.startsWith(n))) ||
                                   (isAlim && alimentNafs.some(n => cNafNorm.startsWith(n))) ||
                                   (isDéchet && dechetNafs.some(n => cNafNorm.startsWith(n))) ||
                                   (isPlasti && plastiNafs.some(n => cNafNorm.startsWith(n)));
                  if (hasPhone && hasDept) score = Math.round(score * 1.3);
                  if (hasPhone && hasNaf)  score = Math.round(score * 1.2);

                  return { candidat: c, score };
                }).sort((a, b) => b.score - a.score);

                log("info", null, "ENRICH_ADDR_SCORING", {
                  rueNom, cp,
                  top3: scored.slice(0,3).map(s => ({
                    nom: s.candidat.nom_raison_sociale,
                    score: s.score,
                    naf: s.candidat.activite_principale,
                    cp: s.candidat.siege?.code_postal,
                    adr: s.candidat.siege?.adresse || "",
                    enseigne: s.candidat.enseigne1 || s.candidat.nom_commercial || "",
                  }))
                });

                // Seuil 80 — scoring renforcé exige plus de confiance
                const best = scored[0];
                if (best && best.score >= 80) {
                  const siege = best.candidat.siege || {};
                  log("info", null, "ENRICH_ADDR_MATCH", { nom: best.candidat.nom_raison_sociale, score: best.score, siret: siege.siret });
                  const foundName3 = best.candidat.nom_raison_sociale || best.candidat.nom || e.nom;
                  gouvData = {
                    ...e,
                    nom:              foundName3,
                    nom_commercial:   e.nom_commercial || (normalizeForMatch(foundName3) !== normalizeForMatch(e.nom) ? e.nom : ""),
                    siret:            siege.siret || "",
                    naf:              best.candidat.activite_principale || "",
                    adresse:          siege.adresse || e.adresse,
                    codePostal:       siege.code_postal || cp,
                    ville:            siege.libelle_commune || ville,
                    effectif:         best.candidat.tranche_effectif_salarie || "",
                    date_creation:    best.candidat.date_creation || "",
                    dirigeants:       best.candidat.dirigeants || [],
                    secteur_activite: best.candidat.libelle_activite_principale || "",
                  };
                  dirigeant = extractDirigeantFromSearch(gouvData) || dirigeant;

                  // Mémoriser dans le cache terrain
                  await setEnseigneCache(cacheKey, {
                    nom: gouvData.nom, siret: gouvData.siret,
                    naf: gouvData.naf, ville: gouvData.ville,
                    codePostal: gouvData.codePostal, dirigeants: gouvData.dirigeants || [],
                  });
                } else if (!gouvData.siret && !isFranchise) {
                  // Score insuffisant ET pas franchise → Brave en dernier recours
                  log("info", null, "ENRICH_BRAVE_LAST_RESORT", { nom: e.nom, topScore: best?.score || 0 });
                  try {
                    const brave = await findJuridicalFromEnseigne(e.nom, ville, cp, phone);
                    // Cas 1 : Brave trouve un SIRET directement
                    if (brave.siret) {
                      const etab = await searchEtablissementBySiret(brave.siret);
                      if (etab && (etab.codePostal || "").slice(0,2) === cp.slice(0,2)) {
                        gouvData = { ...e, ...etab };
                        dirigeant = extractDirigeantFromSearch(gouvData) || dirigeant;
                        log("info", null, "ENRICH_BRAVE_LAST_RESORT_OK", { siret: brave.siret, nom: etab.nom });
                        await setEnseigneCache(cacheKey, {
                          nom: gouvData.nom, siret: gouvData.siret,
                          naf: gouvData.naf, ville: gouvData.ville,
                          codePostal: gouvData.codePostal, dirigeants: gouvData.dirigeants || [],
                        });
                      }
                    // Cas 2 : Brave trouve seulement un SIREN → chercher établissement local via API Gouv
                    } else if (brave.siren) {
                      try {
                        const sirenUrl = `https://recherche-entreprises.api.gouv.fr/search?q=${brave.siren}&limite=20`;
                        const sirenR = await safeFetchJson(sirenUrl, {}, { timeoutMs: 5000, retries: 0, label: "GOUV_SIREN" });
                        const candidates = sirenR.data?.results || [];
                        const local = candidates.find(r => {
                          const rCp = (r.siege?.code_postal || r.codePostal || "");
                          return rCp.slice(0,2) === cp.slice(0,2);
                        });
                        if (local) {
                          const siege = local.siege || {};
                          gouvData = {
                            ...e,
                            nom:              local.nom_raison_sociale || local.nom || e.nom,
                            siret:            siege.siret || local.siret || "",
                            naf:              local.activite_principale || "",
                            adresse:          siege.adresse || e.adresse,
                            codePostal:       siege.code_postal || cp,
                            ville:            siege.libelle_commune || ville,
                            effectif:         local.tranche_effectif_salarie || "",
                            date_creation:    local.date_creation || "",
                            dirigeants:       local.dirigeants || [],
                            secteur_activite: local.libelle_activite_principale || "",
                          };
                          dirigeant = extractDirigeantFromSearch(gouvData) || dirigeant;
                          log("info", null, "ENRICH_BRAVE_SIREN_LOCAL_OK", { siren: brave.siren, siret: gouvData.siret, nom: gouvData.nom });
                          await setEnseigneCache(cacheKey, {
                            nom: gouvData.nom, siret: gouvData.siret,
                            naf: gouvData.naf, ville: gouvData.ville,
                            codePostal: gouvData.codePostal, dirigeants: gouvData.dirigeants || [],
                          });
                        }
                      } catch (_) {}
                    }
                  } catch (err) {
                    log("warn", null, "ENRICH_BRAVE_LAST_RESORT_ERR", { err: err?.message });
                  }
                }
              }
          } catch (_) {}
      }
    }
  }

  const place = await googlePlacesEnrichCached(gouvData.nom || e.nom, gouvData.ville || e.ville);
  // CSE remplacé par Brave Search
  const cse = {};

  let email = place.website ? await scrapeEmailCached(place.website) : "";
  if (!email) email = cse.email || "";
  if (!email && place.website) {
    const guessed = await guessEmailFromDomain(place.website).catch(() => "");
    if (guessed) { email = `~${guessed}`; }  // ~ = email deviné non confirmé
  }
  // Brave Search pour email si toujours vide
  if (!email && (gouvData.nom || e.nom)) {
    try {
      const braveEmail = await searchViaBrave(`"${gouvData.nom || e.nom}" "${gouvData.ville || e.ville}" "@"`);
      if (braveEmail.email) email = braveEmail.email;
    } catch (_) {}
  }

  let phone = normalizePhone(place.phone || e._places_phone || "");
  if (!phone) phone = normalizePhone(cse.phone || "");

  const website = place.website || cse.website || e._places_website || "";
  const finalNaf = normalizeNAF(gouvData.naf || e.naf || "");

  return ensureContactFallback({
    name:             normalizeSpaces(gouvData.nom || e.nom || ""),
    siret:            validateSIRET(String(gouvData.siret || e.siret || "").replace(/\D/g, "")),
    address:          stripPostalCityFromAddress(gouvData.adresse || e.adresse || "", gouvData.codePostal || e.codePostal || "", gouvData.ville || e.ville || ""),
    postal_code:      normalizeSpaces(gouvData.codePostal || e.codePostal || "").replace(/\s+/g, ""),
    city:             normalizeSpaces(gouvData.ville || e.ville || ""),
    naf:              finalNaf,
    secteur_activite: gouvData.secteur_activite || e.secteur_activite || nafLabel(finalNaf),
    effectif:         gouvData.effectif || e.effectif || "",
    date_creation:    gouvData.date_creation || e.date_creation || "",
    capital:          gouvData.capital || e.capital || "",
    dirigeant,
    interlocuteur: "", contact_civility: "", contact_firstname: "", contact_lastname: "",
    website: website,
    phone:   phone,
    phone2:  "",
    email:   email || "",
    resume: "", commande: "", qualification: "",
    card_photo_file_id: "", card_photo_url: "", facade_photo_file_id: "", facade_photo_url: "",
  });
}


function normalizeFlexibleSearchQuery(query) {
  return normalizeSpaces(query).replace(/[;,]+/g, " ").replace(/\s+/g, " ").trim();
}

function tokenizeFlexibleQuery(query) {
  return normalizeFlexibleSearchQuery(query).split(" ").map((x) => x.trim()).filter(Boolean);
}

function detectLikelyPostalCodeToken(token) {
  return /^\d{5}$/.test(String(token || "").trim());
}

function detectLikelyPhoneToken(token) {
  return !!normalizePhone(token || "");
}

function buildFlexibleSearchCandidates(query) {
  const clean = normalizeFlexibleSearchQuery(query);
  if (!clean) return [];
  const tokens = tokenizeFlexibleQuery(clean);
  const phones      = tokens.filter(detectLikelyPhoneToken).map((t) => normalizePhone(t));
  const postalCodes = tokens.filter(detectLikelyPostalCodeToken);
  const words       = tokens.filter((t) => !detectLikelyPhoneToken(t) && !detectLikelyPostalCodeToken(t));

  const candidates = [];
  const seen = new Set();
  function push(q) {
    const s = normalizeFlexibleSearchQuery(q);
    if (!s) return;
    const k = s.toLowerCase();
    if (seen.has(k)) return;
    seen.add(k); candidates.push(s);
  }

  push(clean);
  if (words.length >= 2) push(words.join(" "));
  if (words.length >= 1 && postalCodes.length >= 1) push(`${words.join(" ")} ${postalCodes[0]}`);
  if (words.length >= 1 && phones.length >= 1) push(`${words.join(" ")} ${phones[0]}`);
  if (words.length >= 2) push(words.slice(0, 2).join(" "));
  if (words.length >= 3) { push(words.slice(0, 3).join(" ")); push(words.slice(-2).join(" ")); }

  return candidates.slice(0, 6);
}

async function enrichDraftFromCurrentData(draft) {
  const d = normalizeDraft(draft || {});
  log("info", null, "ENRICH_DRAFT_ENTRY", { name: d.name, city: d.city, cp: d.postal_code, siret: d.siret, isFacade: !!d.facade_photo_file_id, active_visit_id: draft?.active_visit_id });
  if (!d.name && !d.city) return d;

  // Si on a déjà un SIRET valide dans le bon département → ne pas ré-enrichir
  // EXCEPTION : façades (source_mode=photo) car l'adresse GPS peut être fausse
  const dDept = (d.postal_code || "").slice(0, 2);
  // Façade = a une photo de façade → toujours ré-enrichir car adresse GPS peut être fausse
  const isFacade = !!(d.facade_photo_file_id);
  if (d.siret && d.siret.length >= 9 && dDept && !isFacade) {
    log("info", null, "ENRICH_SKIP_HAS_SIRET", { siret: d.siret, dept: dDept });
    // Compléter seulement les données manquantes
    if (!d.email && (d.name || d.website)) {
      try {
        const brave = await searchViaBrave(`"${d.name}" "${d.city}" "@"`);
        if (brave.email) d.email = brave.email;
      } catch (_) {}
    }
    return d;
  }

  // Construire un objet Places-like sans SIRET
  const enrichInput = {
    nom:             d.name,
    codePostal:      d.postal_code || "",
    ville:           d.city || "",
    adresse:         d.facade_photo_file_id ? "" : (d.address || ""),
    _places_phone:   normalizePhone(d.phone || ""),
    _places_website: d.website || "",
    _places_lat:     d.gps_lat || null,
    _places_lon:     d.gps_lon || null,
  };

  const enriched = await enrichCompany(enrichInput);
  return applyCompanyToDraft(d, enriched);
}

function applyCompanyToDraft(draft, enriched) {
  const d = normalizeDraft(draft || {});
  const e = normalizeDraft(enriched || {});
  for (const k of [
    "name", "address", "postal_code", "city", "siret", "naf",
    "secteur_activite", "effectif", "date_creation", "capital",
    "dirigeant", "website",
  ]) {
    if (e[k]) d[k] = e[k];
  }
  // phone : enrichissement → fixe société. Si OCR a trouvé un portable → le préserver dans phone2
  if (e.phone) {
    if (!d.phone2 && d.phone && d.phone !== e.phone) d.phone2 = d.phone; // garder l'ancien en portable
    d.phone = e.phone;
  }
  // phone2 : ne jamais écraser si déjà rempli par OCR (portable interlocuteur)
  if (e.phone2 && !d.phone2) d.phone2 = e.phone2;
  // email : enrichissement → email générique société
  if (e.email) {
    if (d.email && d.email !== e.email) d.email_card = d.email_card || d.email; // sauvegarder l'email OCR
    d.email = e.email;
  }
  // interlocuteur, titre : jamais écraser si déjà rempli par OCR carte de visite
  if (!d.interlocuteur && e.interlocuteur) d.interlocuteur = e.interlocuteur;
  if (!d.titre && e.titre) d.titre = e.titre;
  if (e.nom_commercial && !d.nom_commercial) d.nom_commercial = e.nom_commercial;
  if (!d.contact_firstname && e.contact_firstname) d.contact_firstname = e.contact_firstname;
  if (!d.contact_lastname && e.contact_lastname) d.contact_lastname = e.contact_lastname;
  // Validation adresse BAN en arrière-plan (non bloquant)
  return normalizeDraft(ensureContactFallback(d));
}

// Vérifier si la fiche est assez complète pour auto-save
function isFicheCompleteForAutoSave(d) {
  return !!(
    d.name?.trim() &&
    (d.city?.trim() || d.postal_code?.trim()) &&
    (normalizePhone(d.phone) || normalizePhone(d.phone2) || normalizeEmail(d.email) || d.website?.trim())
  );
}

function buildAutoSaveSummary(d) {
  const lines = [
    `✅ <b>${escapeHtml(d.name)}</b> — ${escapeHtml(d.city || d.postal_code)}`,
  ];
  if (d.nom_commercial) lines.push(`🏪 Enseigne : ${escapeHtml(d.nom_commercial)}`);
  if (d.phone) lines.push(`📞 ${escapeHtml(formatPhone(d.phone))}`);
  if (d.email) lines.push(`✉️ ${escapeHtml(d.email)}`);
  if (d.website) lines.push(`🌐 ${escapeHtml(d.website)}`);
  if (d.siret) lines.push(`🆔 ${escapeHtml(d.siret)}`);
  if (d.dirigeant) lines.push(`👔 ${escapeHtml(d.dirigeant)}`);
  if (d.effectif) lines.push(`👥 ${escapeHtml(d.effectif)}`);
  return lines.join("\n");
}


// =====================================================
// SEARCH
// =====================================================

// Recherche via Google Places quand API Gouv ne trouve rien
async function searchViaPlaces(query) {
  const apiKey = String(globalThis.GOOGLE_PLACES_API_KEY || "").trim();
  if (!apiKey) return [];
  try {
    const r = await safeFetchJson(
      "https://places.googleapis.com/v1/places:searchText",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-Goog-Api-Key": apiKey,
          "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,places.websiteUri,places.location",
        },
        body: JSON.stringify({
          textQuery: query,
          maxResultCount: 5,
          languageCode: "fr",
          regionCode: "FR",
        }),
      },
      { timeoutMs: 8000, retries: 0, label: "PLACES_TEXT_SEARCH" }
    );
    if (!r.ok || !r.data) return [];
    const places = r.data.places || [];
    log("info", null, "PLACES_SEARCH_RAW", {
      query,
      count: places.length,
      first: places[0] ? {
        name: places[0].displayName?.text,
        addr: places[0].formattedAddress,
        phone: places[0].nationalPhoneNumber,
      } : null,
    });
    return places.map(p => {
      const addr = p.formattedAddress || "";
      // Adresse FR typique: "1450 Rue Aristide Berges, 38430 Moirans, France"
      const cpMatch = addr.match(/\b(\d{5})\b/);
      const cp = cpMatch ? cpMatch[1] : "";
      // Extraire ville : texte entre le CP et la virgule/fin
      let city = "";
      if (cp) {
        const afterCp = addr.slice(addr.indexOf(cp) + cp.length).replace(/^[\s,]+/, "");
        city = normalizeSpaces(afterCp.split(",")[0].replace(/France$/i, "").trim());
      }
      if (!city) {
        // Fallback : dernier segment avant "France"
        const parts = addr.split(",").map(s => s.trim()).filter(s => s && !/^France$/i.test(s));
        city = parts[parts.length - 1] || "";
        city = city.replace(/^\d{5}\s*/, "").trim();
      }
      return {
        nom: p.displayName?.text || "",
        siret: "",
        naf: "",
        adresse: stripPostalCityFromAddress(addr, cp, city),
        codePostal: cp,
        ville: city,
        secteur_activite: "",
        effectif: "",
        date_creation: "",
        capital: "",
        dirigeants: [],
        _source: "places",
        _places_phone: normalizePhone(p.nationalPhoneNumber || ""),
        _places_website: p.websiteUri || "",
        _places_lat: p.location?.latitude || null,
        _places_lon: p.location?.longitude || null,
      };
    }).filter(p => p.nom);
  } catch (e) {
    log("warn", null, "PLACES_SEARCH_ERROR", { err: e?.message, query });
    return [];
  }
}

async function handleSearch(chatId, s, query) {
  const q = normalizeFlexibleSearchQuery(query);
  s.last_search_query = q;
  await setSession(chatId, s);

  const qNorm = normalizeForMatch(q);
  const isFranchiseQuery = [...KNOWN_BRANDS].some(b => qNorm.includes(normalizeForMatch(b)));

  const candidates = buildFlexibleSearchCandidates(q);
  let list = [];
  let lastErr = null;

  // Pour les enseignes connues → Places en priorité (évite les résultats nationaux)
  if (isFranchiseQuery) {
    try {
      const placeResults = await searchViaPlaces(q);
      if (placeResults?.length) {
        list = placeResults;
        log("info", chatId, "SEARCH_FRANCHISE_PLACES", { query: q, count: list.length });
      }
    } catch (_) {}
  }

  if (!list.length) {
    for (const candidate of candidates) {
      try {
        const partial = await searchCompanyCached(candidate);
        if (partial?.length) { list = partial; break; }
      } catch (e) {
        lastErr = e;
      }
    }
  }

  // Filtrer par département si CP détecté dans la query
  const cpInQuery = q.match(/\b(\d{5})\b/)?.[1] || "";
  const deptInQuery = cpInQuery.slice(0, 2);
  if (deptInQuery && list.length > 0) {
    const localList = list.filter(r => (r.codePostal || "").slice(0, 2) === deptInQuery);
    if (localList.length > 0) list = localList;
  }

  // Filtrer les résultats dont la ville ne correspond pas aux mots de la query
  if (list.length > 0) {
    const queryWords = q.split(/\s+/).filter(w => w.length >= 4 && !/^\d+$/.test(w));
    if (queryWords.length >= 2) {
      const localList = list.filter(r => {
        const rVille = normalizeForMatch(r.ville || "");
        const rAdresse = normalizeForMatch(r.adresse || "");
        return queryWords.some(w => {
          const wn = normalizeForMatch(w);
          return rVille.includes(wn) || rAdresse.includes(wn);
        });
      });
      if (localList.length > 0) {
        list = localList;
        log("info", chatId, "SEARCH_CITY_FILTER", { count: localList.length });
      } else {
        // Aucun résultat local → fallback Places
        try {
          const placeResults = await searchViaPlaces(q);
          if (placeResults?.length) {
            list = placeResults;
            log("info", chatId, "SEARCH_NO_LOCAL_PLACES", { query: q, count: placeResults.length });
          }
        } catch (_) {}
      }
    }
  }

  // Fallback 1 — Google Places si Gouv ne trouve rien
  if (!list.length) {
    try {
      const placeResults = await searchViaPlaces(q);
      if (placeResults?.length) {
        list = placeResults;
        log("info", chatId, "SEARCH_PLACES_FALLBACK", { query: q, count: list.length });
      }
    } catch (_) {}
  }

  // Fallback 2 — Relancer Gouv sans la ville (ex: "4s voreppe" → "4s")
  if (!list.length) {
    const tokens = tokenizeFlexibleQuery(q);
    const nameOnly = tokens.filter(t => !detectLikelyPostalCodeToken(t) && t.length > 2).slice(0,2).join(" ");
    if (nameOnly && nameOnly !== q) {
      try {
        const results2 = await searchCompanyCached(nameOnly);
        if (results2?.length) {
          list = results2;
          log("info", chatId, "SEARCH_NAME_ONLY_FALLBACK", { query: nameOnly, count: list.length });
        }
      } catch (_) {}
    }
  }

  // Fallback 3 — Places sans la ville
  if (!list.length) {
    const nameOnly = tokenizeFlexibleQuery(q).filter(t => !detectLikelyPostalCodeToken(t) && t.length > 2).slice(0,2).join(" ");
    if (nameOnly && nameOnly !== q) {
      try {
        const placeResults2 = await searchViaPlaces(nameOnly);
        if (placeResults2?.length) {
          list = placeResults2;
          log("info", chatId, "SEARCH_PLACES_NAME_ONLY", { query: nameOnly, count: list.length });
        }
      } catch (_) {}
    }
  }

  if (!list.length && lastErr) {
    log("error", chatId, "SEARCH_FAILED", { err: lastErr?.message || String(lastErr), query: q });
    return sendMessage(chatId, "⚠️ API entreprises indisponible.\n\nTu peux réessayer ou passer en saisie manuelle.", failSafeKeyboard());
  }

  s.companyChoices = list;
  s.step = list.length ? "PICK" : "SEARCH";
  await setSession(chatId, s);

  if (!list.length) {
    const tokens = tokenizeFlexibleQuery(q);
    const tail = tokens.length ? tokens[tokens.length - 1] : "";
    s._manual_city = detectLikelyPostalCodeToken(tail) || detectLikelyPhoneToken(tail) ? "" : tail;
    await setSession(chatId, s);
    return sendMessage(chatId, `Aucun résultat pour "<b>${escapeHtml(q)}</b>".\n\nTu peux :`, failSafeKeyboard());
  }

  const maybeShort = q.split(/\s+/).length < 2;
  return sendMessage(
    chatId,
    `${maybeShort ? "⚠️ Astuce : ajoute une 2e information pour être plus précis.\n\n" : ""}Choisis l'entreprise :`,
    companyPickKeyboard(list)
  );
}


// =====================================================
// DÉDUPLICATION — ÉLARGIE 30 JOURS
// =====================================================

function scoreDuplicateCandidate(target, existing) {
  let score = 0;
  const tSiret = String(target?.siret || "").replace(/\D/g, "");
  const eSiret = String(existing?.siret || "").replace(/\D/g, "");
  if (tSiret && eSiret && tSiret === eSiret) score += 100;
  const tName = normalizeForMatch(target?.name || "");
  const eName = normalizeForMatch(existing?.name || "");
  if (tName && eName && tName === eName) score += 40;
  const tCity = normalizeForMatch(target?.city || "");
  const eCity = normalizeForMatch(existing?.city || "");
  if (tCity && eCity && tCity === eCity) score += 15;
  const tPostal = String(target?.postal_code || "").replace(/\s+/g, "");
  const ePostal = String(existing?.postal_code || "").replace(/\s+/g, "");
  if (tPostal && ePostal && tPostal === ePostal) score += 10;
  const tPhone = normalizePhone(target?.phone || "");
  const ePhone = normalizePhone(existing?.phone || "");
  if (tPhone && ePhone && tPhone === ePhone) score += 10;
  const tEmail = normalizeEmail(target?.email || "");
  const eEmail = normalizeEmail(existing?.email || "");
  if (tEmail && eEmail && tEmail === eEmail) score += 20;
  const tInter = normalizeForMatch(target?.interlocuteur || "");
  const eInter = normalizeForMatch(existing?.interlocuteur || "");
  if (tInter && eInter && tInter === eInter) score += 8;
  return score;
}

function buildFingerprint(r) {
  return [
    normalizeForMatch(r?.name || ""),
    normalizePhone(r?.phone || ""),
    normalizeEmail(r?.email || ""),
    normalizeForMatch(r?.city || ""),
  ].join("|");
}

async function checkDuplicateWindow(agency, siret, name, row = null) {
  const target = row || { siret, name };
  let best = { duplicate: false, reason: "", record: null, score: 0, fingerprint: "" };

  // Calcul du fingerprint pour déduplication rapide
  const fp = buildFingerprint(target);

  // Vérification rapide via index fingerprints (O(1))
  const allFingerprints = await Promise.all(
    Array.from({ length: DEDUP_WINDOW_DAYS }, (_, i) =>
      getDupFingerprints(ymdFROffset(i), agency)
    )
  );
  const fpExists = allFingerprints.flat().includes(fp);
  if (fpExists) best.fingerprint = fp;

  // Lecture des clés via index dédié (O(1) par jour)
  const allKeyArrays = await Promise.all(
    Array.from({ length: DEDUP_WINDOW_DAYS }, (_, i) =>
      getDayIndexedKeys(ymdFROffset(i), agency)
    )
  );
  const allKeys = allKeyArrays.flat();

  if (!allKeys.length) {
    return fpExists
      ? { duplicate: true, reason: "fingerprint", record: null, score: 50, fingerprint: fp }
      : best;
  }

  // Lecture parallèle de tous les prospects de la fenêtre
  const allRecords = await batchGet(allKeys);

  for (const p of allRecords) {
    if (!p) continue;
    const score = scoreDuplicateCandidate(target, p);
    if (score > best.score) {
      best = {
        duplicate: score >= 40 || fpExists,
        reason: score >= 100 ? "siret" : fpExists ? "fingerprint" : "match",
        record: p,
        score,
        fingerprint: fp,
      };
    }
  }

  return best.duplicate ? best : { duplicate: false, reason: "", record: null, score: 0, fingerprint: fp };
}


// =====================================================
// STORAGE SHAPE
// =====================================================

function finalizeRecordForStorage(s, date, chatId) {
  const d = normalizeDraft((s && s.draft) || {});
  return {
    created_at:    new Date().toISOString(),
    chatId:        String(chatId),
    date,
    agency:        s.agency,
    initials:      s.initials,

    visit_id:      d._visit_id || s.active_visit_id || "",
    visit_no:      sanitizeVisitNo(s.active_visit_no || 0),
    source_mode:   d.source_mode || s.mode || "company",

    name:             d.name || "",
    nom_commercial:   d.nom_commercial || "",
    address:          d.address || "",
    postal_code:      d.postal_code || "",
    city:             d.city || "",
    phone:            d.phone || "",
    phone2:           d.phone2 || "",
    email:            d.email || "",
    email_card:       d.email_card || "",     // Email personnel carte de visite
    siret:            d.siret || "",
    naf:              normalizeNAF(d.naf || ""),
    secteur_activite: d.secteur_activite || nafLabel(d.naf || ""),
    effectif:         d.effectif || "",
    date_creation:    d.date_creation || "",
    capital:          d.capital || "",
    website:          d.website || "",
    dirigeant:        d.dirigeant || "",
    interlocuteur:    d.interlocuteur || "",
    contact_civility: d.contact_civility || "",
    contact_firstname: d.contact_firstname || "",
    contact_lastname:  d.contact_lastname || "",
    titre:            d.titre || "",           // fonction / poste
    resume:           d.resume || "",
    commande:         d.commande || "",
    qualification:    d.qualification || "",

    card_photo_file_id:  d.card_photo_file_id || "",
    card_photo_url:      d.card_photo_url || "",
    facade_photo_file_id: d.facade_photo_file_id || "",
    facade_photo_url:    d.facade_photo_url || "",

    card_kv_key:        d.card_kv_key || "",
    linked_card_keys:   Array.isArray(d.linked_card_keys) ? d.linked_card_keys : [],
    linked_photo_keys:  Array.isArray(d.linked_photo_keys) ? d.linked_photo_keys : [],

    // GPS — depuis le draft ou la session courante
    gps_lat:  d.gps_lat || (s.current_geo?.lat || null),
    gps_lon:  d.gps_lon || (s.current_geo?.lon || null),

    _ocr_confidence: d._ocr_confidence || "",
  };
}


// =====================================================
// SAVE HELPERS
// =====================================================

async function saveCurrentVisitAsProspect(chatId, s) {
  if (!s?.agency || !s?.initials) return { ok: false, reason: "session" };

  const d = normalizeDraft(s.draft || {});

  // Validation minimum
  const errors = validateDraftMinimum(d);
  if (errors.length) return { ok: false, reason: "validation", errors };

  const date   = ymdFR();
  const record = finalizeRecordForStorage(s, date, chatId);

  // Déduplication 30 jours
  const duplicate = await checkDuplicateWindow(s.agency, record.siret, record.name, record);
  if (duplicate.duplicate && !s.draft?._edit_key) {
    s._pending_duplicate_record = record;
    await setSession(chatId, s);
    return { ok: false, reason: "duplicate", duplicate };
  }

  const editKey = s.draft?._edit_key;
  const key = editKey || `prospect:${date}:${s.agency}:${Date.now()}:${chatId}`;
  await kv().put(key, JSON.stringify(record), { expirationTtl: TTL_PROSPECT });
  await saveLastProspectKey(chatId, key);

  // Mise à jour des index KV (non bloquant)
  if (!editKey) {
    addDayIndex(date, s.agency, key).catch(() => {});
    const fp = duplicate.fingerprint || buildFingerprint(record);
    if (fp) addDupIndex(date, s.agency, fp).catch(() => {});
  }

  s.draft._source   = "";
  s.draft._edit_key = "";
  s._pending_duplicate_record = null;

  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);

  return { ok: true, key, record, edit: !!editKey };
}

async function saveCurrentVisitAsProspectForced(chatId, s) {
  if (!s?.agency || !s?.initials) return { ok: false, reason: "session" };
  const date   = ymdFR();
  const record = s._pending_duplicate_record || finalizeRecordForStorage(s, date, chatId);
  const editKey = s.draft?._edit_key;

  // Vérifier doublon si pas en mode édition
  if (!editKey && record.name) {
    const dup = await checkDuplicateWindow(s.agency, record.siret, record.name, record);
    if (dup.duplicate && dup.score >= 80) {
      log("info", chatId, "SAVE_FORCED_DEDUP_SKIP", { name: record.name, score: dup.score, reason: dup.reason });
      return { ok: false, reason: "duplicate", score: dup.score };
    }
  }
  const key = editKey || `prospect:${date}:${s.agency}:${Date.now()}:${chatId}`;
  await kv().put(key, JSON.stringify(record), { expirationTtl: TTL_PROSPECT });
  await saveLastProspectKey(chatId, key);

  // Alimentation des index (même logique que saveCurrentVisitAsProspect)
  if (!editKey) {
    addDayIndex(date, s.agency, key).catch(() => {});
    const fp = buildFingerprint(record);
    if (fp) addDupIndex(date, s.agency, fp).catch(() => {});
  }

  s.draft._source   = "";
  s.draft._edit_key = "";
  s._pending_duplicate_record = null;
  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);
  return { ok: true, key, record, edit: !!editKey };
}

function attachCardKeyToDraft(draft, cardKey) {
  const d = normalizeDraft(draft || {});
  const arr = Array.isArray(d.linked_card_keys) ? [...d.linked_card_keys] : [];
  if (cardKey && !arr.includes(cardKey)) arr.push(cardKey);
  d.linked_card_keys = arr;
  d.card_kv_key = d.card_kv_key || cardKey || "";
  return d;
}

function attachPhotoKeyToDraft(draft, photoKey) {
  const d = normalizeDraft(draft || {});
  const arr = Array.isArray(d.linked_photo_keys) ? [...d.linked_photo_keys] : [];
  if (photoKey && !arr.includes(photoKey)) arr.push(photoKey);
  d.linked_photo_keys = arr;
  return d;
}


// =====================================================
// PROSPECT FROM FACADE / CARD
// =====================================================

async function createProspectFromFacade(chatId, s, file_id, cityHint = "", comment = null) {
  await createNewActiveVisit(chatId, s, "photo");

  // ── Étape 1 : OCR facade ───────────────────────────
  let facade = {};
  try {
    const imgBytes = await withTimeout(getBytesFromFileId(file_id), 10_000, new Uint8Array());
    if (imgBytes.length) facade = await hybridExtractFacade(imgBytes, cityHint);
  } catch (_) {}

  const company = facade.company || "";
  const city    = facade.city    || cityHint || "";

  s.draft.name             = company;
  s.draft.city             = city;
  s.draft.phone            = normalizePhone(facade.phone    || "");
  s.draft.website          = facade.website   || "";
  s.draft.secteur_activite = facade.activity  || "";
  s.draft._ocr_confidence  = facade.confidence || "low";
  s.draft.facade_photo_file_id = file_id;
  s.draft.facade_photo_url     = await tgGetFileUrl(file_id).catch(() => "");
  if (comment) s.draft.resume = trunc(comment, 300);

  // GPS depuis la session courante si disponible
  const geoSnap = getFreshGeoOrNull(s);
  if (geoSnap) {
    s.draft.gps_lat = geoSnap.lat;
    s.draft.gps_lon = geoSnap.lon;
    // Reverse geocoding → ville et CP uniquement (PAS l'adresse — elle appartient à l'entreprise, pas au GPS)
    if (geoSnap.lat && geoSnap.lon) {
      const geo = await reverseGeocode(geoSnap.lat, geoSnap.lon).catch(() => null);
      if (geo) {
        if (!s.draft.postal_code && geo.postal_code) s.draft.postal_code = geo.postal_code;
        if (!s.draft.city && geo.city) s.draft.city = geo.city;
        // Ne jamais écraser s.draft.address avec l'adresse GPS
      }
    }
  }

  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);

  // ── Étape 2 : enrichissement via pipeline complet ─────────────
  if (company) {
    await sendMessage(chatId, `🔍 Enrichissement de <b>${escapeHtml(company)}</b>…`);
    try {
      const enrichInput = {
        nom:             company,
        codePostal:      s.draft.postal_code || "",
        ville:           s.draft.city || city || "",
        adresse:         "", // Jamais l'adresse GPS pour une façade
        _places_phone:   normalizePhone(s.draft.phone || ""),
        _places_website: s.draft.website || "",
        _places_lat:     s.draft.gps_lat || null,
        _places_lon:     s.draft.gps_lon || null,
      };
      const enriched = await enrichCompany(enrichInput);
      s.draft = applyCompanyToDraft(s.draft, enriched);
    } catch (err) {
      log("warn", chatId, "FACADE_ENRICH_ERROR", { err: err?.message });
    }
  }

  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);

  s.step = "VISIT_ACTIVE";
  s.mode = "company";
  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);

  return { draft: s.draft, ocrConf: facade.confidence || "low" };
}

async function createProspectFromCard(chatId, s, file_id, comment = null) {
  await createNewActiveVisit(chatId, s, "card");

  // ── Étape 1 : OCR carte ───────────────────────────
  let card = {};
  try {
    const imgBytes = await withTimeout(getBytesFromFileId(file_id), 10_000, new Uint8Array());
    if (imgBytes.length) card = await hybridExtractCard(imgBytes);
  } catch (_) {}

  // Fusion champs OCR dans le draft
  if (card.company) s.draft.name = card.company;
  if (card.name) {
    s.draft.interlocuteur = card.name;
    const sp = splitHumanName(card.name);
    s.draft.contact_firstname = sp.first || "";
    s.draft.contact_lastname  = sp.last  || card.name;
  }
  if (card.contact_civility) s.draft.contact_civility = card.contact_civility;
  if (card.email) {
    s.draft.email_card  = normalizeEmail(card.email);  // Email personnel carte
    s.draft.email       = normalizeEmail(card.email);  // Sera potentiellement remplacé par email générique
  }
  if (card.phone)       s.draft.phone       = normalizePhone(card.phone);
  if (card.phone2)      s.draft.phone2      = normalizePhone(card.phone2);
  if (card.website)     s.draft.website     = card.website;
  if (card.city)        s.draft.city        = card.city;
  if (card.postal_code) s.draft.postal_code = card.postal_code;
  if (card.address)     s.draft.address     = card.address;
  if (card.title)       s.draft.titre       = card.title;
  s.draft._ocr_confidence   = card._confidence || "low";
  s.draft.card_photo_file_id = file_id;
  s.draft.card_photo_url     = await tgGetFileUrl(file_id).catch(() => "");
  if (comment) s.draft.resume = trunc(comment, 300);

  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);

  // ── Étape 2 : enrichissement Gouv + Places ─────────
  const company = s.draft.name || "";
  const city    = s.draft.city || "";

  if (company) {
    await sendMessage(chatId, `🔍 Enrichissement de <b>${escapeHtml(company)}</b>…`);

    try {
      // Construction des queries Gouv — combinaison de tous les signaux de la carte
      const d = s.draft;
      const cp = d.postal_code || "";
      const emailDomain = (d.email || "").split("@")[1]?.replace(/\.(fr|com|net|org|eu)$/, "").replace(/[\-_.]/g, " ").trim() || "";
      const siteDomain  = (d.website || "").replace(/https?:\/\/(www\.)?/, "").split("/")[0].replace(/\.(fr|com|net|org|eu)$/, "").replace(/[\-_.]/g, " ").trim();

      const queries = new Set();
      // 1. Nom + ville (priorité)
      if (company && city)        queries.add(`${company} ${city}`);
      // 2. Nom + code postal (très précis)
      if (company && cp)          queries.add(`${company} ${cp}`);
      // 3. Domaine email + code postal
      if (emailDomain && cp)      queries.add(`${emailDomain} ${cp}`);
      // 4. Domaine site + code postal
      if (siteDomain && cp && siteDomain !== emailDomain) queries.add(`${siteDomain} ${cp}`);
      // 5. Domaine email + ville
      if (emailDomain && city)    queries.add(`${emailDomain} ${city}`);
      // 6. Téléphone seul (API Gouv supporte la recherche par tel)
      if (d.phone)                queries.add(normalizePhone(d.phone));
      // 7. Nom seul en dernier recours
      if (company)                queries.add(company);

      let best = null;
      for (const q of queries) {
        if (!q.trim()) continue;
        const results = await searchCompany(q.trim()).catch(() => []);
        if (results[0]) { best = results[0]; break; }
      }

      if (best) {
        if (!d.name          && best.nom)            d.name            = best.nom;
        if (best.siret)                              d.siret           = best.siret;
        if (best.naf)                                d.naf             = best.naf;
        if (!d.address       && best.adresse)        d.address         = best.adresse;
        if (!d.postal_code   && best.codePostal)     d.postal_code     = best.codePostal;
        if (!d.city          && best.ville)          d.city            = best.ville;
        if (best.effectif)                           d.effectif        = best.effectif;
        if (best.date_creation)                      d.date_creation   = best.date_creation;
        if (best.capital)                            d.capital         = best.capital;
        if (!d.secteur_activite && best.secteur_activite) d.secteur_activite = best.secteur_activite;
        if (best.dirigeants?.length && !d.dirigeant) {
          const dir = best.dirigeants[0];
          const nom = typeof dir === "string" ? dir
            : cleanPersonName([dir.prenom?.split(" ")[0], dir.nom || dir.nom_usage].filter(Boolean).join(" "));
          if (nom) d.dirigeant = normalizeSpaces(nom);
        }
      } else {
        log("info", chatId, "CARD_ENRICH_NO_RESULT", { queries: [...queries].slice(0, 4) });
      }

      // Google Places : complète ce qui manque encore
      if (!s.draft.phone || !s.draft.website) {
        const place = await googlePlacesEnrich(s.draft.name || company, s.draft.city || city).catch(() => ({}));
        if (place.phone   && !s.draft.phone)   s.draft.phone   = place.phone;
        if (place.website && !s.draft.website)  s.draft.website = place.website;
      }

      // Email : scraping site → fallback Google CSE
      if (!s.draft.email) {
        if (s.draft.website) {
          s.draft.email = await scrapeEmailCached(s.draft.website).catch(() => "");
        }
        if (!s.draft.email) s.draft.email = await searchEmailViaGoogle(s.draft.name || company, s.draft.city || city).catch(() => "");
        if (!s.draft.email && s.draft.website) s.draft.email = await guessEmailFromDomain(s.draft.website).catch(() => "");
      }

    } catch (e) {
      log("warn", chatId, "CARD_ENRICH_ERROR", { err: e?.message });
    }

    await persistActiveVisit(chatId, s);
    await setSession(chatId, s);

    // Feedback enrichissement complet
    const lines = [`✅ <b>${escapeHtml(s.draft.name || company)}</b>`];
    if (s.draft.interlocuteur)    lines.push(`👤 ${escapeHtml(s.draft.interlocuteur)}`);
    if (s.draft.contact_civility) lines.push(`🎩 ${escapeHtml(s.draft.contact_civility)}`);
    if (s.draft.siret)            lines.push(`🆔 SIRET : ${escapeHtml(s.draft.siret)}`);
    if (s.draft.secteur_activite) lines.push(`🏭 ${escapeHtml(s.draft.secteur_activite)}`);
    if (s.draft.effectif)         lines.push(`👥 ${escapeHtml(s.draft.effectif)}`);
    if (s.draft.dirigeant)        lines.push(`👔 ${escapeHtml(s.draft.dirigeant)}`);
    if (s.draft.phone)            lines.push(`📞 ${escapeHtml(formatPhone(s.draft.phone))}`);
    if (s.draft.email)            lines.push(`✉️ ${escapeHtml(s.draft.email)}`);
    if (s.draft.website)          lines.push(`🌐 ${escapeHtml(s.draft.website)}`);
    if (s.draft.city || s.draft.postal_code)
      lines.push(`📍 ${escapeHtml([s.draft.postal_code, s.draft.city].filter(Boolean).join(" "))}`);
    await sendMessage(chatId, lines.join("\n"));

    // Si enrichissement incomplet → proposer recherche ou voir fiche
    if (!s.draft.siret) {
      await sendMessage(chatId,
        `⚠️ <b>Entreprise non trouvée automatiquement.</b>\n\n` +
        `Tu peux :\n` +
        `• 🔎 Lancer une recherche pour compléter le SIRET, l'effectif, etc.\n` +
        `• ✏️ Voir et modifier la fiche telle quelle\n` +
        `• ✅ Sauvegarder directement sans enrichissement`,
        {
          inline_keyboard: [
            [{ text: "🔎 Rechercher l'entreprise", callback_data: "search_company" }],
            [{ text: "✏️ Voir / modifier la fiche", callback_data: "visit:resume_active" }],
            [{ text: "✅ Sauvegarder maintenant", callback_data: "save" }],
          ],
        }
      );
    } else {
      // SIRET trouvé → ouvrir la fiche
      s.step = "VISIT_ACTIVE";
      s.mode = "company";
      await persistActiveVisit(chatId, s);
      await setSession(chatId, s);
      return postOrReplaceForm(chatId, s);
    }
  }

  s.step = "VISIT_ACTIVE";
  s.mode = "company";
  await persistActiveVisit(chatId, s);
  await setSession(chatId, s);

  return { draft: s.draft, ocrConf: card._confidence || "low", card };
}

async function buildVisitsListMessage(chatId, s) {
  const items = await buildVisitsList(chatId, s);
  return { text: renderVisitsListText(items), keyboard: visitsListKeyboard(items), items };
}


// =====================================================
// RÉSUMÉ MANAGER AUTOMATIQUE
// =====================================================

async function sendManagerDailySummary(date, agencyStats) {
  const managerId = String(globalThis.MANAGER_CHAT_ID || "").trim();
  if (!managerId) return;

  const lines = [`📊 <b>Bilan prospection — ${escapeHtml(date)}</b>\n`];

  for (const [agency, data] of Object.entries(agencyStats)) {
    lines.push(`\n🏷️ <b>${escapeHtml(agency)}</b>`);
    for (const [initials, stats] of Object.entries(data.users)) {
      const hot  = stats.hot  || 0;
      const warm = stats.warm || 0;
      const cold = stats.cold || 0;
      const total = stats.total || 0;
      const closed = stats.closed ? "✅" : "⚠️ pas de clôture";
      lines.push(
        `👤 ${escapeHtml(initials)} : ${total} fiche(s) · ` +
        `${hot} 🔥 · ${warm} 📅 · ${cold} ❄️ — ${closed}`
      );
    }
    lines.push(`Total agence : <b>${data.total}</b> prospects`);
  }

  lines.push(`\n🕐 Export CRM : ✅ déclenché`);

  try {
    await sendMessage(managerId, lines.join("\n"));
  } catch (e) {
    log("error", null, "MANAGER_SUMMARY_FAILED", { err: e?.message });
  }
}


// =====================================================
// GITHUB DISPATCH
// =====================================================

async function dispatchGitHub(date, agency, initials) {
  const token    = String(globalThis.GITHUB_TOKEN || "").trim();
  const owner    = String(globalThis.GITHUB_OWNER || "").trim();
  const repo     = String(globalThis.GITHUB_REPO || "").trim();
  const workflow = String(globalThis.GITHUB_WORKFLOW_FILE || "prospection-close.yml").trim();
  if (!token || !owner || !repo) throw new Error("GitHub config missing");

  const res = await safeFetchText(
    `https://api.github.com/repos/${owner}/${repo}/actions/workflows/${workflow}/dispatches`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/vnd.github+json",
        "Content-Type": "application/json",
        "User-Agent": "prospection-worker/8.4",
      },
      body: JSON.stringify({
        ref: "main",
        inputs: { mode: "individual", run_date: date, agency, initials },
      }),
    },
    { timeoutMs: GITHUB_FETCH_TIMEOUT_MS, retries: 1, label: "GITHUB_DISPATCH", acceptStatuses: [204] }
  );

  if (!res.ok) {
    throw new Error(`GitHub dispatch failed (${res.status}): ${(res.text || "").slice(0, 200)}`);
  }
}

async function dispatchGitHubWithRetry(date, agency, initials, maxRetries = 3) {
  let lastErr = null;
  for (let i = 0; i < maxRetries; i++) {
    try {
      await dispatchGitHub(date, agency, initials);
      log("info", null, "GITHUB_DISPATCH_OK", { date, agency, initials, attempt: i + 1 });
      return true;
    } catch (e) {
      lastErr = e;
      if (i < maxRetries - 1) await new Promise((r) => setTimeout(r, Math.pow(2, i) * 1000));
    }
  }
  const queueKey = `queue:github:${Date.now()}:${agency}:${initials}`;
  await kv().put(queueKey, JSON.stringify({ date, agency, initials, queued_at: new Date().toISOString() }), {
    expirationTtl: GITHUB_QUEUE_TTL,
  });
  log("error", null, "GITHUB_DISPATCH_FAILED_FINAL", { date, agency, initials, err: lastErr?.message });
  return false;
}

async function processGitHubQueue() {
  const queueKeys = await listAllKeys("queue:github:");
  let processed = 0, removed = 0, failed = 0;
  for (const k of queueKeys) {
    const payload = await kv().get(k.name, { type: "json" });
    if (!payload?.date || !payload?.agency || !payload?.initials) { await kv().delete(k.name); continue; }
    processed++;
    try {
      await dispatchGitHub(payload.date, payload.agency, payload.initials);
      await kv().delete(k.name);
      removed++;
    } catch (e) {
      failed++;
      log("error", null, "GITHUB_QUEUE_DISPATCH_FAIL", { key: k.name, err: e?.message });
    }
  }
  return { processed, removed, failed };
}


// =====================================================
// ENDPOINT /update — Enrichissement OCR GitHub Actions
// =====================================================

async function handleOCRUpdate(req) {
  const ok = await checkExportToken(req);
  if (!ok) return new Response("Unauthorized", { status: 401 });

  let body;
  try { body = await req.json(); }
  catch (_) { return json({ ok: false, error: "invalid JSON" }, 400); }

  const { kv_key, chat_id, ocr_result } = body || {};
  if (!kv_key || !ocr_result) return json({ ok: false, error: "missing fields" }, 400);

  const existing = await kv().get(kv_key, { type: "json" });
  if (!existing) return json({ ok: false, error: "record not found" }, 404);

  const enriched = { ...existing };
  const improvements = [];

  const ENRICHABLE = ["name", "phone", "phone2", "email", "website", "city",
    "address", "postal_code", "interlocuteur", "contact_firstname",
    "contact_lastname", "contact_civility", "secteur_activite"];

  for (const k of ENRICHABLE) {
    const v = String(ocr_result[k] || "").trim();
    if (v && !enriched[k]) {
      enriched[k] = v;
      improvements.push(k);
    }
  }

  enriched.ocr_enriched_at = new Date().toISOString();
  enriched._ocr_confidence  = ocr_result._confidence || enriched._ocr_confidence || "";

  await kv().put(kv_key, JSON.stringify(enriched), { expirationTtl: TTL_PROSPECT });

  if (improvements.length > 0 && chat_id) {
    try {
      await sendMessage(
        chat_id,
        `🔍 <b>Fiche enrichie par OCR avancé</b>\n` +
        `Champs ajoutés : ${improvements.join(", ")}`
      );
    } catch (_) {}
  }

  log("info", chat_id, "OCR_ENRICHED", { kv_key, improvements });
  return json({ ok: true, improvements });
}


// =====================================================
// TELEGRAM ROUTER
// =====================================================

// ── Pin / Unpin fiche active ──────────────────────
async function pinFormMessage(chatId, messageId) {
  // Pin désactivé — génère des messages système parasites dans le chat
  return false;
}

async function unpinFormMessage(chatId, messageId = null) {
  // Unpin désactivé — cohérent avec pinFormMessage
  return false;
}

// ── Spinner pendant les enrichissements ──────────
async function sendSpinner(chatId, text = "⏳ En cours…") {
  try {
    return await tg("sendMessage", { chat_id: chatId, text, parse_mode: "HTML" });
  } catch (_) { return null; }
}

async function deleteSpinner(chatId, msg) {
  if (msg?.message_id) await deleteMessage(chatId, msg.message_id).catch(() => {});
}

async function handleTelegram(update, event = null) {
  if (update.message) return handleMessage(update.message, event);
  if (update.callback_query) return handleCallback(update.callback_query, event);
}


// =====================================================
// MESSAGE HANDLER
// =====================================================

async function handleMessage(msg, event = null) {
  const chatId = msg.chat.id;
  const text   = String(msg.text || "").trim();
  let s = await getSession(chatId);

  // /n = nouvelle visite, /s = save, /l = liste, /g = GPS
  if (text === "/n" && s?.agency && s?.initials) {
    return handleCallback({ id: "", data: "visit:new", message: { chat: { id: chatId } } }, event);
  }
  if (text === "/s" && s?.agency && s?.initials) {
    return handleCallback({ id: "", data: "save", message: { chat: { id: chatId } } }, event);
  }
  if (text === "/l" && s?.agency && s?.initials) {
    return handleCallback({ id: "", data: "visits:list", message: { chat: { id: chatId } } }, event);
  }
  if (text === "/g") {
    return sendMessage(chatId, "📍 Envoie ta position GPS.", locationRequestKeyboard());
  }

  // Bouton "⏭ Passer (sans GPS)" dans le flow façade simple
  if (text === "⏭ Passer (sans GPS)" && s?.step === "AWAIT_LOCATION_THEN_PHOTO" && s?.awaitingField === "facade") {
    s.step = "AWAIT_PHOTO"; s.awaitingField = "facade";
    await setSession(chatId, s);
    return sendMessage(chatId,
      "📸 Envoie la photo de la façade / enseigne.\n<i>Utilise le trombone 📎 pour meilleure qualité.</i>",
      failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]])
    );
  }

  // Bouton menu texte
  if (text === "⬅️ Menu") {
    if (!s) return sendMessage(chatId, "Tape /start pour démarrer.");
    // Si en mode lot, nettoyer le reply keyboard avant de revenir
    if (s.step === "PHOTO_COLLECT" || s.step === "PHOTO_COMMENT") {
      resetPhotoState(s);
      await sendMessage(chatId, "Lot photo abandonné.", removeReplyKeyboard());
    } else if (s.step === "CARD_COLLECT" || s.step === "CARD_COMMENT") {
      resetCardState(s);
      await sendMessage(chatId, "Lot cartes abandonné.", removeReplyKeyboard());
    }
    s.step = "MENU"; s.mode = "company"; s.awaitingField = null;
    resetPhotoState(s); resetCardState(s);
    await setSession(chatId, s);
    return showMainMenu(chatId, s);
  }

  // Bouton "✅ Terminer" du reply keyboard photo (pendant PHOTO_COMMENT ou CARD_COMMENT)
  if (text === "✅ Terminer") {
    if (s?.step === "PHOTO_COLLECT" || s?.step === "PHOTO_COMMENT") {
      return handleCallback({ id: "", data: "photo_finish", message: { chat: { id: chatId } } }, event);
    }
    if (s?.step === "CARD_COLLECT" || s?.step === "CARD_COMMENT") {
      return handleCallback({ id: "", data: "card_finish", message: { chat: { id: chatId } } }, event);
    }
  }

  // /start
  if (text === "/start") {
    const prefs = await getPrefs(chatId);
    if (prefs?.agency && prefs?.initials) {
      await resetSession(chatId);
      const fresh = initSession(prefs.agency, prefs.initials);
      fresh.step = "PROFILE_CONFIRM";
      await setSession(chatId, fresh);
      return sendMessage(
        chatId,
        `👋 Bon retour !\n\nDernier profil : <b>${escapeHtml(prefs.agency)}</b> / <b>${escapeHtml(prefs.initials)}</b>\n\nContinuer avec ce profil ?`,
        confirmProfileKeyboard(prefs.agency, prefs.initials)
      );
    }
    await resetSession(chatId);
    await setSession(chatId, initSession());
    return sendMessage(chatId, "👋 Bienvenue ! Choisis ton agence :", agenciesKeyboard());
  }

  // /debug
  if (text === "/clearcache") {
    // Purge tous les caches Gemini/Places/Gouv en listant les clés cache:
    try {
      const listed = await kv().list({ prefix: "cache:" });
      const keys = (listed?.keys || []).map(k => k.name);
      let deleted = 0;
      for (const k of keys) {
        await kv().delete(k).catch(() => {});
        deleted++;
      }
      return sendMessage(chatId, `🗑️ Cache purgé : ${deleted} entrée(s) supprimée(s).\n\nTu peux renvoyer ta photo.`);
    } catch (e) {
      return sendMessage(chatId, `❌ Erreur purge cache : ${e?.message}`);
    }
  }

  // /debug
  if (text === "/debug") {
    const hasGemini = !!String(globalThis.GEMINI_API_KEY || "").trim();
    const hasKV     = !!globalThis.PROSPECTION_KV;
    const step      = s?.step || "(pas de session)";
    const agency    = s?.agency || "";
    const initials  = s?.initials || "";
    const awaitingField = s?.awaitingField || "";
    return sendMessage(chatId,
      `🔍 <b>Debug session</b>\n\n` +
      `<b>KV :</b> ${hasKV ? "✅" : "❌ MANQUANT"}\n` +
      `<b>Gemini :</b> ${hasGemini ? "✅" : "❌ MANQUANT"}\n` +
      `<b>Agence :</b> ${agency || "(vide)"}\n` +
      `<b>Initiales :</b> ${initials || "(vide)"}\n` +
      `<b>Step :</b> <code>${step}</code>\n` +
      `<b>Champ attendu :</b> ${awaitingField || "(aucun)"}\n\n` +
      `Envoie une photo <b>maintenant</b> puis refais /debug pour voir le changement.`
    );
  }

  // /aide
  if (text === "/aide") {
    return sendMessage(chatId,
      `📖 <b>Commandes</b>\n\n` +
      `/start — Démarrer / changer profil\n` +
      `/aide — Aide\n` +
      `/p &lt;recherche&gt; — Recherche souple\n` +
      `/c — Clôture immédiate\n` +
      `/resume — Résumé session\n\n` +
      `<b>Raccourcis terrain :</b>\n` +
      `/n — Nouvelle visite\n` +
      `/s — Enregistrer la fiche\n` +
      `/l — Liste du jour\n` +
      `/g — Demander position GPS\n\n` +
      `<b>Astuce photo :</b>\n` +
      `📎 Envoie la carte via le <b>trombone</b> pour une meilleure qualité OCR\n\n` +
      `<b>Astuce résumé :</b>\n` +
      `🎤 Envoie un message vocal pour dicter le résumé d'entretien`
    );
  }

  if (text === "/resume") {
    if (!s?.agency || !s?.initials) return sendMessage(chatId, "Tape /start d'abord.");
    const hasVisits = Array.isArray(s.visits_index) && s.visits_index.length > 0;
    return sendMessage(chatId, await buildSessionSummary(chatId, s),
      mainMenuKeyboard(!!(await getLastProspectKey(chatId)), hasVisits));
  }

  if (text === "/p") {
    if (!s?.agency || !s?.initials) return sendMessage(chatId, "Tape /start d'abord.");
    s.step = "SEARCH"; await setSession(chatId, s);
    return sendMessage(chatId, "🔎 Recherche entreprise :", {
      force_reply: true,
      selective: true,
      input_field_placeholder: "Ex : Lely Lyon / 0476123456 / recyclage 38",
    });
  }

  if (text.startsWith("/p ") || text.startsWith("/p\n")) {
    const query = text.slice(3).trim();
    if (!query || query.length < 2) return sendMessage(chatId, "Usage : /p Dupont Lyon", failSafeKeyboard());
    if (!s?.agency || !s?.initials) return sendMessage(chatId, "Profil incomplet. Tape /start.");
    s.step = "SEARCH"; await setSession(chatId, s);
    return handleSearch(chatId, s, query);
  }

  if (text === "/c") {
    if (!s?.agency || !s?.initials) return sendMessage(chatId, "Tape /start d'abord.");
    s.step = "CLOSE_CLIENTS";
    s.close_stats = { visits_clients: "", visits_prospects: "", commandes: "" };
    await setSession(chatId, s);
    return sendMessage(chatId, "🔢 Combien de <b>CLIENTS</b> visités aujourd'hui ?");
  }

  if (!s) return sendMessage(chatId, "Tape /start pour démarrer.");

  // GPS
  if (msg.location && typeof msg.location.latitude === "number") {
    const lat = msg.location.latitude;
    const lon = msg.location.longitude;
    s.current_geo = { lat, lon, at: Date.now() };
    await setSession(chatId, s);

    // Reverse geocoding automatique
    const spinner = await sendMessage(chatId, "📍 Localisation en cours…");
    const place = await reverseGeocode(lat, lon);
    await deleteMessage(chatId, spinner?.message_id).catch(() => {});

    if (place?.city) {
      const cityLabel = [place.postal_code, place.city].filter(Boolean).join(" ");

      // Mode façade simple → GPS reçu → confirmer + demander photo
      if (s.step === "AWAIT_LOCATION_THEN_PHOTO" && s.awaitingField === "facade") {
        s.draft.city        = place.city;
        s.draft.postal_code = place.postal_code || "";
        s.draft.address     = place.address     || "";
        s.draft.gps_lat     = lat;
        s.draft.gps_lon     = lon;
        s.step = "AWAIT_PHOTO"; s.awaitingField = "facade";
        await setSession(chatId, s);
        await sendMessage(chatId, `📍 Position enregistrée : <b>${escapeHtml(cityLabel)}</b>`);
        return sendMessage(chatId,
          "📸 Envoie maintenant la photo de la façade / enseigne.\n<i>Utilise le trombone 📎 pour meilleure qualité.</i>",
          failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]])
        );
      }

      // Si on attend un GPS pour le mode photo lot → basculer directement
      if (s.mode === "photo" && (s.step === "PHOTO_CITY_INPUT" || !s.photo_city)) {
        s.photo_city        = cityLabel;
        s._last_photo_city  = cityLabel;
        s._gps_city         = place.city;
        s._gps_postal_code  = place.postal_code;
        s.step = "PHOTO_COLLECT";
        await setSession(chatId, s);
        return sendMessage(chatId,
          `✅ Position : <b>${escapeHtml(cityLabel)}</b>\n\n` +
          `📎 <b>Envoie tes photos via le TROMBONE</b>`,
          photoReplyKeyboard("📸 Envoyer une façade", false)
        );
      }

      // Sinon → juste confirmer et proposer de l'utiliser pour le prochain lot
      s._gps_city        = place.city;
      s._gps_postal_code = place.postal_code;
      await setSession(chatId, s);
      return sendMessage(chatId,
        `📍 Position enregistrée : <b>${escapeHtml(cityLabel)}</b>\n` +
        `<i>Utilisée automatiquement pour le prochain lot de façades.</i>`
      );
    }

    return sendMessage(chatId, `✅ Position GPS enregistrée (${lat.toFixed(4)}, ${lon.toFixed(4)}).`);
  }

  // ── MESSAGE VOCAL → résumé d'entretien ────────────
  if (msg.voice && s.step === "VISIT_ACTIVE" && s.active_visit_id) {
    const voice = msg.voice;
    await sendMessage(chatId, "🎤 Transcription en cours…");
    try {
      const file = await tgGetFile(voice.file_id);
      if (file?.file_path) {
        const audioBytes = await tgDownloadFileBytes(file.file_path);
        if (audioBytes.length) {
          const transcription = await geminiTranscribeAudio(audioBytes, voice.mime_type || "audio/ogg");
          if (transcription) {
            s.draft.resume = trunc(transcription, 500);
            await setSession(chatId, s);
            await postOrReplaceForm(chatId, s);
            return sendMessage(chatId, `✅ Résumé transcrit :\n<i>${escapeHtml(transcription.slice(0, 200))}</i>`);
          }
        }
      }
    } catch (_) {}
    return sendMessage(chatId, "⚠️ Impossible de transcrire le vocal. Saisis le résumé manuellement.");
  }

  // Initiales
  if (s.step === "INITIALS" && text) {
    const ini = text.trim().toUpperCase();
    if (!/^[A-Z]{1,6}$/.test(ini)) return sendMessage(chatId, "Initiales invalides. Ex : SL, JL, CZ…");
    s.initials = ini; s.step = "MENU";
    await setSession(chatId, s);
    await savePrefs(chatId, s.agency, s.initials);
    return showMainMenu(chatId, s, "✅ Profil enregistré !");
  }

  // Close flow — validation numérique
  if (s.step === "CLOSE_CLIENTS") {
    const n = parsePositiveInt(text);
    if (n === null) return sendMessage(chatId, "⚠️ Entre un nombre valide (ex: 3).");
    s.close_stats.visits_clients = String(n);
    s.step = "CLOSE_PROSPECTS"; await setSession(chatId, s);
    return sendMessage(chatId, "🔢 Combien de <b>PROSPECTS</b> visités ?");
  }

  if (s.step === "CLOSE_PROSPECTS") {
    const n = parsePositiveInt(text);
    if (n === null) return sendMessage(chatId, "⚠️ Entre un nombre valide (ex: 5).");
    s.close_stats.visits_prospects = String(n);
    s.step = "CLOSE_COMMANDES"; await setSession(chatId, s);
    return sendMessage(chatId, "🔢 Combien de <b>COMMANDES</b> signées ?");
  }

  if (s.step === "CLOSE_COMMANDES") {
    const n = parsePositiveInt(text);
    if (n === null) return sendMessage(chatId, "⚠️ Entre un nombre valide (ex: 0).");
    s.close_stats.commandes = String(n);
    s.step = "CLOSE_CONFIRM"; await setSession(chatId, s);
    return sendMessage(chatId,
      `🔎 <b>Récapitulatif clôture</b>\n\n` +
      `🏷️ <b>${escapeHtml(s.agency)}</b> | 👤 <b>${escapeHtml(s.initials)}</b>\n\n` +
      `👥 Clients visités : <b>${s.close_stats.visits_clients}</b>\n` +
      `🎯 Prospects visités : <b>${s.close_stats.visits_prospects}</b>\n` +
      `📦 Commandes signées : <b>${s.close_stats.commandes}</b>\n\n` +
      `Confirmer ?`,
      {
        inline_keyboard: [
          [{ text: "✅ Confirmer", callback_data: "close_confirm" }],
          [{ text: "❌ Annuler", callback_data: "close_cancel" }],
        ],
      }
    );
  }

  // Photo lot: city
  if (s.step === "PHOTO_CITY_INPUT") {
    const city = normalizeSpaces(text);
    if (!city || city.length < 2) return sendMessage(chatId, "Ville invalide.", failSafeKeyboard());
    s.photo_city = trunc(city, 100);
    s._last_photo_city = s.photo_city;
    s.step = "PHOTO_COLLECT";
    await setSession(chatId, s);
    return sendMessage(chatId,
      `✅ Secteur : <b>${escapeHtml(s.photo_city)}</b>\n\n` +
      `📎 <b>IMPORTANT — Envoie via le TROMBONE</b>\n` +
      `Trombone → Fichier → Galerie → choisir la photo\n\n` +
      `<i>L'envoi normal compresse et recadre l'image → enseigne invisible.\n` +
      `Via trombone = qualité originale = OCR fiable.</i>\n\n` +
      `Tape <b>.</b> pour passer le commentaire.`,
      photoReplyKeyboard("📸 Envoyer une façade (via trombone SVP)", false)
    );
  }

  // Photo lot: receive photo
  if (s.step === "PHOTO_COLLECT") {
    const photos = msg.photo || [];
    const doc    = msg.document;
    let file_id  = "";

    if (doc && doc.mime_type && doc.mime_type.startsWith("image/")) {
      file_id = doc.file_id;
    } else if (photos.length) {
      file_id = photos[photos.length - 1].file_id;
    }

    if (!file_id) return sendMessage(chatId, "📸 Envoie une photo de façade.",
      photoReplyKeyboard("📸 Prendre / envoyer une façade", s.photo_batch?.length > 0));

    if (!isValidFileId(file_id)) return sendMessage(chatId, "⚠️ Photo invalide.", failSafeKeyboard());
    if (!Array.isArray(s.photo_batch)) s.photo_batch = [];
    if (!s.photo_ts0) s.photo_ts0 = Date.now();
    if (s.photo_batch.length >= PHOTO_LIMIT) return sendMessage(chatId, `⚠️ Limite ${PHOTO_LIMIT} atteinte.`,
      failSafeKeyboard([[{ text: "✅ Terminer", callback_data: "photo_finish" }]], true));

    const geoSnap = getFreshGeoOrNull(s);

    // Spinner pendant OCR
    const spinnerPhoto = await sendMessage(chatId, "⏳ Analyse de la façade…");

    let facade = {};
    try {
      const imgBytes = await withTimeout(getBytesFromFileId(file_id), 10_000, new Uint8Array());
      if (imgBytes.length) {
        facade = await hybridExtractFacade(imgBytes, s.photo_city);
      }
    } catch (e) {
      log("warn", chatId, "PHOTO_OCR_ERROR", { err: e?.message });
    } finally {
      await deleteMessage(chatId, spinnerPhoto?.message_id).catch(() => {});
    }

    const gemini_company  = facade.company  || "";
    const gemini_city     = facade.city     || "";
    const gemini_activity = facade.activity || "";

    s.photo_batch.push({
      file_id, ts: Date.now(), geo: geoSnap, comment: null,
      gemini_company, gemini_city, gemini_activity,
    });
    s.step = "PHOTO_COMMENT";
    await setSession(chatId, s);

    // Feedback OCR : affiche ce qui a été détecté (ou guidance si rien)
    const n = s.photo_batch.length;
    if (gemini_company) {
      await sendMessage(chatId,
        `✅ Photo ${n}/${PHOTO_LIMIT} — <b>${escapeHtml(s.photo_city)}</b>\n` +
        `🏢 <b>${escapeHtml(gemini_company)}</b>\n` +
        (gemini_activity ? `🏭 ${escapeHtml(gemini_activity)}\n` : "") +
        (facade.phone    ? `📞 ${escapeHtml(formatPhone(facade.phone))}\n`    : "") +
        (facade.website  ? `🌐 ${escapeHtml(facade.website)}\n`  : "")
      );
    } else {
      await sendMessage(chatId,
        `✅ Photo ${n}/${PHOTO_LIMIT} — <b>${escapeHtml(s.photo_city)}</b>\n` +
        `<i>Enseigne non détectée — le nom sera enrichi lors de l'export.</i>`
      );
    }

    return sendMessage(chatId, `💬 Commentaire ? <i>(<b>.</b> pour passer)</i>`);
  }

  // Photo lot: comment
  if (s.step === "PHOTO_COMMENT") {
    const comment = parseComment(text);
    if (Array.isArray(s.photo_batch) && s.photo_batch.length > 0) {
      s.photo_batch[s.photo_batch.length - 1].comment = comment;
    }
    s.step = "PHOTO_COLLECT"; await setSession(chatId, s);
    const n = s.photo_batch.length;
    return sendMessage(chatId,
      `${comment ? `📝 <i>${escapeHtml(comment)}</i>\n\n` : ""}📸 ${n}/${PHOTO_LIMIT}. Continue ou termine.`,
      photoReplyKeyboard("📸 Photo suivante", true)
    );
  }

  // Card lot: receive card
  if (s.step === "CARD_COLLECT") {
    // Support document non compressé + photo
    const photos = msg.photo || [];
    const doc    = msg.document;
    let file_id  = "";

    if (doc && doc.mime_type && doc.mime_type.startsWith("image/")) {
      file_id = doc.file_id;
    } else if (photos.length) {
      file_id = photos[photos.length - 1].file_id;
    }

    if (!file_id) return sendMessage(chatId,
      "🪪 Envoie une photo de carte.\n<i>📎 Astuce : utilise le trombone pour la qualité originale.</i>",
      photoReplyKeyboard("📸 Prendre / envoyer une carte", s.card_batch?.length > 0));

    if (!isValidFileId(file_id)) return sendMessage(chatId, "⚠️ Photo invalide.", failSafeKeyboard());
    if (!Array.isArray(s.card_batch)) s.card_batch = [];
    if (!s.card_ts0) s.card_ts0 = Date.now();
    if (s.card_batch.length >= CARD_LIMIT) return sendMessage(chatId, `⚠️ Limite ${CARD_LIMIT} atteinte.`,
      failSafeKeyboard([[{ text: "✅ Terminer", callback_data: "card_finish" }]], true));

    // Sauvegarder l'état AVANT l'OCR pour éviter de bloquer le bot si timeout
    s.step = "CARD_OCR_PENDING";
    await setSession(chatId, s);

    const spinnerMsg = await sendMessage(chatId, "⏳ Lecture de la carte en cours…");

    // OCR carte — synchrone mais dans la même requête
    let gemini = {};
    try {
      const imgBytes = await getBytesFromFileId(file_id).catch(() => new Uint8Array());
      if (imgBytes.length) {
        gemini = await hybridExtractCard(imgBytes);
      }
    } catch (e) {
      log("warn", chatId, "CARD_OCR_ERROR", { err: e?.message });
    }

    await deleteMessage(chatId, spinnerMsg?.message_id).catch(() => {});

    // Enrichissement Gouv immédiat + ouverture fiche active
    const cardCompany = gemini.company || "";
    const cardCity    = gemini.city || "";
    const cardCp      = gemini.postal_code || "";

    // Créer une nouvelle fiche active
    await createNewActiveVisit(chatId, s, "card");
    s = await getSession(chatId);

    // Remplir avec les données OCR
    if (gemini.company) s.draft.name = gemini.company;
    if (gemini.name) {
      s.draft.interlocuteur = gemini.name;
      const sp = splitHumanName(gemini.name);
      s.draft.contact_firstname = sp.first || "";
      s.draft.contact_lastname  = sp.last || gemini.name;
    }
    if (gemini.contact_civility) s.draft.contact_civility = gemini.contact_civility;
    if (gemini.title)       s.draft.titre       = gemini.title;
    if (gemini.email)       s.draft.email       = normalizeEmail(gemini.email);
    if (gemini.phone)       s.draft.phone       = normalizePhone(gemini.phone);
    if (gemini.phone2)      s.draft.phone2      = normalizePhone(gemini.phone2);
    if (gemini.website)     s.draft.website     = gemini.website;
    if (gemini.city)        s.draft.city        = gemini.city;
    if (gemini.postal_code) s.draft.postal_code = gemini.postal_code;
    if (gemini.address)     s.draft.address     = gemini.address;
    s.draft.card_photo_file_id = file_id;
    s.draft._ocr_confidence = gemini._confidence || "medium";

    // Enrichissement Gouv
    if (cardCompany) {
      try {
        const emailDomain = (gemini.email || "").split("@")[1]?.replace(/\.(fr|com|net|org|eu)$/, "").replace(/[\-_.]/g, " ").trim() || "";
        const siteDomain  = (gemini.website || "").replace(/https?:\/\/(www\.)?/, "").split("/")[0].replace(/\.(fr|com|net|org|eu)$/, "").replace(/[\-_.]/g, " ").trim();
        const queries = new Set();
        if (cardCompany && cardCity) queries.add(`${cardCompany} ${cardCity}`);
        if (cardCompany && cardCp)   queries.add(`${cardCompany} ${cardCp}`);
        if (emailDomain && cardCp)   queries.add(`${emailDomain} ${cardCp}`);
        if (siteDomain && cardCp && siteDomain !== emailDomain) queries.add(`${siteDomain} ${cardCp}`);
        if (emailDomain && cardCity) queries.add(`${emailDomain} ${cardCity}`);
        if (normalizePhone(gemini.phone || "")) queries.add(normalizePhone(gemini.phone));
        queries.add(cardCompany);

        for (const q of queries) {
          if (!q.trim()) continue;
          const r = await searchCompany(q.trim()).catch(() => []);
          const best = r[0];
          if (best) {
            if (!s.draft.name && best.nom)           s.draft.name            = best.nom;
            if (best.siret)                          s.draft.siret           = best.siret;
            if (best.naf)                            s.draft.naf             = best.naf;
            if (!s.draft.address && best.adresse)    s.draft.address         = best.adresse;
            if (!s.draft.postal_code && best.codePostal) s.draft.postal_code = best.codePostal;
            if (!s.draft.city && best.ville)         s.draft.city            = best.ville;
            if (best.effectif)                       s.draft.effectif        = best.effectif;
            if (best.date_creation)                  s.draft.date_creation   = best.date_creation;
            if (best.capital)                        s.draft.capital         = best.capital;
            if (!s.draft.secteur_activite && best.secteur_activite) s.draft.secteur_activite = best.secteur_activite;
            if (best.dirigeants?.length && !s.draft.dirigeant) {
              const dir = best.dirigeants[0];
              const nom = typeof dir === "string" ? dir
                : cleanPersonName([dir.prenom?.split(" ")[0], dir.nom || dir.nom_usage].filter(Boolean).join(" "));
              if (nom) s.draft.dirigeant = normalizeSpaces(nom);
            }
            break;
          }
        }

        // Places pour tel/site si manquants
        if (!s.draft.phone || !s.draft.website) {
          const place = await googlePlacesEnrich(s.draft.name || cardCompany, s.draft.city || cardCity).catch(() => ({}));
          if (place.phone   && !s.draft.phone)   s.draft.phone   = place.phone;
          if (place.website && !s.draft.website) s.draft.website = place.website;
        }
      } catch (_) {}
    }

    // Stocker dans le batch pour card_finish
    if (!Array.isArray(s.card_batch)) s.card_batch = [];
    s.card_batch.push({ file_id, ts: Date.now(), comment: null, gemini });
    s.step = "VISIT_ACTIVE";
    s.mode = "card";
    await persistActiveVisit(chatId, s);
    await setSession(chatId, s);

    // Ouvrir la fiche active avec tous les boutons d'édition
    return postOrReplaceForm(chatId, s);
  }

  // Sécurité : si le bot était en CARD_OCR_PENDING (crash/timeout précédent)
  // → remettre en CARD_COLLECT pour débloquer
  if (s.step === "CARD_OCR_PENDING") {
    s.step = "CARD_COLLECT";
    await setSession(chatId, s);
    return sendMessage(chatId,
      "⚠️ La lecture précédente a expiré. Renvoie la carte.",
      photoReplyKeyboard("📸 Renvoyer la carte", s.card_batch?.length > 0)
    );
  }

  // Card lot: comment
  if (s.step === "CARD_COMMENT") {
    const comment = parseComment(text);
    if (Array.isArray(s.card_batch) && s.card_batch.length > 0) {
      s.card_batch[s.card_batch.length - 1].comment = comment;
    }
    s.step = "CARD_COLLECT"; await setSession(chatId, s);
    const n = s.card_batch.length;
    return sendMessage(chatId,
      `${comment ? `📝 <i>${escapeHtml(comment)}</i>\n\n` : ""}🪪 ${n}/${CARD_LIMIT}. Continue ou termine.`,
      photoReplyKeyboard("📸 Carte suivante", true)
    );
  }

  // Active visit: awaiting photo (card) — support document non compressé
  if (s.step === "AWAIT_PHOTO" && s.awaitingField === "card") {
    const photos = msg.photo || [];
    const doc    = msg.document;
    let file_id  = "";
    if (doc && doc.mime_type && doc.mime_type.startsWith("image/")) {
      file_id = doc.file_id;
    } else if (photos.length) {
      file_id = photos[photos.length - 1].file_id;
    }
    if (!file_id) return sendMessage(chatId,
      "📎 Envoie la photo de la carte.\n<i>Astuce : utilise le trombone 📎 pour la qualité originale.</i>",
      failSafeKeyboard());
    if (!isValidFileId(file_id)) return sendMessage(chatId, "⚠️ Photo invalide.", failSafeKeyboard());

    const spinnerCard = await sendMessage(chatId, "⏳ Lecture de la carte en cours…");
    let card = {};
    try {
      const imgBytes = await withTimeout(getBytesFromFileId(file_id), 10_000, new Uint8Array());
      if (imgBytes.length) {
        card = await hybridExtractCard(imgBytes);
      }
    } catch (e) {
      log("warn", chatId, "AWAIT_PHOTO_CARD_OCR_ERROR", { err: e?.message });
    } finally {
      await deleteMessage(chatId, spinnerCard?.message_id).catch(() => {});
    }
    await sendOCRFeedback(chatId, card, "card");

    // Fusion dans le draft
    if (card.company) s.draft.name = s.draft.name || card.company;
    if (card.name) {
      s.draft.interlocuteur = card.name;
      const sp = splitHumanName(card.name);
      s.draft.contact_firstname = sp.first || "";
      s.draft.contact_lastname  = sp.last || card.name;
    }
    if (card.contact_civility) s.draft.contact_civility = card.contact_civility;
    if (card.email)       s.draft.email   = normalizeEmail(card.email) || s.draft.email;
    if (card.phone)       s.draft.phone   = normalizePhone(card.phone) || s.draft.phone;
    if (card.phone2)      s.draft.phone2  = normalizePhone(card.phone2) || s.draft.phone2;
    if (card.website)     s.draft.website = card.website || s.draft.website;
    if (card.city)        s.draft.city    = s.draft.city || card.city;
    if (card.postal_code) s.draft.postal_code = s.draft.postal_code || card.postal_code;
    if (card.address)     s.draft.address = s.draft.address || card.address;
    s.draft.card_photo_file_id = file_id;
    s.draft.card_photo_url     = await tgGetFileUrl(file_id);
    s.draft._ocr_confidence    = card._confidence || "low";

    s.step = "VISIT_ACTIVE"; s.awaitingField = null;
    await setSession(chatId, s);
    return postOrReplaceForm(chatId, s);
  }

  // Active visit: awaiting photo (facade)
  if (s.step === "AWAIT_PHOTO" && s.awaitingField === "facade") {
    const photos = msg.photo || [];
    const doc    = msg.document;
    let file_id  = "";

    if (doc?.mime_type?.startsWith("image/")) file_id = doc.file_id;
    else if (photos.length) file_id = photos[photos.length - 1].file_id;

    if (!file_id) return sendMessage(chatId, "📸 Envoie la photo de la façade (via trombone pour meilleure qualité).", failSafeKeyboard());
    if (!isValidFileId(file_id)) return sendMessage(chatId, "⚠️ Photo invalide.", failSafeKeyboard());

    // Créer la visite active AVANT d'écrire les données OCR
    // (createNewActiveVisit réinitialise s.draft : le faire après effacerait l'OCR)
    if (!s.active_visit_id) await createNewActiveVisit(chatId, s, "photo");

    // GPS → ville + code postal uniquement. JAMAIS l'adresse : elle appartient
    // à l'entreprise (trouvée par l'enrichissement), pas à la position du commercial.
    const geoSnap = getFreshGeoOrNull(s);
    if (geoSnap) {
      s.draft.gps_lat = geoSnap.lat;
      s.draft.gps_lon = geoSnap.lon;
      if (!s.draft.city || !s.draft.postal_code) {
        const geo = await reverseGeocode(geoSnap.lat, geoSnap.lon).catch(() => null);
        if (geo) {
          if (!s.draft.city)        s.draft.city        = geo.city || "";
          if (!s.draft.postal_code) s.draft.postal_code = geo.postal_code || "";
        }
      }
    }

    const spinnerFacade = await sendMessage(chatId, "⏳ Analyse de la façade…");
    let facade = {};
    try {
      const imgBytes = await withTimeout(getBytesFromFileId(file_id), 10_000, new Uint8Array());
      if (imgBytes.length) {
        facade = await hybridExtractFacade(imgBytes, s.draft.city || "");
      }
    } catch (e) {
      log("warn", chatId, "AWAIT_PHOTO_FACADE_OCR_ERROR", { err: e?.message });
    } finally {
      await deleteMessage(chatId, spinnerFacade?.message_id).catch(() => {});
    }
    await sendOCRFeedback(chatId, facade, "facade");

    if (facade.company && !s.draft.name) s.draft.name = facade.company;
    if (facade.city && !s.draft.city)    s.draft.city = facade.city;
    if (facade.phone && !s.draft.phone)  s.draft.phone = normalizePhone(facade.phone);
    if (facade.website && !s.draft.website) s.draft.website = facade.website;
    if (facade.activity && !s.draft.secteur_activite) s.draft.secteur_activite = facade.activity;
    s.draft.facade_photo_file_id = file_id;
    s.draft.facade_photo_url     = await tgGetFileUrl(file_id);
    s.draft._ocr_confidence      = facade.confidence || "low";

    // Enrichissement automatique (Gouv/Places/Brave), calqué sur createProspectFromFacade.
    // adresse vide + ville/CP GPS → la Stratégie 3 cherche par nom+CP sans l'adresse du commercial.
    if (s.draft.name) {
      const spinnerEnrich = await sendMessage(chatId, `🔍 Enrichissement de <b>${escapeHtml(s.draft.name)}</b>…`);
      try {
        const enriched = await enrichCompany({
          nom:             s.draft.name,
          codePostal:      s.draft.postal_code || "",
          ville:           s.draft.city || "",
          adresse:         "", // Jamais l'adresse GPS pour une façade
          _places_phone:   normalizePhone(s.draft.phone || ""),
          _places_website: s.draft.website || "",
          _places_lat:     s.draft.gps_lat || null,
          _places_lon:     s.draft.gps_lon || null,
        });
        s.draft = applyCompanyToDraft(s.draft, enriched);
      } catch (err) {
        log("warn", chatId, "AWAIT_PHOTO_FACADE_ENRICH_ERROR", { err: err?.message });
      } finally {
        await deleteMessage(chatId, spinnerEnrich?.message_id).catch(() => {});
      }
    }

    s.step = "VISIT_ACTIVE"; s.awaitingField = null;
    await setSession(chatId, s);
    return postOrReplaceForm(chatId, s);
  }

  // Search
  if (s.step === "SEARCH") {
    if (!text || text.length < 2) {
      return sendMessage(chatId, "Recherche trop courte. Ex : <i>Lely Lyon</i>, <i>0476123456 Grenoble</i>", {
        force_reply: true,
        selective: true,
        input_field_placeholder: "Ex : Lely Lyon / 0476123456 / recyclage 38",
      });
    }
    return handleSearch(chatId, s, text);
  }

  // Manual company name
  if (s.step === "MANUAL_NAME") {
    const name = trunc(text, 200);
    if (!name || name.length < 2) return sendMessage(chatId, "Nom invalide.", failSafeKeyboard());
    if (!s.active_visit_id) { await createNewActiveVisit(chatId, s, "manual"); s = await getSession(chatId); }
    s.draft = normalizeDraft({ ...s.draft, _source: "manual", name, city: s._manual_city || s.draft.city || "", source_mode: "manual" });
    s._manual_city = ""; s.step = "VISIT_ACTIVE";
    await setSession(chatId, s);
    return postOrReplaceForm(chatId, s);
  }

  // Await text field
  if (s.step === "AWAIT_TEXT" && s.awaitingField) {
    const field = s.awaitingField;
    const value = trunc(text);
    if (field === "interlocuteur") {
      s.draft.interlocuteur = normalizeSpaces(value);
      const sp = splitHumanName(value);
      s.draft.contact_firstname = sp.first && sp.last ? sp.first : "";
      s.draft.contact_lastname  = sp.last || value;
    } else if (field === "phone" || field === "phone2") {
      s.draft[field] = normalizePhone(value);
    } else if (field === "email") {
      s.draft[field] = normalizeEmail(value);
    } else if (field === "postal_code") {
      s.draft[field] = normalizeSpaces(value).replace(/\s+/g, "");
    } else {
      s.draft[field] = normalizeSpaces(value);
    }
    s.draft = normalizeDraft(s.draft);
    s.step = "VISIT_ACTIVE"; s.awaitingField = null;
    await setSession(chatId, s);
    return postOrReplaceForm(chatId, s);
  }

  if (s.step === "MENU" || s.step === "PROFILE_CONFIRM" || s.step === "VISIT_ACTIVE") {
    if (s.step === "VISIT_ACTIVE" && s.active_visit_id) return postOrReplaceForm(chatId, s);
    const hasVisits = Array.isArray(s.visits_index) && s.visits_index.length > 0;
    return sendMessage(chatId, "Utilise le menu ou tape /aide.", mainMenuKeyboard(!!(await getLastProspectKey(chatId)), hasVisits));
  }

  return sendMessage(chatId, "Je ne comprends pas. Tape /aide ou /start.", failSafeKeyboard());
}


// =====================================================
// CALLBACK HANDLER
// =====================================================

async function handleCallback(cb, event = null) {
  const chatId = cb.message.chat.id;
  await answerCb(cb.id);
  let s = await getSession(chatId);
  if (!s) return sendMessage(chatId, "Tape /start pour démarrer.");
  const data = cb.data || "";

  // ── Summary ───────────────────────────────────────
  if (data === "summary") {
    const hasVisits = Array.isArray(s.visits_index) && s.visits_index.length > 0;
    return sendMessage(chatId, await buildSessionSummary(chatId, s),
      mainMenuKeyboard(!!(await getLastProspectKey(chatId)), hasVisits));
  }

  if (data === "help:inline") {
    return sendMessage(chatId,
      `📖 <b>Aide rapide</b>\n\n` +
      `• <b>Nouvelle visite</b> : ouvre une fiche active\n` +
      `• <b>Liste du jour</b> : rouvre une fiche par numéro\n` +
      `• <b>Carte</b> : crée / enrichit une fiche\n` +
      `• <b>Façade</b> : crée / enrichit une fiche\n` +
      `• <b>Enrichir</b> : relance recherche + Places\n` +
      `• <b>Qualification</b> : 🔥 chaud / 📅 tiède / ❄️ froid\n\n` +
      `🎤 Message vocal → résumé d'entretien automatique\n` +
      `📎 Trombone → carte de visite en qualité originale`
    );
  }

  // ── Profile ───────────────────────────────────────
  if (data === "profile:keep") {
    s.step = "MENU"; await setSession(chatId, s);
    return showMainMenu(chatId, s, `✅ Profil <b>${escapeHtml(s.agency)}</b> / <b>${escapeHtml(s.initials)}</b> actif.`);
  }

  if (data === "profile:change") {
    await resetSession(chatId); await setSession(chatId, initSession());
    return sendMessage(chatId, "Choisis ton agence :", agenciesKeyboard());
  }

  if (data.startsWith("agency:")) {
    const agency = data.split(":")[1];
    if (!VALID_AGENCIES.includes(agency)) return sendMessage(chatId, "Agence inconnue.", failSafeKeyboard());
    const fresh = initSession(agency, ""); fresh.step = "INITIALS";
    await setSession(chatId, fresh);
    return sendMessage(chatId, `Agence <b>${escapeHtml(agency)}</b> ✅\n\nTes initiales ?`);
  }

  // ── Navigation ────────────────────────────────────
  if (data === "menu") {
    s.step = "MENU"; s.mode = "company"; s.awaitingField = null;
    resetPhotoState(s); resetCardState(s); await setSession(chatId, s);
    return showMainMenu(chatId, s);
  }

  if (data === "retry_search") {
    s.step = "SEARCH"; await setSession(chatId, s);
    return sendMessage(chatId, "🔎 Recherche entreprise :", {
      force_reply: true,
      selective: true,
      input_field_placeholder: "Ex : Lely Lyon / 0476123456 / recyclage 38",
    });
  }

  if (data === "request_geo_hint") {
    return sendMessage(chatId, "📍 Utilise le bouton ci-dessous.", locationRequestKeyboard());
  }

  if (data === "menu_with_warn") {
    const nPhoto = s.photo_batch?.length || 0;
    const nCard  = s.card_batch?.length  || 0;
    if (nPhoto > 0) return sendMessage(chatId, `⚠️ Tu as ${nPhoto} photo(s) non validée(s). Abandonner ?`, {
      inline_keyboard: [
        [{ text: "✅ Oui, abandonner", callback_data: "photo_abandon" }],
        [{ text: "❌ Non, continuer",  callback_data: "photo_back" }],
      ],
    });
    if (nCard > 0) return sendMessage(chatId, `⚠️ Tu as ${nCard} carte(s) non validée(s). Abandonner ?`, {
      inline_keyboard: [
        [{ text: "✅ Oui, abandonner", callback_data: "card_abandon" }],
        [{ text: "❌ Non, continuer",  callback_data: "card_back" }],
      ],
    });
    s.step = "MENU"; s.mode = "company"; await setSession(chatId, s);
    return showMainMenu(chatId, s);
  }

  if (data === "photo_abandon") { resetPhotoState(s); s.step = "MENU"; s.mode = "company"; await setSession(chatId, s); return showMainMenu(chatId, s, "Lot photo abandonné."); }
  if (data === "photo_back") {
    s.step = "PHOTO_COLLECT"; await setSession(chatId, s);
    const n = s.photo_batch?.length || 0;
    return sendMessage(chatId, `📸 ${n}/${PHOTO_LIMIT}. Continue ou termine.`, {
      inline_keyboard: [
        [{ text: `✅ Terminer (${n})`, callback_data: "photo_finish" }],
        [{ text: "⬅️ Menu", callback_data: "menu_with_warn" }],
      ],
    });
  }
  if (data === "card_abandon") { resetCardState(s); s.step = "MENU"; s.mode = "company"; await setSession(chatId, s); return showMainMenu(chatId, s, "Lot cartes abandonné."); }

  // Recherche manuelle après OCR carte sans SIRET trouvé
  if (data === "search_company") {
    s.step = "SEARCH"; s.mode = "company";
    await setSession(chatId, s);
    return sendMessage(chatId, "🔎 Recherche entreprise :", {
      force_reply: true,
      selective: true,
      input_field_placeholder: "Ex : ETS Poncin 38870 / poncin versoirs",
    });
  }

  // Reprendre la fiche active sans enrichissement supplémentaire
  if (data === "visit:resume_active") {
    s.step = "VISIT_ACTIVE"; s.mode = "company";
    await setSession(chatId, s);
    return postOrReplaceForm(chatId, s);
  }

  // Réessai photo carte après échec OCR
  if (data === "card_retry_photo") {
    // Retire la dernière carte du batch si elle existe (la mauvaise)
    if (Array.isArray(s.card_batch) && s.card_batch.length > 0) {
      s.card_batch.pop();
    }
    s.step = "CARD_COLLECT";
    await setSession(chatId, s);
    return sendMessage(chatId,
      "📷 Renvoie la carte :\n" +
      "• Carte à plat, bien cadrée\n" +
      "• 📎 Via le <b>trombone</b> pour la qualité originale",
      photoReplyKeyboard("📸 Renvoyer la carte", s.card_batch?.length > 0)
    );
  }
  if (data === "card_back") {
    s.step = "CARD_COLLECT"; await setSession(chatId, s);
    const n = s.card_batch?.length || 0;
    return sendMessage(chatId, `🪪 ${n}/${CARD_LIMIT}. Continue ou termine.`, {
      inline_keyboard: [
        [{ text: `✅ Terminer (${n})`, callback_data: "card_finish" }],
        [{ text: "⬅️ Menu", callback_data: "menu" }],
      ],
    });
  }

  // ── Visits ────────────────────────────────────────
  if (data === "visit:new") {
    s.step = "VISIT_START"; s.mode = "company"; await setSession(chatId, s);
    return sendMessage(chatId, "➕ <b>Nouvelle visite</b>\n\nChoisis comment démarrer :", startVisitKeyboard());
  }

  if (data === "visit:start_search") {
    await createNewActiveVisit(chatId, s, "company");
    s.step = "SEARCH"; s.mode = "company"; await setSession(chatId, s);
    return sendMessage(chatId, "🔎 Recherche entreprise :", {
      force_reply: true,
      selective: true,
      input_field_placeholder: "Ex : Lely Lyon / 0476123456 / recyclage 38",
    });
  }

  if (data === "visit:start_card") {
    await createNewActiveVisit(chatId, s, "card");
    s.step = "AWAIT_PHOTO"; s.awaitingField = "card"; await setSession(chatId, s);
    return sendMessage(chatId, "📸 Envoie la photo de la carte.\n<i>📎 Trombone = qualité originale.</i>",
      failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]]));
  }

  if (data === "visit:start_facade") {
    await createNewActiveVisit(chatId, s, "photo");
    s.step = "AWAIT_LOCATION_THEN_PHOTO"; s.awaitingField = "facade";
    await setSession(chatId, s);
    return sendMessage(chatId,
      "📍 Commence par définir le lieu de la visite :",
      {
        keyboard: [
          [{ text: "📍 Définir le lieu", request_location: true }],
          [{ text: "⏭ Passer (sans GPS)", }],
        ],
        resize_keyboard: true,
        one_time_keyboard: true,
      }
    );
  }

  if (data === "visit:show_form") {
    if (s.active_visit_id) return postOrReplaceForm(chatId, s);
    return showMainMenu(chatId, s);
  }

  if (data === "visit:resume") {
    if (!s.active_visit_id) {
      const items = await buildVisitsList(chatId, s);
      if (!items.length) return showMainMenu(chatId, s, "Aucune fiche active.");
      const ok = await openVisitById(chatId, s, items[items.length - 1].visit_id);
      if (!ok) return showMainMenu(chatId, s, "Impossible de rouvrir la fiche.");
      s = await getSession(chatId);
    }
    return postOrReplaceForm(chatId, s);
  }

  if (data === "visits:list") {
    const pack = await buildVisitsListMessage(chatId, s);
    return sendMessage(chatId, pack.text, pack.keyboard);
  }

  if (data === "visit:back_to_active_or_menu") {
    if (s.active_visit_id) return postOrReplaceForm(chatId, s);
    return showMainMenu(chatId, s);
  }

  if (data.startsWith("visit:open_no:")) {
    const no = sanitizeVisitNo(data.split(":").pop());
    const visitId = getVisitIdFromNo(s, no);
    if (!visitId) return sendMessage(chatId, "Fiche introuvable.", failSafeKeyboard());
    const ok = await openVisitById(chatId, s, visitId);
    if (!ok) return sendMessage(chatId, "Impossible de rouvrir cette fiche.", failSafeKeyboard());
    s = await getSession(chatId);
    return postOrReplaceForm(chatId, s);
  }

  if (data === "visit:delete_active") {
    if (!s.active_visit_id) return showMainMenu(chatId, s, "Aucune fiche active.");
    s.pending_delete_visit_id = s.active_visit_id; await setSession(chatId, s);
    return sendMessage(chatId, `🗑 Supprimer la fiche <b>${escapeHtml(currentVisitLabel(s))}</b> ?`, {
      inline_keyboard: [
        [{ text: "✅ Oui, supprimer", callback_data: "visit:delete_confirm" }],
        [{ text: "❌ Non", callback_data: "visit:delete_cancel" }],
      ],
    });
  }

  if (data === "visit:delete_confirm") {
    const visitId = s.pending_delete_visit_id || s.active_visit_id;
    if (!visitId) return showMainMenu(chatId, s, "Aucune fiche à supprimer.");
    await deleteVisitById(chatId, s, visitId);
    s = await getSession(chatId);
    const items = await buildVisitsList(chatId, s);
    if (items.length) {
      const ok = await openVisitById(chatId, s, items[items.length - 1].visit_id);
      if (ok) { s = await getSession(chatId); return postOrReplaceForm(chatId, s); }
    }
    return showMainMenu(chatId, s, "✅ Fiche supprimée.");
  }

  if (data === "visit:delete_cancel") {
    s.pending_delete_visit_id = ""; await setSession(chatId, s);
    if (s.active_visit_id) return postOrReplaceForm(chatId, s);
    return showMainMenu(chatId, s);
  }

  if (data === "visit:add_card") {
    if (!s.active_visit_id) await createNewActiveVisit(chatId, s, "card");
    s.step = "AWAIT_PHOTO"; s.awaitingField = "card"; await setSession(chatId, s);
    return sendMessage(chatId, "📸 Envoie la photo de la carte.\n<i>📎 Trombone = qualité originale.</i>",
      failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]]));
  }

  if (data === "visit:add_facade") {
    if (!s.active_visit_id) await createNewActiveVisit(chatId, s, "photo");
    s.step = "AWAIT_PHOTO"; s.awaitingField = "facade"; await setSession(chatId, s);
    return sendMessage(chatId, "📸 Envoie la photo de la façade.", failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]]));
  }

  if (data === "visit:enrich") {
    log("info", chatId, "ENRICH_BUTTON_CLICKED", { active_visit_id: s.active_visit_id, step: s.step, name: s.draft?.name });
    if (!s.active_visit_id) return showMainMenu(chatId, s, "Aucune fiche active.");
    const spinner = await sendSpinner(chatId, "⏳ Enrichissement en cours…");
    s.draft = await enrichDraftFromCurrentData(s.draft || {});

    // Validation adresse BAN
    if (s.draft.address || s.draft.city) {
      const ban = await banNormalizeAddress(s.draft.address, s.draft.postal_code, s.draft.city).catch(() => null);
      if (ban && ban.score > 0.7) {
        if (ban.address) s.draft.address     = ban.address;
        if (ban.postal_code) s.draft.postal_code = ban.postal_code;
        if (ban.city) s.draft.city          = ban.city;
      }
    }

    s.step = "VISIT_ACTIVE"; await setSession(chatId, s);
    await deleteSpinner(chatId, spinner);
    return postOrReplaceForm(chatId, s);
  }

  // ── Qualification ─────────────────────────────────
  if (data.startsWith("qualif:")) {
    const q = data.split(":")[1];
    if (q === "chaud" || q === "tiede" || q === "froid") {
      s.draft.qualification = q;
    } else {
      s.draft.qualification = "";
    }
    await setSession(chatId, s);

    // Si on était en attente de qualification avant save
    if (s._pending_qualif_save) {
      s._pending_qualif_save = false;
      await setSession(chatId, s);
      // Lancer le save
      const res = await saveCurrentVisitAsProspect(chatId, s);
      if (!res.ok && res.reason === "duplicate") {
        return sendMessage(chatId,
          `⚠️ Entreprise similaire déjà saisie (${res.duplicate.reason}).\n\nConfirmer ?`,
          duplicateConfirmKeyboard());
      }
      if (!res.ok && res.reason === "validation") {
        return sendMessage(chatId, `⚠️ Impossible d'enregistrer :\n• ${res.errors.join("\n• ")}`, failSafeKeyboard());
      }
      s = await getSession(chatId);
      return sendMessage(chatId, res.edit ? "✅ Fiche mise à jour." : "✅ Fiche enregistrée.", {
        inline_keyboard: [
          [{ text: "➕ Nouvelle visite", callback_data: "visit:new" }],
          [{ text: "📋 Liste du jour",   callback_data: "visits:list" }],
          [{ text: "⬅️ Menu",            callback_data: "menu" }],
        ],
      });
    }

    return postOrReplaceForm(chatId, s);
  }

  // ── Search / Pick / Manual ────────────────────────
  if (data.startsWith("company:")) {
    const idx = parseInt(data.split(":")[1], 10);
    const chosen = s.companyChoices?.[idx];
    if (!chosen) return sendMessage(chatId, "Choix invalide.", failSafeKeyboard());

    if (!s.active_visit_id) await createNewActiveVisit(chatId, s, "company");

    const spinner = await sendSpinner(chatId, "⏳ Enrichissement en cours…");

    // Sauvegarder le nom Places comme nom commercial (avant enrichissement)
    const placesNom = chosen._source === "places" ? (chosen.nom || "") : "";

    const enriched = await enrichCompany(chosen);
    s.draft = applyCompanyToDraft(s.draft || {}, enriched);

    // Si le nom juridique diffère du nom Places → conserver le nom Places comme nom_commercial
    // Mais seulement si le nom Places ressemble à une enseigne connue ou contient le nom recherché
    if (placesNom && s.draft.name && normalizeForMatch(placesNom) !== normalizeForMatch(s.draft.name)) {
      const qNorm2 = normalizeForMatch(s.last_search_query || "");
      const pNorm2 = normalizeForMatch(placesNom);
      const isKB2 = [...KNOWN_BRANDS].some(b => pNorm2.includes(normalizeForMatch(b)));
      const hasCommon2 = !qNorm2 || pNorm2.includes(qNorm2.slice(0,2));
      if (isKB2 || hasCommon2) s.draft.nom_commercial = placesNom;
    }

    // Validation adresse BAN
    if (s.draft.address || s.draft.city) {
      const ban = await banNormalizeAddress(s.draft.address, s.draft.postal_code, s.draft.city).catch(() => null);
      if (ban && ban.score > 0.7) {
        if (ban.address) s.draft.address = ban.address;
        if (ban.postal_code) s.draft.postal_code = ban.postal_code;
        if (ban.city) s.draft.city = ban.city;
      }
    }

    await deleteSpinner(chatId, spinner);
    s.step = "VISIT_ACTIVE"; s.awaitingField = null;
    await setSession(chatId, s);
    await persistActiveVisit(chatId, s); // Sauvegarder nom_commercial dans session_visit

    // Auto-save si fiche complète
    if (isFicheCompleteForAutoSave(s.draft)) {
      const summary = buildAutoSaveSummary(s.draft);
      return sendMessage(chatId, summary + "\n\nSauvegarder cette fiche ?", autoSaveKeyboard());
    }

    return postOrReplaceForm(chatId, s);
  }

  if (data === "save_direct") {
    // Auto-save 1-clic : demander qualification d'abord
    s._pending_qualif_save = true;
    await setSession(chatId, s);
    return sendMessage(chatId, "⭐ Comment qualifier ce prospect ?", qualificationKeyboard());
  }

  if (data === "company:none") {
    s.step = "SEARCH"; await setSession(chatId, s);
    return sendMessage(chatId, "OK. Refais la recherche ou saisis manuellement.", failSafeKeyboard());
  }

  if (data === "manual_entry") {
    if (!s.active_visit_id) await createNewActiveVisit(chatId, s, "manual");
    s.step = "MANUAL_NAME"; await setSession(chatId, s);
    const hint = s.last_search_query ? `\n\n<i>(dernier essai : ${escapeHtml(s.last_search_query)})</i>` : "";
    return sendMessage(chatId, `✏️ Saisis le <b>nom exact</b> de l'entreprise :${hint}`,
      failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]]));
  }

  // ── Edit last ─────────────────────────────────────
  if (data === "edit:last") {
    const lastKey = await getLastProspectKey(chatId);
    if (!lastKey) return sendMessage(chatId, "Aucun prospect récent.", failSafeKeyboard());
    const record = await kv().get(lastKey, { type: "json" });
    if (!record) return sendMessage(chatId, "Prospect introuvable (expiré).", failSafeKeyboard());
    if (!s.active_visit_id) await createNewActiveVisit(chatId, s, "edit");
    s.draft = normalizeDraft({
      ...s.draft, _source: "edit", _edit_key: lastKey,
      name: record.name || "", address: record.address || "",
      postal_code: record.postal_code || "", city: record.city || "",
      siret: record.siret || "", naf: record.naf || "",
      secteur_activite: record.secteur_activite || "", effectif: record.effectif || "",
      date_creation: record.date_creation || "", capital: record.capital || "",
      dirigeant: record.dirigeant || "", interlocuteur: record.interlocuteur || "",
      contact_civility: record.contact_civility || "",
      contact_firstname: record.contact_firstname || "", contact_lastname: record.contact_lastname || "",
      phone: record.phone || "", phone2: record.phone2 || "",
      email: record.email || "", website: record.website || "",
      resume: record.resume || "", commande: record.commande || "",
      qualification: record.qualification || "",
      card_photo_file_id: record.card_photo_file_id || "", card_photo_url: record.card_photo_url || "",
      facade_photo_file_id: record.facade_photo_file_id || "", facade_photo_url: record.facade_photo_url || "",
      source_mode: record.source_mode || "edit",
    });
    s.step = "VISIT_ACTIVE"; s.awaitingField = null;
    await setSession(chatId, s);
    return postOrReplaceForm(chatId, s);
  }

  // ── Field edit ────────────────────────────────────
  if (data.startsWith("field:")) {
    const field = data.split(":")[1];
    if (!s.active_visit_id) await createNewActiveVisit(chatId, s, "company");

    if (field === "card") {
      s.step = "AWAIT_PHOTO"; s.awaitingField = "card"; await setSession(chatId, s);
      return sendMessage(chatId, "📸 Envoie la photo de la carte.\n<i>📎 Trombone = qualité originale.</i>",
        failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]]));
    }

    if (field === "facade") {
      s.step = "AWAIT_PHOTO"; s.awaitingField = "facade"; await setSession(chatId, s);
      return sendMessage(chatId, "📸 Envoie la photo de la façade.", failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]]));
    }

    if (field === "qualification") {
      s._pending_qualif_save = false; await setSession(chatId, s);
      return sendMessage(chatId, "⭐ Qualifie ce prospect :", qualificationKeyboard());
    }

    const ALLOWED = [
      "name", "interlocuteur", "phone", "phone2", "email", "website",
      "address", "postal_code", "city", "resume", "commande", "contact_civility",
    ];
    if (!ALLOWED.includes(field)) return sendMessage(chatId, "Champ inconnu.", failSafeKeyboard());

    s.step = "AWAIT_TEXT"; s.awaitingField = field; await setSession(chatId, s);

    const labels = {
      name: "le nom de l'entreprise",
      interlocuteur: "le nom de l'interlocuteur (ex: Jean Dupont)",
      phone: "le téléphone fixe",
      phone2: "le portable",
      email: "l'email",
      website: "le site web",
      address: "l'adresse",
      postal_code: "le code postal",
      city: "la ville",
      resume: "le résumé d'entretien",
      commande: "la commande",
      contact_civility: "la civilité (M. ou Mme)",
    };
    return sendMessage(
      chatId,
      `✏️ Saisis ${labels[field]} :`,
      {
        force_reply: true,
        selective: true,
        input_field_placeholder: labels[field] || "Saisis la valeur",
      }
    );
  }

  // ── Save ──────────────────────────────────────────
  if (data === "save") {
    if (!s.agency || !s.initials) return sendMessage(chatId, "Session incomplète. Tape /start.", failSafeKeyboard());
    if (!s.active_visit_id) return showMainMenu(chatId, s, "Aucune fiche active.");

    // Validation minimum
    const errors = validateDraftMinimum(normalizeDraft(s.draft || {}));
    if (errors.length) {
      return sendMessage(chatId,
        `⚠️ <b>Impossible d'enregistrer</b>\n\n` +
        `Champs manquants :\n• ${errors.join("\n• ")}`
      );
    }

    // Qualification avant save
    s._pending_qualif_save = true;
    await setSession(chatId, s);
    return sendMessage(chatId, "⭐ Comment qualifier ce prospect ?", qualificationKeyboard());
  }

  if (data === "save_force") {
    const res = await saveCurrentVisitAsProspectForced(chatId, s);
    if (!res.ok) return sendMessage(chatId, "Impossible d'enregistrer la fiche.", failSafeKeyboard());
    s = await getSession(chatId);
    return sendMessage(chatId, "✅ Fiche enregistrée malgré doublon probable.", {
      inline_keyboard: [
        [{ text: "➕ Nouvelle visite", callback_data: "visit:new" }],
        [{ text: "📋 Liste du jour",   callback_data: "visits:list" }],
        [{ text: "⬅️ Menu",            callback_data: "menu" }],
      ],
    });
  }

  if (data === "save_cancel_duplicate") {
    s._pending_duplicate_record = null; await setSession(chatId, s);
    return postOrReplaceForm(chatId, s);
  }

  // ── Photo lot ─────────────────────────────────────
  if (data === "photo_finish") {
    const batch = Array.isArray(s.photo_batch) ? [...s.photo_batch] : [];
    if (!batch.length) return sendMessage(chatId, "⚠️ Aucune photo reçue.", failSafeKeyboard());
    if (!s.photo_city) return sendMessage(chatId, "⚠️ Ville/secteur non défini.", failSafeKeyboard());
    if (!s.agency || !s.initials) return sendMessage(chatId, "⚠️ Session incomplète.", failSafeKeyboard());

    const date = ymdFR();
    const ts0  = s.photo_ts0 || Date.now();

    for (let i = 0; i < batch.length; i++) {
      const item = batch[i];
      const key  = `photo:${date}:${s.agency}:${ts0}:${chatId}:${i + 1}`;

      await kv().put(key, JSON.stringify({
        type: "photo", kind: "location",
        chatId: String(chatId), agency: String(s.agency), user: String(s.initials), initials: String(s.initials),
        date, ts: Number(item.ts || Date.now()),
        file_id: String(item.file_id || ""), city: String(s.photo_city),
        geo: item.geo || null, comment: item.comment || null,
        gemini_company: item.gemini_company || "", gemini_city: item.gemini_city || "",
        gemini_activity: item.gemini_activity || "", status: "pending",
      }), { expirationTtl: TTL_PROSPECT });

      const { draft } = await createProspectFromFacade(chatId, s, item.file_id, s.photo_city, item.comment || null);
      s = await getSession(chatId);
      s.draft = attachPhotoKeyToDraft(s.draft, key);
      await setSession(chatId, s);
      const saveRes = await saveCurrentVisitAsProspectForced(chatId, s);
      if (!saveRes.ok) log("error", chatId, "AUTO_SAVE_FACADE_FAILED", { key });
      s = await getSession(chatId);
    }

    s._last_photo_city = s.photo_city;
    scheduleBackgroundTask(event, dispatchGitHubWithRetry(date, s.agency, s.initials), "GITHUB_DISPATCH_BG_FAILED");

    resetPhotoState(s); s.step = "MENU"; s.mode = "company";
    await setSession(chatId, s);
    // Supprime le clavier photo bas d'écran avant d'afficher le menu
    await sendMessage(chatId, `✅ ${batch.length} photo(s) enregistrée(s).`, removeReplyKeyboard());
    return showMainMenu(chatId, s);
  }

  // ── Card lot ──────────────────────────────────────
  if (data === "card_finish") {
    const batch = Array.isArray(s.card_batch) ? [...s.card_batch] : [];
    if (!s.agency || !s.initials) return sendMessage(chatId, "⚠️ Session incomplète.", failSafeKeyboard());

    // Sauvegarder la fiche active en cours si elle existe
    if (s.active_visit_id && s.draft?.name) {
      await saveCurrentVisitAsProspectForced(chatId, s);
      s = await getSession(chatId);
    }

    const date = ymdFR();
    scheduleBackgroundTask(event, dispatchGitHubWithRetry(date, s.agency, s.initials), "GITHUB_DISPATCH_BG_FAILED");
    resetCardState(s); s.step = "MENU"; s.mode = "company";
    await setSession(chatId, s);
    await sendMessage(chatId, `✅ ${batch.length || 1} carte(s) enregistrée(s).`, removeReplyKeyboard());
    return showMainMenu(chatId, s);
  }

  // ── Mode photo lot ───────────────────────────
  if (data === "mode:photo") {
    s.mode = "photo";
    resetPhotoState(s);
    s.photo_ts0 = Date.now();
    await setSession(chatId, s);

    const lastCity = s._last_photo_city || "";
    const geo = s.current_geo;
    const geoAge = geo ? Math.round((Date.now() - (geo.at || 0)) / 60000) : null;
    const geoFresh = geoAge !== null && geoAge < 60; // GPS de moins d'1h

    // Si GPS disponible et récent → reverse geocode automatique
    if (geoFresh && geo.lat && geo.lon) {
      const spinner = await sendMessage(chatId, "📍 Localisation en cours…");
      const place = await reverseGeocode(geo.lat, geo.lon);
      await deleteMessage(chatId, spinner?.message_id).catch(() => {});

      if (place?.city) {
        const cityLabel = [place.postal_code, place.city].filter(Boolean).join(" ");
        const kb = {
          inline_keyboard: [
            [{ text: `✅ ${cityLabel} (GPS ${geoAgeLabel(geo.at)})`, callback_data: "photo_city_gps" }],
          ],
        };
        if (lastCity && lastCity !== place.city) {
          kb.inline_keyboard.push([{ text: `🔁 Réutiliser : ${escapeHtml(lastCity)}`, callback_data: "photo_city_reuse" }]);
        }
        kb.inline_keyboard.push([{ text: "🏙️ Autre ville / secteur", callback_data: "photo_city_new" }]);
        kb.inline_keyboard.push([{ text: "⬅️ Menu", callback_data: "menu" }]);
        s._gps_city        = place.city;
        s._gps_postal_code = place.postal_code;
        await setSession(chatId, s);
        return sendMessage(chatId,
          `📸 <b>Mode Photo lieu</b>\n\n📍 Position détectée :`,
          kb
        );
      }

      // Reverse geocoding échoué → fallback saisie manuelle
      return sendMessage(chatId,
        `📸 <b>Mode Photo lieu</b>\n\n⚠️ Ville non déterminée depuis le GPS.\nSaisis le secteur manuellement :`,
        {
          inline_keyboard: [
            ...(lastCity ? [[{ text: `🔁 Réutiliser : ${escapeHtml(lastCity)}`, callback_data: "photo_city_reuse" }]] : []),
            [{ text: "🏙️ Saisir ville / secteur", callback_data: "photo_city_new" }],
            [{ text: "⬅️ Menu", callback_data: "menu" }],
          ],
        }
      );
    }

    // Pas de GPS ou trop vieux → proposer d'envoyer la position ou saisir manuellement
    const kb = { inline_keyboard: [] };
    if (lastCity) {
      kb.inline_keyboard.push([{ text: `🔁 Même secteur : ${escapeHtml(lastCity)}`, callback_data: "photo_city_reuse" }]);
    }
    kb.inline_keyboard.push([{ text: "📍 Envoyer ma position GPS", callback_data: "photo_request_gps" }]);
    kb.inline_keyboard.push([{ text: "🏙️ Saisir ville / secteur", callback_data: "photo_city_new" }]);
    kb.inline_keyboard.push([{ text: "⬅️ Menu", callback_data: "menu" }]);

    return sendMessage(chatId, `📸 <b>Mode Photo lieu</b>\n\nQuel secteur pour ce lot ?`, kb);
  }

  if (data === "photo_request_gps") {
    // Demander la position GPS via le bouton Telegram
    return sendMessage(chatId,
      "📍 Envoie ta position pour détecter la ville automatiquement :",
      locationRequestKeyboard()
    );
  }

  if (data === "photo_city_gps") {
    // Utiliser la ville détectée par GPS
    s.photo_city = [s._gps_postal_code, s._gps_city].filter(Boolean).join(" ") || s._gps_city || "";
    s._last_photo_city = s.photo_city;
    s.step = "PHOTO_COLLECT";
    await setSession(chatId, s);
    return sendMessage(chatId,
      `✅ Secteur GPS : <b>${escapeHtml(s.photo_city)}</b>\n\n` +
      `📎 <b>Envoie via le TROMBONE</b> pour la qualité optimale.`,
      photoReplyKeyboard("📸 Envoyer une façade", false)
    );
  }

  if (data === "photo_city_reuse") {
    s.photo_city = s._last_photo_city || "";
    s.step = "PHOTO_COLLECT";
    await setSession(chatId, s);
    return sendMessage(chatId,
      `✅ Secteur : <b>${escapeHtml(s.photo_city)}</b>\n\n` +
      `📎 <b>Envoie via le TROMBONE</b> (Trombone → Fichier → Galerie)\n` +
      `<i>Evite l'envoi normal qui recadre l'enseigne.</i>`,
      photoReplyKeyboard("📸 Envoyer une façade (via trombone SVP)", false)
    );
  }

  if (data === "photo_city_new") {
    s.step = "PHOTO_CITY_INPUT";
    await setSession(chatId, s);
    return sendMessage(chatId, `🏙️ Quelle ville / secteur ?`,
      failSafeKeyboard([[{ text: "⬅️ Menu", callback_data: "menu" }]]));
  }

  // ── Mode card lot ─────────────────────────────
  if (data === "mode:card") {
    s.mode = "card";
    resetCardState(s);
    s.card_ts0 = Date.now();
    s.step = "CARD_COLLECT";
    await setSession(chatId, s);
    return sendMessage(chatId,
      `🪪 <b>Mode Cartes de visite</b>\n\n` +
      `📎 <b>Utilise le TROMBONE</b> pour envoyer chaque carte :\n` +
      `<i>Trombone → Photo → choisir dans la galerie ou appareil photo</i>\n\n` +
      `La qualité d'image est cruciale pour la lecture automatique.\n` +
      `Les photos compressées (envoi normal) donnent de mauvais résultats.\n\n` +
      `Limite : ${CARD_LIMIT} cartes par lot. Tape <b>.</b> pour passer le commentaire.`,
      photoReplyKeyboard("📸 Prendre / envoyer une carte", false)
    );
  }

  // ── Close ─────────────────────────────────────────
  if (data === "admin:reset_cache") {
    // Réservé à SL uniquement
    if ((s.initials || "").toUpperCase() !== "SL") {
      return sendMessage(chatId, "⛔ Action réservée à l'administrateur.");
    }
    const spinner = await sendSpinner(chatId, "🔄 Nettoyage du cache en cours…");
    let deleted = 0;
    try {
      // Vider enseigne_cache:*
      const enseigneKeys = await listAllKeys("enseigne_cache:");
      await Promise.all(enseigneKeys.map(k => kv().delete(typeof k === "string" ? k : k.name)));
      deleted += enseigneKeys.length;

      // Vider session_visit:*
      const visitKeys = await listAllKeys(`session_visit:${chatId}:`);
      await Promise.all(visitKeys.map(k => kv().delete(typeof k === "string" ? k : k.name)));
      deleted += visitKeys.length;

      // Réinitialiser la session (garder agence/initiales)
      s.visits_index   = [];
      s.active_visit_id = "";
      s.draft          = {};
      await setSession(chatId, s);

    } catch (err) {
      log("error", chatId, "ADMIN_RESET_CACHE_ERROR", { err: err?.message });
    }
    await deleteSpinner(chatId, spinner);
    return showMainMenu(chatId, s, `✅ Cache nettoyé — ${deleted} entrée(s) supprimée(s). Bonne journée !`);
  }

  if (data === "close") {
    if (!s.agency || !s.initials) return sendMessage(chatId, "Session incomplète.", failSafeKeyboard());
    s.step = "CLOSE_CLIENTS";
    s.close_stats = { visits_clients: "", visits_prospects: "", commandes: "" };
    await setSession(chatId, s);
    return sendMessage(chatId, "🔢 Combien de <b>CLIENTS</b> visités aujourd'hui ?");
  }

  if (data === "close_cancel") {
    s.step = "MENU"; s.close_stats = null; await setSession(chatId, s);
    return showMainMenu(chatId, s, "❌ Clôture annulée.");
  }

  if (data === "close_confirm") {
    const date = ymdFR();

    // Sauvegarder TOUTES les visites de la session (optimisé — minimum de KV puts)
    const visitsIndex = Array.isArray(s.visits_index) ? s.visits_index : [];
    let savedCount = 0;

    // Récupérer toutes les visites en parallèle pour réduire les appels KV
    const storedVisits = await Promise.all(
      visitsIndex.map(entry => entry?.visit_id
        ? kv().get(buildVisitStoreKey(chatId, entry.visit_id), { type: "json" }).catch(() => null)
        : Promise.resolve(null)
      )
    );

    for (let i = 0; i < visitsIndex.length; i++) {
      const entry = visitsIndex[i];
      const stored = storedVisits[i];
      if (!stored?.draft?.name) continue;

      // Vérifier si déjà sauvegardé (évite les KV puts inutiles)
      const existingKey = stored.prospect_key || stored.draft?._edit_key;
      if (existingKey) { savedCount++; continue; }

      // Sauvegarder comme prospect
      try {
        const tempS = { ...s, draft: stored.draft, active_visit_id: entry.visit_id, active_visit_no: entry.no || 0 };
        const saveRes = await saveCurrentVisitAsProspectForced(chatId, tempS);
        if (saveRes.ok) {
          savedCount++;
          log("info", chatId, "CLOSE_SAVE_VISIT", { visitId: entry.visit_id, name: stored.draft?.name });
        }
      } catch (err) {
        log("error", chatId, "CLOSE_SAVE_VISIT_FAILED", { visitId: entry.visit_id, err: err?.message });
      }
    }

    // Sauvegarder aussi la fiche active courante si pas dans l'index
    if (s.active_visit_id && s.draft?.name) {
      const alreadySaved = visitsIndex.some(e => e?.visit_id === s.active_visit_id);
      if (!alreadySaved) {
        await saveCurrentVisitAsProspectForced(chatId, s);
        savedCount++;
      }
      s = await getSession(chatId);
    }

    log("info", chatId, "CLOSE_ALL_VISITS_SAVED", { count: savedCount });

    const closeKey = `close:${date}:${s.agency}:${s.initials}:${Date.now()}:${chatId}`;
    await kv().put(closeKey, JSON.stringify({
      date, agency: s.agency, initials: s.initials,
      visits_clients:   s.close_stats?.visits_clients   || "",
      visits_prospects: s.close_stats?.visits_prospects || "",
      commandes:        s.close_stats?.commandes        || "",
      declaratif: true,
      closed_at: new Date().toISOString(),
    }), { expirationTtl: TTL_PROSPECT });

    scheduleBackgroundTask(event, dispatchGitHubWithRetry(date, s.agency, s.initials), "GITHUB_DISPATCH_BG_FAILED");

    s.draft = normalizeDraft({});
    s.companyChoices = []; s.step = "MENU"; s.awaitingField = null;
    s.close_stats = null; s._pending_duplicate_record = null;
    await setSession(chatId, s);
    return showMainMenu(chatId, s, "✅ Session clôturée. Export individuel lancé.");
  }

  log("warn", chatId, "UNKNOWN_CB", { data });
  return sendMessage(chatId, "Action inconnue. Tape /aide.", failSafeKeyboard());
}


// =====================================================
// REMIND
// =====================================================

async function handleRemind(req) {
  const ok = await checkExportToken(req);
  if (!ok) return new Response("Unauthorized", { status: 401 });

  const queueStats = await processGitHubQueue();
  const date = ymdFR();
  const sessionKeys = await listAllKeys("session:");
  let sent = 0;

  // Stats pour résumé manager
  const agencyStats = {};

  // Lecture parallèle de toutes les sessions
  const sessions = await batchGet(sessionKeys);

  for (let idx = 0; idx < sessionKeys.length; idx++) {
    const s      = sessions[idx];
    const k      = sessionKeys[idx];
    if (!s || !s.agency || !s.initials) continue;

    const parts  = String(k.name || "").split(":");
    const chatId = parts[1];
    if (!chatId) continue;

    // Stats agence pour résumé manager
    const ag  = s.agency;
    const ini = s.initials;
    if (!agencyStats[ag]) agencyStats[ag] = { total: 0, users: {} };
    if (!agencyStats[ag].users[ini]) agencyStats[ag].users[ini] = { total: 0, hot: 0, warm: 0, cold: 0, closed: false };

    // Compter les fiches du jour
    const dayKeys = await listAllKeys(`prospect:${date}:${ag}:`);
    const dayRecs = await batchGet(dayKeys);
    for (const rec of dayRecs) {
      if (!rec) continue;
      if (String(rec.initials || "").toUpperCase() !== ini) continue;
      agencyStats[ag].users[ini].total++;
      agencyStats[ag].total++;
      if (rec.qualification === QUALIF_HOT)  agencyStats[ag].users[ini].hot++;
      if (rec.qualification === QUALIF_WARM) agencyStats[ag].users[ini].warm++;
      if (rec.qualification === QUALIF_COLD) agencyStats[ag].users[ini].cold++;
    }

    // Vérifier clôture
    const closeList = await listAllKeys(`close:${date}:${ag}:${ini}:`);
    if (closeList.length > 0) { agencyStats[ag].users[ini].closed = true; continue; }

    // Vérifier activité
    const hasActivity = agencyStats[ag].users[ini].total > 0;
    if (!hasActivity) continue;

    // Envoyer rappel
    try {
      await sendMessage(chatId,
        `⏰ <b>Rappel clôture</b>\n\nTu as ${agencyStats[ag].users[ini].total} fiche(s) enregistrée(s) aujourd'hui mais la session n'est pas clôturée.\n\n💡 Pense à clôturer avant 18h.`,
        { inline_keyboard: [
          [{ text: "✅ Clore maintenant", callback_data: "close" }],
          [{ text: "📋 Résumé session",   callback_data: "summary" }],
        ]}
      );
      sent++;
    } catch (e) {
      log("error", chatId, "REMIND_SEND_FAILED", { err: e?.message });
    }
  }

  // Résumé manager
  await sendManagerDailySummary(date, agencyStats);

  log("info", null, "REMIND_DONE", { sent, date, queueStats });
  return json({ ok: true, sent, date, queueStats });
}


// =====================================================
// DUMP
// =====================================================

async function handleDump(req, url) {
  const ok = await checkExportToken(req);
  if (!ok) return new Response("Unauthorized", { status: 401 });

  const date = url.searchParams.get("date");
  const kind = (url.searchParams.get("kind") || "prospects").toLowerCase();
  if (!date) return new Response("Missing date", { status: 400 });

  const PREFIX_MAP = {
    prospects: `prospect:${date}:`,
    closes:    `close:${date}:`,
    photos:    `photo:${date}:`,
    cards:     `card:${date}:`,
  };
  const prefix = PREFIX_MAP[kind];
  if (!prefix) return new Response("Invalid kind (prospects|closes|photos|cards)", { status: 400 });

  let keys = [];
  try { keys = await listAllKeys(prefix); }
  catch (e) {
    log("error", null, "DUMP_LIST_FAILED", { prefix, err: e?.message });
    return new Response("Dump list failed", { status: 500 });
  }

  // Lecture parallèle
  const values = await Promise.all(
    keys.map(async (k) => {
      try { return await kv().get(k.name); }
      catch (_) { return null; }
    })
  );

  const out = [];
  for (const v of values) {
    if (!v) continue;
    if (typeof v === "string") { const t = v.trim(); if (t) out.push(t); continue; }
    try { out.push(JSON.stringify(v)); } catch (_) {}
  }

  return new Response(out.join("\n"), {
    status: 200,
    headers: {
      "content-type": "text/plain; charset=utf-8",
      "x-record-count": String(out.length),
      "x-list-complete": "true",
      "cache-control": "no-store",
    },
  });
}