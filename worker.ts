/**
 * Prospection Intake Bot (Telegram) - Cloudflare Worker
 * V1: Saisie minimale + validation -> création d'une GitHub Issue (inbox)
 */
export interface Env {
  TELEGRAM_TOKEN: string;
  TELEGRAM_WEBHOOK_SECRET: string;
  KV: KVNamespace;
  GITHUB_REPO: string; // "owner/repo"
  GITHUB_TOKEN: string; // PAT
  // Optionnel: liste blanche d'utilisateurs Telegram autorisés (IDs séparés par virgule)
  ALLOWED_USERS?: string;
}

type Draft = {
  id: string;
  agency?: string; // GR / VR / GRS / SLS
  createdAt: string;
  chatId: number;
  userId: number;
  userName?: string;
  name?: string;
  city?: string;
  phone?: string;
  interlocuteur?: string;
  interlocuteur_email?: string;
  resume?: string;
  commande?: string;
  photoFileId?: string;
  enriched?: {
    address?: string;
    postalCode?: string;
    city?: string;
    phone1?: string;
    phone2?: string;
    email?: string;
    siret?: string;
    naf?: string;
    website?: string;
  };
};

type ConvState = {
  mode?: "NEW";
  step?: "TYPE"|"WAIT_NAME"|"WAIT_CITY"|"WAIT_PHONE"|"WAIT_INTERLOC"|"WAIT_EMAIL"|"WAIT_RESUME"|"WAIT_CMD"|"WAIT_PHOTO";
  kind?: "MANUAL"|"PHONE"|"PHOTO";
  data?: Record<string, any>;
};

function convKey(chatId: number, userId: number) {
  return `conv:${chatId}:${userId}`;
}
async function loadConv(env: Env, chatId: number, userId: number): Promise<ConvState> {
  const raw = await env.KV.get(convKey(chatId, userId));
  if (!raw) return {};
  try { return JSON.parse(raw); } catch { return {}; }
}
async function saveConv(env: Env, chatId: number, userId: number, s: ConvState) {
  await env.KV.put(convKey(chatId, userId), JSON.stringify(s));
}
async function clearConv(env: Env, chatId: number, userId: number) {
  await env.KV.delete(convKey(chatId, userId));
}

async function sendMainMenu(env: Env, chatId: number, text: string) {
  await tg("sendMessage", env.TELEGRAM_TOKEN, {
    chat_id: chatId,
    text,
    reply_markup: {
      keyboard: [
        [{ text: "➕ Nouvelle entrée" }],
        [{ text: "📊 Bilan session" }, { text: "🔒 Clore session" }],
      ],
      resize_keyboard: true,
      one_time_keyboard: false,
    },
  });
}

async function askNewEntryType(env: Env, chatId: number) {
  await tg("sendMessage", env.TELEGRAM_TOKEN, {
    chat_id: chatId,
    text: "➕ Nouvelle entrée — choisis :",
    reply_markup: {
      keyboard: [
        [{ text: "✏️ Saisie manuelle" }],
        [{ text: "📞 Nom + téléphone" }],
        [{ text: "📷 Photo carte de visite" }],
        [{ text: "↩️ Menu" }],
      ],
      resize_keyboard: true,
      one_time_keyboard: true,
    },
  });
}



const DRAFT_PREFIX = "draft:";

function nowISO() { return new Date().toISOString(); }
function uid() { return crypto.randomUUID(); }

async function tg(method: string, token: string, payload: any) {
  const res = await fetch(`https://api.telegram.org/bot${token}/${method}`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await res.json() as { ok: boolean; result?: any };
  if (!data.ok) throw new Error(`Telegram API error: ${method} -> ${JSON.stringify(data)}`);
  return data.result;
}

function parseQuickLine(text: string) {
  const t = text.trim();
  const phoneMatch = t.match(/(\+33\s?[\d\s\.]{8,}|0[\d\s\.]{8,})/);
  let phone = phoneMatch ? phoneMatch[1].replace(/\s|\./g, "") : undefined;

  let name: string | undefined;
  let city: string | undefined;

  const parts = t.split(/;|,|-|–/).map(s => s.trim()).filter(Boolean);
  if (parts.length >= 2) {
    name = parts[0];
    city = parts.slice(1).join(" ");
  } else {
    if (phone) {
      name = t.replace(phoneMatch![0], "").replace(/tel\s?:?/i, "").trim();
    } else {
      name = t;
    }
  }
  return { name, city, phone };
}


function isAllowedUser(env: Env, userId: number) {
  const raw = (env.ALLOWED_USERS || "").trim();
  if (!raw) return true; // si vide, tout le monde est autorisé
  const set = new Set(raw.split(",").map(s => s.trim()).filter(Boolean));
  return set.has(String(userId));
}

function buildPreview(d: Draft) {
  const e = d.enriched || {};
  return [
    `🧾 *Brouillon* #${d.id}`,
    ``,
    `*NOM* : ${d.name || "—"}`,
    `*VILLE* : ${d.city || "—"}`,
    `*TEL* : ${d.phone || "—"}`,
    `*AGENCE* : ${d.agency || "—"}`,
    `*USER* : ${d.userName || "—"} (${d.userId})`,
    `*Interlocuteur* : ${d.interlocuteur || "—"}`,
    `*Mail interlocuteur* : ${d.interlocuteur_email || "—"}`,
    ``,
    `*Résumé entretien* : ${d.resume ? d.resume.slice(0, 800) : "—"}`,
    `*Commande* : ${d.commande ? d.commande.slice(0, 800) : "—"}`,
    ``,
    `*Complétion auto (V1)*`,
    `Adresse : ${e.address || "—"}`,
    `CP : ${e.postalCode || "—"}`,
    `Email : ${e.email || "—"}`,
    `SIRET : ${e.siret || "—"}`,
    `NAF : ${e.naf || "—"}`,
    `Site : ${e.website || "—"}`,
    ``,
    d.photoFileId ? `📸 Carte de visite : *oui* (OCR V2)` : `📸 Carte de visite : *non*`,
  ].join("\n");
}

// Aperçu utilisé par le flow guidé (NEW)
function renderDraftPreview(d: Draft) {
  return buildPreview(d);
}

function keyboard(draftId: string) {
  return {
    inline_keyboard: [
      [{ text: "✅ Valider", callback_data: `validate:${draftId}` },
       { text: "✏️ Ajouter résumé", callback_data: `addresume:${draftId}` }],
      [{ text: "🧾 Ajouter commande", callback_data: `addcmd:${draftId}` },
       { text: "🗑️ Annuler", callback_data: `cancel:${draftId}` }],
    ],
  };
}

async function createOrGetDraft(env: Env, chatId: number, userId: number, userName?: string) {
  const key = `${DRAFT_PREFIX}${chatId}`;
  const raw = await env.KV.get(key);
  if (raw) return JSON.parse(raw) as Draft;
  const d: Draft = { id: uid().slice(0, 8), createdAt: nowISO(), chatId, userId, userName };
  await env.KV.put(key, JSON.stringify(d), { expirationTtl: 60 * 60 * 24 * 7 });
  return d;
}

// Alias sémantique: garantit qu'un brouillon existe pour ce chat.
async function loadDraft(env: Env, chatId: number, userId: number, userName?: string) {
  return createOrGetDraft(env, chatId, userId, userName);
}

async function saveDraft(env: Env, d: Draft) {
  await env.KV.put(`${DRAFT_PREFIX}${d.chatId}`, JSON.stringify(d), { expirationTtl: 60 * 60 * 24 * 7 });
}

async function clearDraft(env: Env, chatId: number) {
  await env.KV.delete(`${DRAFT_PREFIX}${chatId}`);
}

async function enrichV1(d: Draft) {
  const enriched: Draft["enriched"] = {};
  if (d.city) enriched.city = d.city;
  if (d.phone) enriched.phone1 = d.phone;
  return enriched;
}

async function githubCreateIssueWithLabels(env: Env, title: string, body: string, labels: string[]) {
  const [owner, repo] = env.GITHUB_REPO.split("/");
  const res = await fetch(`https://api.github.com/repos/${owner}/${repo}/issues`, {
    method: "POST",
    headers: {
      "accept": "application/vnd.github+json",
      "authorization": `Bearer ${env.GITHUB_TOKEN}`,
      "user-agent": "prospection-intake-worker",
      "content-type": "application/json",
    },
    body: JSON.stringify({ title, body, labels }),
  });
  if (!res.ok) throw new Error(`GitHub issue create failed: ${res.status} ${await res.text()}`);
  return await res.json();
}

async function githubCreateIssue(env: Env, title: string, body: string) {
  return githubCreateIssueWithLabels(env, title, body, ["intake"]);
}

function buildIssueBody(d: Draft) {
  const e = d.enriched || {};
  const safe = (s?: string) => (s ? s.replace(/\r/g, "").trim() : "");
  return [
    `source: telegram`,
    `draft_id: ${d.id}`,
    `created_at: ${d.createdAt}`,
    `user: ${safe(d.userName)}`,
    `agency: ${safe(d.agency)}`,
    `name: ${safe(d.name)}`,
    `city: ${safe(d.city)}`,
    `phone: ${safe(d.phone)}`,
    `interlocuteur: ${safe(d.interlocuteur)}`,
    `interlocuteur_email: ${safe(d.interlocuteur_email)}`,
    `resume_entretien: |`,
    ...safe(d.resume).split("\n").map(l => `  ${l}`),
    `commande: |`,
    ...safe(d.commande).split("\n").map(l => `  ${l}`),
    `photo_file_id: ${safe(d.photoFileId)}`,
    `enriched:`,
    `  address: ${safe(e.address)}`,
    `  postalCode: ${safe(e.postalCode)}`,
    `  city: ${safe(e.city)}`,
    `  phone1: ${safe(e.phone1)}`,
    `  phone2: ${safe(e.phone2)}`,
    `  email: ${safe(e.email)}`,
    `  siret: ${safe(e.siret)}`,
    `  naf: ${safe(e.naf)}`,
    `  website: ${safe(e.website)}`,
  ].join("\n");
}

async function handleMessage(env: Env, msg: any) {
  const chatId = msg.chat.id as number;
  const userId = msg.from.id as number;
  const userName = [msg.from.first_name, msg.from.last_name].filter(Boolean).join(" ");
  const hasPhoto = !!msg.photo;
  const text = (msg.text || "").trim();

  // Sécurité: autorisation par liste blanche
  if (!isAllowedUser(env, userId)) {
    await tg("sendMessage", env.TELEGRAM_TOKEN, {
      chat_id: chatId,
      text: "⛔ Accès non autorisé à ce bot. Contacte l'administrateur.",
    });
    return;
  }

  // Utilitaire: renvoyer l'ID Telegram (pour que l'admin puisse whitelister)
  if (text === "/whoami") {
    await tg("sendMessage", env.TELEGRAM_TOKEN, {
      chat_id: chatId,
      text: `👤 ${userName || "Utilisateur"}\n🆔 Telegram user_id: ${userId}`,
    });
    return;
  }

  if (text === "/start") {
    await tg("sendMessage", env.TELEGRAM_TOKEN, {
      chat_id: chatId,
      parse_mode: "Markdown",
      text:
        "👋 *Prospection Intake*\n\n" +
        "Envoie-moi :\n" +
        "• `Nom ; Ville`\n" +
        "• `Nom 06xxxxxxxx`\n" +
        "• ou une *photo* de carte de visite\n\n" +
        "Puis :\n• `resume: ...`\n• `cmd: ...`\n• `bilan: prospect=6 client=4 cmd=2 (option: visites=10)`\n\n" +
        "Clique ✅ *Valider* pour envoyer au fichier de réception.",
    });
    await sendMainMenu(env, chatId, "Menu.");
    return;
  }

  // Main menu actions
  if (text === "↩️ Menu") {
    await clearConv(env, chatId, userId);
    await sendMainMenu(env, chatId, "Menu.");
    return;
  }
  if (text === "➕ Nouvelle entrée") {
    await saveConv(env, chatId, userId, { mode: "NEW", step: "TYPE" });
    await askNewEntryType(env, chatId);
    return;
  }
  if (text === "📊 Bilan session") {
    await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, parse_mode: "Markdown", text: "📊 Envoie :\n`bilan: prospect=6 client=4 cmd=2`\n(option: `visites=10`)" });
    return;
  }
  if (text === "🔒 Clore session") {
    await clearConv(env, chatId, userId);
    await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "🔒 Session clôturée ✅" });
    await sendMainMenu(env, chatId, "Menu.");
    return;
  }

  // Guided NEW flow
  const conv = await loadConv(env, chatId, userId);
  if (conv?.mode === "NEW") {
    conv.data = conv.data || {};

    if (conv.step === "TYPE") {
      if (text === "✏️ Saisie manuelle") {
        conv.kind = "MANUAL"; conv.step = "WAIT_NAME";
        await saveConv(env, chatId, userId, conv);
        await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Nom entreprise ?" });
        return;
      }
      if (text === "📞 Nom + téléphone") {
        conv.kind = "PHONE"; conv.step = "WAIT_NAME";
        await saveConv(env, chatId, userId, conv);
        await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Nom entreprise ?" });
        return;
      }
      if (text === "📷 Photo carte de visite") {
        conv.kind = "PHOTO"; conv.step = "WAIT_PHOTO";
        await saveConv(env, chatId, userId, conv);
        await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Envoie la photo de la carte de visite. (OCR automatique non inclus en gratuit)"});
        return;
      }
    }

    if (conv.step === "WAIT_PHOTO") {
      if (!hasPhoto) {
        await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "J’attends une photo 📷"});
        return;
      }
      const photos = msg.photo as any[];
      const best = photos[photos.length - 1];
      conv.data.photoFileId = best.file_id;
      conv.step = "WAIT_NAME";
      await saveConv(env, chatId, userId, conv);
      await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Photo reçue ✅\nNom entreprise ?" });
      return;
    }

    if (conv.step === "WAIT_NAME") {
      conv.data.name = text;
      conv.step = "WAIT_CITY";
      await saveConv(env, chatId, userId, conv);
      await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Ville ?" });
      return;
    }

    if (conv.step === "WAIT_CITY") {
      conv.data.city = text;
      if (conv.kind === "PHONE") {
        conv.step = "WAIT_PHONE";
        await saveConv(env, chatId, userId, conv);
        await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Téléphone ?" });
        return;
      }
      conv.step = "WAIT_INTERLOC";
      await saveConv(env, chatId, userId, conv);
      await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Nom interlocuteur ? (optionnel, réponds - si aucun)"});
      return;
    }

    if (conv.step === "WAIT_PHONE") {
      conv.data.phone = text;
      conv.step = "WAIT_INTERLOC";
      await saveConv(env, chatId, userId, conv);
      await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Nom interlocuteur ? (optionnel, réponds - si aucun)"});
      return;
    }

    if (conv.step === "WAIT_INTERLOC") {
      conv.data.interlocuteur = (text === "-" ? "" : text);
      conv.step = "WAIT_EMAIL";
      await saveConv(env, chatId, userId, conv);
      await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Mail interlocuteur ? (optionnel, réponds - si aucun)"});
      return;
    }

    if (conv.step === "WAIT_EMAIL") {
      conv.data.email = (text === "-" ? "" : text);
      conv.step = "WAIT_RESUME";
      await saveConv(env, chatId, userId, conv);
      await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Résumé d’entretien ?"});
      return;
    }

    if (conv.step === "WAIT_RESUME") {
      conv.data.resume = text;
      conv.step = "WAIT_CMD";
      await saveConv(env, chatId, userId, conv);
      await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Commande ? (optionnel, réponds - si aucune)"});
      return;
    }

    if (conv.step === "WAIT_CMD") {
      conv.data.cmd = (text === "-" ? "" : text);

      // Remplit le draft existant puis propose Valider
      const d = await loadDraft(env, chatId, userId, userName);
      d.name = conv.data.name || d.name;
      d.city = conv.data.city || d.city;
      d.phone = conv.data.phone || d.phone;
      d.interlocuteur = conv.data.interlocuteur || d.interlocuteur;
      d.interlocuteur_email = conv.data.email || d.interlocuteur_email;
      d.resume = conv.data.resume || d.resume;
      d.commande = conv.data.cmd || d.commande;
      d.photoFileId = conv.data.photoFileId || d.photoFileId;

      // enrichissement existant
      d.enriched = await enrichV1(d);

      await saveDraft(env, d);
      await clearConv(env, chatId, userId);

      await tg("sendMessage", env.TELEGRAM_TOKEN, {
        chat_id: chatId,
        parse_mode: "Markdown",
        text: renderDraftPreview(d),
        reply_markup: { inline_keyboard: [[{ text: "✅ Valider", callback_data: "VALIDATE" }], [{ text: "↩️ Menu", callback_data: "MENU" }]] },
      });
      return;
    }
  }

  // ── Saisie rapide (hors flow guidé) ────────────────────────────
  let d = await createOrGetDraft(env, chatId, userId, userName);

  if (msg.photo && Array.isArray(msg.photo) && msg.photo.length) {
    const best = msg.photo[msg.photo.length - 1];
    d.photoFileId = best.file_id;
    d.enriched = await enrichV1(d);
    await saveDraft(env, d);
    await sendMainMenu(env, chatId,
      "Menu prospection ✅\n\n" +
      "1) Choisis l’agence: /ag GR (ou VR/GRS/SLS)\n" +
      "2) ➕ Nouvelle entrée\n" +
      "3) 📊 Bilan session (bilan: prospect=.. client=.. cmd=..)\n" +
      "4) 🔒 Clore session");
    return;
  }

  if (!text) return;

  if (/^resume\s*:/i.test(text)) {
    d.resume = text.replace(/^resume\s*:/i, "").trim();
    d.enriched = await enrichV1(d);
    await saveDraft(env, d);
    await tg("sendMessage", env.TELEGRAM_TOKEN, {
      chat_id: chatId,
      parse_mode: "Markdown",
      text: buildPreview(d),
      reply_markup: keyboard(d.id),
    });
    return;
  }

  if (/^cmd\s*:/i.test(text)) {
    d.commande = text.replace(/^cmd\s*:/i, "").trim();
    d.enriched = await enrichV1(d);
    await saveDraft(env, d);
    await tg("sendMessage", env.TELEGRAM_TOKEN, {
      chat_id: chatId,
      parse_mode: "Markdown",
      text: buildPreview(d),
      reply_markup: keyboard(d.id),
    });
    return;
  }

  // Agency selector (ergonomic)
  // Formats:
  //  - "/ag GR" (définit l'agence pour le brouillon en cours)
  //  - "GR | Nom ; Ville"
  //  - "ag:VR Nom 06..."
  if (/^\/ag\s+/i.test(text)) {
    const ag = text.replace(/^\/ag\s+/i, "").trim().toUpperCase();
    d.agency = ag;
    d.enriched = await enrichV1(d);
    await saveDraft(env, d);
    await tg("sendMessage", env.TELEGRAM_TOKEN, {
      chat_id: chatId,
      parse_mode: "Markdown",
      text: `✅ Agence définie sur *${d.agency}* pour le brouillon en cours.`,
    });
    return;
  }

  // Inline agency prefix
  let mainText = text;
  const m1 = text.match(/^([A-Za-z]{2,3})\s*\|\s*(.+)$/);
  const m2 = text.match(/^ag\s*:\s*([A-Za-z]{2,3})\s+(.+)$/i);
  if (m1) { d.agency = m1[1].toUpperCase(); mainText = m1[2].trim(); }
  if (m2) { d.agency = m2[1].toUpperCase(); mainText = m2[2].trim(); }

  const { name, city, phone } = parseQuickLine(mainText);
  d.name = name || d.name;
  d.city = city || d.city;
  d.phone = phone || d.phone;
  d.enriched = await enrichV1(d);
  await saveDraft(env, d);

  await tg("sendMessage", env.TELEGRAM_TOKEN, {
    chat_id: chatId,
    parse_mode: "Markdown",
    text: buildPreview(d),
    reply_markup: keyboard(d.id),
  });
}

async function handleCallback(env: Env, cb: any) {
  const chatId = cb.message.chat.id as number;
  const userId = cb.from.id as number;
  const data = cb.data as string;
  const d = await createOrGetDraft(env, chatId, userId, [cb.from.first_name, cb.from.last_name].filter(Boolean).join(" "));

  if (data === "MENU") {
    await tg("answerCallbackQuery", env.TELEGRAM_TOKEN, { callback_query_id: cb.id });
    await sendMainMenu(env, chatId, "Menu.");
    return;
  }
  if (data.startsWith("cancel:")) {
    await clearDraft(env, chatId);
    await tg("answerCallbackQuery", env.TELEGRAM_TOKEN, { callback_query_id: cb.id, text: "Annulé ✅" });
    await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Brouillon annulé. Envoie une nouvelle société quand tu veux." });
    return;
  }
  if (data.startsWith("addresume:")) {
    await tg("answerCallbackQuery", env.TELEGRAM_TOKEN, { callback_query_id: cb.id, text: "Envoie: resume: ..." });
    await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Envoie: resume: ton résumé d'entretien" });
    return;
  }
  if (data.startsWith("addcmd:")) {
    await tg("answerCallbackQuery", env.TELEGRAM_TOKEN, { callback_query_id: cb.id, text: "Envoie: cmd: ..." });
    await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: "Envoie: cmd: ta commande (en vrac)" });
    return;
  }
  if (data === "VALIDATE" || data.startsWith("validate:")) {
    if (!d.agency) {
      await tg("answerCallbackQuery", env.TELEGRAM_TOKEN, { callback_query_id: cb.id, text: "Choisis l'agence (ex: /ag GR)" });
      return;
    }
    if (!d.name && !d.phone) {
      await tg("answerCallbackQuery", env.TELEGRAM_TOKEN, { callback_query_id: cb.id, text: "Il manque au moins NOM ou TEL" });
      return;
    }
    d.enriched = await enrichV1(d);
    const issueTitle = `Prospection [${d.agency || 'NA'}]: ${d.name || "SansNom"}${d.city ? " — " + d.city : ""}`;
    const body = "```yaml\n" + buildIssueBody(d) + "\n```";
    const issue = await githubCreateIssue(env, issueTitle, body) as { html_url: string };
    await clearDraft(env, chatId);
    await tg("answerCallbackQuery", env.TELEGRAM_TOKEN, { callback_query_id: cb.id, text: "Validé ✅" });
    await tg("sendMessage", env.TELEGRAM_TOKEN, { chat_id: chatId, text: `✅ Enregistré. Issue: ${issue.html_url}\n\nEnvoie la prochaine société.` });
    return;
  }
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    if (url.pathname !== `/${env.TELEGRAM_WEBHOOK_SECRET}`) return new Response("Not found", { status: 404 });
    if (request.method !== "POST") return new Response("OK", { status: 200 });

    const update = await request.json() as any;
    try {
      if (update.message) await handleMessage(env, update.message);
      else if (update.callback_query) await handleCallback(env, update.callback_query);
      return new Response(JSON.stringify({ ok: true }), { headers: { "content-type": "application/json" } });
    } catch (e: any) {
      return new Response(JSON.stringify({ ok: false, error: e?.message || String(e) }), { status: 500, headers: { "content-type": "application/json" } });
    }
  },
};
