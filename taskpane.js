/* STILGAR by Psomas — the "Start Legal Hold" button (task pane), add-in v1.0.

   The pane is the New Legal Hold page of Stilgar's wizard, in Outlook: the same
   questions, the same project-number tripwire, the same "Looks right" + Start button,
   with Terry's billing-list request composed live underneath (editable, like the app's
   preview). Pressing Start sends ONE request to the "Stilgar Intake" Power Automate
   flow: the flow queues Terry's email for the Legal Hold Email Sender flow (so it goes
   out from Paul Atreides within a minute or two) and files the request for Stilgar,
   which creates the hold with the very code the wizard runs.

   Works in classic Outlook AND new Outlook / the web (Matt's rule of 2026-10-10): plain
   Office.js, no COM. The pane never reads the selected email — it only needs the
   signed-in user's name and address.

   Everything here mirrors the Python package (ids.py, templates.py, mailqueue.py,
   holidays.py/dates.py, guilogic.parse_extra_recipients, state.next_seq_from); the
   constants come from config.js, generated from that package. Keep the two in step.
   Requests go as text/plain so the browser makes no CORS preflight; the flow answers
   with Access-Control-Allow-Origin: *.
*/
(function () {
  "use strict";
  const D = window.SG_DEFAULTS;
  const VERSION = "1.0";
  const $ = (id) => document.getElementById(id);
  const params = new URLSearchParams(location.search);
  const DEV = params.get("dev") === "1";
  const state = {
    flowUrl: "", user: { name: "", email: "" }, testMode: false, busy: false,
    index: null, indexNote: "", docprod: false, keywords: [],
    composed: null, touched: {}, allowAdditional: false, pending: null, pollTimer: null,
  };

  // ------------------------------------------------------------------ small helpers
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const humanDate = (d) => `${WEEKDAYS[d.getDay()]}, ${MONTHS[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
  const hhmm = (d) => { const h12 = d.getHours() % 12 || 12; return `${h12}:${pad(d.getMinutes())} ${d.getHours() < 12 ? "AM" : "PM"}`; };
  const humanDateTime = (d) => `${humanDate(d)} at ${hhmm(d)}`;
  const isoLocal = (d) => { const off = -d.getTimezoneOffset(), s = off >= 0 ? "+" : "-", a = Math.abs(off);
    return `${ymd(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}${s}${pad(Math.floor(a / 60))}:${pad(a % 60)}`; };
  const stamp = (d) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
  const rand = (n) => Array.from(crypto.getRandomValues(new Uint8Array(n))).map((b) => (b % 16).toString(16)).join("");
  // html.escape(s) with quote=True — the Python side's escaping, character for character
  const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#x27;");
  const unesc = (s) => { const t = document.createElement("textarea"); t.innerHTML = s; return t.value; };
  const replaceAll = (s, a, b) => s.split(a).join(b);

  // ------------------------------------------------------------------ holidays + business days (holidays.py / dates.py)
  function psomasHolidaysRule(y) {
    const H = {};
    H["New Year's Day"] = new Date(y, 0, 1);
    let d = new Date(y, 4, 31); while (d.getDay() !== 1) d.setDate(d.getDate() - 1); H["Memorial Day"] = d;
    H["Independence Day"] = new Date(y, 6, 4);
    d = new Date(y, 8, 1); while (d.getDay() !== 1) d.setDate(d.getDate() + 1); H["Labor Day"] = d;
    d = new Date(y, 10, 1); while (d.getDay() !== 4) d.setDate(d.getDate() + 1); d.setDate(d.getDate() + 21);
    H["Thanksgiving Day"] = new Date(d); H["Day after Thanksgiving"] = new Date(y, 10, d.getDate() + 1);
    H["Christmas Day"] = new Date(y, 11, 25);
    return H;
  }
  function observe(name, h) {
    const d = new Date(h);
    if (name === "New Year's Day") { if (d.getDay() === 6) d.setDate(d.getDate() + 2); else if (d.getDay() === 0) d.setDate(d.getDate() + 1); return d; }
    if (d.getDay() === 6) d.setDate(d.getDate() - 1); else if (d.getDay() === 0) d.setDate(d.getDate() + 1);
    return d;
  }
  const holidayCache = {};
  function observedSet(y) {
    if (!holidayCache[y]) {
      const table = (D.holidays || {})[String(y)];
      if (table) holidayCache[y] = new Set(table);                      // the app's own table, exported
      else { const H = psomasHolidaysRule(y); holidayCache[y] = new Set(Object.keys(H).map((k) => ymd(observe(k, H[k])))); }
    }
    return holidayCache[y];
  }
  const isBusinessDay = (d) => d.getDay() !== 0 && d.getDay() !== 6 && !observedSet(d.getFullYear()).has(ymd(d));
  function addBusinessDays(start, n) { const d = new Date(start.getFullYear(), start.getMonth(), start.getDate()); let left = n; while (left > 0) { d.setDate(d.getDate() + 1); if (isBusinessDay(d)) left--; } return d; }
  const terryDeadline = (anchor) => addBusinessDays(anchor || new Date(), D.terryBusinessDays);

  // ------------------------------------------------------------------ project identity (ids.py)
  const sanitizeNumber = (s) => String(s || "").trim().replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  const projectIndicator = (name, number, seq) => { const base = `${String(name).trim()} (${String(number).trim()})`; return seq > 1 ? `${base} - Hold ${seq}` : base; };
  const holdId = (number, seq) => { const s = sanitizeNumber(number); return seq > 1 ? `${s}-h${seq}` : s; };
  const RULES = D.numberRules;
  const CANON = new RegExp(RULES.canonicalToken);
  const SPLIT = new RegExp(RULES.split, "i");
  const RESERVED = new RegExp(RULES.reservedSuffix, "i");
  function validateProjectNumber(raw) {
    const normalized = String(raw || "").split(/\s+/).filter(Boolean).join(" ");
    if (!normalized) return { ok: false, warn: true, normalized: "", reason: "Enter a project number." };
    const bad = [...new Set([...normalized].filter((c) => RULES.forbidden.includes(c)))].sort();
    if (bad.length) return { ok: false, warn: true, normalized, reason: "Remove the character(s) " + bad.join(" ") + " — they can't be used in a folder name." };
    if (RESERVED.test(sanitizeNumber(normalized))) return { ok: false, warn: true, normalized,
      reason: "This number ends in a pattern the app reserves for numbering repeat holds (…-h2). Adjust the separator (e.g. use 'H2' without a space or dash)." };
    const tokens = normalized.split(SPLIT).filter(Boolean);
    if (tokens.length && tokens.every((t) => CANON.test(t))) return { ok: true, warn: false, normalized, reason: "" };
    return { ok: true, warn: true, normalized, reason: `This doesn't look like the usual Psomas project number (e.g. ${RULES.example}). Double-check it before continuing.` };
  }
  // state.next_seq_from + start_hold's active check, over the index Stilgar publishes
  const sameProject = (entry, number) => sanitizeNumber(entry.projectNumber || "").toLowerCase() === sanitizeNumber(number).toLowerCase();
  function nextSeqFrom(holds, number) { const seqs = (holds || []).filter((h) => sameProject(h, number)).map((h) => h.holdSeq || 1); return seqs.length ? Math.max(...seqs) + 1 : 1; }
  const TERMINAL = new Set(D.terminalStatuses || []);
  function activeHolds(holds, number) { return (holds || []).filter((h) => sameProject(h, number) && !TERMINAL.has(h.status)); }

  // ------------------------------------------------------------------ addresses (guilogic.parse_extra_recipients / templates.join_emails)
  function parseExtraRecipients(raw) {
    const out = [], seen = new Set();
    for (const part of String(raw || "").split(/[;,]/)) {
      const addr = part.trim(); if (!addr) continue;
      if (!addr.includes("@") || (addr.match(/@/g) || []).length > 1 || addr.includes(" ") || addr.startsWith("@") || addr.endsWith("@"))
        return ["", `"${addr}" doesn't look like an email address.`];
      const k = addr.toLowerCase(); if (!seen.has(k)) { seen.add(k); out.push(addr); }
    }
    return [out.join(";"), ""];
  }
  function joinEmails(list) { const seen = [], out = []; for (const e of list) { const t = (e || "").trim(); if (t && !seen.includes(t.toLowerCase())) { seen.push(t.toLowerCase()); out.push(t); } } return out.join(";"); }
  function splitTerms(raw) { const out = [], seen = new Set(); for (const p of String(raw || "").split(/[\n,]/)) { const t = p.trim(); if (t && !seen.has(t.toLowerCase())) { seen.add(t.toLowerCase()); out.push(t); } } return out; }

  // ------------------------------------------------------------------ Terry's email (templates.billing_request, as data)
  function fillParagraph(template, fields) {
    let text = template;
    for (const k of Object.keys(fields)) text = replaceAll(text, "{" + k + "}", esc(fields[k]));
    return `<p style="${D.prose}">${text.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>")}</p>`;
  }
  function composeBilling(indicator, number, deadlineHuman, extraTo) {
    const body = [...D.billing.paragraphs.map((t) => fillParagraph(t, { projectNumber: number, deadline: deadlineHuman })), D.spacer, D.signaturePaul].join("\n");
    return { purpose: D.billing.purpose, to: joinEmails([D.cast.terry, ...String(extraTo || "").split(";")]), cc: D.cast.chad, bcc: "",
      subject: replaceAll(D.billing.subject, "{indicator}", indicator), bodyHtml: body, from: D.cast.paulEmail, importance: "Normal", threadSubject: "" };
  }

  // ------------------------------------------------------------------ the editable projection (editablemail.py, the parts Terry's email needs)
  const BLOCK_RE = /<(p|ol|ul)\b[\s\S]*?<\/\1>/gi;
  function splitSignature(bodyHtml) {
    const i = bodyHtml.indexOf(D.signaturePaul); if (i < 0) return [bodyHtml, ""];
    let prefix = bodyHtml.slice(0, i);
    const m = /(?:<p[^>]*>(?:&nbsp;|\s)*<\/p>\s*)+$/i.exec(prefix);
    const cut = m ? m.index : i;
    return [bodyHtml.slice(0, cut), bodyHtml.slice(cut)];
  }
  function inlineToText(h) {
    h = h.replace(/<br\s*\/?>/gi, "\n");
    h = h.replace(/<a\b[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, (m, url, text) => `[${text}]\u0000${unesc(url)}\u0001`);
    h = h.replace(/<b\b[^>]*>([\s\S]*?)<\/b>/gi, "**$1**");
    h = h.replace(/<[^>]+>/g, "");
    return unesc(h).replace(/\u0000/g, "<").replace(/\u0001/g, ">").trim();
  }
  function blockToText(block) {
    if (/^<(ol|ul)\b/i.test(block)) { const items = [...block.matchAll(/<li\b[^>]*>([\s\S]*?)<\/li>/gi)].map((m) => inlineToText(m[1])); return items.map((t, i) => `${i + 1}. ${t}`).join("\n"); }
    return inlineToText(block.replace(/^<p\b[^>]*>|<\/p>$/gi, ""));
  }
  const blocksOf = (html) => [...html.matchAll(BLOCK_RE)].map((m) => m[0]);
  const toText = (bodyHtml) => blocksOf(splitSignature(bodyHtml)[0]).map(blockToText).join("\n\n").trim();
  function escInline(text) {
    return esc(text).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/(?<!["=>])(https?:\/\/[^\s<]+)/g, (u) => `<a href="${u}" style="${D.linkStyle}">${u}</a>`);
  }
  function inlineToHtml(text) {
    const re = /\[([^\]]*)\]<(https?:\/\/[^>\s]+)>/g; let out = "", last = 0, m;
    while ((m = re.exec(text)) !== null) { out += escInline(text.slice(last, m.index)); out += `<a href="${esc(unesc(m[2]))}" style="${D.linkStyle}">${escInline(m[1])}</a>`; last = m.index + m[0].length; }
    return out + escInline(text.slice(last));
  }
  function renderBlock(text) {
    const lines = text.split("\n").filter((l) => l.trim());
    if (lines.length && lines.every((l) => /^\s*\d+[.)]\s+/.test(l))) return `<ol style="${D.prose}">\n${lines.map((l) => `  <li>${inlineToHtml(l.replace(/^\s*\d+[.)]\s+/, "").trim())}</li>`).join("\n")}\n</ol>`;
    return `<p style="${D.prose}">${inlineToHtml(text.trim()).replace(/\n/g, "<br>")}</p>`;
  }
  const norm = (t) => t.split(/\s+/).filter(Boolean).join(" ");
  function rebuild(bodyHtml, editedText) {
    editedText = String(editedText || "").replace(/\r\n?/g, "\n").trim();
    if (editedText === toText(bodyHtml)) return bodyHtml;
    const [prefix, suffix] = splitSignature(bodyHtml);
    const originals = blocksOf(prefix).map((b) => ({ html: b, key: norm(blockToText(b)), used: false }));
    const out = [];
    for (const t of editedText.split(/\n\s*\n/).map((s) => s.trim()).filter(Boolean)) {
      const o = originals.find((x) => !x.used && x.key === norm(t));       // untouched paragraphs keep their exact markup
      if (o) { o.used = true; out.push(o.html); } else out.push(renderBlock(t));
    }
    const rebuilt = out.join("\n");
    return suffix ? rebuilt + "\n" + suffix : rebuilt;
  }

  // ------------------------------------------------------------------ the queue payload (mailqueue.py)
  function makeQueueId(key, purpose, when) { const num = sanitizeNumber(key) || "hold"; const p = purpose.replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, ""); return `${num}-${p}-${stamp(when)}-${rand(4)}`; }
  function buildQueuePayload(email, number, indicator, when, hid, seq) {
    return { queueId: makeQueueId(hid || number, email.purpose, when), createdAt: isoLocal(when), projectNumber: number, projectIndicator: indicator,
      holdId: hid, holdSeq: seq, purpose: email.purpose, to: email.to, cc: email.cc, bcc: email.bcc || "", subject: email.subject, bodyHtml: email.bodyHtml,
      importance: email.importance || "Normal", threadSubject: email.threadSubject || "", from: email.from || null };
  }
  const recipientSet = (p) => new Set(`${p.to || ""};${p.cc || ""}`.split(";").map((x) => x.trim().toLowerCase()).filter(Boolean));
  const appendCc = (p, addrs) => { if (!addrs.length) return p; const out = Object.assign({}, p); out.cc = p.cc ? [p.cc, ...addrs].join(";") : addrs.join(";"); return out; };
  function ccSupervisors(p, sender) {   // Matt's rule: Matt AND Chad on every email they didn't send themselves
    const who = (sender || p.from || "").trim().toLowerCase(), have = recipientSet(p);
    return appendCc(p, (D.supervisors || []).filter((s) => s.toLowerCase() !== who && !have.has(s.toLowerCase())));
  }
  function ccOperator(p, operatorEmail) {   // 2.9: Paul-persona mail also CCs whoever opened the hold
    const paul = (D.cast.paulEmail || "").toLowerCase(), op = (operatorEmail || "").trim();
    if (!paul || !op || (p.from || "").toLowerCase() !== paul || recipientSet(p).has(op.toLowerCase())) return p;
    return appendCc(p, [op]);
  }
  function applyTestMode(p, redirect) {
    const note = replaceAll(replaceAll(D.testNote, "{to}", p.to || "—"), "{cc}", p.cc || "—");
    return Object.assign({}, p, { to: redirect, cc: "", bcc: "", subject: "[TEST] " + (p.subject || ""), bodyHtml: note + (p.bodyHtml || "") });
  }
  function validatePayload(p) {
    const problems = [];
    for (const f of ["queueId", "purpose", "subject", "bodyHtml"]) if (!String(p[f] || "").trim()) problems.push(`missing ${f}`);
    if (!String(p.to || "").trim()) problems.push("no recipients (to is empty)");
    for (const f of ["to", "cc", "bcc"]) { const v = String(p[f] || ""); if (v.includes("<") || v.includes(">")) problems.push(`${f} uses a 'Display Name <email>' wrapper — use plain emails`); if (v.includes(",")) problems.push(`${f} is comma-separated — must be semicolon-separated`); }
    if (!["Normal", "High", "Low"].includes(p.importance || "Normal")) problems.push("importance must be Normal/High/Low");
    return problems;
  }

  // ------------------------------------------------------------------ the flow
  const GENERIC = ["An action failed", "ActionFailed", "The execution of template action"];
  function shortError(raw) {
    if (raw === null || raw === undefined) return "The flow reported a failure.";
    const s = typeof raw === "string" ? raw.trim() : JSON.stringify(raw); let data = null;
    try { data = /^[\[{]/.test(s) ? JSON.parse(s) : null; } catch (e) { data = null; }
    const pick = (o) => { if (!o || typeof o !== "object") return ""; for (const k of ["message", "error"]) { let v = o[k]; if (v && typeof v === "object") v = v.message; if (typeof v === "string" && v.trim() && !GENERIC.some((g) => v.includes(g))) return v.trim(); } return ""; };
    if (Array.isArray(data)) { for (const it of data) { const m = it && it.error && it.error.message; if (m) return `${(it.name || "a step").replace(/_/g, " ")}: ${m}`.slice(0, 400); } }
    else if (data) { const m = pick(data); if (m) return m.slice(0, 400); }
    return s.slice(0, 400) || "The flow reported a failure.";
  }
  async function callFlow(payload, timeoutMs) {
    if (!state.flowUrl) throw new Error("The flow address isn't set — open Settings.");
    const ctrl = new AbortController(); const t = setTimeout(() => ctrl.abort(), timeoutMs || 115000);
    let res;
    try { res = await fetch(state.flowUrl, { method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(payload), signal: ctrl.signal }); }
    catch (e) { throw new Error(e.name === "AbortError" ? "The flow didn't answer in time — try again in a moment." : "Couldn't reach the flow (" + e.message + "). Is the address right and the flow turned on?"); }
    finally { clearTimeout(t); }
    const text = await res.text(); let data;
    try { data = JSON.parse(text); if (typeof data === "string") data = JSON.parse(data); } catch (e) { data = { ok: res.ok, error: text.slice(0, 300) || ("HTTP " + res.status) }; }
    if (!res.ok && data.ok !== true) throw new Error(shortError(data.error || ("HTTP " + res.status)));
    if (data.ok === false) throw new Error(shortError(data.error));
    return data;
  }
  const userKey = () => (state.user.email || "user").split("@")[0].replace(/[^A-Za-z0-9.]+/g, "-");
  const basePayload = (kind, when) => ({ schema: D.schema, kind, id: `sg-${kind}-${stamp(when || new Date())}-${rand(4)}`, app: `Stilgar add-in ${VERSION}`,
    submittedBy: state.user.email, submittedByName: state.user.name, submittedAt: isoLocal(when || new Date()) });
  async function fetchIndex() {
    try {
      const r = await callFlow(basePayload("lookup"), 40000);
      const idx = r.index && typeof r.index === "string" ? JSON.parse(r.index) : r.index;
      state.index = idx && Array.isArray(idx.holds) ? idx.holds : [];
      state.indexNote = idx && idx.writtenAt ? "" : "Stilgar hasn't published its list of holds yet — the hold number may be adjusted when it files this one.";
    } catch (e) { state.index = null; state.indexNote = "Couldn't read Stilgar's list of existing holds (" + e.message + ") — Stilgar will decide the hold number when it files this one."; }
  }

  // ------------------------------------------------------------------ roaming settings
  const mem = {};
  function roam(k, v) {
    if (DEV) { if (v === undefined) return mem[k]; mem[k] = v; return undefined; }
    try { const rs = Office.context.roamingSettings; if (v === undefined) return rs.get(k); rs.set(k, v); rs.saveAsync(() => {}); } catch (e) { return undefined; }
  }

  // ------------------------------------------------------------------ UI plumbing
  function show(id, on) { $(id).classList.toggle("hidden", !on); }
  function setStatus(text, kind) { const el = $("status"); if (!text) { el.classList.add("hidden"); return; } el.className = "status " + (kind || ""); el.textContent = text; el.classList.remove("hidden"); }
  function page(name) {
    ["sec-setup", "sec-form", "sec-email", "sec-confirm", "sec-result", "sec-settings"].forEach((id) => show(id, false));
    if (name === "setup") show("sec-setup", true);
    else if (name === "settings") { show("sec-settings", true); renderSettings(); }
    else if (name === "confirm") show("sec-confirm", true);
    else if (name === "result") show("sec-result", true);
    else { show("sec-form", true); show("sec-email", true); }
  }
  function renderHeader() { const pill = $("testpill"); if (state.testMode) { pill.textContent = `TEST MODE — Terry's email comes to ${state.user.email}; the hold is a practice hold`; pill.classList.remove("hidden"); } else pill.classList.add("hidden"); }

  // --- the wizard page
  function setDocprod(on) { state.docprod = on; $("dp-no").classList.toggle("active", !on); $("dp-yes").classList.toggle("active", on); show("docframe", on); recompute(); }
  function renderChips() {
    const host = $("kw-chips"); host.innerHTML = "";
    state.keywords.forEach((k, i) => { const c = document.createElement("span"); c.className = "chip"; c.textContent = k; const x = document.createElement("button"); x.type = "button"; x.title = "Remove"; x.textContent = "×"; x.addEventListener("click", () => { state.keywords.splice(i, 1); renderChips(); recompute(); }); c.appendChild(x); host.appendChild(c); });
  }
  function addKeywords(raw) { for (const t of splitTerms(raw)) if (!state.keywords.some((k) => k.toLowerCase() === t.toLowerCase())) state.keywords.push(t); renderChips(); $("kw-input").value = ""; recompute(); }

  function validation() {
    const numRaw = $("f-number").value, name = $("f-name").value.trim();
    const check = validateProjectNumber(numRaw);
    const [extra, extraProblem] = parseExtraRecipients($("f-extra").value);
    return { numRaw, name, check, extra, extraProblem, valid: !!(check.ok && name && !extraProblem) };
  }
  function tripwire(v) {
    const el = $("tripwire");
    if (!v.name || !v.numRaw.trim()) { el.className = "tripwire neutral"; el.textContent = "Enter a project name and number."; return null; }
    if (!v.check.ok) { el.className = "tripwire err"; el.textContent = "✗ " + v.check.reason; return null; }
    if (v.extraProblem) { el.className = "tripwire err"; el.textContent = "✗ " + v.extraProblem; return null; }
    const seq = nextSeqFrom(state.index, v.check.normalized);
    const indicator = projectIndicator(v.name, v.check.normalized, seq);
    el.className = "tripwire " + (v.check.warn ? "warn" : "ok");
    el.textContent = (v.check.warn ? "⚠  " : "") + `"${indicator}"` + (seq > 1 ? `   (this project's hold #${seq})` : "") + (v.check.warn ? "  — " + v.check.reason : "");
    return { seq, indicator };
  }
  function recompute(fromEdit) {
    const v = validation();
    const id = tripwire(v);
    if (fromEdit !== "confirm") $("looks-right").checked = false;      // the app un-ticks "Looks right" on every edit
    updateStart(v);
    if (!v.valid || !id) { state.composed = null; $("email-pill").textContent = "fill in the page above"; $("email-pill").className = "pill neutral"; return; }
    const deadline = humanDate(terryDeadline(new Date()));
    const email = composeBilling(id.indicator, v.check.normalized, deadline, v.extra);
    state.composed = Object.assign({ seq: id.seq, indicator: id.indicator, number: v.check.normalized, deadline }, email);
    $("pv-from").textContent = D.cast.paulEmail + "  (Paul Atreides)";
    if (!state.touched.to) $("pv-to").value = email.to.split(";").join("; ");
    if (!state.touched.cc) $("pv-cc").value = email.cc.split(";").join("; ");
    if (!state.touched.subject) $("pv-subject").value = email.subject;
    if (!state.touched.body) $("pv-body").value = toText(email.bodyHtml);
    $("email-pill").textContent = state.testMode ? `TEST: rerouted to ${state.user.email}` : "goes out when you press Start";
    $("email-pill").className = "pill " + (state.testMode ? "test" : "neutral");
  }
  function updateStart(v) { v = v || validation(); $("btn-start").disabled = !(v.valid && $("looks-right").checked) || state.busy; }

  // --- confirmations (in-pane: Outlook task panes can't rely on window.confirm)
  function ask(title, text) {
    return new Promise((resolve) => {
      $("confirm-title").textContent = title; $("confirm-text").textContent = text;
      const yes = $("btn-confirm-yes"), no = $("btn-confirm-no");
      const done = (v) => { yes.onclick = null; no.onclick = null; resolve(v); };
      yes.onclick = () => done(true); no.onclick = () => done(false);
      page("confirm");
    });
  }

  // --- 2.25: the stilgar: link that opens (or wakes) Stilgar on this PC (launcher.py)
  function stilgarLink(requestId) { return D.launch.intakeUrl.replace("{requestId}", encodeURIComponent(requestId)); }
  function launchStilgar(requestId) {
    const url = stilgarLink(requestId);
    $("res-open").setAttribute("href", url);           // the manual fallback on the result page
    try { location.assign(url); return true; }         // an external scheme never replaces the pane
    catch (e) { return false; }
  }

  // --- Start Legal Hold (one press)
  async function start() {
    if (state.busy) return;
    const v = validation(); if (!v.valid || !$("looks-right").checked || !state.composed) return;
    state.busy = true; updateStart(v); $("form-problem").textContent = "";
    try {
      state.allowAdditional = false;
      const active = activeHolds(state.index, v.check.normalized);
      if (active.length) {
        const a = active[active.length - 1], n = state.composed.seq;
        const ok = await ask("Active hold exists",
          `There's already an ACTIVE hold for this project: ${a.projectIndicator} (${String(a.status || "").replace(/_/g, " ")}).\n\n` +
          `Start an additional, separate hold (Hold ${n}) anyway? It gets its own folder, survey link, and reminders.\n\n` +
          "Note: with two holds running on one project, reminder emails may appear in the same Outlook conversation (they are sent per hold, but Outlook threads by project number).");
        if (!ok) { page("main"); return; }
        state.allowAdditional = true;
      }
      if (state.testMode) {
        const ok = await ask("Test mode is ON — create a PRACTICE hold?",
          "Because test mode is on, this hold will be created as a PRACTICE hold: every email it ever produces (billing request, notices, follow-ups) comes to YOU instead of the real recipients — permanently, even after test mode is turned off.\n\n" +
          "For a real hold, press Cancel, turn test mode off in Settings, then start the hold again.\n\nCreate it as a practice hold?");
        if (!ok) { page("main"); return; }
      }
      page("main");
      await submit(v);
    } finally { state.busy = false; updateStart(); }
  }

  async function submit(v) {
    const c = state.composed;
    // the edited email, validated like the app's preview panel
    const [toNorm, toProb] = parseExtraRecipients($("pv-to").value), [ccNorm, ccProb] = parseExtraRecipients($("pv-cc").value);
    const subject = $("pv-subject").value.trim();
    const problem = toProb || ccProb || (!toNorm ? "The To line needs at least one address." : "") || (!subject ? "The subject can't be empty." : "") || (!$("pv-body").value.trim() ? "The email text is empty." : "");
    if (problem) { $("form-problem").textContent = problem; return; }
    const when = new Date();
    const email = Object.assign({}, c, { to: toNorm, cc: ccNorm, subject, bodyHtml: rebuild(c.bodyHtml, $("pv-body").value) });
    let payload = buildQueuePayload(email, c.number, c.indicator, when, holdId(c.number, c.seq), c.seq);
    payload = ccSupervisors(payload); payload = ccOperator(payload, state.user.email);
    const realTo = payload.to, realCc = payload.cc;
    if (state.testMode) payload = applyTestMode(payload, state.user.email);
    const bad = validatePayload(payload); if (bad.length) { $("form-problem").textContent = bad.join("; "); return; }
    const requestId = `sg-${stamp(when)}-${userKey()}-${rand(4)}`;
    const request = {
      schema: D.schema, kind: D.kindStart, requestId, app: `Stilgar add-in ${VERSION}`, submittedAt: isoLocal(when),
      requester: { name: state.user.name, email: state.user.email },
      projectName: v.name, projectNumber: c.number, docProduction: state.docprod,
      searchParams: state.docprod ? { dateRangeRaw: $("f-daterange").value.trim(), timePeriodDisplay: $("f-daterange").value.trim().replace(/-/g, "–"), keywordLines: state.keywords.slice() } : null,
      terryExtraTo: v.extra, allowAdditional: state.allowAdditional, testHold: state.testMode,
      expectedSeq: c.seq, expectedIndicator: c.indicator,
    };
    // 2.25: open (or wake) Stilgar on this PC right now — still inside the click that pressed
    // Start, which browsers require before they launch a stilgar: link. Stilgar then watches the
    // shared queue for this very request and opens the hold's tab the moment it is filed.
    const opened = launchStilgar(requestId);
    setStatus("Sending Terry's request and handing the hold to Stilgar…", "busy");
    let res;
    try { res = await callFlow(Object.assign(basePayload("start-hold", when), { requestId, request, queue: payload })); }
    catch (e) { setStatus(e.message, "err"); return; }
    setStatus("", "");
    const sent = res.emailQueued !== false;
    $("res-sent").textContent = sent
      ? (state.testMode ? `✓ Terry's billing-list request was rerouted to you (TEST MODE). It would have gone To: ${realTo.split(";").join("; ")}; CC: ${realCc.split(";").join("; ")}.`
                        : `✓ Terry's billing-list request is on its way from Paul Atreides — To: ${realTo.split(";").join("; ")}; CC: ${realCc.split(";").join("; ")}.`)
      : "⚠ Terry's request was NOT queued — Stilgar will open the hold on her email preview so it can be sent from there.";
    $("res-filed").textContent = `⏳ Waiting for Stilgar to file the hold "${c.indicator}"…`;
    $("res-note").textContent = (opened ? "Stilgar is opening on this PC to file it (about a minute). " : "")
      + "If Stilgar did not open, press Open Stilgar. Any other open Stilgar files it after a few minutes. Safe to close this panel.";
    state.pending = { requestId, indicator: c.indicator, started: Date.now() };
    page("result");
    poll();
  }

  async function poll() {
    if (state.pollTimer) { clearTimeout(state.pollTimer); state.pollTimer = null; }
    const p = state.pending; if (!p) return;
    try {
      const r = await callFlow(Object.assign(basePayload("status"), { requestId: p.requestId }), 40000);
      const receipt = r.receipt && typeof r.receipt === "string" ? JSON.parse(r.receipt) : r.receipt;
      if (receipt && receipt.ok) {
        let when = ""; try { when = humanDateTime(new Date(receipt.filedAt)); } catch (e) { when = ""; }
        $("res-filed").textContent = `✓ Filed in Stilgar as "${receipt.indicator || p.indicator}" by ${receipt.machine || "a Stilgar"}${receipt.operator ? " (" + receipt.operator + ")" : ""}${when ? " — " + when : ""}.`;
        $("res-note").textContent = receipt.billingQueued === false ? "Terry's request still needs to be sent: open the hold in Stilgar and press Send on her preview." : "Terry's reply deadline and the automatic reminders now run from Stilgar, as usual.";
        state.pending = null; return;
      }
      if (receipt && receipt.ok === false) {
        $("res-filed").textContent = `⚠ Stilgar could not file the hold: ${receipt.error || "see the request in intake\\needs-attention"}.`;
        $("res-note").textContent = "The request is parked for Matt and you've been emailed the details.";
        state.pending = null; return;
      }
    } catch (e) { /* keep polling — a slow flow answer is not an outcome */ }
    if (Date.now() - p.started > 4 * 60 * 1000) {
      $("res-filed").textContent = `⏳ No Stilgar has picked up "${p.indicator}" yet (none open right now?).`;
      $("res-note").textContent = "Nothing is lost: press Open Stilgar, or the next Stilgar to open — or the 11 AM daily check — files it.";
      state.pending = null; return;
    }
    state.pollTimer = setTimeout(poll, 10000);
  }

  function resetForm() {
    ["f-number", "f-name", "f-extra", "f-daterange", "kw-input"].forEach((id) => { $(id).value = ""; });
    state.keywords = []; renderChips(); state.touched = {}; state.composed = null; state.allowAdditional = false;
    setDocprod(false); $("looks-right").checked = false; $("form-problem").textContent = ""; $("res-open").setAttribute("href", "#");
  }

  // --- settings
  function renderSettings() {
    $("set-url").value = state.flowUrl || ""; $("set-testmode").checked = !!state.testMode;
    $("about").textContent = `Stilgar add-in ${VERSION} · wording from Stilgar v${D.appVersion} · signed in as ${state.user.email}`;
  }
  async function connect(url, fromSetup) {
    const msg = fromSetup ? $("setup-msg") : $("set-msg");
    url = (url || "").trim(); if (!/^https?:\/\//i.test(url) || (!DEV && !/^https:\/\//i.test(url))) { msg.textContent = "That doesn't look like a flow address (it starts with https://)."; return false; }
    state.flowUrl = url; msg.textContent = "Connecting…";
    try { const r = await callFlow(basePayload("ping"), 40000); msg.textContent = `Connected — the flow answered (${r.service || "ok"}).`; }
    catch (e) { msg.textContent = e.message; return false; }
    roam("stilgarFlowUrl", url); return true;
  }

  // ------------------------------------------------------------------ init
  function wire() {
    $("dp-no").addEventListener("click", () => setDocprod(false));
    $("dp-yes").addEventListener("click", () => setDocprod(true));
    ["f-number", "f-name", "f-extra", "f-daterange"].forEach((id) => $(id).addEventListener("input", () => recompute("field")));
    $("kw-add").addEventListener("click", () => addKeywords($("kw-input").value));
    $("kw-input").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); addKeywords($("kw-input").value); } });
    $("kw-input").addEventListener("paste", (e) => { const t = (e.clipboardData || window.clipboardData).getData("text"); if (/[\n,]/.test(t)) { e.preventDefault(); addKeywords(t); } });
    $("looks-right").addEventListener("change", () => updateStart());
    $("btn-start").addEventListener("click", start);
    [["pv-to", "to"], ["pv-cc", "cc"], ["pv-subject", "subject"], ["pv-body", "body"]].forEach(([id, k]) => $(id).addEventListener("input", () => { state.touched[k] = true; }));
    $("pv-reset").addEventListener("click", (e) => { e.preventDefault(); state.touched = {}; recompute("confirm"); });
    $("btn-again").addEventListener("click", async () => { state.pending = null; resetForm(); page("main"); await fetchIndex(); recompute(); });
    $("res-open").addEventListener("click", (e) => { if (!/^[a-z][a-z0-9+.-]*:/i.test($("res-open").getAttribute("href") || "")) e.preventDefault(); });
    $("setup-save").addEventListener("click", async () => { if (await connect($("setup-url").value, true)) { page("main"); await fetchIndex(); recompute(); } });
    $("set-save").addEventListener("click", async () => { if (await connect($("set-url").value, false)) setStatus("Flow address saved.", "ok"); });
    $("set-test").addEventListener("click", async () => { await connect($("set-url").value || state.flowUrl, false); });
    $("set-testmode").addEventListener("change", () => { state.testMode = $("set-testmode").checked; roam("stilgarTestMode", state.testMode ? "1" : ""); renderHeader(); recompute("confirm"); });
    $("nav-main").addEventListener("click", (e) => { e.preventDefault(); page(state.flowUrl ? "main" : "setup"); });
    $("nav-settings").addEventListener("click", (e) => { e.preventDefault(); page("settings"); });
  }

  // a few pure helpers exposed for checking from a browser console (no data, no secrets)
  window.__sg = { validateProjectNumber, projectIndicator, holdId, sanitizeNumber, nextSeqFrom, activeHolds, parseExtraRecipients,
    terryDeadline, humanDate, isBusinessDay, composeBilling, fillParagraph, toText, rebuild, buildQueuePayload, ccSupervisors, ccOperator,
    applyTestMode, validatePayload, shortError, splitTerms, stilgarLink, version: VERSION };

  async function init() {
    wire();
    if (DEV) { state.user = { name: "Dev User", email: "dev.user@psomas.com" }; }
    else { try { const p = Office.context.mailbox.userProfile; state.user = { name: p.displayName || "", email: (p.emailAddress || "").toLowerCase() }; } catch (e) { state.user = { name: "", email: "" }; } }
    // the manifest IT deploys carries the flow address (taskpane.html?flow=...), so nobody types it;
    // a person who added the button themselves pastes it once and it stays in their roaming settings
    let fromManifest = ""; try { fromManifest = (params.get("flow") || "").trim(); } catch (e) { fromManifest = ""; }
    if (/^https?:\/\//i.test(fromManifest)) { state.flowUrl = fromManifest; if (roam("stilgarFlowUrl") !== fromManifest) roam("stilgarFlowUrl", fromManifest); }
    else state.flowUrl = roam("stilgarFlowUrl") || "";
    state.testMode = roam("stilgarTestMode") === "1";
    renderHeader(); renderChips(); setDocprod(false);
    if (!state.flowUrl) { page("setup"); return; }
    page("main");
    setStatus("Reading Stilgar's list of existing holds…", "busy");
    await fetchIndex();
    setStatus(state.indexNote, state.indexNote ? "warn" : "");
    recompute();
  }

  if (DEV || typeof Office === "undefined") init();
  else Office.onReady(() => { init(); });
})();
