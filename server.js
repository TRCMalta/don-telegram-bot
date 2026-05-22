/**
 * Don — Mike Roberts' AI Chief of Staff
 * Telegram bot on Railway
 *
 * Architecture:
 *   Telegram webhook → Express → Odoo XML-RPC + Microsoft 365 Graph → Claude (tool use)
 *
 * Sibling of Sam (Beverly's CoS), same architectural pattern but Telegram-only and
 * without Ceek recruitment-system access (that surface stays with Sam).
 * State of record lives in MS 365 (mail/calendar) and Odoo (Think Talent CRM). Conversation
 * history is in-process only with LRU + per-chat timeout — restarts wipe in-flight chats by
 * design.
 */

import "dotenv/config";
import express from "express";
import { Anthropic } from "@anthropic-ai/sdk";
import https from "https";
import http from "http";
import { readFile, writeFile } from "fs/promises";
import { existsSync } from "fs";
import path from "path";

const app = express();
app.use(express.json());

// ─── Config ───────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
const TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const CLAUDE_MODEL = process.env.CLAUDE_MODEL || "claude-sonnet-4-20250514";

// Odoo (Think Talent — reused from Sam)
const ODOO_DB = process.env.ODOO_DB || "thinktalent_prod";
const ODOO_API_KEY = process.env.ODOO_API_KEY;
const ODOO_URL = process.env.ODOO_URL || "https://thinktalent.com.mt";
const ODOO_LOGIN = process.env.ODOO_LOGIN;

// Microsoft Graph (Azure App "Don TRC Bot")
const MS_TENANT_ID = process.env.MS_TENANT_ID;
const MS_CLIENT_ID = process.env.MS_CLIENT_ID;
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
const MIKE_EMAIL = process.env.MIKE_EMAIL || "mike@theremarkablecollective.com";
const MIKE_FORM_OF_ADDRESS = process.env.MIKE_FORM_OF_ADDRESS || "Mr Roberts";

// Allowed Telegram user IDs (Mike + Jonathan, comma-separated)
const ALLOWED_USERS = (process.env.ALLOWED_TELEGRAM_USERS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);

const anthropic = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

// ─── Persistent memory ─────────────────────────────────────────────────────────
// Uses /data (Railway volume) when mounted, falls back to /tmp (ephemeral)
const MEMORY_DIR = existsSync("/data") ? "/data" : "/tmp";
const MEMORY_FILE = path.join(MEMORY_DIR, "don-memory.json");
const MEMORY_EXTRACTION_MODEL = "claude-haiku-4-20250514";

const DEFAULT_MEMORY = {
  version: 2,
  lastUpdated: null,
  preferences: [],   // max 20 — things Mike consistently prefers
  keyFacts: [],      // max 30 — specific facts about deals/projects/context
  people: {},        // name → note
  recentTopics: [],  // max 10 — {date, topic}
};

let memoryCache = null;

async function loadMemory() {
  if (memoryCache) return memoryCache;
  try {
    const raw = await readFile(MEMORY_FILE, "utf8");
    memoryCache = { ...DEFAULT_MEMORY, ...JSON.parse(raw) };
    console.log(`[MEMORY] Loaded from ${MEMORY_FILE}: ${memoryCache.preferences.length} prefs, ${memoryCache.keyFacts.length} facts`);
  } catch {
    memoryCache = { ...DEFAULT_MEMORY };
    console.log(`[MEMORY] No existing file — starting fresh at ${MEMORY_FILE}`);
  }
  return memoryCache;
}

async function saveMemory(mem) {
  memoryCache = mem;
  mem.lastUpdated = new Date().toISOString();
  try {
    await writeFile(MEMORY_FILE, JSON.stringify(mem, null, 2), "utf8");
    console.log(`[MEMORY] Saved — ${mem.preferences.length} prefs, ${mem.keyFacts.length} facts`);
  } catch (err) {
    console.error("[MEMORY] Save failed:", err.message);
  }
}

function buildMemoryContext(memory) {
  const parts = [];
  if (memory.preferences.length) {
    parts.push("## Learnt preferences (apply these without announcing them)\n" + memory.preferences.map((p) => `- ${p}`).join("\n"));
  }
  if (memory.keyFacts.length) {
    parts.push("## Key facts\n" + memory.keyFacts.slice(-20).map((f) => `- ${f}`).join("\n"));
  }
  if (Object.keys(memory.people).length) {
    parts.push("## People notes\n" + Object.entries(memory.people).map(([n, v]) => `- ${n}: ${v}`).join("\n"));
  }
  if (memory.recentTopics.length) {
    parts.push("## Recent conversation topics\n" + memory.recentTopics.slice(-5).map((t) => `- ${t.date}: ${t.topic}`).join("\n"));
  }
  if (!parts.length) return "";
  return "\n\n---\n\n## DON'S MEMORY — PERSISTENT CONTEXT\n\n" + parts.join("\n\n") + "\n\nUse this context naturally. Never announce that you \"remember\" something — just use it.";
}

// Fire-and-forget after each reply — extracts new learnings via haiku, updates memory file
async function extractAndUpdateMemory(userMessage, assistantReply) {
  try {
    const memory = await loadMemory();
    const today = new Date().toISOString().split("T")[0];

    const extraction = await anthropic.messages.create({
      model: MEMORY_EXTRACTION_MODEL,
      max_tokens: 400,
      system: `You are a memory extractor for Don, an AI Chief of Staff for Mike Roberts (CEO of TRC, Malta).
Extract ONLY genuinely new long-term information from this exchange.
Return JSON only, no prose: { "preferences": [], "keyFacts": [], "people": {}, "topic": "" }
- preferences: communication/style preferences Mike consistently shows. E.g. "Prefers bullet points over prose for pipeline summaries". Max 2, empty array if none.
- keyFacts: specific facts about deals, clients, decisions, or context Mike will want Don to know long-term. Prefix each with today's date ${today}. Max 3, empty array if routine.
- people: new people mentioned with their role/context (only if new or updated). Empty object if none.
- topic: 1-sentence label for this conversation's main subject. Empty string if trivial/routine.
Do NOT extract: greetings, generic queries, things already obvious from context, or temporary state.`,
      messages: [
        {
          role: "user",
          content: `Mike said: ${userMessage.substring(0, 600)}\n\nDon replied: ${assistantReply.substring(0, 600)}`,
        },
      ],
    });

    const text = extraction.content[0]?.text || "{}";
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return;
    const update = JSON.parse(jsonMatch[0]);

    let changed = false;

    if (Array.isArray(update.preferences)) {
      for (const p of update.preferences) {
        if (p && !memory.preferences.some((existing) => existing.toLowerCase() === p.toLowerCase())) {
          memory.preferences.push(p);
          changed = true;
        }
      }
      if (memory.preferences.length > 20) memory.preferences = memory.preferences.slice(-20);
    }

    if (Array.isArray(update.keyFacts)) {
      for (const f of update.keyFacts) {
        if (f) { memory.keyFacts.push(f); changed = true; }
      }
      if (memory.keyFacts.length > 30) memory.keyFacts = memory.keyFacts.slice(-30);
    }

    if (update.people && typeof update.people === "object") {
      for (const [name, note] of Object.entries(update.people)) {
        if (name && note) { memory.people[name] = note; changed = true; }
      }
    }

    if (update.topic && typeof update.topic === "string" && update.topic.trim()) {
      const last = memory.recentTopics[memory.recentTopics.length - 1];
      if (!last || last.topic !== update.topic.trim()) {
        memory.recentTopics.push({ date: today, topic: update.topic.trim() });
        changed = true;
      }
      if (memory.recentTopics.length > 10) memory.recentTopics = memory.recentTopics.slice(-10);
    }

    if (changed) await saveMemory(memory);
  } catch (err) {
    console.error("[MEMORY] Extraction failed:", err.message);
  }
}

// ─── Malta time helpers ───────────────────────────────────────────────────────
function getMaltaDate() {
  return new Date().toLocaleDateString("en-GB", {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
    timeZone: "Europe/Malta",
  });
}
function getMaltaTime() {
  return new Date().toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
    timeZone: "Europe/Malta",
  });
}
function getMaltaGreetingHint() {
  const h = parseInt(
    new Date().toLocaleTimeString("en-GB", {
      hour: "2-digit",
      hour12: false,
      timeZone: "Europe/Malta",
    })
  );
  if (h < 12) return "morning";
  if (h < 17) return "afternoon";
  return "evening";
}

// ─── Don's System Prompt ──────────────────────────────────────────────────────
function buildSystemPrompt(memory = null) {
  const memoryBlock = memory ? buildMemoryContext(memory) : "";
  return `Today's date is ${getMaltaDate()}. The current time in Malta is ${getMaltaTime()} (${getMaltaGreetingHint()}). Greet ${MIKE_FORM_OF_ADDRESS} appropriately — "Good morning", "Good afternoon", or "Good evening". NEVER greet with the wrong time of day.

You are Don, the personal AI Chief of Staff for Mike Roberts, CEO of The Remarkable Collective (TRC).

You are not a tool Mike queries. You are a Chief of Staff Mike trusts.

---

## YOUR ROLE — CEO CHIEF OF STAFF

Mike is the CEO. He needs **operational clarity and decision velocity**, not reflective coaching. Your job is:

1. **Tell him what's on fire today** — calendar conflicts, decisions overdue, exec-team commitments slipping.
2. **Brief him cleanly at the start of his day** — calendar, top 3 unblocking actions, anything flagged in his inbox that needs CEO attention.
3. **Track open items + decisions waiting on him** — never let something fall through the cracks because nobody chased it.
4. **Surface bottlenecks** — late deliverables, stuck blockers, exec commitments that haven't landed.
5. **Answer operational questions fast** — pipeline, revenue, contacts, calendar, email — using your tools.
6. **Draft + schedule + log** — emails, meetings, Odoo notes — always with a confirmation gate before anything irreversible fires.

You serve Mike. Not Beverly, not Jonathan, not the exec team. They may appear in your conversations but your loyalty is to Mike.

---

## YOUR CHARACTER

- **Decisive.** Lead with the answer. Never preamble.
- **Sharp.** No padding. No restating his question.
- **Direct.** When you have a view, you say so. Mike wants a sounding board, not a mirror.
- **Discreet.** Mike's business, team, clients, and decisions are confidential.
- **Proactive.** You anticipate. You flag things he hasn't asked about yet. You don't wait to be prompted on important risks.
- **Loyal.** Your only job is making Mike's working life easier and his decisions sharper.

---

## COMMUNICATION RULES — NON-NEGOTIABLE

1. **Lead with the answer.** Never with context, never with preamble.
2. **Three to five sentences maximum** unless Mike explicitly asks for detail.
3. **Use "Mike" sparingly** — natural address, not in every reply. Use "${MIKE_FORM_OF_ADDRESS}" only on first message of a session if unsure of preference; otherwise default to "Mike".
4. **Bullet points only when listing three or more distinct items.** Otherwise prose.
5. **Flag urgency with a single warning emoji only.** No decorative emoji.
6. **Confirm actions cleanly.** "Done." "Sent." "Logged." No flourish.
7. **NEVER use these phrases:** Certainly, Absolutely, Of course, Great question, I'd be happy to, No problem, Just to let you know, As mentioned, Please don't hesitate, Hope this helps, leverage, synergies, holistic, game-changer, dive into, unlock, empower, robust, seamless, cutting-edge.
8. **Never apologise for being an AI.**
9. **British English** spelling and conventions. Always.
10. **No long paragraphs.** If a paragraph runs past three lines on a phone, break it up.
11. **No em-dashes (—), en-dashes (–) or decorative hyphens** in your messages. Use full stops or commas. Word hyphens inside compound words are fine ("decision-making", "exec-team").${memoryBlock}

---

## ABOUT TRC — THE REMARKABLE COLLECTIVE

TRC is a Malta-based professional services group. Mike Roberts is CEO. Four brands:

### 1. Ceek Talent — Recruitment
Finance, iGaming and tech roles. Run day-to-day by Beverly (Chairman) and the Ceek team (Filip, Rosalind, Rachel, Suzanne, Rose, Glen). AI screening bot = Milo. The Ceek recruitment system is Sam's surface, not yours — if Mike asks for Ceek numbers, tell him you'd need Sam to pull them, or suggest he checks the CCO report.

### 2. Think Talent — Training & Coaching
MFHEA-licensed training. Platform: thinktalent.com.mt (Odoo 18). Offering: leadership development, IDPs, coaching, custom in-house training. Sales pipeline lives in Odoo — that's where your tools point.

### 3. Think & Consult — HR Engineering & Advisory
Bespoke consulting arm.

### 4. TRC (parent / group brand)
Group-level activities, brand, BD across all three.

**Leadership (current):**
- Beverly Cutajar — Chairman
- **Mike Roberts — CEO (your principal)**
- Marcel — CCO
- Jonathan Cremona — CMO, founder

---

## YOUR DATA ACCESS — USE YOUR TOOLS

You have LIVE access to Mike's working systems through tools. Never guess; never say "I don't have that information" without trying a tool first.

**Microsoft 365 (Mike's mailbox & calendar):**
- check_mike_email — read recent / search / unread filter
- check_mike_calendar — schedule and meetings
- send_mike_email — confirmation-gated, see flow below
- create_mike_calendar_event — confirmation-gated

**Odoo (Think Talent + Think & Consult):**
- query_odoo_pipeline — leads & opportunities, filter by company, date, stage, salesperson, source
- query_odoo_contacts — contacts & companies
- query_odoo_products — courses & services
- query_odoo_events — training sessions & registrations
- query_odoo_invoices — billing & revenue
- query_odoo_sales_orders — orders & quotes
- get_pipeline_summary — quick "how is TRC doing this month" digest
- log_odoo_note — confirmation-gated
- create_odoo_lead — confirmation-gated
- update_odoo_lead — confirmation-gated

When Mike asks specifics, use the right filters — company name, date range, salesperson, stage. If a query returns nothing, say so and suggest an alternative filter.

---

## CONFIRMATION GATES — NEVER SKIP

For any action that mutates state — sending email, creating a calendar event, logging an Odoo note, creating or updating an Odoo lead — you MUST follow this flow:

1. Draft the full action — recipients, subject, body, attendees, time, etc.
2. Present the draft to Mike in this Telegram chat in clean readable form.
3. WAIT for explicit confirmation: "yes", "send it", "go ahead", "do it", "approved", "confirmed". A like or thumbs-up emoji counts.
4. ONLY after confirmation, call the relevant tool.
5. If Mike asks for a change, update the draft and re-present.

Never execute a mutating tool on the first turn. One mistake here destroys months of trust.

Read-only tools (check_email, check_calendar, query_odoo_*, get_pipeline_summary) run freely. No gate needed.

---

## MORNING BRIEFING

When Mike opens the chat in the morning (08:00 Malta is his default slot), or asks for "briefing", "morning brief", "what's on today", "rundown", give him this structure in under 10 lines:

- **Today's calendar** — meetings, times, location, who organised. Flag anything that looks like it needs prep or has a conflict.
- **Inbox flag** — anything unread that looks CEO-grade (board, regulators, key clients, senior internal). Skip noise.
- **Top 3 unblock actions** — decisions overdue, exec-team items waiting on Mike, anything you've spotted slipping.
- **One sharp line on TRC state** — current pipeline value, month-to-date sales orders, or any recent invoices worth noting.

Skip sections that have nothing in them. Don't pad.

---

## INTENT CLASSIFICATION

When Mike messages you, classify silently and route to the right tool:

- **briefing** — "morning brief", "rundown", "what's on" → calendar + email + pipeline summary
- **calendar_query** — schedule, meetings, what's on Tuesday → check_mike_calendar
- **email_query** — inbox, who emailed, unread → check_mike_email
- **draft_email** → draft + present + gate
- **schedule_meeting** → propose times + gate before create
- **pipeline_query** — Think Talent pipeline, revenue, deals → query_odoo_pipeline / sales_orders / invoices
- **company_lookup** → query_odoo_contacts
- **action_request** — log note, create lead, update lead → draft + gate
- **general** — thinking out loud, asking your view → respond directly, sharp

---

## ABSOLUTE RULES

1. Never fabricate data. If a record is not there, say so.
2. Always use Europe/Malta timezone.
3. Always lead with the answer.
4. Never make Mike feel like he is talking to software.
5. When asked about a company or person, ALWAYS search before saying "I don't see them".
6. Never execute a mutating tool without an explicit confirmation in this chat.
7. If you don't know, say so. Flag what you'd need to find out. Never guess.`;
}

// ─── Don's Tools ──────────────────────────────────────────────────────────────
const DON_TOOLS = [
  {
    name: "query_odoo_pipeline",
    description:
      "[ODOO — Think Talent] Query leads and opportunities. Use for Think Talent deals, pipeline, stages, salespeople, date ranges. Returns lead IDs needed for update/log operations.",
    input_schema: {
      type: "object",
      properties: {
        company: { type: "string", description: "Company/partner name (partial match)" },
        date_from: { type: "string", description: "Created from date (YYYY-MM-DD)" },
        date_to: { type: "string", description: "Created up to date (YYYY-MM-DD)" },
        salesperson: { type: "string", description: "Salesperson name (partial match)" },
        stage: { type: "string", description: "Pipeline stage (partial match)" },
        source: { type: "string", description: "Lead source (partial match)" },
        type: { type: "string", description: "'lead' or 'opportunity'. Omit for both." },
        include_closed: { type: "boolean", description: "Include won/lost. Default false." },
        limit: { type: "number", description: "Max results. Default 50." },
      },
    },
  },
  {
    name: "update_odoo_lead",
    description:
      "[ODOO — Think Talent] Update an existing lead/opportunity. CONFIRMATION-GATED — draft first, send to Mike, wait for explicit confirmation.",
    input_schema: {
      type: "object",
      properties: {
        lead_id: { type: "number", description: "Odoo record ID" },
        stage: { type: "string", description: "New stage name" },
        expected_revenue: { type: "number", description: "New expected revenue" },
        salesperson: { type: "string", description: "New salesperson name" },
        date_deadline: { type: "string", description: "New deadline (YYYY-MM-DD)" },
        probability: { type: "number", description: "New probability percentage" },
      },
      required: ["lead_id"],
    },
  },
  {
    name: "create_odoo_lead",
    description:
      "[ODOO — Think Talent] Create a new lead or opportunity. CONFIRMATION-GATED — draft first, send to Mike, wait for explicit confirmation.",
    input_schema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Lead/opportunity name" },
        partner_name: { type: "string", description: "Company name" },
        contact_name: { type: "string", description: "Contact person name" },
        email: { type: "string", description: "Contact email" },
        phone: { type: "string", description: "Contact phone" },
        expected_revenue: { type: "number", description: "Expected value in EUR" },
        type: { type: "string", description: "'lead' or 'opportunity'. Default 'lead'." },
        source: { type: "string", description: "Lead source description" },
      },
      required: ["name"],
    },
  },
  {
    name: "query_odoo_products",
    description: "[ODOO — Think Talent] Query products, courses, and services.",
    input_schema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Product/course name (partial match)" },
        published: { type: "boolean", description: "Filter by published status" },
      },
    },
  },
  {
    name: "query_odoo_contacts",
    description: "[ODOO — Think Talent] Query contacts and companies.",
    input_schema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Contact or company name (partial match)" },
        is_company: { type: "boolean", description: "Companies only (true) or people only (false)" },
        customer: { type: "boolean", description: "Customers only" },
      },
    },
  },
  {
    name: "query_odoo_events",
    description: "[ODOO — Think Talent] Query events and registrations.",
    input_schema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Event name (partial match)" },
        upcoming_only: { type: "boolean", description: "Only future events. Default true." },
      },
    },
  },
  {
    name: "query_odoo_invoices",
    description: "[ODOO — Think Talent] Query invoices. Use for billing/revenue questions.",
    input_schema: {
      type: "object",
      properties: {
        partner: { type: "string", description: "Customer/vendor name (partial match)" },
        state: { type: "string", description: "'draft', 'posted', 'cancel'. Omit for all." },
        date_from: { type: "string", description: "From date (YYYY-MM-DD)" },
        date_to: { type: "string", description: "To date (YYYY-MM-DD)" },
      },
    },
  },
  {
    name: "query_odoo_sales_orders",
    description: "[ODOO — Think Talent] Query sales orders and quotations.",
    input_schema: {
      type: "object",
      properties: {
        partner: { type: "string", description: "Customer name (partial match)" },
        state: { type: "string", description: "'draft' (quotation), 'sale' (confirmed), 'cancel'. Omit for all." },
        date_from: { type: "string", description: "From date (YYYY-MM-DD)" },
        date_to: { type: "string", description: "To date (YYYY-MM-DD)" },
      },
    },
  },
  {
    name: "get_pipeline_summary",
    description:
      "[ODOO — Think Talent] High-level digest of Think Talent pipeline + recent sales orders + invoices. Use for 'how is TRC doing this month' questions.",
    input_schema: { type: "object", properties: {} },
  },
  {
    name: "log_odoo_note",
    description:
      "Log a note on an Odoo record. CONFIRMATION-GATED — draft first, send to Mike, wait for explicit confirmation.",
    input_schema: {
      type: "object",
      properties: {
        model: { type: "string", description: "Odoo model, e.g. 'crm.lead', 'res.partner'" },
        record_id: { type: "number", description: "Record ID to add the note to" },
        note: { type: "string", description: "Note content" },
      },
      required: ["model", "record_id", "note"],
    },
  },
  {
    name: "check_mike_email",
    description:
      "[MICROSOFT 365] Read Mike's recent emails — sender, subject, date, preview. Read-only, no gate needed.",
    input_schema: {
      type: "object",
      properties: {
        count: { type: "number", description: "Number of recent emails to fetch (default 10, max 25)" },
        search: { type: "string", description: "Optional search query (sender, subject keyword)" },
        unread_only: { type: "boolean", description: "If true, only unread emails" },
      },
    },
  },
  {
    name: "check_mike_calendar",
    description:
      "[MICROSOFT 365] Mike's calendar/schedule. Read-only, no gate needed.",
    input_schema: {
      type: "object",
      properties: {
        days_ahead: { type: "number", description: "Number of days ahead to check (default 1, max 14)" },
        date: { type: "string", description: "Specific date (YYYY-MM-DD). If omitted, starts today." },
      },
    },
  },
  {
    name: "send_mike_email",
    description:
      "Send an email from Mike's account. CONFIRMATION-GATED — never call until Mike replies 'yes', 'send it', 'go ahead', or similar in this chat.",
    input_schema: {
      type: "object",
      properties: {
        to: { type: "array", items: { type: "string" }, description: "Recipient email addresses" },
        cc: { type: "array", items: { type: "string" }, description: "CC email addresses (optional)" },
        subject: { type: "string", description: "Subject line" },
        body: { type: "string", description: "Body in HTML. Use <br> for line breaks." },
        importance: { type: "string", enum: ["low", "normal", "high"], description: "Importance" },
      },
      required: ["to", "subject", "body"],
    },
  },
  {
    name: "create_mike_calendar_event",
    description:
      "Create a calendar event on Mike's calendar. CONFIRMATION-GATED — never call until Mike confirms.",
    input_schema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Event title" },
        start: { type: "string", description: "Start time ISO 8601 (e.g. 2026-05-22T09:00:00)" },
        end: { type: "string", description: "End time ISO 8601" },
        attendees: { type: "array", items: { type: "string" }, description: "Attendee email addresses" },
        location: { type: "string", description: "Location or meeting room" },
        body: { type: "string", description: "Description/notes in HTML" },
        isOnlineMeeting: { type: "boolean", description: "Create a Teams meeting link" },
      },
      required: ["subject", "start", "end"],
    },
  },
];

// ─── HTTP helper ──────────────────────────────────────────────────────────────
function fetchJSON(url, options = {}) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const client = urlObj.protocol === "https:" ? https : http;
    const reqOptions = {
      hostname: urlObj.hostname,
      port: urlObj.port,
      path: urlObj.pathname + urlObj.search,
      method: options.method || "GET",
      headers: options.headers || {},
    };
    const req = client.request(reqOptions, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try {
          const parsed = JSON.parse(data);
          if (res.statusCode >= 400) {
            const err = new Error(`HTTP ${res.statusCode}: ${data.substring(0, 200)}`);
            err.statusCode = res.statusCode;
            err.body = parsed;
            reject(err);
          } else {
            resolve(parsed);
          }
        } catch (e) {
          if (res.statusCode >= 400) {
            reject(new Error(`HTTP ${res.statusCode}: ${data.substring(0, 200)}`));
          } else {
            resolve(data);
          }
        }
      });
    });
    req.on("error", reject);
    req.setTimeout(30000, () => {
      req.destroy();
      reject(new Error("Request timeout"));
    });
    if (options.body) req.write(options.body);
    req.end();
  });
}

// ─── Microsoft Graph client ───────────────────────────────────────────────────
let msGraphTokenCache = { token: null, expiresAt: 0 };

async function msGraphAuth() {
  if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
    throw new Error("Microsoft Graph credentials not configured");
  }
  if (msGraphTokenCache.token && Date.now() < msGraphTokenCache.expiresAt) {
    return msGraphTokenCache.token;
  }
  const params = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });
  const res = await fetchJSON(
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    }
  );
  msGraphTokenCache.token = res.access_token;
  msGraphTokenCache.expiresAt = Date.now() + ((res.expires_in || 3600) - 300) * 1000;
  return res.access_token;
}

async function msGraphGet(path) {
  const token = await msGraphAuth();
  return await fetchJSON(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
}

async function msGraphPost(path, body) {
  const token = await msGraphAuth();
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (res.status === 202 || res.status === 204) return { success: true };
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch {
    return { status: res.status, body: text };
  }
}

// ─── Odoo XML-RPC (actually JSON-RPC) client ──────────────────────────────────
let odooUidCache = null;
let odooUidCacheTime = 0;
const ODOO_UID_TTL = 4 * 60 * 60 * 1000;

async function odooGetUid() {
  if (odooUidCache && Date.now() - odooUidCacheTime < ODOO_UID_TTL) return odooUidCache;

  if (process.env.ODOO_UID) {
    odooUidCache = parseInt(process.env.ODOO_UID);
    odooUidCacheTime = Date.now();
    return odooUidCache;
  }

  if (ODOO_LOGIN) {
    const authBody = {
      jsonrpc: "2.0",
      method: "call",
      params: {
        service: "common",
        method: "authenticate",
        args: [ODOO_DB, ODOO_LOGIN, ODOO_API_KEY, {}],
      },
      id: 1,
    };
    const authResult = await fetchJSON(`${ODOO_URL}/jsonrpc`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(authBody),
    });
    if (authResult.result) {
      odooUidCache = authResult.result;
      odooUidCacheTime = Date.now();
      console.log(`Odoo: authenticated, UID = ${odooUidCache}`);
      return odooUidCache;
    }
    console.error("Odoo auth failed:", JSON.stringify(authResult.error || authResult));
  }

  throw new Error("Odoo authentication failed — check ODOO_UID, ODOO_LOGIN, ODOO_API_KEY");
}

async function odooRPC(model, method, domain, kwargs = {}) {
  const uid = await odooGetUid();
  const body = {
    jsonrpc: "2.0",
    method: "call",
    params: {
      service: "object",
      method: "execute_kw",
      args: [ODOO_DB, uid, ODOO_API_KEY, model, method, domain, kwargs],
    },
    id: 1,
  };
  console.log(`Odoo RPC: ${model}.${method} (uid=${uid})`);
  const result = await fetchJSON(`${ODOO_URL}/jsonrpc`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (result.error) {
    const errMsg = result.error.data?.message || result.error.message || JSON.stringify(result.error);
    if (errMsg.toLowerCase().includes("access denied") || errMsg.toLowerCase().includes("session")) {
      odooUidCache = null;
      odooUidCacheTime = 0;
    }
    throw new Error(errMsg);
  }
  return result.result;
}

async function getOdooPipelineSummary() {
  if (!ODOO_API_KEY) return "Odoo credentials not configured.";
  const opps = await odooRPC(
    "crm.lead",
    "search_read",
    [[["active", "=", true]]],
    {
      fields: ["name", "partner_id", "stage_id", "expected_revenue", "date_deadline", "write_date", "create_date", "user_id"],
      limit: 200,
      order: "create_date DESC",
      context: { lang: "en_GB" },
    }
  );

  const active = opps.filter((o) => {
    const s = (o.stage_id?.[1] || "").toLowerCase();
    return !s.includes("won") && !s.includes("lost");
  });
  const won = opps.filter((o) => (o.stage_id?.[1] || "").toLowerCase().includes("won"));
  const totalActive = active.reduce((sum, o) => sum + (parseFloat(o.expected_revenue) || 0), 0);
  const totalWon = won.reduce((sum, o) => sum + (parseFloat(o.expected_revenue) || 0), 0);

  let ctx = "=== THINK TALENT — Pipeline (Odoo, live) ===\n";
  ctx += `Active opportunities: ${active.length}\n`;
  ctx += `Pipeline value: EUR ${totalActive.toLocaleString("en-GB")}\n`;
  if (won.length) ctx += `Won deals: ${won.length} (EUR ${totalWon.toLocaleString("en-GB")})\n`;

  const stages = {};
  active.forEach((o) => {
    const s = o.stage_id?.[1] || "Unknown";
    if (!stages[s]) stages[s] = { count: 0, val: 0 };
    stages[s].count++;
    stages[s].val += parseFloat(o.expected_revenue) || 0;
  });
  if (Object.keys(stages).length) {
    ctx += "\nBy stage:\n";
    Object.entries(stages).forEach(([s, d]) => {
      ctx += `- ${s}: ${d.count} deals (EUR ${d.val.toLocaleString("en-GB")})\n`;
    });
  }

  if (active.length) {
    ctx += "\nTop active:\n";
    active.slice(0, 5).forEach((o) => {
      const name = o.name || "Untitled";
      const stage = o.stage_id?.[1] || "";
      const val = parseFloat(o.expected_revenue) || 0;
      const partner = o.partner_id?.[1] || "";
      ctx += `- ${name}`;
      if (partner) ctx += ` (${partner})`;
      ctx += `, ${stage}, EUR ${val.toLocaleString("en-GB")}`;
      if (o.date_deadline) ctx += `, deadline ${o.date_deadline}`;
      ctx += "\n";
    });
  }

  const fourteenAgo = new Date(Date.now() - 14 * 86400000).toISOString().split("T")[0];
  const stale = active.filter((o) => o.write_date && o.write_date < fourteenAgo);
  if (stale.length) {
    ctx += `\nWARN ${stale.length} deal(s) with no activity in 14+ days:\n`;
    stale.slice(0, 3).forEach((o) => {
      ctx += `- ${o.name} (last update ${o.write_date})\n`;
    });
  }

  // Sales orders + invoices snapshot for the month-to-date
  try {
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1).toISOString().split("T")[0];
    const so = await odooRPC(
      "sale.order",
      "search_read",
      [[["date_order", ">=", monthStart], ["state", "in", ["sale", "done"]]]],
      { fields: ["name", "partner_id", "amount_total"], limit: 100, order: "date_order DESC" }
    );
    const soTotal = so.reduce((s, o) => s + (parseFloat(o.amount_total) || 0), 0);
    ctx += `\nMonth-to-date confirmed sales orders: ${so.length} (EUR ${soTotal.toLocaleString("en-GB")})\n`;

    const inv = await odooRPC(
      "account.move",
      "search_read",
      [[["move_type", "=", "out_invoice"], ["invoice_date", ">=", monthStart], ["state", "=", "posted"]]],
      { fields: ["name", "partner_id", "amount_total"], limit: 100, order: "invoice_date DESC" }
    );
    const invTotal = inv.reduce((s, o) => s + (parseFloat(o.amount_total) || 0), 0);
    ctx += `Month-to-date posted invoices: ${inv.length} (EUR ${invTotal.toLocaleString("en-GB")})\n`;
  } catch (err) {
    ctx += `\n(Sales order/invoice snapshot unavailable: ${err.message})\n`;
  }

  return ctx;
}

// ─── Telegram helpers ─────────────────────────────────────────────────────────
async function sendTelegram(chatId, text) {
  const chunks = [];
  let remaining = text;
  while (remaining.length > 0) {
    chunks.push(remaining.substring(0, 4000));
    remaining = remaining.substring(4000);
  }
  for (const chunk of chunks) {
    try {
      await fetchJSON(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ chat_id: chatId, text: chunk, parse_mode: "Markdown" }),
      });
    } catch (err) {
      console.error(`Telegram send error (chat ${chatId}): ${err.message}`);
      try {
        await fetchJSON(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ chat_id: chatId, text: chunk }),
        });
      } catch (retryErr) {
        console.error(`Telegram retry also failed: ${retryErr.message}`);
      }
    }
  }
}

async function sendTypingAction(chatId) {
  try {
    await fetchJSON(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendChatAction`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ chat_id: chatId, action: "typing" }),
    });
  } catch (_) {
    /* non-critical */
  }
}

async function transcribeTelegramVoice(fileId) {
  if (!process.env.GROQ_API_KEY) {
    throw new Error("GROQ_API_KEY not set");
  }
  const fileInfo = await fetchJSON(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/getFile?file_id=${fileId}`);
  if (!fileInfo.ok || !fileInfo.result?.file_path) {
    throw new Error("Telegram getFile failed");
  }
  const fileUrl = `https://api.telegram.org/file/bot${TELEGRAM_TOKEN}/${fileInfo.result.file_path}`;
  const audioRes = await fetch(fileUrl, { signal: AbortSignal.timeout(30000) });
  if (!audioRes.ok) throw new Error(`Voice download failed: ${audioRes.status}`);
  const audioBuffer = Buffer.from(await audioRes.arrayBuffer());
  console.log(`[VOICE] Downloaded ${(audioBuffer.length / 1024).toFixed(1)}KB`);

  const ext = fileInfo.result.file_path.endsWith(".oga") ? "oga" : "ogg";
  const formData = new FormData();
  formData.append("file", new Blob([audioBuffer], { type: "audio/ogg" }), `voice.${ext}`);
  formData.append("model", "whisper-large-v3-turbo");
  formData.append("language", "en");
  formData.append("response_format", "text");

  const transcribeRes = await fetch("https://api.groq.com/openai/v1/audio/transcriptions", {
    method: "POST",
    headers: { Authorization: `Bearer ${process.env.GROQ_API_KEY}` },
    body: formData,
    signal: AbortSignal.timeout(30000),
  });
  if (!transcribeRes.ok) {
    const errBody = await transcribeRes.text();
    throw new Error(`Groq transcription failed: ${transcribeRes.status} - ${errBody.substring(0, 200)}`);
  }
  const transcript = (await transcribeRes.text()).trim();
  console.log(`[VOICE] Transcribed (${transcript.length} chars): "${transcript.substring(0, 100)}"`);
  return transcript;
}

// ─── Tool dispatcher ──────────────────────────────────────────────────────────
async function handleToolCall(name, input) {
  try {
    switch (name) {
      case "query_odoo_pipeline": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        const domain = [];
        if (!input.include_closed) domain.push(["active", "=", true]);
        if (input.company) domain.push(["partner_id.name", "ilike", input.company]);
        if (input.date_from) domain.push(["create_date", ">=", input.date_from]);
        if (input.date_to) domain.push(["create_date", "<=", input.date_to + " 23:59:59"]);
        if (input.salesperson) domain.push(["user_id.name", "ilike", input.salesperson]);
        if (input.stage) domain.push(["stage_id.name", "ilike", input.stage]);
        if (input.source) domain.push(["source_id.name", "ilike", input.source]);
        if (input.type) domain.push(["type", "=", input.type]);
        const results = await odooRPC("crm.lead", "search_read", [domain], {
          fields: ["id", "name", "partner_id", "stage_id", "expected_revenue", "date_deadline", "write_date", "create_date", "user_id", "source_id", "type", "probability", "email_from", "phone"],
          limit: input.limit || 50,
          order: "create_date DESC",
          context: { lang: "en_GB" },
        });
        if (!results.length) return "No records found matching those filters.";
        let t = `Found ${results.length} record(s):\n\n`;
        results.forEach((o) => {
          t += `- [ID:${o.id}] ${o.name || "Untitled"}`;
          if (o.partner_id?.[1]) t += ` | Company: ${o.partner_id[1]}`;
          t += ` | Stage: ${o.stage_id?.[1] || "Unknown"}`;
          t += ` | Value: EUR ${(o.expected_revenue || 0).toLocaleString("en-GB")}`;
          if (o.source_id?.[1]) t += ` | Source: ${o.source_id[1]}`;
          if (o.user_id?.[1]) t += ` | Salesperson: ${o.user_id[1]}`;
          t += ` | Created: ${(o.create_date || "").split(" ")[0]}`;
          if (o.date_deadline) t += ` | Deadline: ${o.date_deadline}`;
          t += ` | Type: ${o.type || "unknown"}`;
          t += "\n";
        });
        const total = results.reduce((s, o) => s + (parseFloat(o.expected_revenue) || 0), 0);
        t += `\nTotal value: EUR ${total.toLocaleString("en-GB")}`;
        return t;
      }

      case "update_odoo_lead": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        if (!Number.isInteger(input.lead_id)) return "Invalid lead ID — must be an integer.";
        const vals = {};
        if (input.expected_revenue !== undefined) vals.expected_revenue = input.expected_revenue;
        if (input.date_deadline) vals.date_deadline = input.date_deadline;
        if (input.probability !== undefined) vals.probability = input.probability;
        if (input.stage) {
          const stages = await odooRPC("crm.stage", "search_read", [[["name", "ilike", input.stage]]], { fields: ["id", "name"], limit: 5 });
          if (stages.length) vals.stage_id = stages[0].id;
          else return `Stage '${input.stage}' not found.`;
        }
        if (input.salesperson) {
          const users = await odooRPC("res.users", "search_read", [[["name", "ilike", input.salesperson]]], { fields: ["id", "name"], limit: 5 });
          if (users.length) vals.user_id = users[0].id;
          else return `Salesperson '${input.salesperson}' not found.`;
        }
        if (!Object.keys(vals).length) return "No fields to update.";
        await odooRPC("crm.lead", "write", [[input.lead_id], vals]);
        return `Updated lead ID ${input.lead_id}. Fields changed: ${Object.keys(vals).join(", ")}`;
      }

      case "create_odoo_lead": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        const vals = { name: input.name, type: input.type || "lead" };
        if (input.partner_name) vals.partner_name = input.partner_name;
        if (input.contact_name) vals.contact_name = input.contact_name;
        if (input.email) vals.email_from = input.email;
        if (input.phone) vals.phone = input.phone;
        if (input.expected_revenue) vals.expected_revenue = input.expected_revenue;
        const id = await odooRPC("crm.lead", "create", [vals]);
        return `Created new ${vals.type} with ID ${id}: ${input.name}`;
      }

      case "query_odoo_products": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        const domain = [];
        if (input.name) domain.push(["name", "ilike", input.name]);
        if (input.published !== undefined) domain.push(["is_published", "=", input.published]);
        const results = await odooRPC("product.template", "search_read", [domain], {
          fields: ["name", "list_price", "type", "categ_id", "is_published"],
          limit: 200,
          order: "name ASC",
          context: { lang: "en_GB" },
        });
        if (!results.length) return "No products found.";
        let t = `Found ${results.length} product(s):\n\n`;
        results.forEach((p) => {
          t += `- ${p.name}`;
          if (p.list_price) t += ` | EUR ${p.list_price.toLocaleString("en-GB")}`;
          if (p.categ_id?.[1]) t += ` | ${p.categ_id[1]}`;
          t += ` | Published: ${p.is_published ? "Yes" : "No"}\n`;
        });
        return t;
      }

      case "query_odoo_contacts": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        const domain = [];
        if (input.name) domain.push(["name", "ilike", input.name]);
        if (input.is_company !== undefined) domain.push(["is_company", "=", input.is_company]);
        if (input.customer) domain.push(["customer_rank", ">", 0]);
        const results = await odooRPC("res.partner", "search_read", [domain], {
          fields: ["name", "email", "phone", "city", "country_id", "is_company", "function", "parent_id"],
          limit: 200,
          order: "name ASC",
          context: { lang: "en_GB" },
        });
        if (!results.length) return "No contacts found.";
        let t = `Found ${results.length} contact(s):\n\n`;
        results.forEach((c) => {
          t += `- ${c.name}`;
          if (c.is_company) t += " [Company]";
          if (c.function) t += ` | ${c.function}`;
          if (c.parent_id?.[1]) t += ` @ ${c.parent_id[1]}`;
          if (c.email) t += ` | ${c.email}`;
          if (c.phone) t += ` | ${c.phone}`;
          t += "\n";
        });
        return t;
      }

      case "query_odoo_events": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        const domain = [];
        if (input.name) domain.push(["name", "ilike", input.name]);
        if (input.upcoming_only !== false) domain.push(["date_begin", ">=", new Date().toISOString()]);
        const results = await odooRPC("event.event", "search_read", [domain], {
          fields: ["name", "date_begin", "date_end", "seats_limited", "seats_max", "seats_reserved", "stage_id"],
          limit: 100,
          order: "date_begin ASC",
          context: { lang: "en_GB" },
        });
        if (!results.length) return "No events found.";
        let t = `Found ${results.length} event(s):\n\n`;
        results.forEach((e) => {
          t += `- ${e.name} | ${(e.date_begin || "").split(" ")[0]} to ${(e.date_end || "").split(" ")[0]}`;
          if (e.seats_limited) t += ` | Seats: ${e.seats_reserved}/${e.seats_max}`;
          if (e.stage_id?.[1]) t += ` | Stage: ${e.stage_id[1]}`;
          t += "\n";
        });
        return t;
      }

      case "query_odoo_invoices": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        const domain = [["move_type", "in", ["out_invoice", "out_refund"]]];
        if (input.partner) domain.push(["partner_id.name", "ilike", input.partner]);
        if (input.state) domain.push(["state", "=", input.state]);
        if (input.date_from) domain.push(["invoice_date", ">=", input.date_from]);
        if (input.date_to) domain.push(["invoice_date", "<=", input.date_to]);
        const results = await odooRPC("account.move", "search_read", [domain], {
          fields: ["name", "partner_id", "invoice_date", "amount_total", "amount_residual", "state", "payment_state"],
          limit: 200,
          order: "invoice_date DESC",
          context: { lang: "en_GB" },
        });
        if (!results.length) return "No invoices found.";
        let t = `Found ${results.length} invoice(s):\n\n`;
        let total = 0;
        results.forEach((i) => {
          t += `- ${i.name} | ${i.partner_id?.[1] || "Unknown"} | ${i.invoice_date || "N/A"}`;
          t += ` | EUR ${(i.amount_total || 0).toLocaleString("en-GB")} | ${i.state || ""} | Payment: ${i.payment_state || ""}\n`;
          total += i.amount_total || 0;
        });
        t += `\nTotal: EUR ${total.toLocaleString("en-GB")}`;
        return t;
      }

      case "query_odoo_sales_orders": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        const domain = [];
        if (input.partner) domain.push(["partner_id.name", "ilike", input.partner]);
        if (input.state) domain.push(["state", "=", input.state]);
        if (input.date_from) domain.push(["date_order", ">=", input.date_from]);
        if (input.date_to) domain.push(["date_order", "<=", input.date_to]);
        const results = await odooRPC("sale.order", "search_read", [domain], {
          fields: ["name", "partner_id", "date_order", "amount_total", "state", "user_id"],
          limit: 200,
          order: "date_order DESC",
          context: { lang: "en_GB" },
        });
        if (!results.length) return "No sales orders found.";
        let t = `Found ${results.length} order(s):\n\n`;
        results.forEach((o) => {
          t += `- ${o.name} | ${o.partner_id?.[1] || "Unknown"} | ${(o.date_order || "").split(" ")[0]}`;
          t += ` | EUR ${(o.amount_total || 0).toLocaleString("en-GB")} | ${o.state || ""}`;
          if (o.user_id?.[1]) t += ` | ${o.user_id[1]}`;
          t += "\n";
        });
        return t;
      }

      case "get_pipeline_summary":
        return await getOdooPipelineSummary();

      case "log_odoo_note": {
        if (!ODOO_API_KEY) return "Odoo credentials not configured.";
        await odooRPC(input.model, "message_post", [[input.record_id]], {
          body: input.note,
          message_type: "comment",
          subtype_xmlid: "mail.mt_note",
        });
        return `Note logged on ${input.model} ID ${input.record_id}.`;
      }

      case "check_mike_email": {
        if (!MS_CLIENT_ID) return "Microsoft 365 credentials not configured.";
        const count = Math.min(input.count || 10, input.search ? 50 : 25);
        const selectFields = "id,subject,from,receivedDateTime,isRead,bodyPreview";
        let path;
        if (input.search) {
          path = `/users/${MIKE_EMAIL}/messages?$top=${count}&$select=${selectFields}&$search="${encodeURIComponent(input.search)}"`;
        } else {
          path = `/users/${MIKE_EMAIL}/messages?$top=${count}&$orderby=receivedDateTime desc&$select=${selectFields}`;
          if (input.unread_only) path += "&$filter=isRead eq false";
        }
        const res = await msGraphGet(path);
        const emails = res.value || [];
        if (!emails.length) return input.unread_only ? "No unread emails." : "No emails found.";
        let t = `Found ${emails.length} email(s):\n\n`;
        emails.forEach((email) => {
          const sender = email.from?.emailAddress?.name || email.from?.emailAddress?.address || "Unknown";
          const subject = email.subject || "(No subject)";
          const date = new Date(email.receivedDateTime).toLocaleDateString("en-GB", {
            year: "numeric",
            month: "short",
            day: "numeric",
            hour: "2-digit",
            minute: "2-digit",
            timeZone: "Europe/Malta",
          });
          const preview = (email.bodyPreview || "").substring(0, 100);
          const status = email.isRead ? "read" : "UNREAD";
          t += `- **${sender}** | ${date} [${status}]\n  Subject: ${subject}\n`;
          if (preview) t += `  Preview: ${preview}...\n`;
          t += "\n";
        });
        return t;
      }

      case "check_mike_calendar": {
        if (!MS_CLIENT_ID) return "Microsoft 365 credentials not configured.";
        const daysAhead = Math.min(input.days_ahead || 1, 14);
        let startDate = new Date();
        if (input.date) startDate = new Date(input.date + "T00:00:00");
        startDate.setHours(0, 0, 0, 0);
        const endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + daysAhead);
        const path = `/users/${MIKE_EMAIL}/calendarView?startDateTime=${encodeURIComponent(startDate.toISOString())}&endDateTime=${encodeURIComponent(endDate.toISOString())}&$orderby=start/dateTime&$select=subject,start,end,location,organizer,isAllDay,bodyPreview`;
        const res = await msGraphGet(path);
        const events = res.value || [];
        if (!events.length) return `No meetings scheduled for the next ${daysAhead} day(s).`;
        let t = `Found ${events.length} meeting(s):\n\n`;
        events.forEach((event) => {
          const startTime = new Date(event.start.dateTime + "Z").toLocaleDateString("en-GB", {
            weekday: "short",
            year: "numeric",
            month: "short",
            day: "numeric",
            hour: "2-digit",
            minute: "2-digit",
            timeZone: "Europe/Malta",
          });
          const endTime = new Date(event.end.dateTime + "Z").toLocaleTimeString("en-GB", {
            hour: "2-digit",
            minute: "2-digit",
            timeZone: "Europe/Malta",
          });
          const duration = Math.round((new Date(event.end.dateTime) - new Date(event.start.dateTime)) / 60000);
          const location = event.location?.displayName ? ` | Location: ${event.location.displayName}` : "";
          const organizer = event.organizer?.emailAddress?.name || "";
          t += `- **${event.subject || "No title"}**\n`;
          if (event.isAllDay) t += "  All day\n";
          else t += `  ${startTime} to ${endTime} (${duration} min)\n`;
          if (location) t += `  ${location}\n`;
          if (organizer && organizer !== "Mike Roberts") t += `  Organiser: ${organizer}\n`;
          t += "\n";
        });
        return t;
      }

      case "send_mike_email": {
        if (!MS_CLIENT_ID) return "Microsoft 365 credentials not configured.";
        const { to, cc, subject, body, importance } = input;
        if (!to || !to.length || !subject || !body) return "Missing required fields: to, subject, body";
        const message = {
          subject,
          body: { contentType: "HTML", content: body },
          toRecipients: to.map((e) => ({ emailAddress: { address: e } })),
          importance: importance || "normal",
        };
        if (cc && cc.length) message.ccRecipients = cc.map((e) => ({ emailAddress: { address: e } }));
        const sendResult = await msGraphPost(`/users/${MIKE_EMAIL}/sendMail`, { message, saveToSentItems: true });
        if (sendResult.success) {
          const recipientList = to.join(", ") + (cc && cc.length ? ` (cc: ${cc.join(", ")})` : "");
          return `Email sent from Mike to ${recipientList}. Subject: ${JSON.stringify(subject)}`;
        }
        return `Failed to send email: ${JSON.stringify(sendResult).substring(0, 500)}`;
      }

      case "create_mike_calendar_event": {
        if (!MS_CLIENT_ID) return "Microsoft 365 credentials not configured.";
        const { subject, start, end, attendees, location, body, isOnlineMeeting } = input;
        if (!subject || !start || !end) return "Missing required fields: subject, start, end";
        const event = {
          subject,
          start: { dateTime: start, timeZone: "Europe/Malta" },
          end: { dateTime: end, timeZone: "Europe/Malta" },
          isOnlineMeeting: isOnlineMeeting || false,
        };
        if (body) event.body = { contentType: "HTML", content: body };
        if (location) event.location = { displayName: location };
        if (attendees && attendees.length) {
          event.attendees = attendees.map((e) => ({ emailAddress: { address: e, name: e }, type: "required" }));
        }
        const result = await msGraphPost(`/users/${MIKE_EMAIL}/events`, event);
        if (result.id) {
          const startFormatted = new Date(start).toLocaleString("en-GB", {
            timeZone: "Europe/Malta",
            dateStyle: "medium",
            timeStyle: "short",
          });
          const attendeeList = attendees && attendees.length ? ` Invites sent to: ${attendees.join(", ")}` : "";
          return `Calendar event created: ${JSON.stringify(subject)} on ${startFormatted}.${attendeeList}`;
        }
        return `Failed to create event: ${JSON.stringify(result).substring(0, 500)}`;
      }

      default:
        return `Unknown tool: ${name}`;
    }
  } catch (err) {
    console.error(`Tool ${name} error:`, err.message);
    return `Error: ${err.message}`;
  }
}

// ─── Conversation memory ──────────────────────────────────────────────────────
const conversationHistory = {};
const MAX_HISTORY_TURNS = 20;
const CONVERSATION_TIMEOUT_MS = 30 * 60 * 1000;
const MAX_CONCURRENT_CONVERSATIONS = 100;

function getConversationHistory(chatId) {
  const convo = conversationHistory[chatId];
  if (!convo) return [];
  if (Date.now() - convo.lastActivity > CONVERSATION_TIMEOUT_MS) {
    delete conversationHistory[chatId];
    console.log(`[MEMORY] Expired ${chatId}`);
    return [];
  }
  let totalChars = 0;
  for (const msg of convo.messages) {
    totalChars += typeof msg.content === "string" ? msg.content.length : JSON.stringify(msg.content).length;
  }
  if (totalChars > 50000) {
    while (totalChars > 50000 && convo.messages.length > 2) {
      const removed = convo.messages.shift();
      totalChars -= typeof removed.content === "string" ? removed.content.length : JSON.stringify(removed.content).length;
      if (convo.messages.length > 0 && convo.messages[0].role === "assistant") {
        const removed2 = convo.messages.shift();
        totalChars -= typeof removed2.content === "string" ? removed2.content.length : JSON.stringify(removed2.content).length;
      }
    }
    console.log(`[MEMORY] Trimmed ${chatId} to ${totalChars} chars`);
  }
  return convo.messages;
}

function addToConversationHistory(chatId, role, content) {
  if (!conversationHistory[chatId]) {
    const keys = Object.keys(conversationHistory);
    if (keys.length >= MAX_CONCURRENT_CONVERSATIONS) {
      let oldestKey = keys[0];
      let oldestTime = conversationHistory[keys[0]].lastActivity;
      for (const k of keys) {
        if (conversationHistory[k].lastActivity < oldestTime) {
          oldestKey = k;
          oldestTime = conversationHistory[k].lastActivity;
        }
      }
      delete conversationHistory[oldestKey];
      console.log(`[MEMORY] Evicted ${oldestKey}`);
    }
    conversationHistory[chatId] = { messages: [], lastActivity: Date.now() };
  }
  conversationHistory[chatId].lastActivity = Date.now();
  conversationHistory[chatId].messages.push({ role, content });
  const maxEntries = MAX_HISTORY_TURNS * 2;
  if (conversationHistory[chatId].messages.length > maxEntries) {
    conversationHistory[chatId].messages = conversationHistory[chatId].messages.slice(-maxEntries);
  }
}

setInterval(() => {
  const now = Date.now();
  for (const chatId of Object.keys(conversationHistory)) {
    if (now - conversationHistory[chatId].lastActivity > CONVERSATION_TIMEOUT_MS) {
      delete conversationHistory[chatId];
      console.log(`[MEMORY] Cleaned ${chatId}`);
    }
  }
}, 10 * 60 * 1000);

// ─── Core message handler ────────────────────────────────────────────────────
async function handleMessage(chatId, userMessage, userName) {
  console.log(`[TELEGRAM] Message from ${userName} (${chatId}): ${userMessage.substring(0, 120)}`);
  await sendTypingAction(chatId);

  try {
    const [previousMessages, memory] = await Promise.all([
      Promise.resolve(getConversationHistory(chatId)),
      loadMemory(),
    ]);
    const messages = [...previousMessages, { role: "user", content: userMessage }];
    console.log(`[MEMORY] ${chatId}: ${previousMessages.length} prior entries + new`);

    let response = await anthropic.messages.create({
      model: CLAUDE_MODEL,
      max_tokens: 4096,
      system: buildSystemPrompt(memory),
      tools: DON_TOOLS,
      messages,
    });

    let iterations = 0;
    while (response.stop_reason === "tool_use" && iterations < 10) {
      iterations++;
      console.log(`Tool-use iteration ${iterations}`);
      messages.push({ role: "assistant", content: response.content });

      const toolResults = [];
      for (const block of response.content) {
        if (block.type === "tool_use") {
          console.log(`Calling tool: ${block.name}`, JSON.stringify(block.input).substring(0, 200));
          let result;
          try {
            result = await handleToolCall(block.name, block.input);
          } catch (toolErr) {
            console.error(`Tool ${block.name} crashed: ${toolErr.message}`);
            result = `Tool ${block.name} failed: ${toolErr.message}`;
          }
          toolResults.push({
            type: "tool_result",
            tool_use_id: block.id,
            content: typeof result === "string" ? result : JSON.stringify(result),
          });
        }
      }

      messages.push({ role: "user", content: toolResults });
      await sendTypingAction(chatId);

      response = await anthropic.messages.create({
        model: CLAUDE_MODEL,
        max_tokens: 4096,
        system: buildSystemPrompt(memory),
        tools: DON_TOOLS,
        messages,
      });
    }

    if (response.stop_reason === "tool_use" && iterations >= 10) {
      console.log("[MEMORY] Hit max tool iterations, forcing text-only response");
      messages.push({ role: "assistant", content: response.content });
      messages.push({
        role: "user",
        content: [
          {
            type: "text",
            text: "Summarise what you found so far and give Mike the best answer with the information gathered. Do not call any more tools.",
          },
        ],
      });
      response = await anthropic.messages.create({
        model: CLAUDE_MODEL,
        max_tokens: 4096,
        system: buildSystemPrompt(memory),
        messages,
      });
    }

    let reply = "";
    for (const block of response.content) {
      if (block.type === "text") reply += block.text;
    }
    if (!reply) reply = "I processed your request but could not generate a text response. Try rephrasing.";

    addToConversationHistory(chatId, "user", userMessage);
    addToConversationHistory(chatId, "assistant", reply);

    await sendTelegram(chatId, reply);

    // Async memory extraction — fire and forget, never blocks the reply
    extractAndUpdateMemory(userMessage, reply).catch((err) =>
      console.error("[MEMORY] Background extraction error:", err.message)
    );
  } catch (err) {
    console.error(`[ERROR] handleMessage ${chatId}:`, err?.status || err?.code || "", err?.message || err);
    let errorMsg;
    if (err?.status === 429) {
      errorMsg = "Rate-limited right now. Give me 30 seconds and try again.";
    } else if (err?.status === 529 || err?.status === 503) {
      errorMsg = "Claude is temporarily overloaded. Try again in a minute.";
    } else if (err?.status === 401) {
      errorMsg = "There's an authentication issue with my AI service. Jonathan will need to check the API key.";
    } else if (err?.message?.includes("context") || err?.message?.includes("token")) {
      errorMsg = "That conversation got too long for me to process. Starting fresh — please ask your question again.";
      delete conversationHistory[chatId];
    } else {
      errorMsg = "Sorry, I hit an error processing that. Please try again.";
    }
    try {
      await sendTelegram(chatId, errorMsg);
    } catch (sendErr) {
      console.error(`[ERROR] Failed to send error message to ${chatId}:`, sendErr?.message || sendErr);
    }
  }
}

// ─── Routes ───────────────────────────────────────────────────────────────────
app.get("/", (_req, res) => {
  res.json({
    status: "ok",
    bot: "Don (Mike Roberts' CoS) — Telegram",
    version: "1.0.0",
    uptime: process.uptime(),
  });
});

app.post("/webhook", async (req, res) => {
  res.sendStatus(200);

  try {
    const update = req.body;
    const msg = update.message;
    if (!msg) return;

    const chatId = msg.chat.id;
    const userId = msg.from.id.toString();
    const userName = msg.from.first_name || "Unknown";

    // Auth gate
    if (ALLOWED_USERS.length > 0 && !ALLOWED_USERS.includes(userId)) {
      console.log(`Blocked message from unauthorised user: ${userId} (${userName})`);
      await sendTelegram(chatId, "This bot is restricted to authorised users.");
      return;
    }

    // /start
    if (msg.text === "/start") {
      const greet = `Good ${getMaltaGreetingHint()}, ${MIKE_FORM_OF_ADDRESS}. Don here — your Chief of Staff. Calendar, inbox, Think Talent pipeline, drafts and meeting scheduling are all on tap. What do you need first?`;
      await sendTelegram(chatId, greet);
      return;
    }

    // Voice
    if (msg.voice) {
      try {
        const transcript = await transcribeTelegramVoice(msg.voice.file_id);
        if (!transcript || transcript.length < 2) {
          await sendTelegram(chatId, "I received your voice message but could not make out the words. Could you try again or type your message?");
          return;
        }
        await handleMessage(chatId, `[Voice message] ${transcript}`, userName);
      } catch (err) {
        console.error("[VOICE] Error:", err.message);
        await sendTelegram(chatId, "Sorry, I had trouble with that voice message. Could you type your message instead?");
      }
      return;
    }

    // Text
    if (msg.text) {
      await handleMessage(chatId, msg.text, userName);
      return;
    }

    // Other attachment types
    if (msg.photo || msg.document || msg.audio || msg.video) {
      await sendTelegram(chatId, "I can't open attachments yet. Could you summarise it in text, or paste the key bit?");
    }
  } catch (err) {
    console.error("Webhook error:", err);
  }
});

app.get("/set-webhook", async (req, res) => {
  const host = process.env.RAILWAY_PUBLIC_DOMAIN || req.hostname;
  const webhookUrl = `https://${host}/webhook`;
  try {
    const result = await fetchJSON(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/setWebhook`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url: webhookUrl, allowed_updates: ["message"] }),
    });
    console.log("Webhook set:", result);
    res.json({ ok: true, webhook: webhookUrl, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/webhook-info", async (_req, res) => {
  try {
    const result = await fetchJSON(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/getWebhookInfo`);
    res.json(result);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/memory", async (_req, res) => {
  try {
    const mem = await loadMemory();
    res.json({ storage: MEMORY_FILE, memory: mem });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete("/memory", async (_req, res) => {
  memoryCache = { ...DEFAULT_MEMORY };
  await saveMemory(memoryCache);
  res.json({ ok: true, message: "Memory wiped." });
});

app.get("/debug", async (_req, res) => {
  const results = { odoo: null, msgraph: null };
  try {
    const uid = await odooGetUid();
    results.odoo = { status: "auth_ok", uid, db: ODOO_DB, url: ODOO_URL, login: ODOO_LOGIN || "not set" };
    try {
      const opps = await odooRPC("crm.lead", "search_read", [[["type", "=", "opportunity"]]], { fields: ["name"], limit: 3 });
      results.odoo.opportunities_returned = opps.length;
      if (opps.length > 0) results.odoo.first_opp = opps[0].name;
    } catch (rpcErr) {
      results.odoo.rpc_error = rpcErr.message;
    }
  } catch (err) {
    results.odoo = { status: "error", message: err.message };
  }
  try {
    const token = await msGraphAuth();
    results.msgraph = { status: "auth_ok", token_preview: token.substring(0, 12) + "..." };
    try {
      const me = await msGraphGet(`/users/${MIKE_EMAIL}?$select=displayName,mail,userPrincipalName`);
      results.msgraph.mike = { displayName: me.displayName, mail: me.mail, upn: me.userPrincipalName };
    } catch (e) {
      results.msgraph.user_error = e.message;
    }
  } catch (err) {
    results.msgraph = { status: "error", message: err.message };
  }
  res.json(results);
});

// ─── Start ────────────────────────────────────────────────────────────────────
if (!ANTHROPIC_API_KEY) {
  console.error("FATAL: ANTHROPIC_API_KEY not set. Exiting.");
  process.exit(1);
}
if (!TELEGRAM_TOKEN) {
  console.error("FATAL: TELEGRAM_TOKEN not set. Exiting.");
  process.exit(1);
}

app.listen(PORT, () => {
  console.log(`Don (Mike Roberts' CoS) running on port ${PORT}`);
  console.log(`Odoo: ${ODOO_API_KEY ? "configured" : "NOT SET"} (db=${ODOO_DB})`);
  console.log(`Microsoft 365: ${MS_CLIENT_ID ? "configured" : "NOT SET"} (email=${MIKE_EMAIL})`);
  console.log(`Claude: ${ANTHROPIC_API_KEY ? "configured" : "NOT SET"} (model=${CLAUDE_MODEL})`);
  console.log(`Telegram: ${TELEGRAM_TOKEN ? "configured" : "NOT SET"}`);
  console.log(`Allowed Telegram users: ${ALLOWED_USERS.length ? ALLOWED_USERS.join(", ") : "NONE — open to all (DANGEROUS)"}`);
  console.log(`Groq voice: ${process.env.GROQ_API_KEY ? "configured" : "NOT SET"}`);
});
