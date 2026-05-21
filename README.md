# Don — Mike Roberts' AI Chief of Staff

Telegram bot serving Mike Roberts, CEO of The Remarkable Collective (TRC). Step 2 of the TRC Chief of Staff rollout. Sibling to Sam (Beverly Cutajar's CoS) and to the future Marcel / Jonathan agents.

## Stack

- Node 18+, Express, `@anthropic-ai/sdk`
- Claude Sonnet 4 (`claude-sonnet-4-20250514`) via Messages API with tool use
- Hosted on Railway, project `don-cos`
- In-process conversation history (LRU + 30 min timeout) — no volume, no DB
- Telegram only (no WhatsApp — that's Sam's pattern)

## Integrations

- **Microsoft 365 Graph** — Mike's mailbox + calendar via Azure App "Don TRC Bot" (`be63c648-...`). Application permissions: `Mail-Advanced.ReadWrite.All`, `Mail.Send`, `Calendars.ReadWrite`, `Calendars.Read`, `Calendars.ReadBasic.All`, `User.Read.All`.
- **Odoo 18 XML-RPC** — Think Talent's CRM/sales/events/invoices on `thinktalent.com.mt`, db `thinktalent_prod`.
- **Groq Whisper** — voice-note transcription.

No Firefish. Mike does not directly run the Ceek recruitment desk — Beverly and the Ceek team do. Sam covers that surface.

## Tool surface

Read tools: `query_odoo_pipeline`, `query_odoo_contacts`, `query_odoo_products`, `query_odoo_events`, `query_odoo_invoices`, `query_odoo_sales_orders`, `get_pipeline_summary`, `check_mike_email`, `check_mike_calendar`.

Write tools (confirmation-gated): `log_odoo_note`, `create_odoo_lead`, `update_odoo_lead`, `send_mike_email`, `create_mike_calendar_event`.

## Confirmation gates

Persona enforces gates: for `send_mike_email`, `create_mike_calendar_event`, `create_odoo_lead`, `update_odoo_lead`, `log_odoo_note`, Don MUST draft the action, send it to Mike for review in Telegram, and wait for explicit confirmation before calling the tool.

## Env vars

See `.env.example`. All secrets live in Railway, never in this repo.

## Local dev

```bash
npm install
cp .env.example .env
# fill .env with real values
npm start
```

## Deploy

Connect this repo to Railway service `don-cos / don-telegram-bot`. Auto-deploys on push to `main`. Register Telegram webhook with `GET /set-webhook` after first deploy.

## Endpoints

- `GET /` — health check + uptime
- `POST /webhook` — Telegram webhook
- `GET /set-webhook` — re-register Telegram webhook against current Railway domain
- `GET /webhook-info` — Telegram webhook status
- `GET /debug` — Odoo + Microsoft Graph auth check

## Related

- Sam (reference pattern): https://github.com/TRCMalta/sam-telegram-bot
- Programme view: vault note `01-Projects/CoS-Rollout-Roadmap.md`
- Project file: vault note `01-Projects/Don-CoS.md`
- Universal playbook: vault note `01-Projects/AI-Chief-of-Staff.md`
