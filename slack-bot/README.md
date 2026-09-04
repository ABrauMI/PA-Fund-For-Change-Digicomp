# Digital Competitive Report — Slack Bot

Drop an AdHawk CSV export into a Slack channel or DM, and the bot replies
in-thread with a finished multi-tab competitive report workbook (one tab
per race, styled like the GPS Impact template, plus a Summary index).

Runs in **Socket Mode** — the app opens an outbound WebSocket to Slack, so
there's no public URL, no Request URL verification, and no inbound
networking to configure on Railway.

## How it works

- The app maintains a persistent Socket Mode connection to Slack and
  receives a `message` event whenever a file is posted in a channel/DM
  it's a member of.
- It downloads any `.csv` attachment, builds the workbook
  (`report_builder.py`), recalculates formulas with a headless LibreOffice
  so Slack's preview shows real numbers, and posts the `.xlsx` back in the
  same thread.
- Nothing is stored — the temp CSV/XLSX are deleted after each reply.

## 1. Create the Slack app

1. Go to <https://api.slack.com/apps> → **Create New App** → **From scratch**.
2. Name it (e.g. "Comp Report Bot") and pick your workspace.

### Turn on Socket Mode

**Settings → Socket Mode** → toggle it on. Slack will prompt you to create
an **App-Level Token** — name it anything, grant it the `connections:write`
scope, and generate it. Copy the token (starts `xapp-`).

### Bot Token Scopes

**Features → OAuth & Permissions** → **Bot Token Scopes**, add:

| Scope | Why |
|---|---|
| `chat:write` | post status messages and the finished report |
| `files:read` | download the CSV someone uploads |
| `files:write` | upload the generated `.xlsx` back |
| `channels:history` | read messages (with attachments) in public channels |
| `im:history` | read DMs sent to the bot |

Add `groups:history` / `mpim:history` too if you want it to work in
private channels / group DMs.

### Event Subscriptions

**Features → Event Subscriptions** → toggle on. Because Socket Mode is
already enabled, there's no Request URL field — events are delivered over
the socket instead.

Under **Subscribe to bot events**, add:

- `message.channels`
- `message.im`

(add `message.groups` / `message.mpim` if you added those scopes above)

### Install the app

**OAuth & Permissions** → **Install to Workspace** → Allow. Copy the
**Bot User OAuth Token** (starts `xoxb-`).

You now have two tokens: the `xoxb-...` bot token and the `xapp-...`
app-level token from the Socket Mode step.

## 2. Deploy to Railway

1. Push this repo to GitHub (already done if you're reading this from the repo).
2. In Railway: **New Project** → **Deploy from GitHub repo** → pick this repo.
3. Open the new service's **Settings** → set **Root Directory** to
   `slack-bot` (this repo has other stuff at the top level, so tell
   Railway to build only this folder). Railway detects the `Dockerfile`
   there automatically.
4. **Variables** tab → add:
   - `SLACK_BOT_TOKEN` = the `xoxb-...` token
   - `SLACK_APP_TOKEN` = the `xapp-...` token
   - `EXCLUDE_PLATFORMS` (optional) = comma-separated Spend Platform values
     to leave out of every report, e.g. `CTV` (or `CTV,Google` for more
     than one). Leave unset to include everything. The report's footer
     notes which platforms were excluded, so it's never silent about it.
5. Deploy. That's it — **no public domain / networking setup needed**;
   this service only makes outbound connections, so you can skip
   Railway's "Generate Domain" step entirely.

## 3. Try it

1. In Slack, invite the bot to a channel (`/invite @Comp Report Bot`) or
   just open a DM with it.
2. Drop an AdHawk CSV export into that channel/DM.
3. The bot replies "Got it — building..." and, a few seconds later, the
   finished `.xlsx`.

## Notes

- **Cost**: Railway's Hobby plan ($5/mo credit) easily covers a low-traffic
  bot like this.
- **Multiple races/years**: the builder infers each week's year from
  today's date and re-sorts races naturally, so it keeps working as new
  exports arrive — it doesn't hardcode PA's 2026 races.
- **Data-quality spot checks are still on you.** The builder includes
  every race code in the file as-is; it won't catch a mistagged row (like
  the stray Baltimore County row filtered out of the one-off PA report by
  hand). Skim a new report's Summary tab before sending it anywhere.
- **Redeploying**: push to the branch Railway is tracking and it redeploys
  automatically.
- **Logs / troubleshooting**: Railway's **Deployments** → **View Logs**
  tab shows the app's logs, including the full traceback if a build fails
  on a malformed CSV.
