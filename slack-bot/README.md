# Digital Competitive Report — Slack Bot

Drop an AdHawk CSV export into a Slack channel or DM, and the bot replies
in-thread with a finished multi-tab competitive report workbook (one tab
per race, styled like the GPS Impact template, plus a Summary index).

## How it works

- Slack sends a `message` event to this app whenever a file is posted in a
  channel/DM the bot is a member of.
- The app downloads any `.csv` attachment, builds the workbook
  (`report_builder.py`), recalculates formulas with a headless LibreOffice
  so Slack's preview shows real numbers, and posts the `.xlsx` back in the
  same thread.
- Nothing is stored — the temp CSV/XLSX are deleted after each reply.

## 1. Create the Slack app

1. Go to <https://api.slack.com/apps> → **Create New App** → **From scratch**.
2. Name it (e.g. "Comp Report Bot") and pick your workspace.

### Bot Token Scopes

**OAuth & Permissions** → **Scopes** → **Bot Token Scopes**, add:

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

**Event Subscriptions** → toggle on. You'll set the **Request URL** in
step 3, after Railway gives you a domain — Slack won't let you save a URL
it can't verify yet, so leave this tab open and come back to it.

Under **Subscribe to bot events**, add:

- `message.channels`
- `message.im`

(add `message.groups` / `message.mpim` if you added those scopes above)

### Install the app

**OAuth & Permissions** → **Install to Workspace** → Allow. Copy the
**Bot User OAuth Token** (starts `xoxb-`).

Then **Basic Information** → **App Credentials** → copy the **Signing
Secret**.

## 2. Deploy to Railway

1. Push this repo to GitHub (already done if you're reading this from the repo).
2. In Railway: **New Project** → **Deploy from GitHub repo** → pick this repo.
3. Open the new service's **Settings**:
   - **Root Directory**: `slack-bot` (this repo has other stuff at the top
     level, so tell Railway to build only this folder).
   - Railway will detect the `Dockerfile` in `slack-bot/` automatically —
     no other build config needed.
4. **Variables** tab → add:
   - `SLACK_BOT_TOKEN` = the `xoxb-...` token from step 1
   - `SLACK_SIGNING_SECRET` = the signing secret from step 1
5. Deploy. Once it's live, open **Settings** → **Networking** → **Generate
   Domain** to get a public URL like `https://your-app.up.railway.app`.

## 3. Finish the Slack app config

Back in api.slack.com → **Event Subscriptions** → **Request URL**, enter:

```
https://your-app.up.railway.app/slack/events
```

Slack will POST a verification challenge; the app answers it automatically
(you should see a green "Verified" ✅). Save.

## 4. Try it

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
  tab shows the Flask/Bolt logs, including the full traceback if a build
  fails on a malformed CSV.
