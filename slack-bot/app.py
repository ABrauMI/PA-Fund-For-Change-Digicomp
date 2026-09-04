import logging
import os
import tempfile
import threading
from datetime import date

import requests
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

from recalc import recalculate
from report_builder import build_workbook

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("comp-report-bot")

app = App(token=os.environ["SLACK_BOT_TOKEN"], request_verification_enabled=False)

# Comma-separated Spend Platform values to leave out of every report this bot
# builds, e.g. "CTV" or "CTV,Google". Set in Railway's Variables tab — no
# redeploy needed beyond restarting the service after changing it.
EXCLUDE_PLATFORMS = [p.strip() for p in os.environ.get("EXCLUDE_PLATFORMS", "").split(",") if p.strip()]

# Slack retries event delivery at least once; skip a file we've already started.
_seen_file_ids = set()


@app.event("message")
def handle_message_with_file(body, event, client, ack):
    ack()

    files = event.get("files") or []
    csv_files = [
        f for f in files
        if f.get("filetype") == "csv" or f.get("name", "").lower().endswith(".csv")
    ]
    if not csv_files:
        return

    channel = event["channel"]
    thread_ts = event.get("thread_ts", event.get("ts"))

    for file_info in csv_files:
        if file_info["id"] in _seen_file_ids:
            continue
        _seen_file_ids.add(file_info["id"])
        threading.Thread(
            target=_build_and_reply,
            args=(client, channel, thread_ts, file_info),
            daemon=True,
        ).start()


def _build_and_reply(client, channel, thread_ts, file_info):
    in_path = None
    out_path = None
    try:
        client.chat_postMessage(
            channel=channel,
            thread_ts=thread_ts,
            text=f"Got it — building the competitive report from `{file_info['name']}`...",
        )

        resp = requests.get(
            file_info["url_private"],
            headers={"Authorization": f"Bearer {os.environ['SLACK_BOT_TOKEN']}"},
            timeout=30,
        )
        resp.raise_for_status()

        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as in_f:
            in_f.write(resp.content)
            in_path = in_f.name
        out_path = in_path[:-4] + ".xlsx"

        result = build_workbook(
            in_path, out_path,
            reference_date=date.today(),
            exclude_platforms=EXCLUDE_PLATFORMS or None,
        )

        try:
            recalculate(out_path)
        except Exception:
            logger.warning("Recalc failed; delivering file with formulas uncached", exc_info=True)

        client.files_upload_v2(
            channel=channel,
            thread_ts=thread_ts,
            file=out_path,
            filename="Digital_Competitive_Report.xlsx",
            title="Digital Competitive Report",
            initial_comment=f"Done — {len(result['races'])} race tabs built.",
        )
    except Exception as e:
        logger.exception("Failed to build report")
        client.chat_postMessage(
            channel=channel,
            thread_ts=thread_ts,
            text=f"Couldn't build the report from `{file_info['name']}`: {e}",
        )
    finally:
        for p in (in_path, out_path):
            if p and os.path.exists(p):
                os.remove(p)


if __name__ == "__main__":
    SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"]).start()
