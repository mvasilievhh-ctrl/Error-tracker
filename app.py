import os
import re
import sqlite3
import threading
import time
from datetime import datetime

import requests
from flask import Flask, jsonify, request, send_from_directory

# ---------------------------------------------------------------------------
# Config (set these as environment variables on Render/Railway)
# ---------------------------------------------------------------------------
BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")  # optional: restrict to one group
DB_PATH = os.environ.get("DB_PATH", "tracker.db")
POLL_INTERVAL_SECONDS = 3

API_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"

app = Flask(__name__, static_folder="static", template_folder="templates")

# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            time TEXT NOT NULL,
            b2c TEXT,
            gts TEXT,
            provider TEXT,
            error TEXT,
            file_name TEXT,
            file_content TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def get_last_update_id():
    conn = get_db()
    row = conn.execute("SELECT value FROM meta WHERE key = 'last_update_id'").fetchone()
    conn.close()
    return int(row["value"]) if row else 0


def set_last_update_id(update_id):
    conn = get_db()
    conn.execute(
        "INSERT INTO meta (key, value) VALUES ('last_update_id', ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (str(update_id),),
    )
    conn.commit()
    conn.close()


def save_entry(entry):
    conn = get_db()
    conn.execute(
        """
        INSERT INTO entries (time, b2c, gts, provider, error, file_name, file_content)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            entry["time"],
            entry["b2c"],
            entry["gts"],
            entry["provider"],
            entry["error"],
            entry.get("file_name"),
            entry.get("file_content"),
        ),
    )
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Parsing — matches the format:
#   Ошибка выписки
#   Номер заказа B2C: 7385
#   Номер заказа GTS: 88602
#   Провайдер: Тестовый
#   Сообщение ошибки: ...
# ---------------------------------------------------------------------------

def parse_message(text):
    if not text:
        return None

    b2c = re.search(r"Номер заказа B2C:\s*(\S+)", text, re.IGNORECASE)
    gts = re.search(r"Номер заказа GTS:\s*(\S+)", text, re.IGNORECASE)
    provider = re.search(r"Провайдер:\s*(.+)", text, re.IGNORECASE)
    error_msg = re.search(r"Сообщение ошибки:\s*([\s\S]+)", text, re.IGNORECASE)

    if not any([b2c, gts, provider, error_msg]):
        return None

    return {
        "time": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "b2c": b2c.group(1).strip() if b2c else "—",
        "gts": gts.group(1).strip() if gts else "—",
        "provider": provider.group(1).strip() if provider else "—",
        "error": error_msg.group(1).strip() if error_msg else "—",
    }


# ---------------------------------------------------------------------------
# Telegram polling (long polling via getUpdates, runs in a background thread)
# ---------------------------------------------------------------------------

def download_file_content(file_id):
    """Downloads a Telegram file and returns its text content, or None if not text."""
    try:
        r = requests.get(f"{API_URL}/getFile", params={"file_id": file_id}, timeout=15)
        file_path = r.json()["result"]["file_path"]
        file_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
        content_resp = requests.get(file_url, timeout=15)
        return content_resp.content.decode("utf-8", errors="replace"), file_path.split("/")[-1]
    except Exception as e:
        print(f"[file download error] {e}")
        return None, None


def process_update(update):
    message = update.get("channel_post") or update.get("message")
    if not message:
        return

    chat = message.get("chat", {})
    if CHAT_ID and str(chat.get("id")) != str(CHAT_ID):
        return  # ignore messages from other chats if CHAT_ID is set

    text = message.get("text") or message.get("caption") or ""
    parsed = parse_message(text)
    if not parsed:
        return

    # attach file if present (document only, text files)
    document = message.get("document")
    if document:
        content, fname = download_file_content(document["file_id"])
        if content is not None:
            parsed["file_name"] = fname or document.get("file_name")
            parsed["file_content"] = content

    save_entry(parsed)
    print(f"[saved] B2C={parsed['b2c']} GTS={parsed['gts']}")


def poll_loop():
    if not BOT_TOKEN:
        print("[bot] TELEGRAM_BOT_TOKEN not set — polling disabled.")
        return

    print("[bot] Starting polling loop...")
    offset = get_last_update_id()

    while True:
        try:
            resp = requests.get(
                f"{API_URL}/getUpdates",
                params={"offset": offset + 1, "timeout": 30},
                timeout=40,
            )
            data = resp.json()
            if not data.get("ok"):
                print(f"[bot] getUpdates error: {data}")
                time.sleep(POLL_INTERVAL_SECONDS)
                continue

            for update in data.get("result", []):
                offset = update["update_id"]
                process_update(update)

            set_last_update_id(offset)

        except Exception as e:
            print(f"[bot] polling error: {e}")
            time.sleep(POLL_INTERVAL_SECONDS)


# ---------------------------------------------------------------------------
# Web routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return send_from_directory("templates", "index.html")


@app.route("/api/entries", methods=["GET"])
def api_get_entries():
    conn = get_db()
    rows = conn.execute("SELECT * FROM entries ORDER BY id DESC").fetchall()
    conn.close()
    return jsonify([dict(row) for row in rows])


@app.route("/api/entries", methods=["POST"])
def api_add_entry():
    """Allows manual additions from the web page too (same as before)."""
    data = request.get_json(force=True)
    entry = {
        "time": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "b2c": data.get("b2c") or "—",
        "gts": data.get("gts") or "—",
        "provider": data.get("provider") or "—",
        "error": data.get("error") or "—",
        "file_name": data.get("file_name"),
        "file_content": data.get("file_content"),
    }
    save_entry(entry)
    return jsonify({"ok": True})


@app.route("/api/entries/<int:entry_id>", methods=["DELETE"])
def api_delete_entry(entry_id):
    conn = get_db()
    conn.execute("DELETE FROM entries WHERE id = ?", (entry_id,))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


# ---------------------------------------------------------------------------
# Startup
# ---------------------------------------------------------------------------

init_db()

if BOT_TOKEN:
    threading.Thread(target=poll_loop, daemon=True).start()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
