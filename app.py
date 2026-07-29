"""
NCD Scrub Dashboard — Backend API
Storage: CSV file in the GitHub repo (data/entries.csv)
Frontend: Static HTML/React served from /static/
"""

import os
import csv
import io
import base64
import json
from datetime import datetime, date
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import requests as http_requests

app = Flask(__name__, static_folder="static", static_url_path="")
CORS(app)

# ── Config ──
API_KEY = os.environ.get("SCRUB_API_KEY", "")
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO = os.environ.get("GITHUB_REPO", "arojasb12/ncd-scrub-dashboard")
CSV_PATH = "data/entries.csv"
LOCAL_CSV = os.path.join(os.path.dirname(__file__), "data", "entries.csv")

# CSV column headers
HEADERS = ["id", "section", "category", "date", "value", "amount", "source"]


# ── In-memory data store (loaded from CSV on startup) ──
entries = []
next_id = 1


def load_csv():
    """Load entries from the local CSV file into memory."""
    global entries, next_id
    entries = []
    next_id = 1

    if not os.path.exists(LOCAL_CSV):
        return

    with open(LOCAL_CSV, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            entry = {
                "id": int(row["id"]),
                "section": row["section"],
                "category": row["category"],
                "date": row["date"],
                "value": float(row["value"]) if row.get("value") else None,
                "amount": float(row["amount"]) if row.get("amount") else None,
                "source": row.get("source", ""),
            }
            entries.append(entry)
            if entry["id"] >= next_id:
                next_id = entry["id"] + 1


def save_csv():
    """Write current entries to the local CSV file."""
    os.makedirs(os.path.dirname(LOCAL_CSV), exist_ok=True)
    with open(LOCAL_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=HEADERS)
        writer.writeheader()
        for entry in entries:
            writer.writerow({
                "id": entry["id"],
                "section": entry["section"],
                "category": entry["category"],
                "date": entry["date"],
                "value": entry["value"] if entry["value"] is not None else "",
                "amount": entry["amount"] if entry["amount"] is not None else "",
                "source": entry.get("source", ""),
            })


def push_to_github():
    """
    Push the updated CSV back to the GitHub repo so it persists
    across Heroku dyno restarts. Uses the GitHub Contents API.
    """
    if not GITHUB_TOKEN:
        print("No GITHUB_TOKEN set, skipping push")
        return False

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{CSV_PATH}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }

    # Get the current file SHA (needed for updates)
    sha = None
    resp = http_requests.get(url, headers=headers)
    if resp.status_code == 200:
        sha = resp.json().get("sha")

    # Read the local CSV content
    with open(LOCAL_CSV, "r", encoding="utf-8") as f:
        content = f.read()

    # Encode content as base64 (required by GitHub API)
    encoded = base64.b64encode(content.encode("utf-8")).decode("utf-8")

    # Build the request
    payload = {
        "message": f"Update entries.csv ({len(entries)} entries)",
        "content": encoded,
    }
    if sha:
        payload["sha"] = sha

    resp = http_requests.put(url, headers=headers, json=payload)
    if resp.status_code in (200, 201):
        print(f"Pushed CSV to GitHub ({len(entries)} entries)")
        return True
    else:
        print(f"GitHub push failed: {resp.status_code} {resp.text[:200]}")
        return False


def check_api_key():
    """Validate the API key from the request header."""
    if not API_KEY:
        return True  # No key configured, allow all
    key = request.headers.get("x-api-key", "")
    return key == API_KEY


def is_duplicate(section, category, entry_date):
    """Check if an entry already exists for this section+category+date."""
    for e in entries:
        if e["section"] == section and e["category"] == category and e["date"] == entry_date:
            return True
    return False


# ── Load data on startup ──
load_csv()
print(f"Loaded {len(entries)} entries from CSV")


# ── API Routes ──

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "entries": len(entries)})


@app.route("/api/entries", methods=["GET"])
def get_entries():
    """Return all entries, optionally filtered by section."""
    if not check_api_key():
        return jsonify({"error": "Invalid API key"}), 401

    section = request.args.get("section")
    if section:
        filtered = [e for e in entries if e["section"] == section]
    else:
        filtered = entries

    return jsonify(filtered)


@app.route("/api/entries", methods=["POST"])
def add_entry():
    """Add a single entry manually."""
    if not check_api_key():
        return jsonify({"error": "Invalid API key"}), 401

    global next_id
    data = request.get_json()

    if not data:
        return jsonify({"error": "JSON body required"}), 400

    section = data.get("section", "")
    category = data.get("category", "")
    entry_date = data.get("scrub_date") or data.get("date") or data.get("d")
    value = data.get("value") or data.get("v")
    amount = data.get("amount") or data.get("a")
    source = data.get("source", "manual")

    if not section or not category or not entry_date:
        return jsonify({"error": "section, category, and date are required"}), 400

    if is_duplicate(section, category, entry_date):
        return jsonify({"status": "skipped", "reason": "duplicate"}), 200

    entry = {
        "id": next_id,
        "section": section,
        "category": category,
        "date": entry_date,
        "value": float(value) if value is not None else None,
        "amount": float(amount) if amount is not None else None,
        "source": source,
    }
    entries.append(entry)
    next_id += 1

    save_csv()
    push_to_github()

    return jsonify({"status": "added", "id": entry["id"]})


@app.route("/api/entries/bulk", methods=["POST"])
def bulk_add():
    """Add multiple entries at once. Expects a JSON array."""
    if not check_api_key():
        return jsonify({"error": "Invalid API key"}), 401

    global next_id
    data = request.get_json()

    if not isinstance(data, list):
        return jsonify({"error": "Expected a JSON array"}), 400

    added = 0
    skipped = 0

    for row in data:
        section = row.get("section") or row.get("s", "")
        category = row.get("category") or row.get("c", "")
        entry_date = row.get("scrub_date") or row.get("date") or row.get("d")
        value = row.get("value") or row.get("v")
        amount = row.get("amount") or row.get("a")
        source = row.get("source", "seed")

        if not section or not category or not entry_date:
            continue

        if is_duplicate(section, category, entry_date):
            skipped += 1
            continue

        entry = {
            "id": next_id,
            "section": section,
            "category": category,
            "date": entry_date,
            "value": float(value) if value is not None else None,
            "amount": float(amount) if amount is not None else None,
            "source": source,
        }
        entries.append(entry)
        next_id += 1
        added += 1

    if added > 0:
        save_csv()
        push_to_github()

    return jsonify({"added": added, "skipped": skipped})


@app.route("/api/entries/<int:entry_id>", methods=["DELETE"])
def delete_entry(entry_id):
    """Delete an entry by ID."""
    if not check_api_key():
        return jsonify({"error": "Invalid API key"}), 401

    global entries
    before = len(entries)
    entries = [e for e in entries if e["id"] != entry_id]

    if len(entries) == before:
        return jsonify({"error": "Entry not found"}), 404

    save_csv()
    push_to_github()

    return jsonify({"status": "deleted"})


# ── Serve the React frontend ──

@app.route("/")
def serve_frontend():
    return send_from_directory("static", "index.html")


@app.route("/<path:path>")
def serve_static(path):
    if os.path.exists(os.path.join("static", path)):
        return send_from_directory("static", path)
    return send_from_directory("static", "index.html")


# ── Run ──

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
