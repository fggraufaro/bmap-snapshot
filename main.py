"""
BMAP Snapshot API — Railway deployment
=======================================
Flask wrapper around bmap_snapshot.py.
Receives inst_key from context-generator.html,
builds the deck, returns the .pptx as a download.

Endpoints:
  POST /generate        { inst_key, bank_name? }  → .pptx file
  POST /generate-batch  { banks: [{inst_key, name}] } → .zip file
  GET  /health          → { status: ok }
"""

import io
import json
import os
import zipfile
from datetime import datetime

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

import bmap_snapshot as bm
import bmap_board_brief as bb

app = Flask(__name__)
CORS(app)  # Allow calls from context-generator on GitHub Pages

# ── Health check — Railway uses this to confirm the app is up ──
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "BMAP Snapshot API"})


# ── Board Brief PDF ────────────────────────────────────────────
@app.route("/generate-brief", methods=["POST"])
def generate_brief():
    body      = request.get_json(force=True)
    ik        = (body.get("inst_key") or "").strip()
    name_hint = (body.get("bank_name") or "").strip()

    if not ik:
        return jsonify({"error": "inst_key required"}), 400

    try:
        print(f"[brief] {ik} — {name_hint or 'no name hint'}")
        buf = bb.generate_board_brief(ik, name_hint or None)

        safe = "".join(c if c.isalnum() or c in " _-" else "_"
                       for c in (name_hint or ik)).strip()
        date = datetime.now().strftime("%Y%m%d")
        filename = f"Board_Brief_{safe}_{date}.pdf"

        print(f"[brief] ✓ {filename} ({buf.getbuffer().nbytes // 1024}KB)")

        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"[brief] ✗ {e}")
        return jsonify({"error": str(e)}), 500


# ── Single deck ────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
def generate():
    body = request.get_json(force=True)
    ik        = (body.get("inst_key") or "").strip()
    name_hint = (body.get("bank_name") or "").strip()

    if not ik:
        return jsonify({"error": "inst_key required"}), 400

    try:
        no_ai = body.get("no_ai", False)
        print(f"[generate] {ik} — {name_hint or 'no name hint'} — no_ai={no_ai}")

        data = bm.fetch_bank_data(ik)
        if name_hint:
            data["bankName"] = name_hint

        # Temporarily skip AI to stay within Railway's 60s HTTP timeout
        if no_ai:
            import os as _os
            _os.environ["_SKIP_AI"] = "1"
        
        logo  = bm.fetch_logo()
        prs   = bm.build_deck(data, logo)
        
        if no_ai:
            import os as _os
            _os.environ.pop("_SKIP_AI", None)

        # Save to in-memory buffer — no disk writes needed
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        safe = "".join(c if c.isalnum() or c in " _-" else "_"
                       for c in data["bankName"]).strip()
        date = datetime.now().strftime("%Y%m%d")
        filename = f"BMAP_Snapshot_{safe}_{date}.pptx"

        print(f"[generate] ✓ {filename} ({buf.getbuffer().nbytes // 1024}KB)")

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"[generate] ✗ {e}")
        return jsonify({"error": str(e)}), 500


# ── Batch decks → ZIP ──────────────────────────────────────────
@app.route("/generate-batch", methods=["POST"])
def generate_batch():
    body  = request.get_json(force=True)
    banks = body.get("banks", [])

    if not banks:
        return jsonify({"error": "banks array required"}), 400

    print(f"[batch] {len(banks)} banks")

    logo    = bm.fetch_logo()
    zip_buf = io.BytesIO()
    results = []

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in banks:
            ik   = (item.get("inst_key") or "").strip()
            name = (item.get("name") or "").strip()
            if not ik:
                continue
            try:
                data = bm.fetch_bank_data(ik)
                if name:
                    data["bankName"] = name

                prs = bm.build_deck(data, logo)

                deck_buf = io.BytesIO()
                prs.save(deck_buf)
                deck_buf.seek(0)

                safe = "".join(c if c.isalnum() or c in " _-" else "_"
                               for c in data["bankName"]).strip()
                fname = f"BMAP_Snapshot_{safe}.pptx"
                zf.writestr(fname, deck_buf.read())

                print(f"[batch] ✓ {fname}")
                results.append({"bank": data["bankName"], "status": "ok", "file": fname})

            except Exception as e:
                print(f"[batch] ✗ {ik}: {e}")
                results.append({"bank": name or ik, "status": "error", "error": str(e)})

    zip_buf.seek(0)
    date     = datetime.now().strftime("%Y%m%d")
    ok_count = sum(1 for r in results if r["status"] == "ok")
    print(f"[batch] complete — {ok_count}/{len(banks)} ok")

    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"BMAP_Batch_{date}.zip"
    )


# ── Entry point ────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    print(f"BMAP Snapshot API starting on port {port}")
    app.run(host="0.0.0.0", port=port)
