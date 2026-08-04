"""
BMAP Snapshot API — Railway deployment
=======================================
Flask wrapper around bmap_snapshot.py.
Receives inst_key from context-generator.html,
builds the deck, returns the .pptx as a download.

Endpoints:
  POST /generate            { inst_key, bank_name? }  → .pptx file
  POST /generate-batch      { banks: [{inst_key, name}] } → .zip file
  POST /generate-brief      { inst_key, bank_name? }  → .pdf file
  POST /generate-assessment { inst_key, bank_name? }  → .docx file (full-network $10K Assessment)
  GET  /health               → { status: ok }

  -- added by secure_proxy.py (Hub auth + data proxy) --
  POST /auth/login           { password }             → { token }
  GET  /api/<table>          Bearer token required     → Supabase rows
  POST /api/ai/briefing-note Bearer token required     → AI narrative JSON
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
import bmap_assessment_doc as bad
from secure_proxy import secure_proxy_bp

app = Flask(__name__)

# Locked to the Hub's actual origin — previously CORS(app) allowed
# requests from any website, not just context-generator on GitHub Pages.
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "https://fggraufaro.github.io")
CORS(app, origins=[ALLOWED_ORIGIN])

app.register_blueprint(secure_proxy_bp)


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


# ── $10K Assessment Word doc — full network, no top-N slice ────
@app.route("/generate-assessment", methods=["POST"])
def generate_assessment():
    body      = request.get_json(force=True)
    ik        = (body.get("inst_key") or "").strip()
    name_hint = (body.get("bank_name") or "").strip()

    if not ik:
        return jsonify({"error": "inst_key required"}), 400

    try:
        print(f"[assessment] {ik} — {name_hint or 'no name hint'}")

        d = bad.fetch_full_network_data(ik)
        if not d["branches"]:
            print(f"[assessment] ✗ {ik} — no branch data found, refusing to build empty report")
            return jsonify({"error": f"No branch data found for inst_key='{ik}'. "
                                      f"This bank may not be ingested into BMAP yet, "
                                      f"or the inst_key is incorrect."}), 422
        bank_name = name_hint or (d["branches"][0].get("namefull") if d["branches"] else None) or ik
        summary = bad.summarize_network(d)
        dives, deep_mode = bad.build_branch_deep_dives(d["branches"], d.get("branch_strategy") or [])
        narr = bad.get_narratives(bank_name, summary, d["fin"], d["targets"], d.get("branch_strategy"), dives)
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            doc = bad.build_assessment_doc(bank_name, summary, d["fin"], d["targets"], narr,
                                            d["branches"], d.get("branches_geo"),
                                            d.get("branch_strategy"), dives, deep_mode, tmpdir=tmpdir)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)

        safe = "".join(c if c.isalnum() or c in " _-" else "_"
                       for c in bank_name).strip()
        date = datetime.now().strftime("%Y%m%d")
        filename = f"BMAP_Assessment_{safe}_{date}.docx"

        print(f"[assessment] ✓ {filename} ({buf.getbuffer().nbytes // 1024}KB, "
              f"{len(d['branches'])} branches)")

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"[assessment] ✗ {e}")
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
        print(f"[generate] {ik} — {name_hint or 'no name hint'}")

        data = bm.fetch_bank_data(ik)
        if name_hint:
            data["bankName"] = name_hint

        # Fetch or generate personas (checks DB first, generates with Claude if not found)
        personas = bm.fetch_or_generate_personas(
            ik, data["bankName"], data.get("br", []), data)
        data["personas"] = personas

        logo  = bm.fetch_logo()
        prs   = bm.build_deck(data, logo)

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

                personas = bm.fetch_or_generate_personas(
                    ik, data["bankName"], data.get("br", []), data)
                data["personas"] = personas

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
