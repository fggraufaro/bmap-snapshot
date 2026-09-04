"""
BMAP Snapshot API — Railway deployment
=======================================
Flask wrapper around bmap_snapshot.py.
Receives inst_key from context-generator.html,
builds the deck, returns the .pptx as a download.

Endpoints:
  POST /generate                   { inst_key, bank_name? }  → .pptx file
  POST /generate-batch             { banks: [{inst_key, name}] } → .zip file
  POST /generate-brief             { inst_key, bank_name? }  → .pdf file
  POST /generate-assessment        { inst_key, bank_name? }  → .docx file (synchronous — small networks only)
  POST /generate-assessment-async  { inst_key, bank_name? }  → { job_id } (returns instantly, runs in background)
  GET  /assessment-status/<job_id>                            → { status, stage, error_message }
  GET  /assessment-download/<job_id>                          → .docx file once status == "done"
  GET  /health               → { status: ok }

  -- added by secure_proxy.py (Hub auth + data proxy) --
  POST /auth/login           { password }             → { token }
  GET  /api/<table>          Bearer token required     → Supabase rows
  POST /api/ai/briefing-note Bearer token required     → AI narrative JSON
"""

import base64
import io
import json
import os
import threading
import concurrent.futures
import zipfile
from datetime import datetime

import requests

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

import bmap_snapshot as bm
import bmap_board_brief as bb
import bmap_assessment_doc as bad
import bmap_branch_preview as bpv
from secure_proxy import secure_proxy_bp, require_session

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
@require_session
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
@require_session
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
        narr = bad.get_narratives(bank_name, summary, d["fin"], d["targets"], d.get("branch_strategy"), dives,
                                   d.get("capped_yoy"))
        persona_brief, market_offer_brief = None, None
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
            fut_persona = pool.submit(bad.get_persona_signal_brief, bank_name, dives)
            fut_market = pool.submit(bad.get_market_offer_brief, bank_name, dives, d.get("branch_strategy"))
            try:
                persona_brief = fut_persona.result(timeout=90)
            except Exception as ex:
                print(f"[generate-assessment] persona brief failed/timed out: {type(ex).__name__}: {str(ex) if str(ex) else '(no message -- likely a timeout)'}")
            try:
                market_offer_brief = fut_market.result(timeout=90)
            except Exception as ex:
                print(f"[generate-assessment] market offer brief failed/timed out: {type(ex).__name__}: {str(ex) if str(ex) else '(no message -- likely a timeout)'}")
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            doc = bad.build_assessment_doc(bank_name, summary, d["fin"], d["targets"], narr,
                                            d["branches"], d.get("branches_geo"),
                                            d.get("branch_strategy"), dives, deep_mode, tmpdir=tmpdir,
                                            capped_yoy=d.get("capped_yoy"),
                                            persona_brief=persona_brief, market_offer_brief=market_offer_brief,
                                            vulnerability_targets=d.get("vulnerability_targets"))
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


# ── Branch Preview — Hub "Quick Export" live pitch-meeting teaser ──
# One branch, same rendering code as the paid Assessment (render_branch_deep_dive),
# scoped down enough to stay fast for live use in a sales meeting -- no async
# job/polling pattern needed here, unlike /generate-assessment, since this is
# deliberately fast (one branch's worth of AI + Mapbox calls, not the whole network).
@app.route("/generate-branch-preview", methods=["POST"])
@require_session
def generate_branch_preview():
    body      = request.get_json(force=True)
    ik        = (body.get("inst_key") or "").strip()
    name_hint = (body.get("bank_name") or "").strip()
    branch    = (body.get("branch_name") or "").strip() or None

    if not ik:
        return jsonify({"error": "inst_key required"}), 400

    try:
        print(f"[branch-preview] {ik} — {name_hint or 'no name hint'} — "
              f"branch={branch or '(auto-select)'}")

        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            doc, bank_name, branch_label, total_branches = bpv.generate_preview(
                ik, name_hint, branch, tmpdir
            )
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)

        safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in bank_name).strip()
        branch_safe = "".join(c if c.isalnum() or c in " _-" else "_"
                               for c in branch_label.split(" (")[0]).strip()
        date = datetime.now().strftime("%Y%m%d")
        filename = f"BMAP_Preview_{safe}_{branch_safe}_{date}.docx"

        print(f"[branch-preview] ✓ {filename} ({buf.getbuffer().nbytes // 1024}KB, "
              f"{branch_label} of {total_branches} branches)")

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename
        )

    except ValueError as e:
        # Bad/unmatched inst_key or branch name -- a client error, not a
        # server failure, per the same distinction /generate-assessment
        # makes with its 422 for "no branch data found".
        print(f"[branch-preview] ✗ {e}")
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        print(f"[branch-preview] ✗ {e}")
        return jsonify({"error": str(e)}), 500


# ── Async Assessment job — background generation + status polling ──
# The synchronous /generate-assessment route above works for small networks,
# but 25-30 page documents with 15 branch deep-dives, chart generation, and
# an AI narrative call routinely exceed Gunicorn's worker timeout, which
# kills the connection mid-request (ERR_CONNECTION_CLOSED client-side).
# This runs the same pipeline in a background thread instead: the initial
# request returns in milliseconds, and the client polls for completion —
# so the HTTP timeout is no longer a constraint regardless of doc size.

def _job_write(job_id, **fields):
    fields["updated_at"] = datetime.now().isoformat()
    url = f"{bad.SUPA_URL}/rest/v1/assessment_jobs?id=eq.{job_id}"
    headers = {"apikey": bad.SUPA_KEY, "Authorization": f"Bearer {bad.SUPA_KEY}",
               "Content-Type": "application/json", "Prefer": "return=minimal"}
    try:
        requests.patch(url, headers=headers, json=fields, timeout=15)
    except Exception as e:
        print(f"[assessment-job] ⚠ failed to write job status: {e}")


def _job_create(ik, bank_name_hint):
    url = f"{bad.SUPA_URL}/rest/v1/assessment_jobs"
    headers = {"apikey": bad.SUPA_KEY, "Authorization": f"Bearer {bad.SUPA_KEY}",
               "Content-Type": "application/json", "Prefer": "return=representation"}
    payload = {"inst_key": ik, "bank_name": bank_name_hint, "status": "pending"}
    r = requests.post(url, headers=headers, json=payload, timeout=15)
    r.raise_for_status()
    return r.json()[0]["id"]


def _job_read(job_id, select="status,stage,error_message,filename,bank_name"):
    url = f"{bad.SUPA_URL}/rest/v1/assessment_jobs?id=eq.{job_id}&select={select}"
    headers = {"apikey": bad.SUPA_KEY, "Authorization": f"Bearer {bad.SUPA_KEY}"}
    r = requests.get(url, headers=headers, timeout=15)
    r.raise_for_status()
    rows = r.json()
    return rows[0] if rows else None


def _run_assessment_job(job_id, ik, name_hint):
    try:
        _job_write(job_id, status="running", stage="Fetching full branch network...")
        d = bad.fetch_full_network_data(ik)
        if not d["branches"]:
            _job_write(job_id, status="error",
                       error_message=f"No branch data found for inst_key='{ik}'.")
            return

        bank_name = name_hint or (d["branches"][0].get("namefull") if d["branches"] else None) or ik
        _job_write(job_id, stage="Computing branch strategy and plays...", bank_name=bank_name)
        summary = bad.summarize_network(d)
        dives, deep_mode = bad.build_branch_deep_dives(d["branches"], d.get("branch_strategy") or [])

        _job_write(job_id, stage="Generating AI narratives...")
        narr = bad.get_narratives(bank_name, summary, d["fin"], d["targets"],
                                   d.get("branch_strategy"), dives, d.get("capped_yoy"))

        _job_write(job_id, stage="Researching persona, demographic & market signal (live web search)...")
        # Run concurrently, not sequentially -- these are independent calls
        # (different prompts, no shared state) and each can take 20-60+
        # seconds with multiple search rounds. Sequential execution was the
        # likely cause of jobs still running when a polling client expected
        # them done -- this alone can save 20-60+ seconds of real wall time.
        persona_brief, market_offer_brief = None, None
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
            fut_persona = pool.submit(bad.get_persona_signal_brief, bank_name, dives)
            fut_market = pool.submit(bad.get_market_offer_brief, bank_name, dives, d.get("branch_strategy"))
            try:
                persona_brief = fut_persona.result(timeout=90)
            except Exception as ex:
                print(f"[assessment-job] persona brief failed/timed out: {type(ex).__name__}: {str(ex) if str(ex) else '(no message -- likely a timeout)'}")
            try:
                market_offer_brief = fut_market.result(timeout=90)
            except Exception as ex:
                print(f"[assessment-job] market offer brief failed/timed out: {type(ex).__name__}: {str(ex) if str(ex) else '(no message -- likely a timeout)'}")

        _job_write(job_id, stage="Building document (charts, branch deep dives)...")
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            doc = bad.build_assessment_doc(bank_name, summary, d["fin"], d["targets"], narr,
                                            d["branches"], d.get("branches_geo"),
                                            d.get("branch_strategy"), dives, deep_mode, tmpdir=tmpdir,
                                            capped_yoy=d.get("capped_yoy"),
                                            persona_brief=persona_brief, market_offer_brief=market_offer_brief,
                                            vulnerability_targets=d.get("vulnerability_targets"))
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            docx_bytes = buf.getvalue()

        safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in bank_name).strip()
        date = datetime.now().strftime("%Y%m%d")
        filename = f"BMAP_Assessment_{safe}_{date}.docx"
        docx_b64 = base64.b64encode(docx_bytes).decode("ascii")

        _job_write(job_id, status="done", stage="Complete", filename=filename, docx_base64=docx_b64)
        print(f"[assessment-job] ✓ {job_id} — {filename} ({len(docx_bytes)//1024}KB, "
              f"{len(d['branches'])} branches)")

    except Exception as e:
        print(f"[assessment-job] ✗ {job_id}: {e}")
        _job_write(job_id, status="error", error_message=str(e))


@app.route("/generate-assessment-async", methods=["POST"])
@require_session
def generate_assessment_async():
    body = request.get_json(force=True)
    ik = (body.get("inst_key") or "").strip()
    name_hint = (body.get("bank_name") or "").strip()

    if not ik:
        return jsonify({"error": "inst_key required"}), 400

    try:
        job_id = _job_create(ik, name_hint or None)
    except Exception as e:
        print(f"[assessment-job] ✗ failed to create job: {e}")
        return jsonify({"error": f"Could not start job: {e}"}), 500

    thread = threading.Thread(target=_run_assessment_job, args=(job_id, ik, name_hint or None), daemon=True)
    thread.start()
    print(f"[assessment-job] started {job_id} for {ik}")
    return jsonify({"job_id": job_id}), 202


@app.route("/assessment-status/<job_id>", methods=["GET"])
@require_session
def assessment_status(job_id):
    row = _job_read(job_id)
    if not row:
        return jsonify({"error": "job not found"}), 404
    return jsonify(row)


@app.route("/assessment-download/<job_id>", methods=["GET"])
@require_session
def assessment_download(job_id):
    row = _job_read(job_id, select="status,filename,docx_base64")
    if not row:
        return jsonify({"error": "job not found"}), 404
    if row["status"] != "done":
        return jsonify({"error": f"job not ready (status={row['status']})"}), 409

    docx_bytes = base64.b64decode(row["docx_base64"])
    buf = io.BytesIO(docx_bytes)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=row["filename"]
    )


# ── Single deck ────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
@require_session
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
        bm.merge_growth_system_intro(buf)
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
@require_session
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
                bm.merge_growth_system_intro(deck_buf)
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
