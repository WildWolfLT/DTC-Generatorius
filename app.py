import os

from flask import Flask, jsonify, render_template, request

from engine import DTC_DATABASE, LOAD_ERROR, lookup_dtc, suggest_codes

# App Platform reserves /static at the edge, so the files live in static/ but are served from /assets.
app = Flask(__name__, static_url_path="/assets")
app.json.ensure_ascii = False


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/meta")
def meta():
    return jsonify(
        {
            "count": len(DTC_DATABASE),
            "ready": LOAD_ERROR is None,
            "error": LOAD_ERROR,
        }
    )


@app.get("/api/lookup")
def lookup():
    payload = lookup_dtc(request.args.get("code", ""))
    if payload["ok"]:
        status = 200
    elif payload["error"] == "invalid_format":
        status = 400
    else:
        status = 404
    return jsonify(payload), status


@app.get("/api/suggest")
def suggest():
    try:
        limit = int(request.args.get("limit", 8))
    except ValueError:
        limit = 8
    limit = max(1, min(limit, 12))
    return jsonify({"items": suggest_codes(request.args.get("q", ""), limit)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
