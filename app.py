#!/usr/bin/env python3
"""
Flask API wrapper for the Change Order Generator.
Not used locally — deploy this to Render when ready to publish.

Endpoint: POST /generate
Payload:  JSON matching the data dict in generator.py
Returns:  JSON with base64-encoded PDF and Word doc

Deploy steps (Render free tier):
1. Push this repo to GitHub
2. Create new Web Service on render.com, connect the repo
3. Build command: pip install -r requirements.txt
4. Start command: gunicorn app:app
5. Add env var: ANTHROPIC_API_KEY
6. Point n8n HTTP Request node at the Render URL
"""

import os
import base64
import tempfile
from flask import Flask, request, jsonify
from generator import generate

app = Flask(__name__)


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/generate", methods=["POST"])
def generate_change_order():
    data = request.get_json()

    if not data:
        return jsonify({"error": "No JSON payload received"}), 400

    required = ["company_name", "project_name", "scope_description"]
    missing = [f for f in required if not data.get(f)]
    if missing:
        return jsonify({"error": f"Missing required fields: {missing}"}), 400

    if len(data.get("scope_description", "")) < 10:
        return jsonify({
            "error": "Scope description is too short. Please describe the change in detail."
        }), 400

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            result = generate(data, output_dir=tmpdir)

            with open(result["pdf"], "rb") as f:
                pdf_b64 = base64.b64encode(f.read()).decode("utf-8")
            with open(result["word"], "rb") as f:
                word_b64 = base64.b64encode(f.read()).decode("utf-8")

        return jsonify({
            "success": True,
            "generated_text": result["generated_text"],
            "pdf_base64":  pdf_b64,
            "word_base64": word_b64,
            "pdf_filename":  os.path.basename(result["pdf"]),
            "word_filename": os.path.basename(result["word"]),
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5001)
