"""
Application web Flask — Générateur d'Assignation.

Lancez avec :  python app.py
Puis ouvrez  :  http://127.0.0.1:5000
"""

from __future__ import annotations

import os
import secrets
import threading
import uuid
import webbrowser
from typing import Dict

from flask import Flask, redirect, render_template_string, request, send_file, session, url_for

from src.document_generator import DocumentGenerator
from src.question_engine import QuestionEngine

# ---------------------------------------------------------------------------
# Configuration chemins
# ---------------------------------------------------------------------------

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH   = os.path.join(BASE_DIR, "config", "questions.xlsx")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR    = os.path.join(BASE_DIR, "output")
DEFAULT_TPL   = "template.docx"

# ---------------------------------------------------------------------------
# Flask + stockage côté serveur (évite la limite de taille des cookies)
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

_engine: QuestionEngine | None = None
_server_sessions: Dict[str, Dict[str, str]] = {}


def get_engine() -> QuestionEngine:
    global _engine
    if _engine is None:
        _engine = QuestionEngine(CONFIG_PATH)
    return _engine


def _get_answers() -> Dict[str, str]:
    sid = session.get("sid")
    if sid:
        return _server_sessions.get(sid, {})
    return {}


def _set_answers(answers: Dict[str, str]) -> None:
    if "sid" not in session:
        session["sid"] = str(uuid.uuid4())
    _server_sessions[session["sid"]] = answers


# ---------------------------------------------------------------------------
# Styles CSS partagés
# ---------------------------------------------------------------------------

_CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
  background: #f0f4f8;
  color: #1e293b;
  min-height: 100vh;
  display: flex;
  flex-direction: column;
}

header {
  background: #1a3a5c;
  color: white;
  padding: 0 2rem;
  height: 56px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-shrink: 0;
  box-shadow: 0 2px 8px rgba(0,0,0,.2);
}
header h1 { font-size: 1.05rem; font-weight: 600; letter-spacing: .2px; }

.progress-bar-bg {
  height: 4px;
  background: #c7d7ea;
  position: relative;
}
.progress-bar-fill {
  position: absolute;
  top: 0; left: 0; height: 100%;
  background: #2563eb;
  transition: width .35s ease;
}

.meta-bar {
  background: white;
  padding: .65rem 2rem;
  border-bottom: 1px solid #e2e8f0;
  display: flex;
  align-items: center;
  justify-content: space-between;
  font-size: .82rem;
  color: #64748b;
}

.section-badge {
  display: inline-block;
  font-size: .72rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: .06em;
  color: #1d4ed8;
  background: #eff6ff;
  padding: .15rem .65rem;
  border-radius: 9999px;
}

main {
  flex: 1;
  display: flex;
  align-items: flex-start;
  justify-content: center;
  padding: 2rem 1rem;
}

.card {
  background: white;
  border-radius: 12px;
  box-shadow: 0 1px 4px rgba(0,0,0,.08), 0 8px 24px rgba(0,0,0,.06);
  padding: 2.5rem;
  width: 100%;
  max-width: 680px;
}

.question-text {
  font-size: 1.1rem;
  font-weight: 500;
  line-height: 1.6;
  margin-bottom: 1.5rem;
  color: #0f172a;
}
.required-star { color: #ef4444; margin-left: 2px; }
.optional-tag {
  display: inline-block;
  font-size: .75rem;
  font-weight: 400;
  color: #94a3b8;
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  padding: .1rem .5rem;
  border-radius: 4px;
  margin-left: .4rem;
  vertical-align: middle;
}

.hint {
  font-size: .8rem;
  color: #94a3b8;
  margin-bottom: .5rem;
}

input[type=text], input[type=number], textarea {
  display: block;
  width: 100%;
  padding: .7rem 1rem;
  font-size: 1rem;
  font-family: inherit;
  border: 1.5px solid #cbd5e1;
  border-radius: 8px;
  outline: none;
  background: #f8fafc;
  color: #0f172a;
  transition: border-color .15s, box-shadow .15s;
}
input[type=text]:focus, input[type=number]:focus, textarea:focus {
  border-color: #2563eb;
  background: white;
  box-shadow: 0 0 0 3px rgba(37,99,235,.12);
}
textarea { resize: vertical; min-height: 130px; line-height: 1.5; }

.radio-group { display: flex; flex-direction: column; gap: .5rem; }
.radio-option {
  display: flex;
  align-items: center;
  gap: .75rem;
  padding: .75rem 1rem;
  border: 1.5px solid #e2e8f0;
  border-radius: 8px;
  cursor: pointer;
  transition: border-color .15s, background .15s;
  user-select: none;
}
.radio-option.checked {
  border-color: #2563eb;
  background: #eff6ff;
}
.radio-option input[type=radio] {
  accent-color: #2563eb;
  width: 17px; height: 17px;
  cursor: pointer;
  flex-shrink: 0;
}
.radio-option span { font-size: .95rem; cursor: pointer; }

.error-box {
  background: #fef2f2;
  border: 1px solid #fecaca;
  color: #b91c1c;
  padding: .7rem 1rem;
  border-radius: 8px;
  margin-bottom: 1.2rem;
  font-size: .9rem;
}

.nav {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-top: 2rem;
  padding-top: 1.5rem;
  border-top: 1px solid #f1f5f9;
}
.nav-left { display: flex; gap: .5rem; align-items: center; }

button, .btn {
  display: inline-block;
  text-decoration: none;
  font-family: inherit;
  font-size: .93rem;
  font-weight: 500;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  padding: .6rem 1.3rem;
  transition: filter .15s, opacity .15s;
  line-height: 1.4;
}
button:hover, .btn:hover { filter: brightness(.93); }
button:disabled { opacity: .35; cursor: not-allowed; filter: none; pointer-events: none; }

.btn-primary  { background: #2563eb; color: white; }
.btn-success  { background: #16a34a; color: white; }
.btn-neutral  { background: #e2e8f0; color: #475569; }
.btn-ghost    { background: transparent; color: #94a3b8; font-size: .85rem; }

/* Pages success / erreur */
.center-card { text-align: center; padding: 3rem 2.5rem; }
.big-icon    { font-size: 3.5rem; margin-bottom: 1.2rem; }
.page-title  { font-size: 1.4rem; font-weight: 700; margin-bottom: .5rem; }
.page-sub    { color: #64748b; font-size: .9rem; margin-bottom: 1.8rem; word-break: break-all; }
.action-row  { display: flex; gap: 1rem; justify-content: center; flex-wrap: wrap; margin-top: 1.5rem; }
"""

# ---------------------------------------------------------------------------
# Templates HTML
# ---------------------------------------------------------------------------

TMPL_QUESTION = """\
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Générateur d'Assignation</title>
  <style>{{ css | safe }}</style>
</head>
<body>

<header>
  <h1>⚖ Générateur d'Assignation</h1>
  <span style="font-size:.8rem;opacity:.65">{{ index + 1 }} / {{ total }}</span>
</header>

<div class="progress-bar-bg">
  <div class="progress-bar-fill" style="width:{{ pct }}%"></div>
</div>

<div class="meta-bar">
  <span>Question <strong>{{ index + 1 }}</strong> sur <strong>{{ total }}</strong></span>
  <span class="section-badge">{{ q.section }}</span>
</div>

<main>
  <div class="card">

    {% if error %}
    <div class="error-box">{{ error }}</div>
    {% endif %}

    <form method="POST">
      <input type="hidden" name="q_id"     value="{{ q.id }}">
      <input type="hidden" name="q_var"    value="{{ q.variable }}">
      <input type="hidden" name="required" value="{{ '1' if q.required else '0' }}">

      <p class="question-text">
        {{ q.question }}
        {% if q.required %}
          <span class="required-star">*</span>
        {% else %}
          <span class="optional-tag">optionnel</span>
        {% endif %}
      </p>

      {% if q.type in ('text', 'number') %}
        <input
          type="{{ 'number' if q.type == 'number' else 'text' }}"
          name="answer"
          value="{{ saved }}"
          autofocus
          autocomplete="off">

      {% elif q.type == 'date' %}
        <p class="hint">Format : JJ/MM/AAAA — ex : 15/03/2025</p>
        <input type="text" name="answer" value="{{ saved }}"
               placeholder="JJ/MM/AAAA" autofocus autocomplete="off">

      {% elif q.type == 'multiline' %}
        <textarea name="answer" autofocus>{{ saved }}</textarea>

      {% elif q.type == 'yes_no' %}
        <div class="radio-group" id="rg">
          {% for val, label in [('oui','Oui'),('non','Non')] %}
          <label class="radio-option {% if saved == val %}checked{% endif %}">
            <input type="radio" name="answer" value="{{ val }}"
                   {% if saved == val %}checked{% endif %}
                   onchange="document.querySelectorAll('#rg .radio-option').forEach(e=>e.classList.remove('checked'));this.closest('.radio-option').classList.add('checked')">
            <span>{{ label }}</span>
          </label>
          {% endfor %}
        </div>

      {% elif q.type == 'choice' %}
        <div class="radio-group" id="rg">
          {% for opt in q.options %}
          <label class="radio-option {% if saved == opt %}checked{% endif %}">
            <input type="radio" name="answer" value="{{ opt }}"
                   {% if saved == opt %}checked{% endif %}
                   onchange="document.querySelectorAll('#rg .radio-option').forEach(e=>e.classList.remove('checked'));this.closest('.radio-option').classList.add('checked')">
            <span>{{ opt }}</span>
          </label>
          {% endfor %}
        </div>

      {% else %}
        <input type="text" name="answer" value="{{ saved }}" autofocus autocomplete="off">
      {% endif %}

      <div class="nav">
        <div class="nav-left">
          <button type="submit" name="action" value="prev"
                  class="btn-neutral" {% if index == 0 %}disabled{% endif %}>
            ← Précédent
          </button>
          {% if not q.required %}
          <button type="submit" name="action" value="skip" class="btn-ghost">
            Passer
          </button>
          {% endif %}
        </div>

        {% if is_last %}
        <button type="submit" name="action" value="generate" class="btn-success">
          ✓&nbsp;Générer le document
        </button>
        {% else %}
        <button type="submit" name="action" value="next" class="btn-primary">
          Suivant →
        </button>
        {% endif %}
      </div>
    </form>

  </div>
</main>

</body>
</html>
"""

TMPL_SUCCESS = """\
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Document généré</title>
  <style>{{ css | safe }}</style>
</head>
<body>

<header>
  <h1>⚖ Générateur d'Assignation</h1>
</header>

<main>
  <div class="card center-card">
    <div class="big-icon">✅</div>
    <p class="page-title" style="color:#15803d">Document généré avec succès !</p>
    <p class="page-sub">{{ filename }}</p>
    <div class="action-row">
      <a href="/download/{{ filename }}" class="btn btn-success">
        ↓&nbsp;Télécharger le document
      </a>
      <a href="/reset" class="btn btn-neutral">
        + Nouvelle assignation
      </a>
    </div>
  </div>
</main>

</body>
</html>
"""

TMPL_ERROR = """\
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <title>Erreur</title>
  <style>{{ css | safe }}</style>
</head>
<body>

<header>
  <h1>⚖ Générateur d'Assignation</h1>
</header>

<main>
  <div class="card center-card">
    <div class="big-icon">⚠️</div>
    <p class="page-title" style="color:#b91c1c">Erreur</p>
    <p class="page-sub">{{ message }}</p>
    <div class="action-row">
      <a href="/" class="btn btn-primary">Recommencer</a>
    </div>
  </div>
</main>

</body>
</html>
"""

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    _set_answers({})
    return redirect(url_for("question_view", index=0))


@app.route("/question/<int:index>", methods=["GET", "POST"])
def question_view(index: int):
    try:
        eng = get_engine()
    except Exception as exc:
        return render_template_string(TMPL_ERROR, css=_CSS,
                                      message=f"Impossible de charger les questions : {exc}")

    answers = _get_answers()

    if request.method == "POST":
        action   = request.form.get("action", "next")
        q_id     = request.form.get("q_id", "")
        q_var    = request.form.get("q_var", "")
        required = request.form.get("required") == "1"
        answer   = request.form.get("answer", "").strip()

        # Précédent — sauvegarde sans validation
        if action == "prev":
            answers[q_id] = answer
            if q_var:
                answers[q_var] = answer
            _set_answers(answers)
            return redirect(url_for("question_view", index=max(0, index - 1)))

        # Passer — efface la réponse, avance
        if action == "skip":
            answers[q_id] = ""
            if q_var:
                answers[q_var] = ""
            _set_answers(answers)
            visible = eng.get_visible_questions(answers)
            return redirect(url_for("question_view", index=min(index + 1, len(visible) - 1)))

        # Suivant / Générer — validation si requis
        if required and not answer:
            visible = eng.get_visible_questions(answers)
            idx = max(0, min(index, len(visible) - 1))
            q   = visible[idx]
            return render_template_string(TMPL_QUESTION, css=_CSS,
                q=q, index=idx, total=len(visible),
                pct=round((idx + 1) / len(visible) * 100),
                saved=answer, is_last=(idx == len(visible) - 1),
                error="Ce champ est obligatoire.")

        # Sauvegarde
        answers[q_id] = answer
        if q_var:
            answers[q_var] = answer
        _set_answers(answers)

        if action == "generate":
            return redirect(url_for("generate_doc"))

        visible = eng.get_visible_questions(answers)
        nxt = index + 1
        if nxt >= len(visible):
            return redirect(url_for("generate_doc"))
        return redirect(url_for("question_view", index=nxt))

    # GET
    visible = eng.get_visible_questions(answers)
    if not visible:
        return redirect(url_for("generate_doc"))

    idx = max(0, min(index, len(visible) - 1))
    q   = visible[idx]

    return render_template_string(TMPL_QUESTION, css=_CSS,
        q=q, index=idx, total=len(visible),
        pct=round((idx + 1) / len(visible) * 100),
        saved=answers.get(q.id, ""),
        is_last=(idx == len(visible) - 1),
        error=None)


@app.route("/generate")
def generate_doc():
    try:
        eng = get_engine()
    except Exception as exc:
        return render_template_string(TMPL_ERROR, css=_CSS,
                                      message=f"Impossible de charger les questions : {exc}")

    answers = _get_answers()

    try:
        tmpl_file = eng.get_template_for_answers(answers, DEFAULT_TPL)
        gen       = DocumentGenerator(TEMPLATES_DIR, OUTPUT_DIR)
        out_path  = gen.generate(tmpl_file, answers)
        filename  = os.path.basename(out_path)
    except Exception as exc:
        return render_template_string(TMPL_ERROR, css=_CSS,
                                      message=f"Erreur lors de la génération : {exc}")

    return render_template_string(TMPL_SUCCESS, css=_CSS, filename=filename)


@app.route("/download/<filename>")
def download_file(filename: str):
    safe_name = os.path.basename(filename)
    file_path = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.exists(file_path):
        return "Fichier introuvable", 404
    return send_file(file_path, as_attachment=True, download_name=safe_name)


@app.route("/reset")
def reset():
    _set_answers({})
    return redirect(url_for("question_view", index=0))


# ---------------------------------------------------------------------------
# Lancement
# ---------------------------------------------------------------------------

def _open_browser() -> None:
    import time
    time.sleep(0.8)
    webbrowser.open("http://127.0.0.1:5000")


def main() -> None:
    print("=" * 52)
    print("  Générateur d'Assignation — Interface Web")
    print("  Ouvrez votre navigateur sur :")
    print("  http://127.0.0.1:5000")
    print("  Ctrl+C pour arrêter.")
    print("=" * 52)
    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
