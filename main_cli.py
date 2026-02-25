#!/usr/bin/env python3
"""
Interface en ligne de commande — Générateur d'Assignation.

Usage :
    python3 main_cli.py
"""

import os
import sys

from src.document_generator import DocumentGenerator
from src.question_engine import QuestionEngine

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

CONFIG_PATH      = os.path.join("config", "questions.xlsx")
TEMPLATES_DIR    = "templates"
OUTPUT_DIR       = "output"
DEFAULT_TEMPLATE = "template.docx"

# ---------------------------------------------------------------------------
# Couleurs ANSI (désactivées automatiquement si pas de terminal interactif)
# ---------------------------------------------------------------------------

if sys.stdout.isatty():
    BOLD  = "\033[1m"
    BLUE  = "\033[34m"
    GREEN = "\033[32m"
    CYAN  = "\033[36m"
    RED   = "\033[31m"
    DIM   = "\033[2m"
    RESET = "\033[0m"
else:
    BOLD = BLUE = GREEN = CYAN = RED = DIM = RESET = ""


# ---------------------------------------------------------------------------
# Saisie selon le type de question
# ---------------------------------------------------------------------------

def _ask_yes_no(q) -> str:
    while True:
        raw = input("  [oui/non] > ").strip().lower()
        if raw in ("oui", "o", "yes", "y"):
            return "oui"
        if raw in ("non", "n", "no"):
            return "non"
        if not raw and not q.required:
            return ""
        print(f"  {RED}Tapez 'oui' ou 'non'.{RESET}")


def _ask_choice(q) -> str:
    for i, opt in enumerate(q.options, 1):
        print(f"    {DIM}{i}.{RESET} {opt}")
    while True:
        raw = input("  Votre choix (numéro ou texte) > ").strip()
        if not raw and not q.required:
            return ""
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(q.options):
                return q.options[idx]
        except ValueError:
            pass
        for opt in q.options:
            if raw.lower() == opt.lower():
                return opt
        print(f"  {RED}Choix invalide. Entrez un numéro (1-{len(q.options)}) ou le texte exact.{RESET}")


def _ask_multiline(q) -> str:
    print(f"  {DIM}(Saisissez le texte sur plusieurs lignes. Ligne vide pour terminer.){RESET}")
    lines = []
    while True:
        line = input("  > ")
        if line == "" and (lines or not q.required):
            break
        lines.append(line)
    return "\n".join(lines)


def _ask_text(q) -> str:
    hint = f" {DIM}[JJ/MM/AAAA]{RESET}" if q.type == "date" else ""
    while True:
        raw = input(f"  >{hint} ").strip()
        if raw:
            return raw
        if not q.required:
            return ""
        print(f"  {RED}Ce champ est requis. (Appuyez sur Entrée vide pour passer si le champ est optionnel.){RESET}")


def ask_question(q) -> str:
    """Affiche une question et retourne la réponse de l'utilisateur."""
    required_tag = f" {RED}*{RESET}" if q.required else f" {DIM}(optionnel — Entrée pour passer){RESET}"
    print(f"\n  {BOLD}{q.question}{required_tag}{RESET}")

    if q.type == "yes_no":
        return _ask_yes_no(q)
    if q.type == "choice":
        return _ask_choice(q)
    if q.type == "multiline":
        return _ask_multiline(q)
    return _ask_text(q)


# ---------------------------------------------------------------------------
# Boucle principale
# ---------------------------------------------------------------------------

def run() -> None:
    # Vérification du fichier de questions
    if not os.path.exists(CONFIG_PATH):
        print(f"\n{RED}Erreur : fichier introuvable :{RESET}")
        print(f"  {os.path.abspath(CONFIG_PATH)}")
        print(f"\nLancez d'abord  python3 setup_sample_data.py  pour créer les fichiers d'exemple.")
        sys.exit(1)

    engine = QuestionEngine(CONFIG_PATH)

    print(f"\n{BOLD}{BLUE}{'═' * 50}{RESET}")
    print(f"{BOLD}{BLUE}   Générateur d'Assignation{RESET}")
    print(f"{BOLD}{BLUE}{'═' * 50}{RESET}")
    print(f"  {len(engine.questions)} questions chargées.")
    print(f"  {DIM}Ctrl+C pour annuler.{RESET}")

    answers: dict = {}
    index = 0
    last_section = None

    while True:
        visible = engine.get_visible_questions(answers)

        if index >= len(visible):
            break

        q = visible[index]
        total = len(visible)

        # En-tête de section (affiché uniquement à la première question de chaque section)
        if q.section != last_section:
            print(f"\n  {CYAN}── {q.section} {'─' * max(0, 40 - len(q.section))}{RESET}")
            last_section = q.section

        print(f"  {DIM}Question {index + 1}/{total}{RESET}", end="")

        answer = ask_question(q)

        answers[q.id] = answer
        if q.variable:
            answers[q.variable] = answer

        index += 1

    # ---------------------------------------------------------------------------
    # Génération du document
    # ---------------------------------------------------------------------------
    print(f"\n{BOLD}{GREEN}{'═' * 50}{RESET}")
    print(f"{BOLD}{GREEN}  Génération du document…{RESET}")
    print(f"{BOLD}{GREEN}{'═' * 50}{RESET}")

    tmpl_file = engine.get_template_for_answers(answers, DEFAULT_TEMPLATE)
    tmpl_path = os.path.join(TEMPLATES_DIR, tmpl_file)

    if not os.path.exists(tmpl_path):
        print(f"\n{RED}Template introuvable : {os.path.abspath(tmpl_path)}{RESET}")
        print("Vérifiez le dossier 'templates/' ou les règles de variantes.")
        sys.exit(1)

    try:
        gen = DocumentGenerator(TEMPLATES_DIR, OUTPUT_DIR)
        out = gen.generate(tmpl_file, answers)
        print(f"\n{GREEN}✓ Document généré :{RESET}")
        print(f"  {os.path.abspath(out)}")
    except Exception as exc:
        print(f"\n{RED}Erreur lors de la génération : {exc}{RESET}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # Ouvrir le document ?
    raw = input("\n  Ouvrir le document ? [O/n] ").strip().lower()
    if raw in ("", "o", "oui", "y", "yes"):
        DocumentGenerator.open_file(out)

    # Nouvelle assignation ?
    raw = input("\n  Rédiger une nouvelle assignation ? [O/n] ").strip().lower()
    if raw in ("", "o", "oui", "y", "yes"):
        run()


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    try:
        run()
    except KeyboardInterrupt:
        print(f"\n\n{DIM}Annulé.{RESET}\n")
