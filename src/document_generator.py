"""
Générateur de documents Word à partir d'un template et d'un dictionnaire de réponses.

Utilise la bibliothèque docxtpl (basée sur python-docx + Jinja2).
Dans votre fichier .docx, utilisez la syntaxe :
    {{ nom_variable }}          → remplacement simple
    {% if Q005 == 'oui' %}      → bloc conditionnel
    {% endif %}
    {{ date_assignation | upper }} → filtres Jinja2
"""

from __future__ import annotations

import os
import subprocess
import platform
from datetime import datetime
from typing import Dict

from docxtpl import DocxTemplate


class DocumentGenerator:

    def __init__(self, templates_dir: str, output_dir: str) -> None:
        self.templates_dir = templates_dir
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)

    # ------------------------------------------------------------------
    # Génération du document
    # ------------------------------------------------------------------

    def generate(self, template_file: str, answers: Dict[str, str]) -> str:
        """
        Remplace les variables Jinja2 dans le template Word avec les réponses,
        sauvegarde le fichier dans output_dir et retourne le chemin du fichier.
        """
        template_path = os.path.join(self.templates_dir, template_file)

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template introuvable : {template_path}")

        doc = DocxTemplate(template_path)

        # Contexte : on expose les réponses par variable ET par id de question
        context = dict(answers)
        # Ajoute la date de génération automatiquement
        context.setdefault("date_generation", datetime.now().strftime("%d/%m/%Y"))
        context.setdefault("date_generation_long",
                           datetime.now().strftime("%A %d %B %Y").capitalize())

        doc.render(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"assignation_{timestamp}.docx"
        output_path = os.path.join(self.output_dir, output_filename)
        doc.save(output_path)

        return output_path

    # ------------------------------------------------------------------
    # Ouverture du fichier généré
    # ------------------------------------------------------------------

    @staticmethod
    def open_file(path: str) -> None:
        """Ouvre le fichier avec l'application par défaut du système."""
        system = platform.system()
        try:
            if system == "Windows":
                os.startfile(path)  # type: ignore[attr-defined]
            elif system == "Darwin":
                subprocess.call(["open", path])
            else:
                subprocess.call(["xdg-open", path])
        except Exception:
            pass  # Si l'ouverture échoue, on ne bloque pas l'utilisateur
