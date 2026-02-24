"""
Interface graphique principale (tkinter).

Flux :
  1. Chargement des questions depuis config/questions.xlsx
  2. Navigation question par question avec conditions dynamiques
  3. Bouton "Générer" sur la dernière question → crée le .docx dans output/
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from src.document_generator import DocumentGenerator
from src.question_engine import Question, QuestionEngine


# ---------------------------------------------------------------------------
# Constantes de style
# ---------------------------------------------------------------------------

CLR_BG       = "#f4f6f9"
CLR_HEADER   = "#1a3a5c"
CLR_ACCENT   = "#2980b9"
CLR_SUCCESS  = "#27ae60"
CLR_LIGHT    = "#ecf0f1"
CLR_TEXT     = "#2c3e50"
CLR_MUTED    = "#7f8c8d"
CLR_WHITE    = "#ffffff"
CLR_BORDER   = "#d5dbe3"

FONT_TITLE   = ("Helvetica", 16, "bold")
FONT_SECTION = ("Helvetica", 11, "italic")
FONT_QUESTION= ("Helvetica", 13)
FONT_BODY    = ("Helvetica", 11)
FONT_SMALL   = ("Helvetica", 9)


# ---------------------------------------------------------------------------
# Application principale
# ---------------------------------------------------------------------------

class AssignationApp:

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Générateur d'Assignation")
        self.root.geometry("960x720")
        self.root.configure(bg=CLR_BG)
        self.root.resizable(True, True)

        # Chemins par défaut
        self.config_path     = os.path.join("config", "questions.xlsx")
        self.templates_dir   = "templates"
        self.output_dir      = "output"
        self.default_template= "template.docx"

        # État
        self.engine  : Optional[QuestionEngine] = None
        self.answers : Dict[str, str] = {}
        self.visible : List[Question] = []
        self.current_index   = 0
        self._input_widget   = None   # widget d'entrée actif

        self._build_ui()
        self._load_engine()

    # ======================================================================
    # Construction de l'interface
    # ======================================================================

    def _build_ui(self) -> None:
        # ---- Barre de titre --------------------------------------------------
        header = tk.Frame(self.root, bg=CLR_HEADER, height=64)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(
            header, text="  Générateur d'Assignation",
            bg=CLR_HEADER, fg=CLR_WHITE, font=FONT_TITLE, anchor=tk.W
        ).pack(side=tk.LEFT, padx=10)

        tk.Button(
            header, text="⚙  Paramètres",
            bg="#16324f", fg=CLR_WHITE, relief=tk.FLAT,
            font=FONT_BODY, padx=12, pady=6, cursor="hand2",
            command=self._open_settings,
        ).pack(side=tk.RIGHT, padx=12, pady=14)

        # ---- Barre de progression --------------------------------------------
        prog_frame = tk.Frame(self.root, bg=CLR_BG)
        prog_frame.pack(fill=tk.X, padx=24, pady=(14, 0))

        self.lbl_progress = tk.Label(
            prog_frame, text="", bg=CLR_BG, fg=CLR_MUTED, font=FONT_SMALL
        )
        self.lbl_progress.pack(anchor=tk.W)

        self.progressbar = ttk.Progressbar(prog_frame, mode="determinate")
        self.progressbar.pack(fill=tk.X, pady=(4, 0))

        # ---- Étiquette de section --------------------------------------------
        self.lbl_section = tk.Label(
            self.root, text="", bg=CLR_BG, fg=CLR_ACCENT, font=FONT_SECTION
        )
        self.lbl_section.pack(padx=24, pady=(10, 0), anchor=tk.W)

        # ---- Carte question ---------------------------------------------------
        card = tk.Frame(
            self.root, bg=CLR_WHITE,
            highlightbackground=CLR_BORDER, highlightthickness=1,
        )
        card.pack(fill=tk.BOTH, expand=True, padx=24, pady=10)

        self.lbl_question = tk.Label(
            card, text="", bg=CLR_WHITE, fg=CLR_TEXT,
            font=FONT_QUESTION, wraplength=880, justify=tk.LEFT, anchor=tk.NW,
        )
        self.lbl_question.pack(padx=28, pady=(24, 12), fill=tk.X)

        self.answer_frame = tk.Frame(card, bg=CLR_WHITE)
        self.answer_frame.pack(padx=28, pady=(0, 20), fill=tk.BOTH, expand=True)

        # ---- Barre de navigation ---------------------------------------------
        nav = tk.Frame(self.root, bg=CLR_BG)
        nav.pack(fill=tk.X, padx=24, pady=(0, 18))

        self.btn_prev = tk.Button(
            nav, text="← Précédent",
            bg=CLR_LIGHT, fg=CLR_TEXT, relief=tk.FLAT,
            font=FONT_BODY, padx=18, pady=8, cursor="hand2",
            command=self._go_prev,
        )
        self.btn_prev.pack(side=tk.LEFT)

        self.btn_skip = tk.Button(
            nav, text="Passer",
            bg=CLR_BG, fg=CLR_MUTED, relief=tk.FLAT,
            font=FONT_SMALL, padx=10, pady=8, cursor="hand2",
            command=self._skip,
        )
        self.btn_skip.pack(side=tk.LEFT, padx=12)

        # Bouton "Suivant" et "Générer" (on alterne leur visibilité)
        self.btn_generate = tk.Button(
            nav, text="✓  Générer le document",
            bg=CLR_SUCCESS, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 11, "bold"), padx=20, pady=8, cursor="hand2",
            command=self._generate,
        )

        self.btn_next = tk.Button(
            nav, text="Suivant →",
            bg=CLR_ACCENT, fg=CLR_WHITE, relief=tk.FLAT,
            font=FONT_BODY, padx=20, pady=8, cursor="hand2",
            command=self._go_next,
        )
        self.btn_next.pack(side=tk.RIGHT)

        # Variable partagée pour les widgets radio / entry simples
        self._var = tk.StringVar()

    # ======================================================================
    # Chargement des questions
    # ======================================================================

    def _load_engine(self) -> None:
        if not os.path.exists(self.config_path):
            messagebox.showwarning(
                "Fichier manquant",
                f"Le fichier de questions est introuvable :\n{os.path.abspath(self.config_path)}\n\n"
                "Lancez d'abord  setup_sample_data.py  pour créer les fichiers d'exemple,\n"
                "ou utilisez ⚙ Paramètres pour sélectionner votre fichier.",
            )
            return

        try:
            self.engine  = QuestionEngine(self.config_path)
            self.answers = {}
            self._refresh_visible()
            self._show(0)
        except Exception as exc:
            messagebox.showerror("Erreur de chargement", str(exc))

    def _refresh_visible(self) -> None:
        if self.engine:
            self.visible = self.engine.get_visible_questions(self.answers)

    # ======================================================================
    # Affichage d'une question
    # ======================================================================

    def _show(self, index: int) -> None:
        self._refresh_visible()

        if not self.visible:
            self.lbl_question.config(text="Aucune question à afficher.")
            return

        # Borne l'index
        index = max(0, min(index, len(self.visible) - 1))
        self.current_index = index
        q = self.visible[index]

        total = len(self.visible)
        self.lbl_progress.config(text=f"Question {index + 1} / {total}")
        self.progressbar["value"] = (index + 1) / total * 100
        self.lbl_section.config(text=q.section)

        # Texte de la question
        suffix = " *" if q.required else "  (optionnel)"
        self.lbl_question.config(text=q.question + suffix)

        # Vide le cadre réponse
        for w in self.answer_frame.winfo_children():
            w.destroy()
        self._input_widget = None

        # Valeur sauvegardée pour cette question
        saved = self.answers.get(q.id, "")

        # Construit le widget d'entrée selon le type
        builders = {
            "text"     : self._w_text,
            "date"     : self._w_date,
            "number"   : self._w_number,
            "multiline": self._w_multiline,
            "yes_no"   : self._w_yes_no,
            "choice"   : self._w_choice,
        }
        builder = builders.get(q.type, self._w_text)
        builder(q, saved)

        # Boutons de navigation
        self.btn_prev.config(state=tk.NORMAL if index > 0 else tk.DISABLED)
        self.btn_skip.config(state=tk.NORMAL if not q.required else tk.DISABLED)

        is_last = index == len(self.visible) - 1
        if is_last:
            self.btn_next.pack_forget()
            self.btn_generate.pack(side=tk.RIGHT)
        else:
            self.btn_generate.pack_forget()
            self.btn_next.pack(side=tk.RIGHT)

    # ======================================================================
    # Builders de widgets de réponse
    # ======================================================================

    def _w_text(self, q: Question, saved: str) -> None:
        self._var = tk.StringVar(value=saved)
        e = ttk.Entry(self.answer_frame, textvariable=self._var,
                      font=("Helvetica", 12), width=60)
        e.pack(anchor=tk.W, ipady=6)
        e.focus_set()
        e.bind("<Return>", lambda _: self._go_next())

    def _w_date(self, q: Question, saved: str) -> None:
        self._var = tk.StringVar(value=saved)
        tk.Label(self.answer_frame, text="Format : JJ/MM/AAAA",
                 bg=CLR_WHITE, fg=CLR_MUTED, font=FONT_SMALL).pack(anchor=tk.W)
        e = ttk.Entry(self.answer_frame, textvariable=self._var,
                      font=("Helvetica", 12), width=20)
        e.pack(anchor=tk.W, ipady=6)
        e.focus_set()
        e.bind("<Return>", lambda _: self._go_next())

    def _w_number(self, q: Question, saved: str) -> None:
        self._var = tk.StringVar(value=saved)
        e = ttk.Entry(self.answer_frame, textvariable=self._var,
                      font=("Helvetica", 12), width=20)
        e.pack(anchor=tk.W, ipady=6)
        e.focus_set()
        e.bind("<Return>", lambda _: self._go_next())

    def _w_multiline(self, q: Question, saved: str) -> None:
        t = tk.Text(
            self.answer_frame, font=("Helvetica", 12),
            width=72, height=7, wrap=tk.WORD,
            relief=tk.SOLID, borderwidth=1,
            bg=CLR_WHITE, fg=CLR_TEXT,
        )
        t.pack(anchor=tk.W)
        t.insert("1.0", saved)
        t.focus_set()
        self._input_widget = t   # sera lu dans _get_answer()

    def _w_yes_no(self, q: Question, saved: str) -> None:
        self._var = tk.StringVar(value=saved or "")
        f = tk.Frame(self.answer_frame, bg=CLR_WHITE)
        f.pack(anchor=tk.W)
        for val, label in [("oui", "Oui"), ("non", "Non")]:
            tk.Radiobutton(
                f, text=label, variable=self._var, value=val,
                bg=CLR_WHITE, fg=CLR_TEXT, font=FONT_BODY,
                activebackground=CLR_WHITE, cursor="hand2",
            ).pack(side=tk.LEFT, padx=(0, 24))

    def _w_choice(self, q: Question, saved: str) -> None:
        self._var = tk.StringVar(value=saved or "")
        for opt in q.options:
            tk.Radiobutton(
                self.answer_frame, text=opt, variable=self._var, value=opt,
                bg=CLR_WHITE, fg=CLR_TEXT, font=FONT_BODY,
                activebackground=CLR_WHITE, cursor="hand2",
            ).pack(anchor=tk.W, pady=3)

    # ======================================================================
    # Lecture de la réponse courante
    # ======================================================================

    def _get_answer(self) -> str:
        if self._input_widget is not None:
            # Text widget multiline
            return self._input_widget.get("1.0", tk.END).strip()
        return self._var.get().strip()

    # ======================================================================
    # Sauvegarde + validation
    # ======================================================================

    def _save(self, validate: bool = True) -> bool:
        if not self.visible:
            return True

        q = self.visible[self.current_index]
        answer = self._get_answer()

        if validate and q.required and not answer:
            messagebox.showwarning(
                "Champ requis",
                "Veuillez répondre à cette question avant de continuer.\n"
                "(Utilisez 'Passer' pour ignorer une question optionnelle.)",
            )
            return False

        self.answers[q.id] = answer
        if q.variable:
            self.answers[q.variable] = answer
        return True

    # ======================================================================
    # Navigation
    # ======================================================================

    def _go_next(self) -> None:
        if not self._save(validate=True):
            return
        self._refresh_visible()
        if self.current_index + 1 < len(self.visible):
            self._show(self.current_index + 1)

    def _go_prev(self) -> None:
        self._save(validate=False)
        if self.current_index > 0:
            self._show(self.current_index - 1)

    def _skip(self) -> None:
        if not self.visible:
            return
        q = self.visible[self.current_index]
        self.answers[q.id] = ""
        if q.variable:
            self.answers[q.variable] = ""
        self._refresh_visible()
        if self.current_index + 1 < len(self.visible):
            self._show(self.current_index + 1)

    # ======================================================================
    # Génération du document
    # ======================================================================

    def _generate(self) -> None:
        if not self._save(validate=True):
            return
        if not self.engine:
            return

        tmpl_file = self.engine.get_template_for_answers(
            self.answers, self.default_template
        )
        tmpl_path = os.path.join(self.templates_dir, tmpl_file)

        if not os.path.exists(tmpl_path):
            messagebox.showerror(
                "Template introuvable",
                f"Le fichier template n'existe pas :\n{os.path.abspath(tmpl_path)}\n\n"
                "Vérifiez le dossier 'templates/' ou les règles de variantes.",
            )
            return

        try:
            gen = DocumentGenerator(self.templates_dir, self.output_dir)
            out = gen.generate(tmpl_file, self.answers)

            msg = f"Document généré avec succès :\n{os.path.abspath(out)}"
            if messagebox.askyesno("Document généré !", msg + "\n\nOuvrir le document ?"):
                DocumentGenerator.open_file(out)

            if messagebox.askyesno("Nouvelle assignation ?",
                                   "Voulez-vous rédiger une nouvelle assignation ?"):
                self.answers = {}
                self._refresh_visible()
                self._show(0)

        except Exception as exc:
            messagebox.showerror("Erreur lors de la génération", str(exc))

    # ======================================================================
    # Fenêtre Paramètres
    # ======================================================================

    def _open_settings(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Paramètres")
        win.geometry("560x320")
        win.configure(bg=CLR_BG)
        win.grab_set()

        tk.Label(win, text="Paramètres", bg=CLR_BG, font=FONT_TITLE,
                 fg=CLR_HEADER).pack(pady=(20, 10))

        def row(label: str, command, bg: str) -> None:
            f = tk.Frame(win, bg=CLR_BG)
            f.pack(fill=tk.X, padx=30, pady=6)
            tk.Label(f, text=label, bg=CLR_BG, font=FONT_BODY,
                     fg=CLR_TEXT, width=36, anchor=tk.W).pack(side=tk.LEFT)
            tk.Button(f, text="Parcourir…", command=command,
                      bg=CLR_ACCENT, fg=CLR_WHITE, relief=tk.FLAT,
                      font=FONT_SMALL, padx=10, pady=4, cursor="hand2").pack(side=tk.RIGHT)

        def pick_questions() -> None:
            path = filedialog.askopenfilename(
                parent=win, title="Choisir le fichier de questions",
                filetypes=[("Fichiers Excel", "*.xlsx")]
            )
            if path:
                self.config_path = path
                win.destroy()
                self._load_engine()

        def pick_template() -> None:
            path = filedialog.askopenfilename(
                parent=win, title="Choisir le template par défaut",
                filetypes=[("Documents Word", "*.docx")]
            )
            if path:
                self.default_template = os.path.basename(path)
                self.templates_dir    = os.path.dirname(path)
                messagebox.showinfo("Paramètre mis à jour",
                                    f"Template : {self.default_template}",
                                    parent=win)

        def open_output() -> None:
            DocumentGenerator.open_file(os.path.abspath(self.output_dir))

        row("Fichier de questions (.xlsx)", pick_questions, CLR_ACCENT)
        row("Template Word par défaut (.docx)", pick_template, CLR_ACCENT)

        tk.Button(
            win, text="📂  Ouvrir le dossier de sortie",
            bg=CLR_SUCCESS, fg=CLR_WHITE, relief=tk.FLAT,
            font=FONT_BODY, padx=14, pady=8, cursor="hand2",
            command=open_output,
        ).pack(pady=(20, 0))
