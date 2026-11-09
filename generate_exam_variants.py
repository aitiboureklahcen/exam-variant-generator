#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generic exam-variant generator.

Copy this file into any folder that contains .txt quiz files, then run it:
    python generate_exam_variants.py

It reads every .txt file in its own folder, parses questions of the form:

    Question text (one or more lines)
    A. option text
    B. option text
    C. option text
    D. option text
    ANSWER: X

(blank lines separate questions), mixes them into 4 .docx variants with
shuffled question order and shuffled answer options, an answer sheet for the
student at the start of each variant, and a confidential correction grid at
the end. Also writes grille_correction.csv and Grille_Comparative.docx.
"""
import csv
import os
import random
import re
import sys

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("python-docx is required: pip install python-docx")
    sys.exit(1)

BASE = os.path.dirname(os.path.abspath(__file__))
VARIANTS = 4
QUESTIONS_PER_VARIANT = 40

OPTION_RE = re.compile(r"^([A-D])[\.\)\-:]\s*(.*)$", re.IGNORECASE)
ANSWER_RE = re.compile(r"(?:ANSWER|REPONSE|RÉPONSE|REPONSE)\s*[:\-]?\s*([A-D])", re.IGNORECASE)


def read_text(path):
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(path, encoding="latin-1") as f:
        return f.read()


def parse_quiz(text, tag):
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]
    qs = []
    for i, b in enumerate(blocks, 1):
        lines = b.splitlines()
        q_lines = []
        opts = []
        ans = None
        for ln in lines:
            ln = ln.strip()
            if not ln:
                continue
            m = OPTION_RE.match(ln)
            if m:
                opts.append((m.group(1).upper(), m.group(2).strip()))
                continue
            am = ANSWER_RE.search(ln)
            if am and (ln.upper().startswith(("ANSWER", "REPONSE", "RÉPONSE")) or "ANSWER" in ln.upper()):
                ans = am.group(1).upper()
                continue
            # any line containing ANSWER: anywhere
            if "ANSWER" in ln.upper() or "REPONSE" in ln.upper() or "RÉPONSE" in ln.upper():
                am2 = ANSWER_RE.search(ln)
                if am2:
                    ans = am2.group(1).upper()
                    continue
            q_lines.append(ln)
        if not opts or ans is None:
            continue
        # normalize options to A-D in order found; pad if needed
        q_text = " ".join(q_lines).strip()
        # clean leading "Question N :" style labels
        q_text = re.sub(r"^Question\s*\d+\s*[:\-]\s*", "", q_text, flags=re.IGNORECASE)
        qs.append({
            "id": f"{tag}{i}",
            "source": tag,
            "num": i,
            "enonce": q_text,
            "options": opts,
            "correct": ans,
        })
    return qs


def shuffle_options_for_question(q, correct_letter):
    original_map = {letter: text for letter, text in q["options"]}
    if correct_letter not in original_map:
        # fall back: first option
        correct_letter = q["options"][0][0]
    correct_text = original_map[correct_letter]
    texts = [text for _, text in q["options"]]
    random.shuffle(texts)
    new_options = [("A", texts[0]), ("B", texts[1]), ("C", texts[2]), ("D", texts[3])]
    new_correct = None
    for letter, text in new_options:
        if text == correct_text:
            new_correct = letter
            break
    if new_correct is None:
        new_correct = "A"
    return new_options, new_correct


def build_variant(selected, q_order_seed, opt_seed):
    random.seed(q_order_seed)
    order = list(range(len(selected)))
    random.shuffle(order)
    variant = []
    correction_map = {}
    for pos, idx in enumerate(order, 1):
        q = selected[idx]
        qcopy = {
            "source_id": q["id"],
            "enonce": q["enonce"],
            "options": [t for t in q["options"]],
            "correct": q["correct"],
        }
        random.seed(opt_seed + idx)
        new_opts, new_correct = shuffle_options_for_question(qcopy, qcopy["correct"])
        qcopy["options"] = new_opts
        qcopy["correct"] = new_correct
        variant.append(qcopy)
        correction_map[pos] = (qcopy["source_id"], new_correct)
    return variant, correction_map


def create_docx(variant_data, label, correction_map, filename, title):
    doc = Document()
    _sec = doc.sections[0]
    _sec.top_margin = Cm(1.0)
    _sec.bottom_margin = Cm(1.0)
    _sec.left_margin = Cm(1.2)
    _sec.right_margin = Cm(1.2)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(9.5)

    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(title)
    r.bold = True; r.font.size = Pt(15)

    st = doc.add_paragraph(); st.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = st.add_run(f"QCM {len(variant_data)} Questions — {label}")
    r.font.size = Pt(12)

    info = doc.add_paragraph(); info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.add_run("Durée : 1 heure   |   Une seule réponse correcte par question").font.size = Pt(9)
    doc.add_paragraph()

    # --- Grille de réponse à remplir par l'étudiant ---
    gh = doc.add_paragraph()
    r = gh.add_run("GRILLE DE RÉPONSE (à remplir par l'étudiant)")
    r.bold = True; r.font.size = Pt(12)
    gh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph()
    sr = sub.add_run("Écrivez la lettre (A, B, C ou D) correspondant à votre réponse dans la case.")
    sr.italic = True; sr.font.size = Pt(9)
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    nq = len(variant_data)
    rows = (nq + 7) // 8 + 1  # 4 pairs per row -> 8 cols, 10 rows for 40
    # use 4 pairs (8 cols) -> 10 rows for up to 40 questions
    pairs = 4
    per_col = (nq + pairs - 1) // pairs
    table_g = doc.add_table(rows=per_col + 1, cols=pairs * 2)
    table_g.style = "Light Grid Accent 1"
    hdr = table_g.rows[0].cells
    for k in range(pairs):
        p = hdr[k * 2].paragraphs[0]; run = p.add_run("N°"); run.bold = True; run.font.size = Pt(8)
        p = hdr[k * 2 + 1].paragraphs[0]; run = p.add_run("Rép."); run.bold = True; run.font.size = Pt(8)

    for i in range(1, nq + 1):
        col = (i - 1) // per_col
        if col >= pairs:
            col = pairs - 1
        row_in_col = (i - 1) % per_col + 1
        cell_num = table_g.rows[row_in_col].cells[col * 2]
        p = cell_num.paragraphs[0]; run = p.add_run(str(i)); run.font.size = Pt(8)
        cell_ans = table_g.rows[row_in_col].cells[col * 2 + 1]
        pa = cell_ans.paragraphs[0]; ra = pa.add_run(""); ra.font.size = Pt(8)
    for row in table_g.rows:
        for k in range(pairs):
            row.cells[k * 2].width = Cm(1.0)
            row.cells[k * 2 + 1].width = Cm(1.4)

    doc.add_paragraph()

    for i, q in enumerate(variant_data, 1):
        p = doc.add_paragraph()
        run = p.add_run(f"Question {i} : {q['enonce']}")
        run.bold = True; run.font.size = Pt(10)
        p.paragraph_format.space_after = Pt(2)
        for letter, text in q["options"]:
            p2 = doc.add_paragraph()
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(0)
            p2.paragraph_format.left_indent = Cm(1)
            run = p2.add_run(f"{letter}) {text}")
            run.font.size = Pt(9.5)

    doc.add_page_break()
    tc = doc.add_paragraph(); tc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = tc.add_run(f"GRILLE DE CORRECTION — {label}")
    r.bold = True; r.font.size = Pt(14)
    note = doc.add_paragraph(); note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    nr = note.add_run("(Confidentiel — Réservé au correcteur)")
    nr.italic = True; nr.font.size = Pt(10)
    doc.add_paragraph()

    table = doc.add_table(rows=nq + 1, cols=2)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    hdr[0].text = "Question"; hdr[1].text = "Réponse"
    for cell in hdr:
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True; r.font.size = Pt(10)
    for pos in range(1, nq + 1):
        src_id, letter = correction_map[pos]
        row = table.rows[pos]
        row.cells[0].text = str(pos)
        row.cells[1].text = letter
    for row in table.rows:
        row.cells[0].width = Cm(3)
        row.cells[1].width = Cm(3)

    filepath = os.path.join(BASE, filename)
    doc.save(filepath)
    print(f"[OK] {filename}")
    return filepath, correction_map


def main():
    # Collect .txt quiz files (exclude this script and outputs)
    txt_files = []
    for name in sorted(os.listdir(BASE)):
        if name.lower().endswith(".txt"):
            full = os.path.join(BASE, name)
            if os.path.isfile(full):
                txt_files.append(name)
    if not txt_files:
        print("No .txt files found in", BASE)
        sys.exit(1)

    print("Source files:")
    questions_source = []
    for idx, fname in enumerate(txt_files, 1):
        tag = f"F{idx}-"
        qs = parse_quiz(read_text(os.path.join(BASE, fname)), tag)
        print(f"  - {fname}: {len(qs)} questions")
        questions_source.extend(qs)

    if not questions_source:
        print("No questions parsed.")
        sys.exit(1)

    folder_name = os.path.basename(BASE.rstrip(os.sep))
    title = f"EXAMEN — {folder_name}"

    n_per = min(QUESTIONS_PER_VARIANT, len(questions_source))
    if n_per < QUESTIONS_PER_VARIANT:
        print(f"Warning: only {len(questions_source)} questions available; variants will have {n_per} questions.")

    all_variants = {}
    for v in range(1, VARIANTS + 1):
        random.seed(100 + v * 17)
        selected = random.sample(questions_source, n_per)
        variant_data, cmap = build_variant(selected, q_order_seed=1000 + v * 42, opt_seed=2000 + v * 100)
        fname = f"Examen_Variante_{v}.docx"
        _, cm = create_docx(variant_data, f"Variante {v}", cmap, fname, title)
        all_variants[v] = (selected, variant_data, cmap)

    # CSV
    csv_path = os.path.join(BASE, "grille_correction.csv")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Variante", "Position", "ID_Source", "Reponse"])
        for v in range(1, VARIANTS + 1):
            _, _, cmap = all_variants[v]
            for pos in range(1, n_per + 1):
                sid, letter = cmap[pos]
                w.writerow([v, pos, sid, letter])
    print(f"[OK] grille_correction.csv")

    # Comparative grid
    source_to_var = {}
    for v in range(1, VARIANTS + 1):
        _, _, cmap = all_variants[v]
        for pos in range(1, n_per + 1):
            sid, letter = cmap[pos]
            source_to_var.setdefault(sid, {})[v] = (pos, letter)

    doc_comp = Document()
    style = doc_comp.styles["Normal"]
    style.font.name = "Calibri"; style.font.size = Pt(9)
    t = doc_comp.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("GRILLE DE CORRECTION COMPARATIVE — Variantes 1 à %d" % VARIANTS)
    r.bold = True; r.font.size = Pt(13)
    note = doc_comp.add_paragraph(); note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    nr = note.add_run("(Confidentiel — Réservé au correcteur)")
    nr.italic = True; nr.font.size = Pt(9)
    doc_comp.add_paragraph()

    headers = ["ID Source"] + [f"V{v} Pos" for v in range(1, VARIANTS + 1)] + [f"V{v} Rép." for v in range(1, VARIANTS + 1)]
    # interleave pos/rep
    headers = ["ID Source"]
    for v in range(1, VARIANTS + 1):
        headers += [f"V{v} Pos", f"V{v} Rép."]
    table = doc_comp.add_table(rows=1 + len(source_to_var), cols=len(headers))
    table.style = "Light Grid Accent 1"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        p = cell.paragraphs[0]; run = p.add_run(h); run.bold = True; run.font.size = Pt(8)

    for row_i, sid in enumerate(sorted(source_to_var.keys()), 1):
        row = table.rows[row_i]
        row.cells[0].text = sid
        for v in range(1, VARIANTS + 1):
            info = source_to_var[sid].get(v)
            if info:
                pos, letter = info
                row.cells[1 + (v - 1) * 2].text = str(pos)
                row.cells[2 + (v - 1) * 2].text = letter
            else:
                row.cells[1 + (v - 1) * 2].text = "-"
                row.cells[2 + (v - 1) * 2].text = "-"

    widths = [Cm(1.8)] + [Cm(1.2), Cm(1.0)] * VARIANTS
    for row in table.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = w

    comp_path = os.path.join(BASE, "Grille_Comparative.docx")
    doc_comp.save(comp_path)
    print(f"[OK] Grille_Comparative.docx")

    print("\n=== GENERATION TERMINEE ===")
    for v in range(1, VARIANTS + 1):
        print(f"Variante {v}: Examen_Variante_{v}.docx")


if __name__ == "__main__":
    main()