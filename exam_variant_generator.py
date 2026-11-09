#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import random
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

from docx import Document
from docx.shared import Cm, Pt

OPTION_RE = re.compile(r"^\s*([A-Da-d])\s*(?:[.)]|-\))\s*(.+?)\s*$")
ANSWER_RE = re.compile(r"^\s*(?:ANSWER|R[ÉE]PONSE)\s*:\s*([A-Da-d])\s*$", re.IGNORECASE)


@dataclass
class Question:
    source_id: str
    question_text: str
    options: Dict[str, str]
    answer: str


@dataclass
class VariantQuestion:
    source_id: str
    question_text: str
    options: List[Tuple[str, str]]
    answer: str


def split_blocks(text: str) -> List[str]:
    return [block.strip() for block in re.split(r"\n\s*\n", text) if block.strip()]


def parse_block(block: str, source_id: str) -> Question:
    question_lines: List[str] = []
    options: Dict[str, str] = {}
    answer: str | None = None

    for raw_line in block.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        option_match = OPTION_RE.match(line)
        if option_match:
            label = option_match.group(1).upper()
            options[label] = option_match.group(2).strip()
            continue

        answer_match = ANSWER_RE.match(line)
        if answer_match:
            answer = answer_match.group(1).upper()
            continue

        question_lines.append(line)

    if len(options) != 4 or set(options.keys()) != {"A", "B", "C", "D"}:
        raise ValueError(f"{source_id}: options A-D are required")
    if not answer or answer not in options:
        raise ValueError(f"{source_id}: ANSWER/REPONSE/RÉPONSE must point to A-D")
    if not question_lines:
        raise ValueError(f"{source_id}: missing question text")

    return Question(source_id=source_id, question_text="\n".join(question_lines), options=options, answer=answer)


def load_question_pool(directory: Path) -> List[Question]:
    txt_files = sorted(directory.glob("*.txt"))
    if not txt_files:
        raise FileNotFoundError("No .txt quiz files were found in the current folder")

    pool: List[Question] = []
    for txt_file in txt_files:
        content = txt_file.read_text(encoding="utf-8")
        for idx, block in enumerate(split_blocks(content), start=1):
            source_id = f"{txt_file.name}#{idx}"
            pool.append(parse_block(block, source_id))
    return pool


def create_variant_questions(pool: List[Question], max_questions: int) -> List[VariantQuestion]:
    count = min(max_questions, len(pool))
    chosen = random.sample(pool, count)
    random.shuffle(chosen)

    variant: List[VariantQuestion] = []
    for question in chosen:
        shuffled = list(question.options.items())
        random.shuffle(shuffled)

        answer_letter = ""
        for idx, (old_letter, _) in enumerate(shuffled):
            new_letter = chr(ord("A") + idx)
            if old_letter == question.answer:
                answer_letter = new_letter

        remapped = [(chr(ord("A") + idx), text) for idx, (_, text) in enumerate(shuffled)]
        variant.append(
            VariantQuestion(
                source_id=question.source_id,
                question_text=question.question_text,
                options=remapped,
                answer=answer_letter,
            )
        )
    return variant


def apply_print_layout(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    for section in doc.sections:
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)


def add_answer_sheet(doc: Document, question_count: int) -> None:
    doc.add_paragraph("Feuille de réponses (à détacher)")
    columns = 4
    rows = (question_count + columns - 1) // columns
    table = doc.add_table(rows=rows + 1, cols=columns * 2)
    table.style = "Table Grid"

    for col in range(columns):
        table.cell(0, col * 2).text = "N°"
        table.cell(0, col * 2 + 1).text = "Réponse"

    for i in range(question_count):
        row = (i % rows) + 1
        col = i // rows
        table.cell(row, col * 2).text = str(i + 1)
        table.cell(row, col * 2 + 1).text = ""


def write_variant_docx(variant_number: int, questions: List[VariantQuestion], output_dir: Path) -> None:
    doc = Document()
    apply_print_layout(doc)

    doc.add_heading(f"Exam Variant {variant_number}", level=0)
    doc.add_paragraph("Durée: 60 minutes")
    add_answer_sheet(doc, len(questions))

    doc.add_page_break()
    for idx, question in enumerate(questions, start=1):
        doc.add_paragraph(f"{idx}. {question.question_text}")
        for letter, option in question.options:
            doc.add_paragraph(f"   {letter}. {option}")

    doc.add_page_break()
    doc.add_paragraph("CONFIDENTIEL — Grille de correction")
    correction_table = doc.add_table(rows=len(questions) + 1, cols=2)
    correction_table.style = "Table Grid"
    correction_table.cell(0, 0).text = "N°"
    correction_table.cell(0, 1).text = "Réponse"

    for idx, question in enumerate(questions, start=1):
        correction_table.cell(idx, 0).text = str(idx)
        correction_table.cell(idx, 1).text = question.answer

    doc.save(output_dir / f"Variant_{variant_number}.docx")


def write_correction_csv(variants: List[List[VariantQuestion]], pool: List[Question], output_dir: Path) -> None:
    rows = []
    for question in pool:
        row = {"source_question": question.source_id}
        for variant_idx, variant_questions in enumerate(variants, start=1):
            position = ""
            answer = ""
            for q_index, variant_question in enumerate(variant_questions, start=1):
                if variant_question.source_id == question.source_id:
                    position = str(q_index)
                    answer = variant_question.answer
                    break
            row[f"variant_{variant_idx}_position"] = position
            row[f"variant_{variant_idx}_answer"] = answer
        rows.append(row)

    fieldnames = ["source_question"]
    for idx in range(1, len(variants) + 1):
        fieldnames.extend([f"variant_{idx}_position", f"variant_{idx}_answer"])

    with (output_dir / "grille_correction.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_comparative_docx(variants: List[List[VariantQuestion]], pool: List[Question], output_dir: Path) -> None:
    doc = Document()
    apply_print_layout(doc)
    doc.add_heading("Grille Comparative", level=0)

    table = doc.add_table(rows=len(pool) + 1, cols=1 + len(variants) * 2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "Question source"

    col = 1
    for idx in range(1, len(variants) + 1):
        table.cell(0, col).text = f"V{idx} Pos"
        table.cell(0, col + 1).text = f"V{idx} Rep"
        col += 2

    for row_idx, source_question in enumerate(pool, start=1):
        table.cell(row_idx, 0).text = source_question.source_id
        col = 1
        for variant_questions in variants:
            position = ""
            answer = ""
            for q_index, variant_question in enumerate(variant_questions, start=1):
                if variant_question.source_id == source_question.source_id:
                    position = str(q_index)
                    answer = variant_question.answer
                    break
            table.cell(row_idx, col).text = position
            table.cell(row_idx, col + 1).text = answer
            col += 2

    doc.save(output_dir / "Grille_Comparative.docx")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate shuffled exam variants from .txt quiz files")
    parser.add_argument("--variants", type=int, default=4, help="Number of variants to generate (default: 4)")
    parser.add_argument(
        "--questions-per-variant",
        type=int,
        default=40,
        help="Maximum number of questions per variant (default: 40)",
    )
    parser.add_argument("--seed", type=int, default=None, help="Optional random seed")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.variants <= 0 or args.questions_per_variant <= 0:
        print("--variants and --questions-per-variant must be positive integers", file=sys.stderr)
        return 1

    if args.seed is not None:
        random.seed(args.seed)

    try:
        pool = load_question_pool(Path.cwd())
    except (FileNotFoundError, ValueError, UnicodeDecodeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    variants = [create_variant_questions(pool, args.questions_per_variant) for _ in range(args.variants)]

    output_dir = Path.cwd()
    for idx, variant_questions in enumerate(variants, start=1):
        write_variant_docx(idx, variant_questions, output_dir)
    write_correction_csv(variants, pool, output_dir)
    write_comparative_docx(variants, pool, output_dir)

    print(f"Generated {len(variants)} variants and correction files in {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
