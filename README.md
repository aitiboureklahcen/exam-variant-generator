# Exam Variant Generator

A small Python tool that turns a folder of `.txt` quiz files into **4 printable exam variants** (`.docx`), each with a student answer sheet and a confidential correction grid. Designed for teachers who need anti-cheating variants (shuffled question order and shuffled answer options) without retyping anything.

## How it works

1. Drop `generate_exam_variants.py` into a folder containing one or more `.txt` quiz files.
2. Run `python generate_exam_variants.py`.
3. The script reads every `.txt` file in that folder, builds a question pool, and generates:
   - `Examen_Variante_1.docx` … `Examen_Variante_4.docx` — 40 questions each (or fewer if the pool is smaller)
   - `grille_correction.csv` — machine-readable answer keys for all variants
   - `Grille_Comparative.docx` — maps each source question to its position/answer in every variant

## Input format

Questions are separated by blank lines. Each question block:

```
Question text (one or more lines)
A. option text
B. option text
C. option text
D. option text
ANSWER: C
```

Supported option labels: `A.`, `A)`, `A-`. Supported answer keys: `ANSWER:`, `REPONSE:`, `RÉPONSE:`. Leading `Question N:` labels are stripped automatically.

## Output layout (print-optimized)

Each variant starts with:
- Title + duration
- A compact **answer sheet** (grille de réponse) for the student to fill in A/B/C/D
- The 40 questions with shuffled order and shuffled options

Ends with a confidential correction grid (question → answer letter).

## Requirements

- Python 3.9+
- `python-docx` (`pip install python-docx`)

## Usage

```bash
cd /path/to/your/quiz/folder
python generate_exam_variants.py
```

## License

MIT
"
