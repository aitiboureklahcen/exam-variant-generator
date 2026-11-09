# exam-variant-generator

Python CLI tool that reads all `.txt` quiz files in the current directory and generates randomized printable exam variants.

## Features

- Parses quiz blocks separated by blank lines.
- Supports options labeled `A.`, `A)`, or `A-)`.
- Supports answer lines with `ANSWER:`, `REPONSE:`, or `RÉPONSE:`.
- Combines all parsed questions into one pool.
- Generates 4 variants by default, each with up to 40 questions.
- Shuffles question order and option order while preserving the correct answer.
- Produces:
  - `Variant_1.docx` ... `Variant_4.docx`
  - `grille_correction.csv`
  - `Grille_Comparative.docx`

## Installation

```bash
python -m pip install .
```

## Usage

From a folder containing your `.txt` quiz files:

```bash
exam-variant-generator
```

Optional arguments:

```bash
exam-variant-generator --variants 4 --questions-per-variant 40 --seed 123
```

## Quiz block format

Each question block is separated by a blank line and should contain:

1. Question text (one or more lines)
2. Four options (`A` to `D`)
3. Answer line

Example:

```text
What is 2 + 2?
A. 3
B) 4
C-) 5
D. 6
ANSWER: B
```
