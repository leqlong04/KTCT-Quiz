# Extract .docx (bold = correct) -> JSON

Script: `extract_docx_to_json.py`

## Install

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
python extract_docx_to_json.py "Cải kinh tế chính trị.docx" -o "Cải kinh tế chính trị.json" --pretty
```

## Output format

Each item:

- `id`: question number in the doc
- `question`: question text
- `choices`: array of `{ label, text, isCorrect }`
- `answerKey`: label of the correct choice if exactly 1 correct answer is detected, otherwise `null`

