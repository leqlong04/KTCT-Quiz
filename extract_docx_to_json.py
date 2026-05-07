import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from docx import Document


QUESTION_START_RE = re.compile(r"^\s*(?:Câu\s*)?(\d+)[\.\)\:\-]\s*(.+)\s*$", re.IGNORECASE)
ANSWER_START_RE = re.compile(r"^\s*([A-H])[\.\)\:\-]\s*(.+)\s*$", re.IGNORECASE)
CHAPTER_RE = re.compile(r"^\s*(CHƯƠNG|Chương)\s*(\d+)\s*$")
SECTION_RE = re.compile(r"^\s*(KIỂM\s*TRA|Kiểm\s*tra)\s*$")


def _normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _paragraph_text(p) -> str:
    return _normalize_space(p.text or "")


def _paragraph_has_bold_content(p) -> bool:
    """
    True if paragraph contains any bold run with non-whitespace text.
    """
    for run in getattr(p, "runs", []) or []:
        if run is None:
            continue
        txt = (run.text or "").strip()
        if not txt:
            continue
        # run.bold can be None (inherit). Treat only explicit True as bold.
        if run.bold is True:
            return True
    return False


def _split_inline_answers(text: str) -> List[Tuple[str, str]]:
    """
    Handle cases like: "A. ... B. ... C. ... D. ..."
    Returns list of (label, answer_text). If not detected, returns [].
    """
    # Find all label positions.
    matches = list(re.finditer(r"(^|\s)([A-H])[\.\)\:\-]\s+", text))
    if len(matches) < 2:
        return []

    parts: List[Tuple[str, str]] = []
    for i, m in enumerate(matches):
        label = m.group(2).upper()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        chunk = _normalize_space(text[start:end])
        if chunk:
            parts.append((label, chunk))
    # Ensure at least 2 answers to be considered a split.
    return parts if len(parts) >= 2 else []


@dataclass
class PendingAnswer:
    label: str
    text: str
    is_correct: bool


@dataclass
class PendingQuestion:
    qid: str
    text: str
    answers: List[PendingAnswer]
    chapter: str
    uid: str


def extract_questions(docx_path: Path) -> List[Dict[str, Any]]:
    doc = Document(str(docx_path))

    questions: List[PendingQuestion] = []
    current: Optional[PendingQuestion] = None
    auto_label_idx: int = 0
    current_chapter: str = "Chưa phân loại"
    global_idx: int = 0
    chapter_fallback_idx: int = 0

    def flush_current():
        nonlocal current, auto_label_idx
        if current is None:
            return
        # Drop empty answers
        current.answers = [a for a in current.answers if a.text]
        if current.text and current.answers:
            questions.append(current)
        current = None
        auto_label_idx = 0

    for p in doc.paragraphs:
        raw = p.text or ""
        text = _paragraph_text(p)
        if not text:
            continue

        cm = CHAPTER_RE.match(text)
        if cm:
            flush_current()
            current_chapter = f"Chương {cm.group(2)}"
            chapter_fallback_idx = 0
            continue

        sm = SECTION_RE.match(text)
        if sm:
            flush_current()
            current_chapter = "Kiểm tra"
            chapter_fallback_idx = 0
            continue

        qmatch = QUESTION_START_RE.match(text)
        if qmatch:
            flush_current()
            qid = qmatch.group(1)
            qtext = _normalize_space(qmatch.group(2))
            global_idx += 1
            uid = f"{re.sub(r'[^0-9A-Za-z]+', '-', current_chapter.lower()).strip('-')}-{global_idx}"
            current = PendingQuestion(
                qid=str(qid),
                text=qtext,
                answers=[],
                chapter=current_chapter,
                uid=uid,
            )
            continue

        # Fallback question start while we're already parsing a fallback-style section:
        # If a line ends with '?' and we already collected a few answers for the current question,
        # treat this as the next question.
        if current is not None and text.endswith("?") and len(current.answers) >= 3:
            flush_current()
            chapter_fallback_idx += 1
            global_idx += 1
            chap_num = None
            m = re.search(r"(\d+)$", current_chapter)
            if m:
                chap_num = m.group(1)
            qid = f"{chap_num}.{chapter_fallback_idx}" if chap_num else str(chapter_fallback_idx)
            uid = f"{re.sub(r'[^0-9A-Za-z]+', '-', current_chapter.lower()).strip('-')}-{global_idx}"
            current = PendingQuestion(
                qid=str(qid),
                text=text,
                answers=[],
                chapter=current_chapter,
                uid=uid,
            )
            continue

        if current is None:
            # Fallback: some sections (e.g. Chương 6) don't prefix questions with "Câu x:"
            # Heuristic: treat any line ending with '?' as a question.
            if text.endswith("?"):
                flush_current()
                chapter_fallback_idx += 1
                global_idx += 1
                chap_num = None
                m = re.search(r"(\d+)$", current_chapter)
                if m:
                    chap_num = m.group(1)
                qid = f"{chap_num}.{chapter_fallback_idx}" if chap_num else str(chapter_fallback_idx)
                uid = f"{re.sub(r'[^0-9A-Za-z]+', '-', current_chapter.lower()).strip('-')}-{global_idx}"
                current = PendingQuestion(
                    qid=str(qid),
                    text=text,
                    answers=[],
                    chapter=current_chapter,
                    uid=uid,
                )
                continue
            # Ignore leading content until first question
            continue

        # Try inline answers first (multiple answers in one paragraph)
        inline = _split_inline_answers(text)
        if inline:
            # If any run is bold, we can't reliably map bold to a specific inline option,
            # so mark correctness using a conservative heuristic: check if the paragraph is bold
            # AND only one option contains bold markers-like in the text (rare). Otherwise false.
            # (Most docs put each option on its own line; inline is a fallback.)
            p_is_bold = _paragraph_has_bold_content(p)
            for label, ans_text in inline:
                current.answers.append(
                    PendingAnswer(
                        label=label,
                        text=ans_text,
                        is_correct=bool(p_is_bold and len(inline) == 1),
                    )
                )
            continue

        amatch = ANSWER_START_RE.match(text)
        if amatch:
            label = amatch.group(1).upper()
            ans_text = _normalize_space(amatch.group(2))
            is_correct = _paragraph_has_bold_content(p)
            current.answers.append(PendingAnswer(label=label, text=ans_text, is_correct=is_correct))
            continue

        # Unlabeled answers (common in Vietnamese docs): each non-empty paragraph is an option.
        # Heuristic: if we're in a question context, treat each paragraph as a new choice unless it
        # clearly looks like a continuation of the previous line.
        def looks_like_continuation(s: str) -> bool:
            s = s.lstrip()
            if not s:
                return True
            # Continuation markers / punctuation
            if s[0] in ".,;:)]}":
                return True
            # Lowercase start tends to be continuation in Vietnamese text wrapping
            return s[0].islower()

        if current.answers:
            if looks_like_continuation(text):
                current.answers[-1].text = _normalize_space(current.answers[-1].text + " " + text)
                if _paragraph_has_bold_content(p):
                    current.answers[-1].is_correct = True
            else:
                label = chr(ord("A") + (auto_label_idx % 26))
                auto_label_idx += 1
                current.answers.append(
                    PendingAnswer(
                        label=label,
                        text=text,
                        is_correct=_paragraph_has_bold_content(p),
                    )
                )
        else:
            # If no answers yet, assume this is the first unlabeled option (not part of the question stem)
            label = chr(ord("A") + (auto_label_idx % 26))
            auto_label_idx += 1
            current.answers.append(
                PendingAnswer(
                    label=label,
                    text=text,
                    is_correct=_paragraph_has_bold_content(p),
                )
            )

    flush_current()

    # Convert to JSON-friendly dicts and set answerKey when unique
    out: List[Dict[str, Any]] = []
    for q in questions:
        correct_labels = [a.label for a in q.answers if a.is_correct]
        answer_key = correct_labels[0] if len(correct_labels) == 1 else None
        out.append(
            {
                "uid": q.uid,
                "id": q.qid,
                "chapter": q.chapter,
                "question": q.text,
                "choices": [{"label": a.label, "text": a.text, "isCorrect": a.is_correct} for a in q.answers],
                "answerKey": answer_key,
            }
        )
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Extract MCQ questions from a .docx into JSON (bold = correct).")
    ap.add_argument("input", type=str, help="Path to input .docx")
    ap.add_argument(
        "-o",
        "--output",
        type=str,
        default="output.json",
        help="Output JSON path (default: output.json)",
    )
    ap.add_argument("--pretty", action="store_true", help="Pretty-print JSON")
    args = ap.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"Input not found: {input_path}")

    data = extract_questions(input_path)
    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with out_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2 if args.pretty else None)
        if args.pretty:
            f.write("\n")

    print(f"Wrote {len(data)} questions to {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
