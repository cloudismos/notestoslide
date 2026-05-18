# PowerPoint Notes Inserter

Python script to insert speaker notes into an existing PowerPoint (.pptx) file, automatically mapping each note to the correct slide based on the slide number in the text file.

## How it works

The script reads a `.txt` file with notes organized under `Slide N` headers and inserts each text block as the speaker note for the corresponding slide in the `.pptx`. The original PowerPoint is not modified — a new file is generated.

## Installation

```bash
pip install python-pptx
```

## Usage

```bash
python insert_notes_to_pptx.py <notes_file.txt> <file.pptx> [--output output.pptx]
```

### Examples

```bash
# Automatically generates "deck_com_notas.pptx"
python insert_notes_to_pptx.py notes.txt deck.pptx

# Specify the output file name
python insert_notes_to_pptx.py notes.txt deck.pptx --output result.pptx
```

## Notes file format

The text file must have headers in the format `Slide N` (accepts variations with `—`, `-`, or `:`). All text between one header and the next is inserted as that slide's speaker note.

```text
Slide 1 — Introduction

Good morning everyone. In this slide we'll introduce the main topic
of the course and the learning objectives.

Slide 2 — Agenda

Here we have the module agenda divided into three blocks.
The first block covers X, the second Y, and the third Z.

Slide 5 — Conclusion

Note for slide 5. Slides without a matching header remain unchanged.
```

### Rules

- The slide number in the text must correspond to the slide in the PowerPoint
- Slides without a note in the `.txt` remain unchanged
- If a slide number doesn't exist in the PowerPoint, the note is skipped with a warning
- The original file is never modified

## Requirements

- Python 3.10+
- [python-pptx](https://python-pptx.readthedocs.io/)
