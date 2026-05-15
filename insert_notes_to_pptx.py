"""
Script to insert speaker notes into a PowerPoint deck.

Usage:
    python insert_notes_to_pptx.py <notes_file> <pptx_file> [--output output_file.pptx]

Notes file format:
    Notes must be organized with "Slide N" headers (e.g., "Slide 1 — Title").
    The script identifies the slide number in the header and inserts the subsequent
    text as the speaker note for the corresponding slide in the PowerPoint.

Example notes file (notes.txt):
    Slide 1 — Introduction

    This is the note for slide 1.
    It can have multiple lines.

    Slide 3 — Conclusion

    Note for slide 3. Slide 2 will remain without a note.

If --output is not provided, the file is saved with a "_com_notas" suffix.
"""

import argparse
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
except ImportError:
    print("Error: the python-pptx library is required.")
    print("Install with: pip install python-pptx")
    sys.exit(1)


def parse_notes_by_slide(notes_path: str) -> dict[int, str]:
    """
    Reads the notes file and returns a dictionary {slide_number: note_text}.

    Identifies headers in the format "Slide N" (with variations like "Slide N —",
    "Slide N -", "Slide N:") and groups the subsequent text as that slide's note.
    """
    path = Path(notes_path)
    if not path.exists():
        print(f"Error: notes file not found: {notes_path}")
        sys.exit(1)

    content = path.read_text(encoding="utf-8")
    lines = content.splitlines()

    # Regex to identify slide header lines
    # Accepts: "Slide 1", "Slide 1 —", "Slide 1 –", "Slide 1 -", "Slide 1:",
    # "## Slide 1:", "# Slide 1 —", etc.
    slide_header_pattern = re.compile(
        r"^\s*#*\s*Slide\s+(\d+)\s*(?:[—–\-:].*)?\s*$", re.IGNORECASE
    )

    notes_by_slide: dict[int, str] = {}
    current_slide_num: int | None = None
    current_lines: list[str] = []

    for line in lines:
        match = slide_header_pattern.match(line)
        if match:
            # Save the previous block if it exists
            if current_slide_num is not None:
                note_text = "\n".join(current_lines).strip()
                if note_text:
                    notes_by_slide[current_slide_num] = note_text

            # Start a new block
            current_slide_num = int(match.group(1))
            current_lines = []
        else:
            if current_slide_num is not None:
                # Skip separator lines (e.g., "---")
                if line.strip() == "---":
                    continue
                current_lines.append(line)

    # Save the last block
    if current_slide_num is not None:
        note_text = "\n".join(current_lines).strip()
        if note_text:
            notes_by_slide[current_slide_num] = note_text

    return notes_by_slide


def insert_notes(pptx_path: str, notes: dict[int, str], output_path: str):
    """Inserts notes into the corresponding slides of an existing PowerPoint."""
    prs = Presentation(pptx_path)
    slides = list(prs.slides)
    total_slides = len(slides)

    print(f"  Total slides in PowerPoint: {total_slides}")
    print(f"  Total notes found: {len(notes)}")
    print()

    inserted = 0
    skipped = 0

    for slide_num, note_text in sorted(notes.items()):
        if slide_num < 1 or slide_num > total_slides:
            print(f"  ⚠ Slide {slide_num}: does not exist in PowerPoint (only {total_slides} slides). Skipped.")
            skipped += 1
            continue

        slide = slides[slide_num - 1]  # 0-based index
        notes_slide = slide.notes_slide
        text_frame = notes_slide.notes_text_frame
        text_frame.text = note_text
        print(f"  ✓ Slide {slide_num}: note inserted ({len(note_text)} characters)")
        inserted += 1

    print(f"\nSummary: {inserted} notes inserted, {skipped} skipped.")
    prs.save(output_path)
    print(f"File saved: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Insert notes from a text file into the corresponding slides of a PowerPoint."
    )
    parser.add_argument(
        "notes",
        help="Path to the notes file (with 'Slide N' headers)",
    )
    parser.add_argument(
        "pptx",
        help="Path to the existing PowerPoint file (.pptx)",
    )
    parser.add_argument(
        "--output", "-o",
        help="Path for the output file (default: appends '_with_notes' to the name)",
        default=None,
    )

    args = parser.parse_args()

    # Validate that the pptx exists
    pptx_path = Path(args.pptx)
    if not pptx_path.exists():
        print(f"Error: PowerPoint file not found: {args.pptx}")
        sys.exit(1)

    # Define output path
    if args.output:
        output_path = args.output
    else:
        output_path = str(pptx_path.with_stem(pptx_path.stem + "_with_notes"))

    print(f"Reading notes from: {args.notes}")
    notes = parse_notes_by_slide(args.notes)

    if not notes:
        print("No notes found in file. Make sure the format uses 'Slide N' as headers.")
        sys.exit(1)

    slide_nums = sorted(notes.keys())
    print(f"Notes found for slides: {slide_nums[0]} to {slide_nums[-1]}")
    print(f"\nInserting notes into: {args.pptx}")
    insert_notes(args.pptx, notes, output_path)


if __name__ == "__main__":
    main()
