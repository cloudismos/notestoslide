"""
Script para inserir notas de slides em um deck PowerPoint existente.

Uso:
    python insert_notes_to_pptx.py <arquivo_notas> <arquivo_pptx> [--output arquivo_saida.pptx]

Formato do arquivo de notas:
    As notas devem ser organizadas com cabeçalhos "Slide N" (ex: "Slide 1 — Título").
    O script identifica o número do slide no cabeçalho e insere o texto subsequente
    como nota do slide correspondente no PowerPoint.

Exemplo de arquivo de notas (notas.txt):
    Slide 1 — Introdução

    Esta é a nota do slide 1.
    Pode ter múltiplas linhas.

    Slide 3 — Conclusão

    Nota do slide 3. O slide 2 ficará sem nota.

Se --output não for fornecido, o arquivo será salvo com sufixo "_com_notas".
"""

import argparse
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
except ImportError:
    print("Erro: a biblioteca python-pptx é necessária.")
    print("Instale com: pip install python-pptx")
    sys.exit(1)


def parse_notes_by_slide(notes_path: str) -> dict[int, str]:
    """
    Lê o arquivo de notas e retorna um dicionário {número_do_slide: texto_da_nota}.
    
    Identifica cabeçalhos no formato "Slide N" (com variações como "Slide N —", 
    "Slide N -", "Slide N:") e agrupa o texto subsequente como nota daquele slide.
    """
    path = Path(notes_path)
    if not path.exists():
        print(f"Erro: arquivo de notas não encontrado: {notes_path}")
        sys.exit(1)

    content = path.read_text(encoding="utf-8")
    lines = content.splitlines()

    # Regex para identificar linhas de cabeçalho de slide
    # Aceita: "Slide 1", "Slide 1 —", "Slide 1 -", "Slide 1:", etc.
    slide_header_pattern = re.compile(
        r"^\s*Slide\s+(\d+)\s*(?:[—\-:].*)?\s*$", re.IGNORECASE
    )

    notes_by_slide: dict[int, str] = {}
    current_slide_num: int | None = None
    current_lines: list[str] = []

    for line in lines:
        match = slide_header_pattern.match(line)
        if match:
            # Salva o bloco anterior se existir
            if current_slide_num is not None:
                note_text = "\n".join(current_lines).strip()
                if note_text:
                    notes_by_slide[current_slide_num] = note_text

            # Inicia novo bloco
            current_slide_num = int(match.group(1))
            current_lines = []
        else:
            if current_slide_num is not None:
                current_lines.append(line)

    # Salva o último bloco
    if current_slide_num is not None:
        note_text = "\n".join(current_lines).strip()
        if note_text:
            notes_by_slide[current_slide_num] = note_text

    return notes_by_slide


def insert_notes(pptx_path: str, notes: dict[int, str], output_path: str):
    """Insere notas nos slides correspondentes de um PowerPoint existente."""
    prs = Presentation(pptx_path)
    slides = list(prs.slides)
    total_slides = len(slides)

    print(f"  Total de slides no PowerPoint: {total_slides}")
    print(f"  Total de notas encontradas: {len(notes)}")
    print()

    inserted = 0
    skipped = 0

    for slide_num, note_text in sorted(notes.items()):
        if slide_num < 1 or slide_num > total_slides:
            print(f"  ⚠ Slide {slide_num}: não existe no PowerPoint (tem apenas {total_slides} slides). Ignorado.")
            skipped += 1
            continue

        slide = slides[slide_num - 1]  # Índice 0-based
        notes_slide = slide.notes_slide
        text_frame = notes_slide.notes_text_frame
        text_frame.text = note_text
        print(f"  ✓ Slide {slide_num}: nota inserida ({len(note_text)} caracteres)")
        inserted += 1

    print(f"\nResumo: {inserted} notas inseridas, {skipped} ignoradas.")
    prs.save(output_path)
    print(f"Arquivo salvo: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Insere notas de um arquivo texto nos slides correspondentes de um PowerPoint."
    )
    parser.add_argument(
        "notas",
        help="Caminho para o arquivo de notas (com cabeçalhos 'Slide N')",
    )
    parser.add_argument(
        "pptx",
        help="Caminho para o arquivo PowerPoint existente (.pptx)",
    )
    parser.add_argument(
        "--output", "-o",
        help="Caminho para o arquivo de saída (padrão: adiciona '_com_notas' ao nome)",
        default=None,
    )

    args = parser.parse_args()

    # Valida que o pptx existe
    pptx_path = Path(args.pptx)
    if not pptx_path.exists():
        print(f"Erro: arquivo PowerPoint não encontrado: {args.pptx}")
        sys.exit(1)

    # Define output path
    if args.output:
        output_path = args.output
    else:
        output_path = str(pptx_path.with_stem(pptx_path.stem + "_com_notas"))

    print(f"Lendo notas de: {args.notas}")
    notes = parse_notes_by_slide(args.notas)

    if not notes:
        print("Nenhuma nota encontrada no arquivo. Verifique se o formato usa 'Slide N' como cabeçalho.")
        sys.exit(1)

    slide_nums = sorted(notes.keys())
    print(f"Notas encontradas para slides: {slide_nums[0]} a {slide_nums[-1]}")
    print(f"\nInserindo notas em: {args.pptx}")
    insert_notes(args.pptx, notes, output_path)


if __name__ == "__main__":
    main()
