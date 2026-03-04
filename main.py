#!/usr/bin/env python3
"""
main.py — Point d'entrée du pipeline rfp-respond
=================================================

Usage :
    python main.py --input RFP.pdf
    python main.py --input RFP.pdf --no-responses
    python main.py --input RFP.pdf --output-dir ./results
"""

import os
import asyncio
import argparse
from pathlib import Path
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from loguru import logger

load_dotenv()

from readers import run_all_readers
from matrix import merge, enrich_with_responses
from exporters.excel_export import export as export_excel
from exporters.word_export import export as export_word

console = Console()


async def run(pdf_path: str, output_dir: str, with_responses: bool):
    pdf_name = Path(pdf_path).stem
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    console.print(Panel.fit(
        f"[bold cyan]📋 RFP Respond[/bold cyan]\n"
        f"Fichier : [bold]{pdf_path}[/bold]\n"
        f"Sortie  : {output_dir}",
        border_style="cyan"
    ))

    # Phase 1 — 5 lectures parallèles
    console.print("\n[bold]Phase 1[/bold] — Lectures parallèles...")
    readings = await run_all_readers(pdf_path)

    # Phase 2 — Fusion et déduplication
    console.print("\n[bold]Phase 2[/bold] — Fusion de la matrice...")
    matrix = merge(readings)

    # Phase 3 — Génération des réponses (optionnel)
    if with_responses:
        console.print("\n[bold]Phase 3[/bold] — Génération des réponses...")
        matrix = await enrich_with_responses(matrix)

    # Phase 4 — Exports
    console.print("\n[bold]Phase 4[/bold] — Exports...")
    excel_path = str(out / f"{pdf_name}_matrice.xlsx")
    word_path  = str(out / f"{pdf_name}_reponse.docx")

    export_excel(matrix, excel_path)
    export_word(matrix, word_path, rfp_name=pdf_name)

    console.print(Panel(
        f"[bold green]✅ Terminé[/bold green]\n\n"
        f"📊 Matrice Excel  : [bold]{excel_path}[/bold]\n"
        f"📝 Réponse Word   : [bold]{word_path}[/bold]\n\n"
        f"[dim]{len(matrix)} exigences — "
        f"{sum(1 for x in matrix if x.get('moscow')=='must')} Must | "
        f"{sum(1 for x in matrix if x.get('moscow')=='should')} Should | "
        f"{sum(1 for x in matrix if x.get('source') in ('implicit','predictive'))} Prédictives[/dim]",
        border_style="green"
    ))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Analyse un RFP et génère une matrice + réponse Word",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemples :
  python main.py --input appel_offres.pdf
  python main.py --input appel_offres.pdf --output-dir ./ma_reponse
  python main.py --input appel_offres.pdf --no-responses
        """
    )
    parser.add_argument("--input",       required=True,  help="Chemin vers le PDF")
    parser.add_argument("--output-dir",  default="output", help="Dossier de sortie (défaut: ./output)")
    parser.add_argument("--no-responses", action="store_true",
                        help="Ne pas générer les propositions de réponse (plus rapide)")

    args = parser.parse_args()

    asyncio.run(run(
        pdf_path=args.input,
        output_dir=args.output_dir,
        with_responses=not args.no_responses,
    ))
