#!/usr/bin/env python3
"""
Script de procesamiento temporal para la aplicación Cartera San Benito Format.

Responsabilidades:
- Leer un archivo Excel (.xlsx) desde un path de entrada.
- Guardar una copia en el path de salida.
- Escribir valores de prueba en celdas específicas (A1 = 1, B1 = 2, C3 = 3).

Este flujo es únicamente de validación y debe reemplazarse por el proceso real.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("processor")


def validate_arguments(args: list[str]) -> tuple[Path, Path]:
    if len(args) < 3:
        raise ValueError("Uso: processor.py <input_path> <output_path>")

    input_path = Path(args[1]).expanduser().resolve()
    output_path = Path(args[2]).expanduser().resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de entrada: {input_path}")

    if input_path.suffix.lower() != ".xlsx":
        raise ValueError("El archivo de entrada debe tener extensión .xlsx")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    return input_path, output_path


def copy_and_annotate_excel(input_path: Path, output_path: Path) -> None:
    logger.info("Leyendo archivo fuente: %s", input_path)
    df = pd.read_excel(input_path)

    logger.info("Escribiendo archivo temporal en: %s", output_path)
    df.to_excel(output_path, index=False)

    workbook = load_workbook(output_path)
    worksheet = workbook.active

    logger.info("Escribiendo valores de prueba en el archivo de salida.")
    worksheet["A1"] = 1
    worksheet["B1"] = 2
    worksheet["C3"] = 3

    workbook.save(output_path)
    logger.info("Archivo procesado correctamente: %s", output_path)


def main() -> int:
    try:
        input_path, output_path = validate_arguments(sys.argv)
        copy_and_annotate_excel(input_path, output_path)
        return 0
    except Exception as exc:  # pylint: disable=broad-except
        logger.error("Error durante el procesamiento: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
