#!/usr/bin/env python3
"""
Script de procesamiento principal para la aplicación Cartera San Benito Format.

Responsabilidades:
- Leer un archivo Excel (.xlsx) desde un path de entrada.
- Transformar la hoja de detalle siguiendo reglas de negocio con pandas.
- Aplicar estilos y crear hojas auxiliares con openpyxl.
- Guardar el resultado en la ruta de salida indicada.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Importar los módulos de procesamiento de hojas
from sheets import SHEET_PROCESSORS, PROCESSING_ORDER


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("processor")

BLUE_FILL = PatternFill(start_color="FF17365D", end_color="FF17365D", fill_type="solid")
WHITE_FONT = Font(color="FFFFFFFF")


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


def unmerge_all_cells(sheet) -> None:
    merged_ranges = list(sheet.merged_cells.ranges)
    for merged in merged_ranges:
        top_left_value = sheet.cell(merged.min_row, merged.min_col).value
        sheet.unmerge_cells(str(merged))
        for row in range(merged.min_row, merged.max_row + 1):
            for col in range(merged.min_col, merged.max_col + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value is None:
                    cell.value = top_left_value


def extract_detail_dataframe(input_path: Path) -> pd.DataFrame:
    workbook = load_workbook(input_path)
    if not workbook.sheetnames:
        workbook.close()
        raise ValueError("El archivo no contiene hojas.")

    worksheet = workbook.active
    worksheet.title = "Cartera_CxC_Det_Comple"
    unmerge_all_cells(worksheet)

    max_column = worksheet.max_column
    if max_column == 0:
        workbook.close()
        return pd.DataFrame()

    columns = [get_column_letter(index) for index in range(1, max_column + 1)]
    rows = [
        [cell for cell in row]
        for row in worksheet.iter_rows(min_row=1, max_col=max_column, values_only=True)
    ]
    workbook.close()

    if not rows:
        return pd.DataFrame(columns=columns)

    dataframe = pd.DataFrame(rows, columns=columns)
    dataframe = dataframe.dropna(axis=1, how="all")
    dataframe.columns = [get_column_letter(index + 1) for index in range(dataframe.shape[1])]
    
    # Asegurar mínimo de 7 filas
    if dataframe.shape[0] < 7:
        missing = 7 - dataframe.shape[0]
        padding = pd.DataFrame([[pd.NA] * dataframe.shape[1]] * missing, columns=dataframe.columns)
        dataframe = pd.concat([dataframe, padding], ignore_index=True)
    
    return dataframe


def write_sheet_to_excel(df: pd.DataFrame, writer: pd.ExcelWriter, sheet_name: str) -> None:
    """Escribe un DataFrame a una hoja específica del archivo Excel."""
    export_df = df.where(pd.notna(df), None)
    export_df.to_excel(
        writer,
        sheet_name=sheet_name,
        index=False,
        header=False,
    )


def apply_styles(output_path: Path) -> None:
    """Aplica estilos a las hojas del archivo Excel."""
    workbook = load_workbook(output_path)
    
    # Aplicar estilos solo a la hoja Cartera_CxC_Det_Comple si existe
    if "Cartera_CxC_Det_Comple" in workbook.sheetnames:
        worksheet = workbook["Cartera_CxC_Det_Comple"]
        max_column = worksheet.max_column or 15  # Columna O

        for row_index in (5, 6):
            for column_index in range(1, max_column + 1):
                cell = worksheet.cell(row=row_index, column=column_index)
                cell.fill = BLUE_FILL
                cell.font = WHITE_FONT

        header_cell = worksheet["G1"]
        header_cell.fill = BLUE_FILL
        header_cell.font = WHITE_FONT

    workbook.save(output_path)
    workbook.close()


def process_workbook(input_path: Path, output_path: Path) -> None:
    """Procesa el archivo Excel usando los módulos de hojas específicas."""
    logger.info("Procesando archivo: %s", input_path)
    
    # Extraer datos de la hoja principal
    detail_df = extract_detail_dataframe(input_path)
    
    # Crear el archivo Excel con todas las hojas
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Procesar cada hoja en el orden especificado
        for sheet_name in PROCESSING_ORDER:
            logger.info("Procesando hoja: %s", sheet_name)
            
            if sheet_name in SHEET_PROCESSORS:
                # Obtener la función de procesamiento para esta hoja
                processor_func = SHEET_PROCESSORS[sheet_name]
                
                # Procesar la hoja (usar datos de detalle para la primera hoja)
                if sheet_name == "Cartera_CxC_Det_Comple":
                    processed_df = processor_func(detail_df)
                else:
                    # Para las demás hojas, crear un DataFrame vacío y procesarlo
                    processed_df = processor_func(pd.DataFrame())
                
                # Escribir la hoja al archivo Excel
                write_sheet_to_excel(processed_df, writer, sheet_name)
                logger.info("Hoja %s procesada y escrita", sheet_name)
            else:
                logger.warning("No se encontró procesador para la hoja: %s", sheet_name)
    
    # Aplicar estilos
    apply_styles(output_path)
    logger.info("Archivo procesado y guardado en: %s", output_path)


def main() -> int:
    try:
        input_path, output_path = validate_arguments(sys.argv)
        process_workbook(input_path, output_path)
        return 0
    except Exception as exc:  # pylint: disable=broad-except
        logger.error("Error durante el procesamiento: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
