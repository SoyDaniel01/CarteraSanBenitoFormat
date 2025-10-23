#!/usr/bin/env python3
"""
Módulo de procesamiento para la hoja Juridico.
Contiene la lógica específica para procesar esta hoja.
"""

import logging
import pandas as pd

logger = logging.getLogger("processor.juridico")

TARGET_SHEET_NAME = "Juridico"


def process_juridico_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """
    Procesa la hoja Juridico con transformaciones específicas.
    Prueba rápida: asigna valores específicos a celdas A1, B2, C3.
    """
    logger.info("Iniciando procesamiento de hoja Juridico")
    
    # Crear un DataFrame con las celdas específicas asignadas
    # Asegurar que tenga al menos 3 filas y 3 columnas
    if df.empty:
        df = pd.DataFrame(index=range(3), columns=['A', 'B', 'C'])
    
    # Asignar valores específicos según la prueba
    df.at[0, 'A'] = "1"  # Celda A1
    df.at[1, 'B'] = "2"  # Celda B2  
    df.at[2, 'C'] = "3"  # Celda C3
    
    logger.info("Valores asignados: A1=1, B2=2, C3=3")
    logger.info("Procesamiento de hoja Juridico completado")
    return df


def get_sheet_name() -> str:
    """Retorna el nombre de la hoja que procesa este módulo."""
    return TARGET_SHEET_NAME
