#!/usr/bin/env python3
"""
Módulo de procesamiento para la hoja Cartera_CxC_Res_Comp.
Contiene la lógica específica para procesar esta hoja.
"""

import logging
import pandas as pd

logger = logging.getLogger("processor.res_comp")

TARGET_SHEET_NAME = "Cartera_CxC_Res_Comp"


def process_res_comp_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """
    Procesa la hoja Cartera_CxC_Res_Comp con transformaciones específicas.
    Prueba rápida: asigna valores específicos a celdas A1, B2, C3.
    """
    logger.info("Iniciando procesamiento de hoja Cartera_CxC_Res_Comp")
    
    # Crear un DataFrame con las celdas específicas asignadas
    # Asegurar que tenga al menos 3 filas y 3 columnas
    if df.empty:
        df = pd.DataFrame(index=range(3), columns=['A', 'B', 'C'])
    
    # Asignar valores específicos según la prueba
    df.at[0, 'A'] = "a"  # Celda A1
    df.at[1, 'B'] = "b"  # Celda B2  
    df.at[2, 'C'] = "c"  # Celda C3
    
    logger.info("Valores asignados: A1=a, B2=b, C3=c")
    logger.info("Procesamiento de hoja Cartera_CxC_Res_Comp completado")
    return df


def get_sheet_name() -> str:
    """Retorna el nombre de la hoja que procesa este módulo."""
    return TARGET_SHEET_NAME
