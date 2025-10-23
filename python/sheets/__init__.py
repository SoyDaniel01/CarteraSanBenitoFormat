"""
Paquete de módulos para el procesamiento de hojas Excel.
Cada hoja tiene su propio módulo con lógica específica.
"""

from .det_comple import process_det_comple_sheet, get_sheet_name as get_det_comple_name
from .res_comp import process_res_comp_sheet, get_sheet_name as get_res_comp_name
from .dse import process_dse_sheet, get_sheet_name as get_dse_name
from .rse import process_rse_sheet, get_sheet_name as get_rse_name
from .especiales import process_especiales_sheet, get_sheet_name as get_especiales_name
from .juridico import process_juridico_sheet, get_sheet_name as get_juridico_name
from .proyeccion import process_proyeccion_sheet, get_sheet_name as get_proyeccion_name
from .recuperacion import process_recuperacion_sheet, get_sheet_name as get_recuperacion_name

# Mapeo de nombres de hojas a sus funciones de procesamiento
SHEET_PROCESSORS = {
    "Cartera_CxC_Det_Comple": process_det_comple_sheet,
    "Cartera_CxC_Res_Comp": process_res_comp_sheet,
    "Cartera_CxC_DSE": process_dse_sheet,
    "Cartera_CxC_RSE": process_rse_sheet,
    "Especiales": process_especiales_sheet,
    "Juridico": process_juridico_sheet,
    "Proyeccion": process_proyeccion_sheet,
    "Recuperacion": process_recuperacion_sheet,
}

# Orden de procesamiento de las hojas
PROCESSING_ORDER = [
    "Cartera_CxC_Det_Comple",
    "Cartera_CxC_Res_Comp", 
    "Cartera_CxC_DSE",
    "Cartera_CxC_RSE",
    "Especiales",
    "Juridico",
    "Proyeccion",
    "Recuperacion"
]
