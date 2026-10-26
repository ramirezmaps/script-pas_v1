# -*- coding: utf-8 -*-
"""
================================================================================
CAJA DE HERRAMIENTAS PYTHON (.pyt) AUTÓNOMA (SINGLE-FILE) PARA ARCGIS PRO
Auditor Cartográfico y Espacial PAS 151 - Protocolo CONAF (Rev 11)
================================================================================
Este archivo .pyt es 100% autónomo: NO requiere carpetas adicionales para
funcionar en ArcGIS Pro. Incluye internamente las 43 plantillas de capas CONAF,
los 31 dominios oficiales, el lector binario DBF/SHP/PRJ, las validaciones de
atributos y las 7 validaciones espaciales e inter-capas.

¿CÓMO ADICIONAR O PERSONALIZAR UN NUEVO ESTÁNDAR DE DATOS?
Opción A (Sin tocar código - vía Excel en ArcGIS Pro):
  1. Ejecute la herramienta '2. Exportar Plantilla Oficial CONAF PAS 151 a Excel (.xlsx)'.
  2. Abra el Excel generado y agregue nuevas capas (columna 'shp'), nuevos campos
     ('nombre', 'tipo', 'longitud', 'decimales') o nuevos valores permitidos
     separados por coma en la columna 'dominio_conaf' (use 'NA' cuando no aplique).
  3. En la herramienta '1. Revisión Integral...', seleccione su archivo Excel y elija si
     desea FUSIONARLO con la plantilla CONAF o USAR SOLO SU EXCEL PERSONALIZADO.

Opción B (En el código de este archivo .pyt):
  Edite los diccionarios `DOMINIOS_CONAF` o `_RAW_TEMPLATES` en la SECCIÓN 1.
================================================================================
"""

from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime
import math
import os
from pathlib import Path
import re
import struct
import sys
import unicodedata
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple, Union

import arcpy
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ==============================================================================
# SECCIÓN 1: PLANTILLAS CONAF PAS 151 Y DOMINIOS OFICIALES
# ==============================================================================
"""
Plantillas oficiales de capas, atributos y dominios del Protocolo CONAF para PAS 151 (Rev 11),
más soporte para carga dinámica desde archivo Excel (Diccionario_PAS151.xlsx) respetando la Regla NA.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
import unicodedata


@dataclass
class FieldSpec:
    """Especificación de un campo según la plantilla CONAF PAS 151."""
    nombre: str
    tipo: str  # 'character', 'integer', 'numeric'
    longitud: Optional[int] = None   # None equivale a NA en el script R
    decimales: Optional[int] = None  # None equivale a NA en el script R
    orden: int = 0


@dataclass
class LayerTemplate:
    """Plantilla estructural de un shapefile del protocolo CONAF."""
    prefijo: str
    campos: List[FieldSpec] = field(default_factory=list)
    origen: str = "PLANTILLA_R_REV11"

    @property
    def nombres_campos(self) -> List[str]:
        return [c.nombre for c in self.campos]

    def obtener_campo(self, nombre: str) -> Optional[FieldSpec]:
        nombre_up = nombre.strip().upper()
        for c in self.campos:
            if c.nombre.strip().upper() == nombre_up:
                return c
        return None


def normalize_layer_key(text: str) -> str:
    """
    Normaliza nombres de prefijos de capas para tolerar diferencias de codificación
    (ej. 'Señaletica_velocidad' vs 'Sealetica_velocidad' o 'Senaletica_velocidad').
    """
    if not text:
        return ""
    s = str(text).strip().replace("\ufffd", "")
    # Reemplazar ñ/Ñ por n/N antes de quitar diacríticos
    s = s.replace("ñ", "n").replace("Ñ", "N")
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    # También normalizar caso donde 'ñ' se perdió como caracter inválido ('Sealetica')
    s_up = s.upper()
    if s_up.startswith("SEALETICA"):
        s_up = "SENALETICA" + s_up[len("SEALETICA"):]
    return s_up


# Definición exacta de las plantillas de Revision_Campos_PAS151_Rev11 2.R + Diccionario_PAS151.xlsx
RAW_TEMPLATES_PAS151: Dict[str, List[Tuple[str, str, Optional[int], Optional[int]]]] = {
    "Area": [
        ("Nom_Predio", "character", None, None),
        ("N_Area", "character", 5, None),
        ("Tipo_Bos", "character", 10, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Limite_Predial": [
        ("Nom_Predio", "character", None, None),
        ("Rol", "character", 50, None),
        ("Comuna", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Suelos": [
        ("Nom_Predio", "character", None, None),
        ("Cat_Erosio", "character", 2, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Rangos_pend": [
        ("Nom_Predio", "character", None, None),
        ("Ran_Pend", "character", 10, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Infra_aislada": [
        ("Nom_Predio", "character", None, None),
        ("Nom_Infra", "character", 50, None),
        ("Coord_X", "integer", 6, 0),
        ("Coord_Y", "integer", 7, 0),
        ("Fuente", "character", None, None),
    ],
    "Parcela": [
        ("Nom_Predio", "character", None, None),
        ("N_Rodal", "character", 5, 0),
        ("N_Parc", "character", 5, 0),
        ("Coord_X", "integer", 6, 0),
        ("Coord_Y", "integer", 7, 0),
        ("Fuente", "character", None, None),
    ],
    "Referencia": [
        ("Nom_Predio", "character", None, None),
        ("Nom_Pto", "character", 100, None),
        ("Coord_X", "integer", 6, 0),
        ("Coord_Y", "integer", 7, 0),
        ("Fuente", "character", None, None),
    ],
    "Caminos": [
        ("Nom_Predio", "character", None, None),
        ("Tipo_Cam", "character", 2, None),
        ("Long_Cam", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Curvas_niv": [
        ("Nom_Predio", "character", None, None),
        ("Cot_Curva", "integer", 9, None),
        ("Long_Curva", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Hidrografia": [
        ("Nom_Predio", "character", None, None),
        ("Tip_Dren", "character", 2, None),
        ("Tipo_Perma", "character", 2, None),
        ("Long_Dren", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Rodales": [
        ("Nom_Predio", "character", None, None),
        ("N_Rodal", "character", 5, None),
        ("Tipo_Bos", "character", 10, None),
        ("Tipo_For", "character", 10, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Predios_vecinos": [
        ("Nom_Predio", "character", None, None),
        ("Rol", "character", 50, None),
        ("Comuna", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "LTE_HVDC": [
        ("Nom_Predio", "character", None, None),
        ("Obra", "character", 80, None),
        ("Nom_obra", "character", 80, None),
        ("Obra_detal", "character", 80, None),
        ("Tipo", "character", 80, None),
        ("Tramo", "character", 80, None),
        ("Long_LTE", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Sendero_peatonal": [
        ("Nom_Predio", "character", None, None),
        ("Nom_Send", "character", 50, None),
        ("Long_Send", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Area_Proyecto": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 80, None),
        ("Obra_detal", "character", 80, None),
        ("Tipo", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Area_proteccion": [
        ("Nom_Predio", "character", None, None),
        ("Tipo", "character", 100, None),
        ("Nombre", "character", 100, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Limite_Comunal": [
        ("Region", "character", 50, None),
        ("Provincia", "character", 50, None),
        ("Comuna", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Servidumbre": [
        ("Nom_Predio", "character", None, None),
        ("Rol", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    # MEDIDAS FAUNA
    "Areas_receptoras_ruido": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 50, None),
        ("Obra_detal", "character", 50, None),
        ("Ambiente", "character", 50, None),
        ("Receptor", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Perturbacion_ruido": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 50, None),
        ("Obra_detal", "character", 50, None),
        ("Ambiente", "character", 50, None),
        ("Receptor", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Perturbacion_baja_mov": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 50, None),
        ("Obra_detal", "character", 50, None),
        ("Perturbaci", "character", 50, None),
        ("Ambiente", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Receptoras_baja_mov": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 50, None),
        ("Obra_detal", "character", 50, None),
        ("Perturbaci", "character", 50, None),
        ("Ambiente", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Perturbacion_curureras": [
        ("Nom_Predio", "character", None, None),
        ("Obra_detal", "character", 50, None),
        ("Ambiente", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Receptoras_curureras": [
        ("Nom_Predio", "character", None, None),
        ("Obra_detal", "character", 50, None),
        ("Ambiente", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Perchas": [
        ("Nom_Predio", "character", None, None),
        ("Medida", "character", 50, None),
        ("Coord_X", "integer", 6, 0),
        ("Coord_Y", "integer", 7, 0),
        ("Fuente", "character", None, None),
    ],
    "Señaletica_velocidad": [
        ("Nom_Predio", "character", None, None),
        ("Tipo_inst", "character", 50, None),
        ("Coord_X", "integer", 6, 0),
        ("Coord_Y", "integer", 7, 0),
        ("Fuente", "character", None, None),
    ],
    # MEDIDAS PLANTAS SUELO
    "Rehabilitacion": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 50, None),
        ("Obra_detal", "character", 50, None),
        ("Medida", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Surcos": [
        ("Nom_Predio", "character", None, None),
        ("Medida", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Rehabilitacion_pendiente_30": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 50, None),
        ("Obra_detal", "character", 50, None),
        ("Medida", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Recoleccion_semillas": [
        ("Nom_Predio", "character", None, None),
        ("Medida", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Sectores_recoleccion_semillas": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 10, None),
        ("Obra_detal", "character", 10, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Sectores_siembra_directa": [
        ("Nom_Predio", "character", None, None),
        ("Tipo_Veg", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    # MEDIDAS PLANTAS
    "Areas_relocalizacion": [
        ("Nom_Predio", "character", None, None),
        ("Medida", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Areas_rescate": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 50, None),
        ("Obra_detal", "character", 50, None),
        ("Especie", "character", 50, None),
        ("Medida", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Reposicion": [
        ("Nom_Predio", "character", None, None),
        ("Medida", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    # REVEGETACION
    "Revegetacion": [
        ("Nom_Predio", "character", None, None),
        ("Nom_Sitio", "character", None, None),
        ("Region", "character", 50, None),
        ("Provincia", "character", 50, None),
        ("Comuna", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Area_FX_singular": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 50, None),
        ("Obra_detal", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    # ENRIQUECIMIENTO
    "Enriquecimiento": [
        ("Nom_Predio", "character", None, None),
        ("Nom_Sitio", "character", None, None),
        ("Region", "character", 50, None),
        ("Provincia", "character", 50, None),
        ("Comuna", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    "Areas_ECC": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", None, None),
        ("Obra_detal", "character", 50, None),
        ("Especie", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
    # MEDIDAS ADICIONALES
    "Infra_incendio": [
        ("Nom_Predio", "character", None, None),
        ("Nom_Infra", "character", 50, None),
        ("Coord_X", "integer", 6, 0),
        ("Coord_Y", "integer", 7, 0),
        ("Fuente", "character", None, None),
    ],
    "Obras_conservacion": [
        ("Nom_Predio", "character", None, None),
        ("Nom_Infra", "character", 50, None),
        ("Coord_X", "integer", 6, 0),
        ("Coord_Y", "integer", 7, 0),
        ("Fuente", "character", None, None),
    ],
    "Badenes": [
        ("Nom_Predio", "character", None, None),
        ("Nom_Infra", "character", 50, None),
        ("Coord_X", "integer", 6, 0),
        ("Coord_Y", "integer", 7, 0),
        ("Fuente", "character", None, None),
    ],
    "Pendientes_caminos": [
        ("Nom_Predio", "character", None, None),
        ("Nom_obra", "character", 80, None),
        ("Obra_detal", "character", 80, None),
        ("Ran_Pend", "character", 50, None),
        ("Medida", "character", 50, None),
        ("Sup_ha", "numeric", 11, 2),
        ("Fuente", "character", None, None),
    ],
}

# Dominios del protocolo CONAF definidos en Revision_Campos_PAS151_Rev11 2.R
DOMINIOS_CONAF: Dict[str, List[str]] = {
    "Tipo_Bos": ["No aplica", "Bosque nativo", "Plantaciones"],
    "Tipo_For": ["No aplica"] + [str(i) for i in range(1, 13)],
    "Tipo_Cam": ["1", "2", "3", "4"],
    "Tip_Dren": ["1", "2", "3", "4", "5"],
    "Tipo_Perma": ["1", "2"],
    "Cat_Erosio": ["1", "2", "3", "4"],
    "Ran_Pend": ["0% - 10%", "10% - 30%", "30% - 45%", "45% - 60%", "60% y mas"],
}

# Categorías de Infra_aislada que no deben ubicarse sobre Area_Proyecto (líneas 817-825 del R)
CATEGORIAS_INFRA_EROSION: List[str] = [
    "Control de erosion",
    "Sacos de tierra",
    "Bordo de piedra",
    "Zanja infiltracion",
    "Empalizada",
    "Fajina",
    "Microterraza",
]

# Capas cuya contención dentro de Limite_Predial y Area_Proyecto se valida (línea 737 del R + Descripcion_general.docx)
CAPAS_CONTENCION_PREDIO: List[str] = [
    "Area",
    "Suelos",
    "Rangos_pend",
    "Pendientes_caminos",
]

# Capas excluidas de la validación transversal de Nom_Predio respecto de Limite_Predial
CAPAS_EXCLUIDAS_NOM_PREDIO: List[str] = [
    "Predios_vecinos",
    "Limite_Comunal",
]


def _parse_optional_int(val: Any) -> Optional[int]:
    """Aplica la Regla NA para longitud y decimales."""
    if val is None:
        return None
    s = str(val).strip()
    if not s or s.upper() in ("NA", "N/A", "NONE", "NULL", "NAN"):
        return None
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return None


def get_builtin_templates() -> Dict[str, LayerTemplate]:
    """Retorna las plantillas oficiales PAS 151 Rev11 indexadas por clave normalizada."""
    result: Dict[str, LayerTemplate] = {}
    for prefijo, fields_raw in RAW_TEMPLATES_PAS151.items():
        campos = [
            FieldSpec(
                nombre=f[0],
                tipo=f[1].lower(),
                longitud=f[2],
                decimales=f[3],
                orden=idx + 1,
            )
            for idx, f in enumerate(fields_raw)
        ]
        key = normalize_layer_key(prefijo)
        result[key] = LayerTemplate(prefijo=prefijo, campos=campos, origen="PLANTILLA_R_REV11")
    return result


def load_templates(
    excel_path: Optional[Union[str, Path]] = None,
    solo_excel: bool = False,
) -> Dict[str, LayerTemplate]:
    """
    Carga las plantillas oficiales PAS 151 Rev11 y, si se provee un archivo Excel
    válido (ej. Diccionario_PAS151.xlsx), actualiza/agrega las definiciones desde el Excel
    (o usa únicamente las del Excel si `solo_excel=True`). También incorpora dominios
    definidos en la columna opcional `dominio_conaf` o `dominio`.
    """
    templates = {} if (excel_path and solo_excel) else get_builtin_templates()
    if not excel_path:
        return templates

    path_obj = Path(excel_path)
    if not path_obj.exists():
        return get_builtin_templates()

    import openpyxl

    wb = openpyxl.load_workbook(path_obj, data_only=True)
    target_ws = wb.active
    required = {"shp", "nombre", "tipo", "longitud", "decimales"}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if first_row:
            cols = {str(c).strip().lower() for c in first_row if c is not None}
            if required.issubset(cols):
                target_ws = ws
                break

    rows = list(target_ws.iter_rows(values_only=True))
    if not rows:
        return templates or get_builtin_templates()

    headers = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
    try:
        idx_shp = headers.index("shp")
        idx_nom = headers.index("nombre")
        idx_tip = headers.index("tipo")
        idx_lon = headers.index("longitud")
        idx_dec = headers.index("decimales")
    except ValueError:
        return templates or get_builtin_templates()

    idx_dom: Optional[int] = None
    for cand in ("dominio_conaf", "dominio", "dominios"):
        if cand in headers:
            idx_dom = headers.index(cand)
            break

    excel_grouped: Dict[str, Tuple[str, List[FieldSpec]]] = {}
    for row in rows[1:]:
        if not row or len(row) <= max(idx_shp, idx_nom, idx_tip):
            continue
        raw_shp = row[idx_shp]
        raw_nom = row[idx_nom]
        raw_tip = row[idx_tip]
        if raw_shp is None or raw_nom is None or raw_tip is None:
            continue

        shp_str = str(raw_shp).strip()
        nom_str = str(raw_nom).strip()
        tip_str = str(raw_tip).strip().lower()
        if not shp_str or not nom_str:
            continue

        # Restaurar 'Señaletica_velocidad' si el Excel tenía caracter corrupto
        if normalize_layer_key(shp_str) == "SENALETICA_VELOCIDAD":
            shp_str = "Señaletica_velocidad"

        key = normalize_layer_key(shp_str)
        lon_val = _parse_optional_int(row[idx_lon] if idx_lon < len(row) else None)
        dec_val = _parse_optional_int(row[idx_dec] if idx_dec < len(row) else None)

        # Si el usuario definió valores de dominio en la columna opcional de Excel, registrarlos
        if idx_dom is not None and idx_dom < len(row) and row[idx_dom] is not None:
            raw_dom_txt = str(row[idx_dom]).strip()
            if raw_dom_txt and raw_dom_txt.upper() != "NA":
                parts = [p.strip() for p in raw_dom_txt.split(",") if p.strip()]
                if parts:
                    if tip_str in ("integer", "int", "long", "short"):
                        parsed_dom: List[Union[str, int]] = []
                        for p in parts:
                            try:
                                parsed_dom.append(int(float(p)))
                            except ValueError:
                                parsed_dom.append(p)
                        DOMINIOS_CONAF[nom_str] = parsed_dom
                    else:
                        DOMINIOS_CONAF[nom_str] = parts

        if key not in excel_grouped:
            excel_grouped[key] = (shp_str, [])
        orden = len(excel_grouped[key][1]) + 1
        excel_grouped[key][1].append(
            FieldSpec(
                nombre=nom_str,
                tipo=tip_str,
                longitud=lon_val,
                decimales=dec_val,
                orden=orden,
            )
        )

    for key, (prefijo_str, campos_list) in excel_grouped.items():
        templates[key] = LayerTemplate(
            prefijo=prefijo_str,
            campos=campos_list,
            origen=f"EXCEL ({path_obj.name})",
        )

    return templates


# ==============================================================================
# SECCIÓN 2: LECTOR HÍBRIDO DE SHAPEFILES (DBF/SHP/PRJ + SHAPELY/ARCPY)
# ==============================================================================
"""
Lector híbrido de Shapefiles (Encabezados binarios DBF/SHP/PRJ + Geometrías Shapely / ArcPy).
Reproduce con exactitud la inspección de `ogrinfo -so -al` y `sf::st_read()` del script R:
- Ancho (length) y decimales (scale) reales desde el descriptor binario dBase (.dbf).
- Detección de dimensiones Z/M (3D/Medida) desde el header binario (.shp) y ArcPy.
- Detección de EPSG (incluyendo reconocimiento directo de WGS_1984_UTM_Zone_19S -> EPSG:32719).
- Lectura de todas las filas de atributos y geometrías para validaciones de contenido y espaciales.
"""

from dataclasses import dataclass, field
from pathlib import Path
import struct
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class DbfFieldInfo:
    """Metadatos exactos de un campo en el archivo .dbf (equivalente a ogrinfo + sf)."""
    nombre: str
    tipo_dbf: str       # 'C', 'N', 'F', 'I', 'D', 'L'
    tipo_ogr: str       # 'String', 'Integer', 'Real', 'Date'
    tipo_r_sf: str      # 'character', 'integer', 'numeric' (como lo clasifica sf/R)
    ancho: int
    decimales: int
    orden: int


@dataclass
class ShapefileData:
    """Contenedor completo de un Shapefile leído (metadatos + registros + geometrías)."""
    filepath: Path
    nombre_shape: str                # Nombre base sin .shp
    prefijo_asociado: Optional[str] = None
    tipo_coincidencia: str = "NO_MATCH"  # EXACT_PREDIO, EXACT_PREFIX, AUTO_PREFIX, NO_MATCH
    read_ok: bool = False
    error_msg: Optional[str] = None

    # Sistema de referencia espacial
    crs_name: str = "NA"
    epsg: Optional[int] = None

    # Geometría y 2D vs Z/M
    shape_type_code: int = 0
    geometry_type_str: str = "Unknown"
    has_zm: bool = False

    # Estructura de campos (en el orden exacto de la tabla de atributos, sin FID/Shape)
    campos: List[DbfFieldInfo] = field(default_factory=list)

    # Filas de atributos (lista de dicts {nombre_campo: valor})
    registros: List[Dict[str, Any]] = field(default_factory=list)

    # Geometrías (objetos Shapely o wrapper compatible con .area, .centroid, .is_valid, etc.)
    geometrias: List[Any] = field(default_factory=list)

    @property
    def nombres_campos(self) -> List[str]:
        return [c.nombre for c in self.campos]

    def obtener_campo(self, nombre: str) -> Optional[DbfFieldInfo]:
        nom_up = nombre.strip().upper()
        for c in self.campos:
            if c.nombre.strip().upper() == nom_up:
                return c
        return None


SHAPE_TYPE_NAMES: Dict[int, Tuple[str, bool]] = {
    0: ("Null Shape", False),
    1: ("Point", False),
    3: ("Line String", False),
    5: ("Polygon", False),
    8: ("Multi Point", False),
    11: ("3D Point (PointZ)", True),
    13: ("3D Line String (PolylineZ)", True),
    15: ("3D Polygon (PolygonZ)", True),
    18: ("3D Multi Point (MultiPointZ)", True),
    21: ("PointM (3D/Measured)", True),
    23: ("Line StringM (3D/Measured)", True),
    25: ("PolygonM (3D/Measured)", True),
    28: ("Multi PointM (3D/Measured)", True),
    31: ("MultiPatch (3D)", True),
}


def _detect_encoding(cpg_path: Path) -> str:
    """Detecta la codificación declarada en el archivo .cpg o usa utf-8 con fallback latin1."""
    if cpg_path.exists():
        try:
            raw = cpg_path.read_text(encoding="ascii", errors="ignore").strip().upper()
            if "UTF-8" in raw or "UTF8" in raw or "65001" in raw:
                return "utf-8"
            if "8859" in raw or "1252" in raw or "ANSI" in raw or "LATIN" in raw:
                return "cp1252"
        except Exception:
            pass
    return "utf-8"


def _parse_dbf_header_and_rows(dbf_path: Path, encoding_hint: str) -> Tuple[List[DbfFieldInfo], List[Dict[str, Any]]]:
    """
    Lee directamente los descriptores de campos y registros del archivo .dbf.
    Garantiza el mismo ancho y decimales que `ogrinfo -so -al` en el script R.
    """
    campos: List[DbfFieldInfo] = []
    registros: List[Dict[str, Any]] = []

    with open(dbf_path, "rb") as f:
        header = f.read(32)
        if len(header) < 32:
            raise ValueError(f"Encabezado DBF inválido en {dbf_path}")

        num_records, header_len, record_len = struct.unpack("<IHH", header[4:12])
        num_fields = (header_len - 33) // 32

        for idx in range(num_fields):
            fdesc = f.read(32)
            if not fdesc or fdesc[0] == 0x0D:
                break
            raw_name = fdesc[:11].split(b"\x00", 1)[0]
            try:
                fname = raw_name.decode(encoding_hint).strip()
            except UnicodeDecodeError:
                fname = raw_name.decode("cp1252", errors="replace").strip()

            ftype = chr(fdesc[11]).upper()
            flength = int(fdesc[16])
            fdec = int(fdesc[17])

            if ftype in ("C", "M"):
                tipo_ogr = "String"
                tipo_r_sf = "character"
                fdec = 0
            elif ftype in ("N", "F"):
                if fdec == 0:
                    # En GDAL/ogrinfo y sf: N con 0 decimales y ancho <= 9 es Integer; >= 10 es Integer64/numeric
                    tipo_ogr = "Integer" if flength <= 9 else "Integer64"
                    tipo_r_sf = "integer" if flength <= 9 else "numeric"
                else:
                    tipo_ogr = "Real"
                    tipo_r_sf = "numeric"
            elif ftype == "I":
                tipo_ogr = "Integer"
                tipo_r_sf = "integer"
                fdec = 0
            elif ftype == "D":
                tipo_ogr = "Date"
                tipo_r_sf = "Date"
                fdec = 0
            elif ftype == "L":
                tipo_ogr = "Integer"
                tipo_r_sf = "logical"
                fdec = 0
            else:
                tipo_ogr = "String"
                tipo_r_sf = "character"

            campos.append(
                DbfFieldInfo(
                    nombre=fname,
                    tipo_dbf=ftype,
                    tipo_ogr=tipo_ogr,
                    tipo_r_sf=tipo_r_sf,
                    ancho=flength,
                    decimales=fdec,
                    orden=idx + 1,
                )
            )

        # Posicionarse en el inicio de los registros
        f.seek(header_len)
        for _ in range(num_records):
            rec_bytes = f.read(record_len)
            if not rec_bytes or len(rec_bytes) < record_len:
                break
            # El primer byte indica borrado lógico (0x2A = '*')
            if rec_bytes[0] == 0x2A:
                continue

            offset = 1
            row_dict: Dict[str, Any] = {}
            for c in campos:
                raw_val = rec_bytes[offset : offset + c.ancho]
                offset += c.ancho
                try:
                    txt_val = raw_val.decode(encoding_hint)
                except UnicodeDecodeError:
                    txt_val = raw_val.decode("cp1252", errors="replace")

                if c.tipo_r_sf == "character":
                    # Preservar espacios internos (para detectar dobles espacios), pero quitar padding derecho de DBF
                    cleaned = txt_val.rstrip(" ")
                    if cleaned == "":
                        row_dict[c.nombre] = None
                    else:
                        row_dict[c.nombre] = cleaned
                elif c.tipo_r_sf == "integer":
                    s_num = txt_val.strip()
                    if not s_num or s_num == "*" * len(s_num):
                        row_dict[c.nombre] = None
                    else:
                        try:
                            row_dict[c.nombre] = int(float(s_num))
                        except ValueError:
                            row_dict[c.nombre] = None
                elif c.tipo_r_sf == "numeric":
                    s_num = txt_val.strip().replace(",", ".")
                    if not s_num or s_num == "*" * len(s_num):
                        row_dict[c.nombre] = None
                    else:
                        try:
                            row_dict[c.nombre] = float(s_num)
                        except ValueError:
                            row_dict[c.nombre] = None
                else:
                    row_dict[c.nombre] = txt_val.strip() or None

            registros.append(row_dict)

    return campos, registros


def _parse_shp_header(shp_path: Path) -> Tuple[int, str, bool]:
    """Lee los primeros 100 bytes del archivo .shp para detectar tipo geométrico y presencia de Z/M."""
    with open(shp_path, "rb") as f:
        header = f.read(100)
        if len(header) < 36:
            raise ValueError(f"Encabezado SHP inválido en {shp_path}")
        shape_type = struct.unpack("<i", header[32:36])[0]
        gname, has_zm = SHAPE_TYPE_NAMES.get(shape_type, (f"ShapeType({shape_type})", shape_type >= 11))
        return shape_type, gname, has_zm


def _parse_prj_epsg(prj_path: Path) -> Tuple[str, Optional[int]]:
    """Extrae el nombre del CRS y determina el código EPSG desde el archivo .prj."""
    if not prj_path.exists():
        return "NA", None

    try:
        wkt = prj_path.read_text(encoding="utf-8", errors="ignore").strip()
    except Exception:
        return "NA", None

    if not wkt:
        return "NA", None

    crs_name = "DEFINIDO"
    if 'PROJCS["' in wkt:
        crs_name = wkt.split('PROJCS["', 1)[1].split('"', 1)[0]
    elif 'GEOGCS["' in wkt:
        crs_name = wkt.split('GEOGCS["', 1)[1].split('"', 1)[0]

    wkt_up = wkt.upper()
    # Reconocimiento directo de EPSG:32719 (WGS 84 / UTM zone 19S) y zonas vecinas de Chile
    if ("WGS_1984_UTM_ZONE_19S" in wkt_up) or (
        "WGS_1984" in wkt_up and "TRANSVERSE_MERCATOR" in wkt_up and '"CENTRAL_MERIDIAN",-69.0' in wkt_up.replace(" ", "")
    ) or ("WGS 84 / UTM ZONE 19S" in wkt_up) or ('AUTHORITY["EPSG","32719"]' in wkt_up.replace(" ", "")):
        return crs_name, 32719

    if "WGS_1984_UTM_ZONE_18S" in wkt_up or "WGS 84 / UTM ZONE 18S" in wkt_up:
        return crs_name, 32718

    if "SIRGAS" in wkt_up and "19S" in wkt_up:
        return crs_name, 5361

    # Intentar con ArcPy si está disponible
    try:
        import arcpy
        sr = arcpy.SpatialReference()
        sr.loadFromString(wkt)
        if sr.factoryCode and sr.factoryCode > 0:
            return sr.name or crs_name, int(sr.factoryCode)
    except Exception:
        pass

    # Intentar con pyproj si está disponible
    try:
        from pyproj import CRS
        crs_obj = CRS.from_wkt(wkt)
        epsg_code = crs_obj.to_epsg()
        if epsg_code:
            return crs_name, int(epsg_code)
    except Exception:
        pass

    return crs_name, None


def _read_geometries(shp_path: Path) -> List[Any]:
    """
    Lee las geometrías de todas las entidades como objetos Shapely (vía Fiona o ArcPy WKB o lector binario SHP).
    """
    # 1. Intentar con Fiona + Shapely (rápido y exacto fuera de ArcGIS Pro)
    try:
        import fiona
        from shapely.geometry import shape

        geoms = []
        with fiona.open(str(shp_path), "r") as src:
            for feat in src:
                geom_dict = feat.get("geometry")
                if geom_dict is None:
                    geoms.append(None)
                else:
                    geoms.append(shape(geom_dict))
        return geoms
    except Exception:
        pass

    # 2. Intentar con ArcPy + Shapely WKB (en el entorno Python de ArcGIS Pro)
    try:
        import arcpy
        from shapely import wkb

        geoms = []
        with arcpy.da.SearchCursor(str(shp_path), ["SHAPE@WKB"]) as cursor:
            for (wkb_bytes,) in cursor:
                if wkb_bytes is None:
                    geoms.append(None)
                else:
                    geoms.append(wkb.loads(bytes(wkb_bytes)))
        return geoms
    except Exception:
        pass

    # 3. Fallback con objetos nativos de ArcPy envueltos en interfaz compatible con Shapely
    try:
        import arcpy

        geoms = []
        with arcpy.da.SearchCursor(str(shp_path), ["SHAPE@"]) as cursor:
            for (arc_geom,) in cursor:
                geoms.append(_ArcPyGeometryWrapper(arc_geom) if arc_geom is not None else None)
        return geoms
    except Exception:
        pass

    return []


class _ArcPyGeometryWrapper:
    """Adaptador ligero para que un objeto `arcpy.Geometry` exponga la misma API que Shapely."""

    def __init__(self, arc_geom: Any):
        self._g = arc_geom

    @property
    def is_empty(self) -> bool:
        return self._g is None or getattr(self._g, "partCount", 0) == 0

    @property
    def is_valid(self) -> bool:
        return not self.is_empty

    @property
    def area(self) -> float:
        return float(getattr(self._g, "area", 0.0) or 0.0)

    @property
    def centroid(self) -> Any:
        pt = self._g.trueCentroid if hasattr(self._g, "trueCentroid") else self._g.centroid
        return type("Point", (), {"x": float(pt.X), "y": float(pt.Y)})()

    def intersects(self, other: "_ArcPyGeometryWrapper") -> bool:
        if self.is_empty or other.is_empty:
            return False
        return not self._g.disjoint(other._g)

    def within(self, other: "_ArcPyGeometryWrapper") -> bool:
        if self.is_empty or other.is_empty:
            return False
        return bool(self._g.within(other._g))

    def intersection(self, other: "_ArcPyGeometryWrapper") -> "_ArcPyGeometryWrapper":
        try:
            # dimension=4 para intersección poligonal (área)
            res = self._g.intersect(other._g, 4)
            return _ArcPyGeometryWrapper(res)
        except Exception:
            return _ArcPyGeometryWrapper(None)

    def difference(self, other: "_ArcPyGeometryWrapper") -> "_ArcPyGeometryWrapper":
        try:
            res = self._g.difference(other._g)
            return _ArcPyGeometryWrapper(res)
        except Exception:
            return _ArcPyGeometryWrapper(None)


def read_shapefile_complete(shp_path: Path) -> ShapefileData:
    """Lee de forma integral un archivo Shapefile (.shp, .dbf, .prj, .cpg)."""
    shp_path = Path(shp_path)
    data = ShapefileData(filepath=shp_path, nombre_shape=shp_path.stem)

    dbf_path = shp_path.with_suffix(".dbf")
    if not dbf_path.exists():
        # Buscar con extensión en mayúsculas (.DBF)
        alt_dbf = shp_path.parent / f"{shp_path.stem}.DBF"
        if alt_dbf.exists():
            dbf_path = alt_dbf
        else:
            data.read_ok = False
            data.error_msg = f"No se encontró el archivo .dbf para {shp_path.name}"
            return data

    try:
        enc = _detect_encoding(shp_path.with_suffix(".cpg"))
        stype, gname, has_zm = _parse_shp_header(shp_path)
        data.shape_type_code = stype
        data.geometry_type_str = gname
        data.has_zm = has_zm

        crs_name, epsg = _parse_prj_epsg(shp_path.with_suffix(".prj"))
        data.crs_name = crs_name
        data.epsg = epsg

        campos, registros = _parse_dbf_header_and_rows(dbf_path, enc)
        data.campos = campos
        data.registros = registros

        data.geometrias = _read_geometries(shp_path)
        data.read_ok = True
    except Exception as exc:
        data.read_ok = False
        data.error_msg = str(exc)

    return data


# ==============================================================================
# SECCIÓN 3: VALIDADORES DE ATRIBUTOS, FORMATO, DOMINIOS Y GEOMETRÍA 2D
# ==============================================================================
"""
Validadores individuales por Shapefile (equivalentes al ciclo principal líneas 342-563
de Revision_Campos_PAS151_Rev11 2.R y Descripcion_general.docx).
"""

from dataclasses import dataclass, field
import re
from typing import Any, Callable, Dict, List, Optional





REGEX_CARACTERES_INVALIDOS = re.compile(r"[áéíóúÁÉÍÓÚ'àèìòùÀÈÌÒÙ´`]")


@dataclass
class RFieldRow:
    """Fila exacta de la tabla `resultado` exportada por el script R (`00_Reporte_revision_campos_<PREDIO>.xlsx`)."""
    shape: str
    nombre: str
    Tipo_Plant: str
    Tipo_Shp: Optional[str]
    Existe_Campo: str               # "Sí" / "No"
    Coin_Tipo_Campo: Optional[bool]
    Long_Plant: Optional[int]
    Long_Shp: Optional[int]
    Long_Correcta: Optional[bool]
    Dec_Plant: Optional[int]
    Dec_Shp: Optional[int]
    Dec_Correcto: Optional[bool]
    orden_correcto: bool


@dataclass
class AuditIssue:
    """Incidencia detallada detectada durante la validación de atributos, geometría o cruces espaciales."""
    issue_id: str
    capa: str
    prefijo_plantilla: str
    categoria: str
    severidad: str          # "ERROR", "ADVERTENCIA", "INFO"
    campo_o_regla: str
    valor_esperado: str
    valor_encontrado: str
    filas_afectadas: str
    mensaje: str
    recomendacion: str


@dataclass
class LayerValidationSummary:
    """Resumen del estado de validación individual de un Shapefile."""
    nombre_shape: str
    prefijo_plantilla: str
    tipo_coincidencia: str
    num_registros: int
    crs_ok: bool
    crs_detalle: str
    es_2d_ok: bool
    geometria_detalle: str
    estructura_campos_ok: bool
    orden_campos_ok: bool
    campos_adicionales: List[str] = field(default_factory=list)
    dominios_ok: bool = True
    campos_sin_vacios_ok: bool = True
    sup_ha_ok: Optional[bool] = None      # None si la capa no tiene Sup_ha
    coord_xy_ok: Optional[bool] = None    # None si la capa no tiene Coord_X/Coord_Y
    texto_caracteres_ok: bool = True
    texto_doble_espacio_ok: bool = True
    geometrias_validas_ok: bool = True
    num_geometrias_invalidas: int = 0
    filas_resultado_r: List[RFieldRow] = field(default_factory=list)
    incidencias: List[AuditIssue] = field(default_factory=list)

    @property
    def estado_global(self) -> str:
        if any(i.severidad == "ERROR" for i in self.incidencias):
            return "CON OBSERVACIONES (ERROR)"
        if any(i.severidad == "ADVERTENCIA" for i in self.incidencias):
            return "CON ADVERTENCIAS"
        return "CONFORME"


def _fmt_rows(indices_1based: List[int], limit: int = 10) -> str:
    if not indices_1based:
        return ""
    head_str = ", ".join(str(i) for i in indices_1based[:limit])
    if len(indices_1based) > limit:
        head_str += f" ... (total: {len(indices_1based)})"
    return head_str


def validate_single_shapefile(
    shp_data: ShapefileData,
    plantilla: LayerTemplate,
    epsg_esperado: int = 32719,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> LayerValidationSummary:
    """
    Ejecuta todas las validaciones individuales del script R sobre un shapefile:
    1. CRS (EPSG:32719)
    2. Geometría 2D vs Z/M
    3. Estructura de campos (Existencia, Tipo, Longitud, Decimales, Orden exacto)
    4. Dominios CONAF y valores vacíos / NA
    5. Superficie declarada (Sup_ha) vs área geométrica recalculada (tolerancia 1%)
    6. Coordenadas (Coord_X, Coord_Y) vs centroide recalculado (redondeado a entero)
    7. Formato de texto (acentos/caracteres inválidos y dobles espacios)
    8. Validez topológica de geometrías (st_is_valid)
    """
    def emit(msg: str, level: str = "INFO") -> None:
        if log_fn:
            log_fn(msg, level)

    nombre_shape = shp_data.nombre_shape
    prefijo = plantilla.prefijo
    emit(f"Revisando shape: {nombre_shape}", "HEADER")

    # 1. Revisar Sistema de Coordenadas (EPSG:32719)
    crs_ok = shp_data.epsg == epsg_esperado
    crs_str = f"EPSG:{shp_data.epsg} ({shp_data.crs_name})" if shp_data.epsg else f"NA ({shp_data.crs_name})"
    if crs_ok:
        emit(f"  [OK] Coordenadas CORRECTAS (WGS84 / UTM 19S, EPSG:{epsg_esperado}) en '{nombre_shape}'.", "OK")
    else:
        emit(
            f"  [ERROR] Coordenadas INCORRECTAS en '{nombre_shape}'. Tiene: "
            f"{shp_data.epsg if shp_data.epsg is not None else 'NA'} ({shp_data.crs_name})",
            "ERROR",
        )

    # 2. Revisar Geometría 2D vs Z/M
    es_2d_ok = not shp_data.has_zm
    geom_str = f"Geometry: {shp_data.geometry_type_str}"
    if shp_data.has_zm:
        emit(f"  [ERROR] El shape '{nombre_shape}' tiene dimensiones Z/M en la geometría. ({geom_str})", "ERROR")
    else:
        emit(f"  [OK] El shape '{nombre_shape}' es 2D (sin dimensiones Z/M). ({geom_str})", "OK")

    summary = LayerValidationSummary(
        nombre_shape=nombre_shape,
        prefijo_plantilla=prefijo,
        tipo_coincidencia=shp_data.tipo_coincidencia,
        num_registros=len(shp_data.registros),
        crs_ok=crs_ok,
        crs_detalle=crs_str,
        es_2d_ok=es_2d_ok,
        geometria_detalle=geom_str,
        estructura_campos_ok=True,
        orden_campos_ok=True,
    )

    if not crs_ok:
        summary.incidencias.append(
            AuditIssue(
                issue_id=f"CRS_{nombre_shape}",
                capa=nombre_shape,
                prefijo_plantilla=prefijo,
                categoria="SISTEMA_COORDENADAS",
                severidad="ERROR",
                campo_o_regla="CRS / EPSG",
                valor_esperado=f"EPSG:{epsg_esperado} (WGS84 / UTM 19S)",
                valor_encontrado=crs_str,
                filas_afectadas="Todas",
                mensaje=f"Coordenadas incorrectas en '{nombre_shape}'. Se encontró {crs_str}.",
                recomendacion=f"Reproyectar la capa a WGS 84 / UTM Zone 19S (EPSG:{epsg_esperado}).",
            )
        )

    if not es_2d_ok:
        summary.incidencias.append(
            AuditIssue(
                issue_id=f"ZM_{nombre_shape}",
                capa=nombre_shape,
                prefijo_plantilla=prefijo,
                categoria="DIMENSIONES_ZM",
                severidad="ERROR",
                campo_o_regla="Geometría 2D",
                valor_esperado="2D (sin Z ni M)",
                valor_encontrado=shp_data.geometry_type_str,
                filas_afectadas="Todas",
                mensaje=f"El shapefile '{nombre_shape}' posee dimensiones Z o M ({shp_data.geometry_type_str}).",
                recomendacion="Exportar el shapefile deshabilitando los valores Z y M en el entorno de geoprocesamiento.",
            )
        )

    # 3. Revisar Campos, Tipos, Longitud, Decimales y Orden Exacto
    campos_plantilla = [c.nombre for c in plantilla.campos]
    campos_shape = [c.nombre for c in shp_data.campos]
    coincide_orden = campos_plantilla == campos_shape
    summary.orden_campos_ok = coincide_orden

    for f_plant in plantilla.campos:
        # En R el join por nombre es exacto (`by = "nombre"`), también soportamos búsqueda case-insensitive para diagnóstico
        f_exact = next((c for c in shp_data.campos if c.nombre == f_plant.nombre), None)
        f_ci = shp_data.obtener_campo(f_plant.nombre)
        f_real = f_exact if f_exact is not None else f_ci

        existe_exacto = f_exact is not None
        existe_str = "Sí" if existe_exacto else "No"

        tipo_shp = f_real.tipo_r_sf if existe_exacto else None
        coin_tipo = (f_plant.tipo == tipo_shp) if (existe_exacto and tipo_shp is not None) else None

        long_shp = f_real.ancho if existe_exacto else None
        long_correcta: Optional[bool] = None
        if existe_exacto and f_plant.longitud is not None and long_shp is not None:
            long_correcta = f_plant.longitud == long_shp

        dec_shp = f_real.decimales if existe_exacto else None
        dec_correcto: Optional[bool] = None
        if existe_exacto and f_plant.decimales is not None and dec_shp is not None:
            dec_correcto = f_plant.decimales == dec_shp

        row_r = RFieldRow(
            shape=nombre_shape,
            nombre=f_plant.nombre,
            Tipo_Plant=f_plant.tipo,
            Tipo_Shp=tipo_shp,
            Existe_Campo=existe_str,
            Coin_Tipo_Campo=coin_tipo,
            Long_Plant=f_plant.longitud,
            Long_Shp=long_shp,
            Long_Correcta=long_correcta,
            Dec_Plant=f_plant.decimales,
            Dec_Shp=dec_shp,
            Dec_Correcto=dec_correcto,
            orden_correcto=coincide_orden,
        )
        summary.filas_resultado_r.append(row_r)

        if not existe_exacto:
            summary.estructura_campos_ok = False
            detalle_ci = f" (existe como '{f_ci.nombre}' con distinta mayúscula/minúscula)" if f_ci else ""
            summary.incidencias.append(
                AuditIssue(
                    issue_id=f"MISSING_FIELD_{nombre_shape}_{f_plant.nombre}",
                    capa=nombre_shape,
                    prefijo_plantilla=prefijo,
                    categoria="ESTRUCTURA_CAMPOS",
                    severidad="ERROR",
                    campo_o_regla=f_plant.nombre,
                    valor_esperado="Campo existente",
                    valor_encontrado=f"No existe{detalle_ci}",
                    filas_afectadas="Estructura",
                    mensaje=f"El campo obligatorio '{f_plant.nombre}' no existe en '{nombre_shape}'{detalle_ci}.",
                    recomendacion=f"Crear el campo '{f_plant.nombre}' de tipo {f_plant.tipo}.",
                )
            )
            continue

        if coin_tipo is False:
            summary.estructura_campos_ok = False
            summary.incidencias.append(
                AuditIssue(
                    issue_id=f"TYPE_FIELD_{nombre_shape}_{f_plant.nombre}",
                    capa=nombre_shape,
                    prefijo_plantilla=prefijo,
                    categoria="ESTRUCTURA_CAMPOS",
                    severidad="ERROR",
                    campo_o_regla=f_plant.nombre,
                    valor_esperado=f_plant.tipo,
                    valor_encontrado=str(tipo_shp),
                    filas_afectadas="Estructura",
                    mensaje=f"Tipo de dato incorrecto en campo '{f_plant.nombre}': esperado '{f_plant.tipo}', encontrado '{tipo_shp}'.",
                    recomendacion=f"Redefinir el campo '{f_plant.nombre}' con tipo '{f_plant.tipo}'.",
                )
            )

        if long_correcta is False:
            summary.estructura_campos_ok = False
            summary.incidencias.append(
                AuditIssue(
                    issue_id=f"LEN_FIELD_{nombre_shape}_{f_plant.nombre}",
                    capa=nombre_shape,
                    prefijo_plantilla=prefijo,
                    categoria="ESTRUCTURA_CAMPOS",
                    severidad="ERROR",
                    campo_o_regla=f_plant.nombre,
                    valor_esperado=str(f_plant.longitud),
                    valor_encontrado=str(long_shp),
                    filas_afectadas="Estructura",
                    mensaje=f"Longitud incorrecta en campo '{f_plant.nombre}': esperada {f_plant.longitud}, encontrada {long_shp}.",
                    recomendacion=f"Ajustar el ancho del campo '{f_plant.nombre}' a {f_plant.longitud}.",
                )
            )

        if dec_correcto is False:
            summary.estructura_campos_ok = False
            summary.incidencias.append(
                AuditIssue(
                    issue_id=f"DEC_FIELD_{nombre_shape}_{f_plant.nombre}",
                    capa=nombre_shape,
                    prefijo_plantilla=prefijo,
                    categoria="ESTRUCTURA_CAMPOS",
                    severidad="ERROR",
                    campo_o_regla=f_plant.nombre,
                    valor_esperado=str(f_plant.decimales),
                    valor_encontrado=str(dec_shp),
                    filas_afectadas="Estructura",
                    mensaje=f"Decimales incorrectos en campo '{f_plant.nombre}': esperados {f_plant.decimales}, encontrados {dec_shp}.",
                    recomendacion=f"Ajustar la escala decimal del campo '{f_plant.nombre}' a {f_plant.decimales}.",
                )
            )

    # Detectar campos adicionales en el shapefile que no están en la plantilla
    extra_fields = [c for c in campos_shape if c not in campos_plantilla]
    summary.campos_adicionales = extra_fields
    if extra_fields:
        summary.estructura_campos_ok = False
        summary.incidencias.append(
            AuditIssue(
                issue_id=f"EXTRA_FIELDS_{nombre_shape}",
                capa=nombre_shape,
                prefijo_plantilla=prefijo,
                categoria="ESTRUCTURA_CAMPOS",
                severidad="ADVERTENCIA",
                campo_o_regla=", ".join(extra_fields),
                valor_esperado="Solo campos de plantilla",
                valor_encontrado=", ".join(extra_fields),
                filas_afectadas="Estructura",
                mensaje=f"El shapefile '{nombre_shape}' contiene campos adicionales fuera de la plantilla: {', '.join(extra_fields)}.",
                recomendacion="Eliminar los campos adicionales antes de la entrega CONAF.",
            )
        )

    if coincide_orden:
        emit(f"  [OK] El shape '{nombre_shape}' tiene el orden de campos CORRECTO.", "OK")
    else:
        emit(f"  [ERROR] El shape '{nombre_shape}' tiene el orden de campos INCORRECTO.", "ERROR")
        emit(f"     Esperado : {' / '.join(campos_plantilla)}", "ERROR")
        emit(f"     Encontrado: {' / '.join(campos_shape)}", "ERROR")
        summary.incidencias.append(
            AuditIssue(
                issue_id=f"ORDER_FIELDS_{nombre_shape}",
                capa=nombre_shape,
                prefijo_plantilla=prefijo,
                categoria="ORDEN_CAMPOS",
                severidad="ERROR",
                campo_o_regla="Orden de Columnas",
                valor_esperado=" / ".join(campos_plantilla),
                valor_encontrado=" / ".join(campos_shape),
                filas_afectadas="Estructura",
                mensaje=f"El orden o conjunto de campos de '{nombre_shape}' no coincide con la plantilla CONAF.",
                recomendacion=f"Reordenar los campos exactamente como: {' / '.join(campos_plantilla)}.",
            )
        )

    # 4. Validación de Dominios CONAF y Valores Vacíos / NA
    dominios_ok = True
    nombres_atributos = shp_data.nombres_campos

    for campo_dom, valores_permitidos in DOMINIOS_CONAF.items():
        if campo_dom in nombres_atributos:
            # Si es Pendientes_caminos, Ran_Pend puede tener un dominio particular según Descripcion_general.docx
            if prefijo == "Pendientes_caminos" and campo_dom == "Ran_Pend":
                continue

            permitidos_norm = {str(v).strip() for v in valores_permitidos}
            idx_invalidos: List[int] = []
            valores_malos_muestra: List[str] = []

            for idx_row, reg in enumerate(shp_data.registros, start=1):
                val = reg.get(campo_dom)
                if val is None or str(val).strip() == "" or str(val).strip().upper() == "NA":
                    idx_invalidos.append(idx_row)
                    valores_malos_muestra.append("NA/Vacío")
                else:
                    s_val = str(val).strip()
                    if s_val not in permitidos_norm:
                        idx_invalidos.append(idx_row)
                        if s_val not in valores_malos_muestra:
                            valores_malos_muestra.append(s_val)

            if idx_invalidos:
                dominios_ok = False
                filas_str = _fmt_rows(idx_invalidos)
                emit(
                    f"  [ERROR] Campo '{campo_dom}' contiene valores fuera del dominio o vacíos (NA/'') en filas: {filas_str}",
                    "ERROR",
                )
                summary.incidencias.append(
                    AuditIssue(
                        issue_id=f"DOM_{nombre_shape}_{campo_dom}",
                        capa=nombre_shape,
                        prefijo_plantilla=prefijo,
                        categoria="DOMINIO_CONAF",
                        severidad="ERROR",
                        campo_o_regla=campo_dom,
                        valor_esperado=", ".join(str(v) for v in valores_permitidos),
                        valor_encontrado=", ".join(valores_malos_muestra[:5]),
                        filas_afectadas=filas_str,
                        mensaje=f"Campo '{campo_dom}' tiene {len(idx_invalidos)} registro(s) fuera del dominio CONAF o vacíos.",
                        recomendacion=f"Ajustar los valores de '{campo_dom}' a uno de: {', '.join(str(v) for v in valores_permitidos)}.",
                    )
                )

    summary.dominios_ok = dominios_ok
    if dominios_ok:
        emit("  [OK] Dominios de atributos OK (sin vacíos ni valores diferentes).", "OK")

    # Revisar también valores vacíos o NA en los demás campos obligatorios de la plantilla
    for f_plant in plantilla.campos:
        if f_plant.nombre in nombres_atributos and f_plant.nombre not in DOMINIOS_CONAF:
            idx_vacios = [
                idx_row
                for idx_row, reg in enumerate(shp_data.registros, start=1)
                if reg.get(f_plant.nombre) is None
                or str(reg.get(f_plant.nombre)).strip() == ""
                or str(reg.get(f_plant.nombre)).strip().upper() == "NA"
            ]
            if idx_vacios:
                summary.campos_sin_vacios_ok = False
                filas_str = _fmt_rows(idx_vacios)
                emit(
                    f"  [ERROR] Campo '{f_plant.nombre}' contiene valores vacíos o NA en filas: {filas_str}",
                    "ERROR",
                )
                summary.incidencias.append(
                    AuditIssue(
                        issue_id=f"EMPTY_{nombre_shape}_{f_plant.nombre}",
                        capa=nombre_shape,
                        prefijo_plantilla=prefijo,
                        categoria="VALORES_VACIOS_NA",
                        severidad="ERROR",
                        campo_o_regla=f_plant.nombre,
                        valor_esperado="Valor no vacío",
                        valor_encontrado="Vacío / NA",
                        filas_afectadas=filas_str,
                        mensaje=f"El campo '{f_plant.nombre}' presenta {len(idx_vacios)} registro(s) vacío(s) o NA.",
                        recomendacion=f"Completar todos los registros vacíos en el campo '{f_plant.nombre}'.",
                    )
                )

    # 5. Verificar Sup_ha vs superficie recalculada desde la geometría (líneas 471-480 del R)
    if "Sup_ha" in nombres_atributos and shp_data.geometrias:
        idx_dif_sup: List[int] = []
        ejemplos_dif: List[str] = []
        for idx_row, (reg, geom) in enumerate(zip(shp_data.registros, shp_data.geometrias), start=1):
            val_sup = reg.get("Sup_ha")
            if geom is None or getattr(geom, "is_empty", False):
                idx_dif_sup.append(idx_row)
                continue
            area_geom = round(float(geom.area) / 10000.0, 2)
            try:
                sup_decl = float(val_sup)
            except (TypeError, ValueError):
                idx_dif_sup.append(idx_row)
                ejemplos_dif.append(f"Fila {idx_row}: decl=NA vs calc={area_geom}")
                continue

            denom = max(area_geom, 1e-6)
            dif_rel = abs(area_geom - sup_decl) / denom
            if dif_rel > 0.01:
                idx_dif_sup.append(idx_row)
                if len(ejemplos_dif) < 5:
                    ejemplos_dif.append(f"Fila {idx_row}: decl={sup_decl} ha vs calc={area_geom} ha")

        if idx_dif_sup:
            summary.sup_ha_ok = False
            filas_str = _fmt_rows(idx_dif_sup)
            emit(
                f"  [ERROR] Hay diferencias (>1%) entre el campo Sup_ha y el área recalculada en filas: {filas_str}",
                "ERROR",
            )
            summary.incidencias.append(
                AuditIssue(
                    issue_id=f"SUP_HA_{nombre_shape}",
                    capa=nombre_shape,
                    prefijo_plantilla=prefijo,
                    categoria="CONSISTENCIA_SUPERFICIE",
                    severidad="ERROR",
                    campo_o_regla="Sup_ha",
                    valor_esperado="Diferencia relativa <= 1% respecto a geometría",
                    valor_encontrado="; ".join(ejemplos_dif) or f"{len(idx_dif_sup)} filas con diferencia",
                    filas_afectadas=filas_str,
                    mensaje=f"Existen {len(idx_dif_sup)} polígonos donde 'Sup_ha' difiere en más de 1% respecto del área geométrica.",
                    recomendacion="Recalcular el campo 'Sup_ha' en hectáreas redondeado a 2 decimales (Área m² / 10,000).",
                )
            )
        else:
            summary.sup_ha_ok = True
            emit("  [OK] No hay diferencias entre Sup_ha con el área recalculada.", "OK")

    # 6. Revisar coordenadas Coord_X y Coord_Y vs Centroide recalculado (líneas 483-505 del R)
    if "Coord_X" in nombres_atributos and "Coord_Y" in nombres_atributos and shp_data.geometrias:
        idx_malos_xy: List[int] = []
        ejemplos_xy: List[str] = []
        for idx_row, (reg, geom) in enumerate(zip(shp_data.registros, shp_data.geometrias), start=1):
            if geom is None or getattr(geom, "is_empty", False):
                idx_malos_xy.append(idx_row)
                continue
            cent = geom.centroid
            # En R: round(coords_cent[, "X"]) y round(coords_cent[, "Y"])
            cent_x = int(round(float(cent.x)))
            cent_y = int(round(float(cent.y)))
            try:
                campo_x = int(round(float(reg.get("Coord_X"))))
                campo_y = int(round(float(reg.get("Coord_Y"))))
                if cent_x != campo_x or cent_y != campo_y:
                    idx_malos_xy.append(idx_row)
                    if len(ejemplos_xy) < 5:
                        ejemplos_xy.append(f"Fila {idx_row}: ({campo_x},{campo_y}) vs ({cent_x},{cent_y})")
            except (TypeError, ValueError):
                idx_malos_xy.append(idx_row)
                if len(ejemplos_xy) < 5:
                    ejemplos_xy.append(f"Fila {idx_row}: coordenadas vacías/no numéricas")

        if not idx_malos_xy:
            summary.coord_xy_ok = True
            emit("  [OK] Coord_X / Coord_Y coherentes con las coordenadas recalculadas.", "OK")
        else:
            summary.coord_xy_ok = False
            filas_str = _fmt_rows(idx_malos_xy)
            emit(
                f"  [ERROR] Coord_X / Coord_Y diferentes del centroide calculado o campos vacíos: {len(idx_malos_xy)}",
                "ERROR",
            )
            emit(f"     Filas (primeros 10): {filas_str}", "ERROR")
            summary.incidencias.append(
                AuditIssue(
                    issue_id=f"COORD_XY_{nombre_shape}",
                    capa=nombre_shape,
                    prefijo_plantilla=prefijo,
                    categoria="CONSISTENCIA_COORDENADAS",
                    severidad="ERROR",
                    campo_o_regla="Coord_X / Coord_Y",
                    valor_esperado="Coincidencia exacta con centroide UTM redondeado a entero",
                    valor_encontrado="; ".join(ejemplos_xy),
                    filas_afectadas=filas_str,
                    mensaje=f"Hay {len(idx_malos_xy)} registro(s) cuyas coordenadas Coord_X/Coord_Y no coinciden con la geometría.",
                    recomendacion="Recalcular Coord_X y Coord_Y como enteros a partir de las coordenadas X/Y del punto/centroide.",
                )
            )

    # 7. Revisar acentos, caracteres inválidos y dobles espacios en campos de texto (líneas 508-553 del R)
    campos_texto = [c.nombre for c in shp_data.campos if c.tipo_r_sf == "character"]
    if campos_texto:
        formato_txt = False
        doble_esp = False
        for campo_txt in campos_texto:
            idx_inv_txt: List[int] = []
            idx_doble: List[int] = []
            for idx_row, reg in enumerate(shp_data.registros, start=1):
                val = reg.get(campo_txt)
                if val is None:
                    continue
                s_val = str(val)
                if REGEX_CARACTERES_INVALIDOS.search(s_val):
                    idx_inv_txt.append(idx_row)
                if "  " in s_val:
                    idx_doble.append(idx_row)

            if idx_inv_txt:
                formato_txt = True
                summary.texto_caracteres_ok = False
                filas_str = _fmt_rows(idx_inv_txt)
                emit(
                    f"  [ERROR] Campo '{campo_txt}' contiene caracteres inválidos (acentos/apóstrofes) en las filas: {filas_str}",
                    "ERROR",
                )
                summary.incidencias.append(
                    AuditIssue(
                        issue_id=f"TXT_CHARS_{nombre_shape}_{campo_txt}",
                        capa=nombre_shape,
                        prefijo_plantilla=prefijo,
                        categoria="FORMATO_TEXTO",
                        severidad="ERROR",
                        campo_o_regla=campo_txt,
                        valor_esperado="Sin tildes, apóstrofes ni comillas invertidas",
                        valor_encontrado=f"{len(idx_inv_txt)} fila(s) con caracteres restringidos",
                        filas_afectadas=filas_str,
                        mensaje=f"El campo '{campo_txt}' contiene tildes o caracteres especiales en las filas: {filas_str}.",
                        recomendacion=f"Reemplazar tildes y caracteres especiales en el campo '{campo_txt}'.",
                    )
                )

            if idx_doble:
                doble_esp = True
                summary.texto_doble_espacio_ok = False
                filas_str = _fmt_rows(idx_doble)
                emit(
                    f"  [ERROR] Campo '{campo_txt}' contiene dobles espacios en las filas: {filas_str}",
                    "ERROR",
                )
                summary.incidencias.append(
                    AuditIssue(
                        issue_id=f"TXT_SPACE_{nombre_shape}_{campo_txt}",
                        capa=nombre_shape,
                        prefijo_plantilla=prefijo,
                        categoria="FORMATO_TEXTO",
                        severidad="ERROR",
                        campo_o_regla=campo_txt,
                        valor_esperado="Espaciado simple",
                        valor_encontrado=f"{len(idx_doble)} fila(s) con dobles espacios",
                        filas_afectadas=filas_str,
                        mensaje=f"El campo '{campo_txt}' contiene dobles espacios en las filas: {filas_str}.",
                        recomendacion=f"Normalizar espacios múltiples a un solo espacio en '{campo_txt}'.",
                    )
                )

        if not formato_txt:
            emit("  [OK] No hay caracteres inválidos en los campos de texto del shape.", "OK")
        if not doble_esp:
            emit("  [OK] No hay dobles espacios en campos de texto.", "OK")
    else:
        emit("  [INFO] El shape no tiene campos de texto para validar caracteres o dobles espacios.", "INFO")

    # 8. Validar geometrías (st_is_valid, líneas 555-561 del R)
    idx_geom_inv: List[int] = []
    for idx_row, geom in enumerate(shp_data.geometrias, start=1):
        if geom is None or getattr(geom, "is_empty", False) or not getattr(geom, "is_valid", True):
            idx_geom_inv.append(idx_row)

    summary.num_geometrias_invalidas = len(idx_geom_inv)
    if idx_geom_inv:
        summary.geometrias_validas_ok = False
        filas_str = _fmt_rows(idx_geom_inv)
        emit(f"  [ERROR] {len(idx_geom_inv)} geometrías inválidas (self-intersection, vacías, etc.) en filas: {filas_str}", "ERROR")
        summary.incidencias.append(
            AuditIssue(
                issue_id=f"GEOM_VALID_{nombre_shape}",
                capa=nombre_shape,
                prefijo_plantilla=prefijo,
                categoria="VALIDEZ_GEOMETRICA",
                severidad="ERROR",
                campo_o_regla="st_is_valid / Topología",
                valor_esperado="Todas las geometrías válidas",
                valor_encontrado=f"{len(idx_geom_inv)} geometría(s) inválida(s)",
                filas_afectadas=filas_str,
                mensaje=f"Se detectaron {len(idx_geom_inv)} geometrías inválidas (autointersección o geometría nula) en '{nombre_shape}'.",
                recomendacion="Ejecutar la herramienta 'Repair Geometry' (Reparar Geometría) en ArcGIS Pro.",
            )
        )
    else:
        emit("  [OK] Todas las geometrías son válidas.", "OK")

    return summary


# ==============================================================================
# SECCIÓN 4: VALIDADORES ESPACIALES E INTER-CAPAS
# ==============================================================================
"""
Validaciones Espaciales y de Consistencia entre Capas (Sección 2 de Revision_Campos_PAS151_Rev11 2.R,
líneas 573-883, y Descripcion_general.docx).
"""

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple






@dataclass
class SpatialCheckResult:
    """Resultado estructurado de un control espacial o inter-capa para el reporte Excel."""
    control_id: str
    grupo: str
    capas_evaluadas: str
    estado: str             # "CONFORME", "ERROR", "ADVERTENCIA", "NO_APLICA"
    resumen_resultado: str
    detalle_tecnico: str


def _unify_geometries(geoms: List[Any]) -> Optional[Any]:
    """Une una lista de geometrías válidas usando shapely.ops.unary_union o unión iterativa."""
    valid_geoms = [g for g in geoms if g is not None and not getattr(g, "is_empty", False)]
    if not valid_geoms:
        return None
    try:
        from shapely.ops import unary_union
        from shapely.validation import make_valid

        repaired = [make_valid(g) if not getattr(g, "is_valid", True) else g for g in valid_geoms]
        return unary_union(repaired)
    except Exception:
        acc = valid_geoms[0]
        for g in valid_geoms[1:]:
            try:
                acc = acc.union(g)
            except Exception:
                pass
        return acc


def _safe_intersection_area(g1: Any, g2: Any) -> float:
    """Calcula el área de intersección en m² manejando posibles errores topológicos."""
    if g1 is None or g2 is None or getattr(g1, "is_empty", False) or getattr(g2, "is_empty", False):
        return 0.0
    try:
        if not g1.intersects(g2):
            return 0.0
        inter = g1.intersection(g2)
        return float(getattr(inter, "area", 0.0) or 0.0)
    except Exception:
        try:
            from shapely.validation import make_valid

            v1 = make_valid(g1)
            v2 = make_valid(g2)
            if not v1.intersects(v2):
                return 0.0
            return float(v1.intersection(v2).area)
        except Exception:
            return 0.0


def _safe_difference_area(geom_list: List[Any], boundary_union: Any) -> float:
    """Calcula la superficie total (m²) de `geom_list` que queda fuera de `boundary_union`."""
    if boundary_union is None:
        return 0.0
    capa_union = _unify_geometries(geom_list)
    if capa_union is None or getattr(capa_union, "is_empty", False):
        return 0.0
    try:
        diff = capa_union.difference(boundary_union)
        return float(getattr(diff, "area", 0.0) or 0.0)
    except Exception:
        try:
            from shapely.validation import make_valid

            diff = make_valid(capa_union).difference(make_valid(boundary_union))
            return float(getattr(diff, "area", 0.0) or 0.0)
        except Exception:
            return 0.0


def run_all_spatial_validations(
    by_prefix: Dict[str, ShapefileData],
    all_shps: List[ShapefileData],
    ruta_shp_area_comp: Optional[Path] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> Tuple[List[SpatialCheckResult], List[AuditIssue], Optional[str]]:
    """
    Ejecuta las 7 validaciones espaciales y entre capas definidas en Revision_Campos_PAS151_Rev11 2.R:
    1. IDs duplicados en Area (N_Area), Referencia (Nom_Pto), Parcela (N_Parc) y Rodales (Nom_Predio / N_Rodal).
    2. Traslape de Area del predio contra el shapefile compilado de áreas de corta.
    3. Consistencia espacial (st_within) y de atributo (N_Rodal) entre Parcela y Rodales.
    4. Contención de Area, Suelos, Rangos_pend y Pendientes_caminos dentro de Limite_Predial y Area_Proyecto.
    5. Traslape entre Predios_vecinos y Limite_Predial.
    6. Elementos de Infra_aislada (categorías de control de erosión) sobre Area_Proyecto.
    7. Consistencia transversal de Nom_Predio en todas las capas respecto de Limite_Predial.
    """
    def emit(msg: str, level: str = "INFO") -> None:
        if log_fn:
            log_fn(msg, level)

    results: List[SpatialCheckResult] = []
    issues: List[AuditIssue] = []

    def get_layer(prefix: str) -> Optional[ShapefileData]:
        return by_prefix.get(normalize_layer_key(prefix))

    emit("\n======================================================================", "HEADER")
    emit("INICIANDO VALIDACIONES ESPACIALES Y DE CONSISTENCIA ENTRE CAPAS", "HEADER")
    emit("======================================================================", "HEADER")

    # -------------------------------------------------------------------------
    # 1. DETECCIÓN DE IDENTIFICADORES DUPLICADOS (líneas 585-639 del R)
    # -------------------------------------------------------------------------
    dup_configs = [
        ("Area", "N_Area", "Areas de corta (N_Area)"),
        ("Referencia", "Nom_Pto", "Puntos de referencia (Nom_Pto)"),
        ("Parcela", "N_Parc", "Parcelas (N_Parc)"),
        ("Rodales", "Nom_Predio+N_Rodal", "Rodales (Nom_Predio / N_Rodal)"),
    ]

    for pref, campo_id, etiqueta in dup_configs:
        layer = get_layer(pref)
        if not layer or not layer.read_ok:
            results.append(
                SpatialCheckResult(
                    control_id=f"DUP_{pref}",
                    grupo="1. Identificadores Duplicados",
                    capas_evaluadas=pref,
                    estado="NO_APLICA",
                    resumen_resultado=f"Capa '{pref}' no presente en el directorio.",
                    detalle_tecnico="No se ejecutó la validación por ausencia de la capa.",
                )
            )
            continue

        valores: List[str] = []
        filas_por_valor: Dict[str, List[int]] = {}
        for idx_row, reg in enumerate(layer.registros, start=1):
            if campo_id == "Nom_Predio+N_Rodal":
                v_pred = str(reg.get("Nom_Predio") or "").strip()
                v_rod = str(reg.get("N_Rodal") or "").strip()
                val_str = f"{v_pred} / {v_rod}"
            else:
                val_str = str(reg.get(campo_id) or "").strip()

            valores.append(val_str)
            filas_por_valor.setdefault(val_str, []).append(idx_row)

        duplicados = {val: idxs for val, idxs in filas_por_valor.items() if len(idxs) > 1}
        if not duplicados:
            msg_ok = f"No hay duplicados en {etiqueta} de la capa: {layer.nombre_shape}.shp"
            emit(f"  [OK] {msg_ok}", "OK")
            results.append(
                SpatialCheckResult(
                    control_id=f"DUP_{pref}",
                    grupo="1. Identificadores Duplicados",
                    capas_evaluadas=layer.nombre_shape,
                    estado="CONFORME",
                    resumen_resultado=msg_ok,
                    detalle_tecnico=f"Total registros evaluados: {len(valores)} (todos únicos).",
                )
            )
        else:
            dup_desc = "; ".join(f"'{k}' (filas {', '.join(map(str, v))})" for k, v in list(duplicados.items())[:10])
            all_dup_rows = sorted(r for v in duplicados.values() for r in v)
            msg_err = f"Hay valores duplicados en la capa {etiqueta} ({layer.nombre_shape}): {dup_desc}"
            emit(f"  [ERROR] {msg_err}", "ERROR")
            results.append(
                SpatialCheckResult(
                    control_id=f"DUP_{pref}",
                    grupo="1. Identificadores Duplicados",
                    capas_evaluadas=layer.nombre_shape,
                    estado="ERROR",
                    resumen_resultado=f"Se detectaron {len(duplicados)} identificador(es) duplicado(s).",
                    detalle_tecnico=dup_desc,
                )
            )
            issues.append(
                AuditIssue(
                    issue_id=f"DUP_ID_{layer.nombre_shape}",
                    capa=layer.nombre_shape,
                    prefijo_plantilla=pref,
                    categoria="ID_DUPLICADO",
                    severidad="ERROR",
                    campo_o_regla=campo_id,
                    valor_esperado="Identificadores únicos sin repetición",
                    valor_encontrado=dup_desc,
                    filas_afectadas=_fmt_rows(all_dup_rows),
                    mensaje=msg_err,
                    recomendacion=f"Corregir la numeración de '{campo_id}' en '{layer.nombre_shape}' para que cada entidad tenga un ID único.",
                )
            )

    # -------------------------------------------------------------------------
    # 2. ÁREA DE CORTA DEL PREDIO VS ÁREAS DE CORTA COMPILADAS (líneas 642-700 del R)
    # -------------------------------------------------------------------------
    shp_area_predio = get_layer("Area")
    if ruta_shp_area_comp and Path(ruta_shp_area_comp).exists():
        if shp_area_predio and shp_area_predio.read_ok:
            emit("Validando que el área de corta del predio no se traslape con el compilado de áreas de corta...", "HEADER")
            shp_comp = read_shapefile_complete(Path(ruta_shp_area_comp))
            if not shp_comp.read_ok:
                msg_w = f"No se pudo leer el shapefile compilado de áreas de corta: {shp_comp.error_msg}"
                emit(f"  [ADVERTENCIA] {msg_w}", "WARN")
                results.append(
                    SpatialCheckResult(
                        control_id="OVERLAP_AREA_COMP",
                        grupo="2. Área de Corta vs Compilado",
                        capas_evaluadas=f"{shp_area_predio.nombre_shape} vs {Path(ruta_shp_area_comp).name}",
                        estado="ADVERTENCIA",
                        resumen_resultado=msg_w,
                        detalle_tecnico=str(ruta_shp_area_comp),
                    )
                )
            else:
                tol_area = 1e-4
                any_touch = False
                real_overlaps: List[str] = []
                sup_m2_total = 0.0

                for i_pred, (reg_p, geom_p) in enumerate(zip(shp_area_predio.registros, shp_area_predio.geometrias), start=1):
                    if geom_p is None:
                        continue
                    for i_comp, (reg_c, geom_c) in enumerate(zip(shp_comp.registros, shp_comp.geometrias), start=1):
                        if geom_c is None:
                            continue
                        try:
                            if not geom_p.intersects(geom_c):
                                continue
                        except Exception:
                            continue
                        any_touch = True
                        a_int = _safe_intersection_area(geom_p, geom_c)
                        if a_int > tol_area:
                            sup_m2_total += a_int
                            n_area_val = reg_c.get("N_Area") or reg_p.get("N_Area") or f"Fila {i_pred}"
                            prov_val = reg_c.get("Provincia", "")
                            pas_val = reg_c.get("PAS", "")
                            exp_val = reg_c.get("EXP", "")
                            rev_val = reg_c.get("Rev", "")
                            real_overlaps.append(
                                f"N_Area={n_area_val} (PAS={pas_val}, EXP={exp_val}, Rev={rev_val}, Prov={prov_val}): {round(a_int, 1)} m²"
                            )

                if not any_touch:
                    msg_ok = "No hay intersección entre Area del predio y las áreas de corta compiladas."
                    emit(f"  [OK] {msg_ok}", "OK")
                    results.append(
                        SpatialCheckResult(
                            control_id="OVERLAP_AREA_COMP",
                            grupo="2. Área de Corta vs Compilado",
                            capas_evaluadas=f"{shp_area_predio.nombre_shape} vs {shp_comp.nombre_shape}",
                            estado="CONFORME",
                            resumen_resultado=msg_ok,
                            detalle_tecnico="0 intersecciones detectadas.",
                        )
                    )
                elif not real_overlaps:
                    msg_ok = "Solo hay contacto en el borde (sin traslape con área) entre Area del predio y las áreas de corta compiladas."
                    emit(f"  [OK] {msg_ok}", "OK")
                    results.append(
                        SpatialCheckResult(
                            control_id="OVERLAP_AREA_COMP",
                            grupo="2. Área de Corta vs Compilado",
                            capas_evaluadas=f"{shp_area_predio.nombre_shape} vs {shp_comp.nombre_shape}",
                            estado="CONFORME",
                            resumen_resultado=msg_ok,
                            detalle_tecnico="Intersección <= 0.0001 m² (contacto de borde).",
                        )
                    )
                else:
                    n_pol = len(real_overlaps)
                    sup_ha = round(sup_m2_total / 10000.0, 2)
                    msg_err = (
                        f"Hay traslape entre las áreas del predio y las áreas de corta compiladas: "
                        f"{n_pol} polígonos, superficie total {round(sup_m2_total, 1)} m² ({sup_ha} ha)."
                    )
                    emit(f"  [ERROR] {msg_err}", "ERROR")
                    for det_line in real_overlaps[:10]:
                        emit(f"     - {det_line}", "ERROR")
                    results.append(
                        SpatialCheckResult(
                            control_id="OVERLAP_AREA_COMP",
                            grupo="2. Área de Corta vs Compilado",
                            capas_evaluadas=f"{shp_area_predio.nombre_shape} vs {shp_comp.nombre_shape}",
                            estado="ERROR",
                            resumen_resultado=msg_err,
                            detalle_tecnico="; ".join(real_overlaps[:10]),
                        )
                    )
                    issues.append(
                        AuditIssue(
                            issue_id="SPATIAL_OVERLAP_AREA_COMP",
                            capa=shp_area_predio.nombre_shape,
                            prefijo_plantilla="Area",
                            categoria="TRASLAPE_AREAS_COMPILADAS",
                            severidad="ERROR",
                            campo_o_regla="st_intersection(Area, Compilado)",
                            valor_esperado="Sin traslape de superficie (<= 0.0001 m²)",
                            valor_encontrado=f"{n_pol} polígonos, {round(sup_m2_total, 1)} m² ({sup_ha} ha)",
                            filas_afectadas=f"{n_pol} intersecciones",
                            mensaje=msg_err,
                            recomendacion="Recortar o ajustar los polígonos de Area para eliminar la superposición con áreas de corta ya declaradas.",
                        )
                    )
        else:
            emit("  [ADVERTENCIA] No se encontró el shapefile Area de corta del predio para validar traslape.", "WARN")
            results.append(
                SpatialCheckResult(
                    control_id="OVERLAP_AREA_COMP",
                    grupo="2. Área de Corta vs Compilado",
                    capas_evaluadas="Area vs Compilado",
                    estado="NO_APLICA",
                    resumen_resultado="No se encontró el shapefile Area de corta del predio.",
                    detalle_tecnico="",
                )
            )
    else:
        results.append(
            SpatialCheckResult(
                control_id="OVERLAP_AREA_COMP",
                grupo="2. Área de Corta vs Compilado",
                capas_evaluadas="Area vs Compilado",
                estado="NO_APLICA",
                resumen_resultado="No se especificó ruta de shapefile compilado de áreas de corta (opcional).",
                detalle_tecnico="Parámetro opcional omitido o archivo no encontrado.",
            )
        )

    # -------------------------------------------------------------------------
    # 3. CONSISTENCIA ENTRE PARCELAS Y RODALES (líneas 703-733 del R)
    # -------------------------------------------------------------------------
    shp_parcela = get_layer("Parcela")
    shp_rodales = get_layer("Rodales")
    if shp_parcela and shp_parcela.read_ok and shp_rodales and shp_rodales.read_ok:
        emit("Validando consistencia entre Parcelas y Rodales...", "HEADER")
        filas_sin_rodal: List[int] = []
        filas_n_rodal_distinto: List[int] = []
        detalles_mismatch: List[str] = []

        for idx_p, (reg_p, geom_p) in enumerate(zip(shp_parcela.registros, shp_parcela.geometrias), start=1):
            if geom_p is None or getattr(geom_p, "is_empty", False):
                filas_sin_rodal.append(idx_p)
                continue

            rodal_match_idx: Optional[int] = None
            for idx_r, geom_r in enumerate(shp_rodales.geometrias):
                if geom_r is None or getattr(geom_r, "is_empty", False):
                    continue
                try:
                    if geom_p.within(geom_r) or geom_r.contains(geom_p):
                        rodal_match_idx = idx_r
                        break
                except Exception:
                    continue

            if rodal_match_idx is None:
                filas_sin_rodal.append(idx_p)
            else:
                n_rod_parc = str(reg_p.get("N_Rodal") or "").strip()
                n_rod_poly = str(shp_rodales.registros[rodal_match_idx].get("N_Rodal") or "").strip()
                if n_rod_parc != n_rod_poly:
                    filas_n_rodal_distinto.append(idx_p)
                    n_parc_id = reg_p.get("N_Parc", idx_p)
                    detalles_mismatch.append(
                        f"Parcela {n_parc_id} (fila {idx_p}): N_Rodal='{n_rod_parc}' vs Rodal contenedor='{n_rod_poly}'"
                    )

        if not filas_sin_rodal:
            emit("  [OK] Todas las parcelas están dentro de algún rodal.", "OK")
            results.append(
                SpatialCheckResult(
                    control_id="PARCELA_WITHIN_RODAL",
                    grupo="3. Parcelas y Rodales",
                    capas_evaluadas=f"{shp_parcela.nombre_shape} / {shp_rodales.nombre_shape}",
                    estado="CONFORME",
                    resumen_resultado="Todas las parcelas están espacialmente dentro de algún rodal.",
                    detalle_tecnico=f"Total parcelas evaluadas: {len(shp_parcela.registros)}",
                )
            )
        else:
            filas_str = _fmt_rows(filas_sin_rodal)
            msg_err = f"Parcelas fuera de rodales: {len(filas_sin_rodal)} (filas: {filas_str})"
            emit(f"  [ERROR] {msg_err}", "ERROR")
            results.append(
                SpatialCheckResult(
                    control_id="PARCELA_WITHIN_RODAL",
                    grupo="3. Parcelas y Rodales",
                    capas_evaluadas=f"{shp_parcela.nombre_shape} / {shp_rodales.nombre_shape}",
                    estado="ERROR",
                    resumen_resultado=msg_err,
                    detalle_tecnico=f"Filas fuera de todo polígono de Rodales: {filas_str}",
                )
            )
            issues.append(
                AuditIssue(
                    issue_id="SPATIAL_PARCELA_OUTSIDE_RODAL",
                    capa=shp_parcela.nombre_shape,
                    prefijo_plantilla="Parcela",
                    categoria="CONSISTENCIA_PARCELA_RODAL",
                    severidad="ERROR",
                    campo_o_regla="st_within(Parcela, Rodales)",
                    valor_esperado="Todas las parcelas dentro de un polígono de Rodales",
                    valor_encontrado=f"{len(filas_sin_rodal)} parcela(s) fuera de rodales",
                    filas_afectadas=filas_str,
                    mensaje=msg_err,
                    recomendacion="Verificar la ubicación espacial de las parcelas o la cobertura de la capa Rodales.",
                )
            )

        if not filas_n_rodal_distinto:
            emit("  [OK] N_Rodal en parcelas coincide con N_Rodal de rodales.", "OK")
            results.append(
                SpatialCheckResult(
                    control_id="PARCELA_NRODAL_MATCH",
                    grupo="3. Parcelas y Rodales",
                    capas_evaluadas=f"{shp_parcela.nombre_shape} / {shp_rodales.nombre_shape}",
                    estado="CONFORME",
                    resumen_resultado="El atributo N_Rodal en parcelas coincide con el N_Rodal del polígono que las contiene.",
                    detalle_tecnico="100% de coincidencia en parcelas contenidas.",
                )
            )
        else:
            filas_str = _fmt_rows(filas_n_rodal_distinto)
            msg_err = f"Hay {len(filas_n_rodal_distinto)} parcela(s) donde N_Rodal no coincide con el número de rodal contenedor."
            emit(f"  [ERROR] {msg_err}", "ERROR")
            results.append(
                SpatialCheckResult(
                    control_id="PARCELA_NRODAL_MATCH",
                    grupo="3. Parcelas y Rodales",
                    capas_evaluadas=f"{shp_parcela.nombre_shape} / {shp_rodales.nombre_shape}",
                    estado="ERROR",
                    resumen_resultado=msg_err,
                    detalle_tecnico="; ".join(detalles_mismatch[:10]),
                )
            )
            issues.append(
                AuditIssue(
                    issue_id="SPATIAL_PARCELA_NRODAL_MISMATCH",
                    capa=shp_parcela.nombre_shape,
                    prefijo_plantilla="Parcela",
                    categoria="CONSISTENCIA_PARCELA_RODAL",
                    severidad="ERROR",
                    campo_o_regla="Parcela.N_Rodal == Rodales.N_Rodal",
                    valor_esperado="Mismo N_Rodal que el polígono contenedor",
                    valor_encontrado="; ".join(detalles_mismatch[:5]),
                    filas_afectadas=filas_str,
                    mensaje=msg_err,
                    recomendacion="Actualizar el campo 'N_Rodal' en la capa Parcela según el rodal que la contiene espacialmente.",
                )
            )
    else:
        results.append(
            SpatialCheckResult(
                control_id="PARCELA_RODAL",
                grupo="3. Parcelas y Rodales",
                capas_evaluadas="Parcela / Rodales",
                estado="NO_APLICA",
                resumen_resultado="No se encontraron simultáneamente las capas Parcela y Rodales.",
                detalle_tecnico="",
            )
        )

    # -------------------------------------------------------------------------
    # 4. CONTENCIÓN EN LIMITE_PREDIAL Y AREA_PROYECTO (líneas 736-780 del R)
    # -------------------------------------------------------------------------
    shp_lim_pred = get_layer("Limite_Predial")
    shp_aproj = get_layer("Area_Proyecto")

    if (shp_lim_pred and shp_lim_pred.read_ok) or (shp_aproj and shp_aproj.read_ok):
        emit("Validando Area / Suelos / Rangos_pend / Pendientes_caminos vs Limite Predial y Area de Proyecto...", "HEADER")
        tol_contencion = 0.005  # 0.005 m² según línea 746 del R

        union_lim = _unify_geometries(shp_lim_pred.geometrias) if (shp_lim_pred and shp_lim_pred.read_ok) else None
        union_aproj = _unify_geometries(shp_aproj.geometrias) if (shp_aproj and shp_aproj.read_ok) else None

        for pref_capa in CAPAS_CONTENCION_PREDIO:
            capa_obj = get_layer(pref_capa)
            if not capa_obj or not capa_obj.read_ok:
                continue

            emit(f" - Capa: {capa_obj.nombre_shape}", "INFO")

            # 4a. Contra Limite_Predial
            if union_lim is not None:
                area_fuera_lim = _safe_difference_area(capa_obj.geometrias, union_lim)
                if area_fuera_lim <= 0.0:
                    msg_ok = "Sin superficie fuera de Limite_Predial"
                    emit(f"   [OK] {msg_ok}", "OK")
                    results.append(
                        SpatialCheckResult(
                            control_id=f"OUT_LIM_{pref_capa}",
                            grupo="4. Contención en Límite Predial",
                            capas_evaluadas=f"{capa_obj.nombre_shape} vs {shp_lim_pred.nombre_shape}",
                            estado="CONFORME",
                            resumen_resultado=msg_ok,
                            detalle_tecnico="0.000 m² fuera del límite predial.",
                        )
                    )
                elif area_fuera_lim <= tol_contencion:
                    msg_ok = f"Sin áreas fuera respecto a Limite_Predial ({round(area_fuera_lim, 3)} m² <= {tol_contencion} m²)"
                    emit(f"   [OK] {msg_ok}", "OK")
                    results.append(
                        SpatialCheckResult(
                            control_id=f"OUT_LIM_{pref_capa}",
                            grupo="4. Contención en Límite Predial",
                            capas_evaluadas=f"{capa_obj.nombre_shape} vs {shp_lim_pred.nombre_shape}",
                            estado="CONFORME",
                            resumen_resultado=msg_ok,
                            detalle_tecnico=f"Diferencia dentro de tolerancia ({round(area_fuera_lim, 4)} m²).",
                        )
                    )
                else:
                    sup_ha_fuera = round(area_fuera_lim / 10000.0, 4)
                    msg_err = f"Superficie fuera de Limite_Predial: {round(area_fuera_lim, 3)} m² ({sup_ha_fuera} ha)"
                    emit(f"   [ERROR] {msg_err}", "ERROR")
                    results.append(
                        SpatialCheckResult(
                            control_id=f"OUT_LIM_{pref_capa}",
                            grupo="4. Contención en Límite Predial",
                            capas_evaluadas=f"{capa_obj.nombre_shape} vs {shp_lim_pred.nombre_shape}",
                            estado="ERROR",
                            resumen_resultado=msg_err,
                            detalle_tecnico=f"Excede la tolerancia de {tol_contencion} m².",
                        )
                    )
                    issues.append(
                        AuditIssue(
                            issue_id=f"SPATIAL_OUT_LIMPRED_{capa_obj.nombre_shape}",
                            capa=capa_obj.nombre_shape,
                            prefijo_plantilla=pref_capa,
                            categoria="CONTENCION_LIMITE_PREDIAL",
                            severidad="ERROR",
                            campo_o_regla=f"st_difference({pref_capa}, Limite_Predial)",
                            valor_esperado=f"<= {tol_contencion} m² fuera de Limite_Predial",
                            valor_encontrado=f"{round(area_fuera_lim, 3)} m² ({sup_ha_fuera} ha)",
                            filas_afectadas="Geometría general",
                            mensaje=f"La capa '{capa_obj.nombre_shape}' sobresale del Límite Predial en {round(area_fuera_lim, 3)} m² ({sup_ha_fuera} ha).",
                            recomendacion="Ajustar o recortar (Clip) los bordes de la capa respecto a Limite_Predial.",
                        )
                    )

            # 4b. Contra Area_Proyecto
            if union_aproj is not None:
                area_fuera_ap = _safe_difference_area(capa_obj.geometrias, union_aproj)
                if area_fuera_ap <= 0.0:
                    msg_ok = "Sin superficie fuera respecto del Area de Proyecto"
                    emit(f"   [OK] {msg_ok}", "OK")
                    results.append(
                        SpatialCheckResult(
                            control_id=f"OUT_APROJ_{pref_capa}",
                            grupo="4. Contención en Área de Proyecto",
                            capas_evaluadas=f"{capa_obj.nombre_shape} vs {shp_aproj.nombre_shape}",
                            estado="CONFORME",
                            resumen_resultado=msg_ok,
                            detalle_tecnico="0.000 m² fuera del área de proyecto.",
                        )
                    )
                elif area_fuera_ap <= tol_contencion:
                    msg_ok = f"Sin áreas fuera respecto a Area del Proyecto ({round(area_fuera_ap, 3)} m² <= {tol_contencion} m²)"
                    emit(f"   [OK] {msg_ok}", "OK")
                    results.append(
                        SpatialCheckResult(
                            control_id=f"OUT_APROJ_{pref_capa}",
                            grupo="4. Contención en Área de Proyecto",
                            capas_evaluadas=f"{capa_obj.nombre_shape} vs {shp_aproj.nombre_shape}",
                            estado="CONFORME",
                            resumen_resultado=msg_ok,
                            detalle_tecnico=f"Diferencia dentro de tolerancia ({round(area_fuera_ap, 4)} m²).",
                        )
                    )
                else:
                    sup_ha_fuera = round(area_fuera_ap / 10000.0, 4)
                    msg_err = f"Superficie fuera de Area_Proyecto: {round(area_fuera_ap, 3)} m² ({sup_ha_fuera} ha)"
                    emit(f"   [ERROR] {msg_err}", "ERROR")
                    results.append(
                        SpatialCheckResult(
                            control_id=f"OUT_APROJ_{pref_capa}",
                            grupo="4. Contención en Área de Proyecto",
                            capas_evaluadas=f"{capa_obj.nombre_shape} vs {shp_aproj.nombre_shape}",
                            estado="ERROR",
                            resumen_resultado=msg_err,
                            detalle_tecnico=f"Excede la tolerancia de {tol_contencion} m².",
                        )
                    )
                    issues.append(
                        AuditIssue(
                            issue_id=f"SPATIAL_OUT_APROJ_{capa_obj.nombre_shape}",
                            capa=capa_obj.nombre_shape,
                            prefijo_plantilla=pref_capa,
                            categoria="CONTENCION_AREA_PROYECTO",
                            severidad="ERROR",
                            campo_o_regla=f"st_difference({pref_capa}, Area_Proyecto)",
                            valor_esperado=f"<= {tol_contencion} m² fuera de Area_Proyecto",
                            valor_encontrado=f"{round(area_fuera_ap, 3)} m² ({sup_ha_fuera} ha)",
                            filas_afectadas="Geometría general",
                            mensaje=f"La capa '{capa_obj.nombre_shape}' sobresale de Area_Proyecto en {round(area_fuera_ap, 3)} m² ({sup_ha_fuera} ha).",
                            recomendacion="Ajustar o recortar (Clip) los polígonos respecto a la capa Area_Proyecto.",
                        )
                    )

    # -------------------------------------------------------------------------
    # 5. PREDIOS VECINOS VS LIMITE PREDIAL (líneas 783-808 del R)
    # -------------------------------------------------------------------------
    shp_vecinos = get_layer("Predios_vecinos")
    if shp_lim_pred and shp_lim_pred.read_ok and shp_vecinos and shp_vecinos.read_ok:
        emit("Validando Predios_vecinos vs Limite_Predial...", "HEADER")
        tol_vecinos = 1e-4
        any_inter = False
        area_traslape_vec = 0.0
        n_traslapes_vec = 0

        for geom_v in shp_vecinos.geometrias:
            if geom_v is None:
                continue
            for geom_l in shp_lim_pred.geometrias:
                if geom_l is None:
                    continue
                try:
                    if not geom_v.intersects(geom_l):
                        continue
                except Exception:
                    continue
                any_inter = True
                a_int = _safe_intersection_area(geom_v, geom_l)
                if a_int > tol_vecinos:
                    area_traslape_vec += a_int
                    n_traslapes_vec += 1

        if not any_inter:
            msg_ok = "Predios_vecinos no se traslapan con Limite_Predial"
            emit(f"  [OK] {msg_ok}", "OK")
            results.append(
                SpatialCheckResult(
                    control_id="VECINOS_VS_LIMPRED",
                    grupo="5. Predios Vecinos vs Límite Predial",
                    capas_evaluadas=f"{shp_vecinos.nombre_shape} vs {shp_lim_pred.nombre_shape}",
                    estado="CONFORME",
                    resumen_resultado=msg_ok,
                    detalle_tecnico="Sin intersección geométrica.",
                )
            )
        elif n_traslapes_vec == 0:
            msg_ok = "Solo hay contacto en el borde (sin traslape con área) entre Predios_vecinos y Limite_Predial"
            emit(f"  [OK] {msg_ok}", "OK")
            results.append(
                SpatialCheckResult(
                    control_id="VECINOS_VS_LIMPRED",
                    grupo="5. Predios Vecinos vs Límite Predial",
                    capas_evaluadas=f"{shp_vecinos.nombre_shape} vs {shp_lim_pred.nombre_shape}",
                    estado="CONFORME",
                    resumen_resultado=msg_ok,
                    detalle_tecnico="Contacto de borde (área <= 0.0001 m²).",
                )
            )
        else:
            sup_ha_vec = round(area_traslape_vec / 10000.0, 2)
            msg_err = (
                f"Hay traslape entre Predios_vecinos y Limite_Predial "
                f"({n_traslapes_vec} intersecciones, {round(area_traslape_vec, 2)} m² / {sup_ha_vec} ha)"
            )
            emit(f"  [ERROR] {msg_err}", "ERROR")
            results.append(
                SpatialCheckResult(
                    control_id="VECINOS_VS_LIMPRED",
                    grupo="5. Predios Vecinos vs Límite Predial",
                    capas_evaluadas=f"{shp_vecinos.nombre_shape} vs {shp_lim_pred.nombre_shape}",
                    estado="ERROR",
                    resumen_resultado=msg_err,
                    detalle_tecnico=f"Superficie superpuesta: {round(area_traslape_vec, 2)} m².",
                )
            )
            issues.append(
                AuditIssue(
                    issue_id="SPATIAL_VECINOS_OVERLAP_LIMPRED",
                    capa=shp_vecinos.nombre_shape,
                    prefijo_plantilla="Predios_vecinos",
                    categoria="TRASLAPE_PREDIOS_VECINOS",
                    severidad="ERROR",
                    campo_o_regla="st_intersection(Predios_vecinos, Limite_Predial)",
                    valor_esperado="Sin traslape de superficie (<= 0.0001 m²)",
                    valor_encontrado=f"{round(area_traslape_vec, 2)} m² ({sup_ha_vec} ha)",
                    filas_afectadas=f"{n_traslapes_vec} polígonos",
                    mensaje=msg_err,
                    recomendacion="Ajustar la topología de Predios_vecinos para que comparta únicamente el borde con Limite_Predial.",
                )
            )

    # -------------------------------------------------------------------------
    # 6. INFRA_AISLADA (CONTROL DE EROSIÓN) NO ESTÉ SOBRE AREA_PROYECTO (líneas 810-849 del R)
    # -------------------------------------------------------------------------
    shp_infra = get_layer("Infra_aislada")
    if shp_infra and shp_infra.read_ok and shp_aproj and shp_aproj.read_ok:
        emit("Validando que Infra_aislada (Control de erosion) NO esté sobre Area_Proyecto...", "HEADER")
        cats_set = set(CATEGORIAS_INFRA_EROSION)
        union_aproj_infra = _unify_geometries(shp_aproj.geometrias)

        filas_evaluadas = 0
        filas_afectadas_infra: List[int] = []
        cats_afectadas: List[str] = []

        for idx_row, (reg_i, geom_i) in enumerate(zip(shp_infra.registros, shp_infra.geometrias), start=1):
            nom_infra = str(reg_i.get("Nom_Infra") or "").strip()
            if nom_infra not in cats_set:
                continue
            filas_evaluadas += 1
            if geom_i is None or union_aproj_infra is None:
                continue
            try:
                if geom_i.intersects(union_aproj_infra):
                    filas_afectadas_infra.append(idx_row)
                    if nom_infra not in cats_afectadas:
                        cats_afectadas.append(nom_infra)
            except Exception:
                pass

        if filas_evaluadas == 0:
            msg_info = "No hay elementos de las categorías de control de erosión a revisar en Infra_aislada."
            emit(f"  [INFO] {msg_info}", "INFO")
            results.append(
                SpatialCheckResult(
                    control_id="INFRA_SOBRE_APROJ",
                    grupo="6. Infra_aislada vs Área de Proyecto",
                    capas_evaluadas=f"{shp_infra.nombre_shape} vs {shp_aproj.nombre_shape}",
                    estado="CONFORME",
                    resumen_resultado=msg_info,
                    detalle_tecnico=f"Categorías buscadas: {', '.join(CATEGORIAS_INFRA_EROSION)}",
                )
            )
        elif not filas_afectadas_infra:
            msg_ok = "Ningún elemento de las categorías revisadas se ubica sobre el Area_Proyecto."
            emit(f"  [OK] {msg_ok}", "OK")
            results.append(
                SpatialCheckResult(
                    control_id="INFRA_SOBRE_APROJ",
                    grupo="6. Infra_aislada vs Área de Proyecto",
                    capas_evaluadas=f"{shp_infra.nombre_shape} vs {shp_aproj.nombre_shape}",
                    estado="CONFORME",
                    resumen_resultado=msg_ok,
                    detalle_tecnico=f"Elementos evaluados: {filas_evaluadas}",
                )
            )
        else:
            filas_str = _fmt_rows(filas_afectadas_infra)
            msg_err = (
                f"Hay {len(filas_afectadas_infra)} elemento(s) de las categorías revisadas sobre el Area_Proyecto "
                f"(Categorías: {', '.join(cats_afectadas)})"
            )
            emit(f"  [ERROR] {msg_err}", "ERROR")
            emit(f"     Filas afectadas en Infra_aislada: {filas_str}", "ERROR")
            results.append(
                SpatialCheckResult(
                    control_id="INFRA_SOBRE_APROJ",
                    grupo="6. Infra_aislada vs Área de Proyecto",
                    capas_evaluadas=f"{shp_infra.nombre_shape} vs {shp_aproj.nombre_shape}",
                    estado="ERROR",
                    resumen_resultado=msg_err,
                    detalle_tecnico=f"Filas: {filas_str} | Categorías: {', '.join(cats_afectadas)}",
                )
            )
            issues.append(
                AuditIssue(
                    issue_id="SPATIAL_INFRA_ON_APROJ",
                    capa=shp_infra.nombre_shape,
                    prefijo_plantilla="Infra_aislada",
                    categoria="INFRA_SOBRE_AREA_PROYECTO",
                    severidad="ERROR",
                    campo_o_regla="st_intersects(Infra_aislada, Area_Proyecto)",
                    valor_esperado="Obras de control de erosión fuera de Area_Proyecto",
                    valor_encontrado=f"Categorías sobre Area_Proyecto: {', '.join(cats_afectadas)}",
                    filas_afectadas=filas_str,
                    mensaje=msg_err,
                    recomendacion="Reubicar o revisar las obras de control de erosión que se superponen con Area_Proyecto.",
                )
            )

    # -------------------------------------------------------------------------
    # 7. CONSISTENCIA DEL NOMBRE DEL PREDIO (Nom_Predio) EN TODAS LAS CAPAS (líneas 852-882 del R)
    # -------------------------------------------------------------------------
    nom_ref: Optional[str] = None
    if shp_lim_pred and shp_lim_pred.read_ok:
        emit("Validando nombre de predio (Nom_Predio) en todas las capas respecto de Limite_Predial...", "HEADER")
        nom_vals = [
            str(r.get("Nom_Predio")).strip()
            for r in shp_lim_pred.registros
            if r.get("Nom_Predio") is not None and str(r.get("Nom_Predio")).strip() != ""
        ]
        if not nom_vals:
            msg_err = "No se pudo determinar Nom_Predio de referencia en Limite_Predial (campo vacío o inexistente)."
            emit(f"  [ERROR] {msg_err}", "ERROR")
            results.append(
                SpatialCheckResult(
                    control_id="NOM_PREDIO_REF",
                    grupo="7. Consistencia de Nom_Predio",
                    capas_evaluadas=shp_lim_pred.nombre_shape,
                    estado="ERROR",
                    resumen_resultado=msg_err,
                    detalle_tecnico="",
                )
            )
        else:
            nom_ref = Counter(nom_vals).most_common(1)[0][0]
            emit(f"   Nom_Predio de referencia (desde Limite_Predial): '{nom_ref}'", "INFO")
            excluidas_norm = {normalize_layer_key(x) for x in CAPAS_EXCLUIDAS_NOM_PREDIO}

            for shp_cap in all_shps:
                if not shp_cap.read_ok:
                    continue
                pref_norm = normalize_layer_key(shp_cap.prefijo_asociado or shp_cap.nombre_shape)
                if pref_norm in excluidas_norm:
                    continue
                if "Nom_Predio" not in shp_cap.nombres_campos:
                    continue

                idx_malos_nom: List[int] = []
                valores_distintos: List[str] = []
                for idx_row, reg in enumerate(shp_cap.registros, start=1):
                    val = reg.get("Nom_Predio")
                    s_val = "" if val is None else str(val).strip()
                    if not s_val or s_val.upper() == "NA" or s_val != nom_ref:
                        idx_malos_nom.append(idx_row)
                        muestra = s_val if s_val else "Vacío/NA"
                        if muestra not in valores_distintos:
                            valores_distintos.append(muestra)

                if not idx_malos_nom:
                    emit(f"  [OK] Capa {shp_cap.nombre_shape}: coincide el nombre del Predio ('{nom_ref}').", "OK")
                    results.append(
                        SpatialCheckResult(
                            control_id=f"NOM_PREDIO_{shp_cap.nombre_shape}",
                            grupo="7. Consistencia de Nom_Predio",
                            capas_evaluadas=shp_cap.nombre_shape,
                            estado="CONFORME",
                            resumen_resultado=f"Coincide Nom_Predio ('{nom_ref}') en todas las filas.",
                            detalle_tecnico=f"{len(shp_cap.registros)} registros verificados.",
                        )
                    )
                else:
                    filas_str = _fmt_rows(idx_malos_nom)
                    msg_err = (
                        f"Capa {shp_cap.nombre_shape}: Nombre del Predio es distinto o vacío en filas: {filas_str} "
                        f"(Encontrado: {', '.join(valores_distintos[:5])})"
                    )
                    emit(f"  [ERROR] {msg_err}", "ERROR")
                    results.append(
                        SpatialCheckResult(
                            control_id=f"NOM_PREDIO_{shp_cap.nombre_shape}",
                            grupo="7. Consistencia de Nom_Predio",
                            capas_evaluadas=shp_cap.nombre_shape,
                            estado="ERROR",
                            resumen_resultado=msg_err,
                            detalle_tecnico=f"Esperado: '{nom_ref}' | Encontrado: {', '.join(valores_distintos[:5])}",
                        )
                    )
                    issues.append(
                        AuditIssue(
                            issue_id=f"NOM_PREDIO_MISMATCH_{shp_cap.nombre_shape}",
                            capa=shp_cap.nombre_shape,
                            prefijo_plantilla=shp_cap.prefijo_asociado or "N/A",
                            categoria="CONSISTENCIA_NOM_PREDIO",
                            severidad="ERROR",
                            campo_o_regla="Nom_Predio",
                            valor_esperado=nom_ref,
                            valor_encontrado=", ".join(valores_distintos[:5]),
                            filas_afectadas=filas_str,
                            mensaje=msg_err,
                            recomendacion=f"Homologar el campo 'Nom_Predio' al valor de referencia '{nom_ref}' de Limite_Predial.",
                        )
                    )

    return results, issues, nom_ref


# ==============================================================================
# SECCIÓN 5: MOTOR ORQUESTADOR DE AUDITORÍA (PAS151AuditEngine)
# ==============================================================================
"""
Motor Orquestador de la Auditoría PAS 151 CONAF (`PAS151AuditEngine`).
Coordina:
1. Carga de plantillas CONAF (integradas Rev11 + Excel opcional).
2. Detección automática o explícita de `nombre_predio` y emparejamiento inteligente de Shapefiles.
3. Inventario de capas esperadas, encontradas, faltantes y adicionales.
4. Ciclo de validación individual por Shapefile (CRS, 2D/ZM, estructura, orden, dominios, Sup_ha, Coord_X/Y, texto, topología).
5. Validaciones espaciales e inter-capas (duplicados, traslape con compilado, Parcela-Rodal, contención en Límite Predial y Área Proyecto, Predios vecinos, Infra_aislada, consistencia de Nom_Predio).
"""

from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple, Union







@dataclass
class PAS151AuditResult:
    """Resultado consolidado de toda la corrida de auditoría PAS 151."""
    timestamp: datetime = field(default_factory=datetime.now)
    directorio: str = ""
    nombre_predio: str = ""
    nom_predio_referencia_atrib: Optional[str] = None
    ruta_shp_area_comp: Optional[str] = None
    archivo_estandar: str = "PLANTILLA_INTERNA_PAS151_REV11"
    epsg_esperado: int = 32719

    # Inventario de archivos
    shapes_en_carpeta: List[str] = field(default_factory=list)
    shapes_auditados: List[str] = field(default_factory=list)
    capas_faltantes: List[str] = field(default_factory=list)
    capas_adicionales: List[str] = field(default_factory=list)

    # Resultados por capa y tabla R
    resumen_capas: List[LayerValidationSummary] = field(default_factory=list)
    filas_tabla_r: List[RFieldRow] = field(default_factory=list)

    # Validaciones espaciales e incidencias consolidadas
    resultados_espaciales: List[SpatialCheckResult] = field(default_factory=list)
    incidencias: List[AuditIssue] = field(default_factory=list)
    plantillas_usadas: Dict[str, LayerTemplate] = field(default_factory=dict)

    @property
    def total_errores(self) -> int:
        return sum(1 for i in self.incidencias if i.severidad == "ERROR")

    @property
    def total_advertencias(self) -> int:
        return sum(1 for i in self.incidencias if i.severidad == "ADVERTENCIA")

    @property
    def total_capas_conformes(self) -> int:
        return sum(1 for c in self.resumen_capas if c.estado_global == "CONFORME")

    @property
    def total_capas_con_error(self) -> int:
        return sum(1 for c in self.resumen_capas if "ERROR" in c.estado_global)


def match_shapefile_to_template(
    shp_stem: str,
    templates: Dict[str, LayerTemplate],
    nombre_predio: str,
) -> Tuple[Optional[LayerTemplate], str, Optional[str]]:
    """
    Asocia el nombre de un archivo .shp (sin extensión) con su plantilla CONAF.
    Retorna: (LayerTemplate | None, tipo_coincidencia, sufijo_predio_detectado | None)
    Orden de prioridad:
    1. Coincidencia exacta `<Prefijo>_<nombre_predio>` -> 'EXACT_PREDIO'
    2. Coincidencia exacta `<Prefijo>` (sin sufijo) -> 'EXACT_PREFIX'
    3. Coincidencia por prefijo más largo `<Prefijo>_<Sufijo>` -> 'AUTO_PREFIX'
    """
    stem_norm = normalize_layer_key(shp_stem)
    predio_norm = normalize_layer_key(nombre_predio) if nombre_predio else ""

    # Ordenar plantillas desde el prefijo más largo al más corto para evitar que 'Area' capture 'Area_Proyecto'
    sorted_items = sorted(templates.items(), key=lambda kv: len(kv[0]), reverse=True)

    # 1. Exacta con `<Prefijo>_<nombre_predio>`
    if predio_norm:
        for key_norm, tmpl in sorted_items:
            expected_norm = f"{key_norm}_{predio_norm}"
            if stem_norm == expected_norm:
                return tmpl, "EXACT_PREDIO", nombre_predio

    # 2. Exacta solo con `<Prefijo>`
    for key_norm, tmpl in sorted_items:
        if stem_norm == key_norm:
            return tmpl, "EXACT_PREFIX", None

    # 3. Coincidencia por prefijo `<Prefijo>_<Sufijo>`
    for key_norm, tmpl in sorted_items:
        prefix_with_sep = f"{key_norm}_"
        if stem_norm.startswith(prefix_with_sep):
            # Extraer el sufijo original conservando mayúsculas/minúsculas del archivo
            suffix_len = len(stem_norm) - len(prefix_with_sep)
            detected_suffix = shp_stem[-suffix_len:] if suffix_len > 0 else None
            return tmpl, "AUTO_PREFIX", detected_suffix

    return None, "NO_MATCH", None


def infer_nombre_predio(shp_files: List[Path], templates: Dict[str, LayerTemplate]) -> str:
    """
    Infiere automáticamente el nombre del predio a partir de los nombres de los archivos .shp
    (priorizando `Limite_Predial_<PREDIO>.shp` o la moda de los sufijos detectados).
    """
    suffixes: List[str] = []
    lim_pred_suffix: Optional[str] = None

    for p in shp_files:
        tmpl, mtype, det_suffix = match_shapefile_to_template(p.stem, templates, "")
        if tmpl and det_suffix:
            suffixes.append(det_suffix)
            if normalize_layer_key(tmpl.prefijo) == "LIMITE_PREDIAL":
                lim_pred_suffix = det_suffix

    if lim_pred_suffix:
        return lim_pred_suffix
    if suffixes:
        return Counter(suffixes).most_common(1)[0][0]
    return "SIN_SUFIJO"


class PAS151AuditEngine:
    """Motor principal de validación PAS 151 CONAF para ArcGIS Pro y ejecución standalone."""

    def run(
        self,
        directorio: Union[str, Path],
        nombre_predio: Optional[str] = None,
        ruta_shp_area_comp: Optional[Union[str, Path]] = None,
        excel_estandar_path: Optional[Union[str, Path]] = None,
        solo_excel: bool = False,
        epsg_esperado: int = 32719,
        buscar_subcarpetas: bool = False,
        log_fn: Optional[Callable[[str, str], None]] = None,
    ) -> PAS151AuditResult:
        def emit(msg: str, level: str = "INFO") -> None:
            if log_fn:
                log_fn(msg, level)

        dir_path = Path(directorio)
        templates = load_templates(excel_estandar_path, solo_excel=solo_excel)

        # Buscar archivos .shp (evitando .shp.xml)
        if buscar_subcarpetas:
            shp_files = sorted(
                p for p in dir_path.rglob("*")
                if p.is_file() and p.suffix.lower() == ".shp" and not p.name.lower().endswith(".shp.xml")
            )
        else:
            shp_files = sorted(
                p for p in dir_path.iterdir()
                if p.is_file() and p.suffix.lower() == ".shp" and not p.name.lower().endswith(".shp.xml")
            )
            # Si la carpeta raíz no tiene .shp directos pero tiene una subcarpeta SHP, buscar recursivamente
            if not shp_files and dir_path.exists():
                shp_files = sorted(
                    p for p in dir_path.rglob("*")
                    if p.is_file() and p.suffix.lower() == ".shp" and not p.name.lower().endswith(".shp.xml")
                )

        predio_clean = (nombre_predio or "").strip()
        if not predio_clean or predio_clean.upper() == "AUTO":
            predio_clean = infer_nombre_predio(shp_files, templates)
            emit(f"Nombre de predio autodetectado desde los archivos: '{predio_clean}'", "INFO")
        else:
            emit(f"Nombre de predio configurado: '{predio_clean}'", "INFO")

        result = PAS151AuditResult(
            directorio=str(dir_path),
            nombre_predio=predio_clean,
            ruta_shp_area_comp=str(ruta_shp_area_comp) if ruta_shp_area_comp else None,
            archivo_estandar=str(excel_estandar_path) if excel_estandar_path else "PLANTILLA_INTERNA_PAS151_REV11",
            epsg_esperado=epsg_esperado,
            shapes_en_carpeta=[p.stem for p in shp_files],
            plantillas_usadas=templates,
        )

        # Clasificar archivos encontrados vs plantillas
        matched_shps: List[Tuple[Path, LayerTemplate, str]] = []
        found_template_keys = set()
        adicionales: List[str] = []

        for shp_path in shp_files:
            tmpl, mtype, _ = match_shapefile_to_template(shp_path.stem, templates, predio_clean)
            if tmpl is not None:
                matched_shps.append((shp_path, tmpl, mtype))
                found_template_keys.add(normalize_layer_key(tmpl.prefijo))
            else:
                adicionales.append(shp_path.stem)

        # Capas esperadas según plantilla que no están en la carpeta
        faltantes: List[str] = []
        for key_norm, tmpl in templates.items():
            if key_norm not in found_template_keys:
                if predio_clean and predio_clean != "SIN_SUFIJO":
                    faltantes.append(f"{tmpl.prefijo}_{predio_clean}")
                else:
                    faltantes.append(tmpl.prefijo)

        result.capas_faltantes = faltantes
        result.capas_adicionales = adicionales

        emit("\n======================================================================", "HEADER")
        emit("REVISIÓN DE ARCHIVOS SHAPEFILES EN EL DIRECTORIO", "HEADER")
        emit("======================================================================", "HEADER")

        if not faltantes:
            emit("[OK] Están todos los shapes de la plantilla a entregar.", "OK")
        else:
            emit(f"[INFO] Capas de la plantilla general no presentes en esta entrega ({len(faltantes)}):", "WARN")
            emit("   " + ", ".join(faltantes[:15]) + (" ..." if len(faltantes) > 15 else ""), "WARN")

        if not adicionales:
            emit("[OK] No hay shapes adicionales fuera de los definidos en la plantilla.", "OK")
        else:
            emit(f"[ERROR] Hay shapes adicionales en la carpeta que no están en la plantilla ({len(adicionales)}):", "ERROR")
            emit("   " + ", ".join(adicionales), "ERROR")
            for adic in adicionales:
                result.incidencias.append(
                    AuditIssue(
                        issue_id=f"EXTRA_SHP_{adic}",
                        capa=adic,
                        prefijo_plantilla="SIN_PLANTILLA",
                        categoria="ARCHIVO_ADICIONAL",
                        severidad="ADVERTENCIA",
                        campo_o_regla="Inventario de Capas",
                        valor_esperado="Solo capas del protocolo CONAF PAS 151",
                        valor_encontrado=f"{adic}.shp",
                        filas_afectadas="Archivo completo",
                        mensaje=f"El archivo '{adic}.shp' no corresponde a ningún prefijo de la plantilla PAS 151.",
                        recomendacion="Verificar el nombre del archivo o retirarlo de la carpeta de entrega CONAF.",
                    )
                )

        # Leer y validar cada shapefile asociado
        emit("\n======================================================================", "HEADER")
        emit("CICLO DE VALIDACIÓN INDIVIDUAL SEGÚN PLANTILLA CONAF", "HEADER")
        emit("======================================================================", "HEADER")

        by_prefix: Dict[str, ShapefileData] = {}
        all_loaded_shps: List[ShapefileData] = []

        for shp_path, tmpl, mtype in matched_shps:
            shp_data = read_shapefile_complete(shp_path)
            shp_data.prefijo_asociado = tmpl.prefijo
            shp_data.tipo_coincidencia = mtype
            all_loaded_shps.append(shp_data)

            # Si hay dos versiones (una con sufijo del predio y otra sin sufijo), priorizar EXACT_PREDIO
            key_norm = normalize_layer_key(tmpl.prefijo)
            if key_norm not in by_prefix or mtype == "EXACT_PREDIO":
                by_prefix[key_norm] = shp_data

            if not shp_data.read_ok:
                emit(f"[ERROR] No se pudo leer '{shp_path.name}': {shp_data.error_msg}", "ERROR")
                result.incidencias.append(
                    AuditIssue(
                        issue_id=f"READ_ERR_{shp_path.stem}",
                        capa=shp_path.stem,
                        prefijo_plantilla=tmpl.prefijo,
                        categoria="ERROR_LECTURA",
                        severidad="ERROR",
                        campo_o_regla="Lectura SHP/DBF",
                        valor_esperado="Archivo íntegro",
                        valor_encontrado=shp_data.error_msg or "Error de lectura",
                        filas_afectadas="Todas",
                        mensaje=f"No fue posible leer '{shp_path.name}': {shp_data.error_msg}",
                        recomendacion="Verificar que existan los archivos .shp, .shx y .dbf y que no estén corruptos.",
                    )
                )
                continue

            # Advertir si el archivo no lleva el sufijo del predio cuando otros sí lo llevan
            if predio_clean and predio_clean != "SIN_SUFIJO" and mtype == "EXACT_PREFIX" and tmpl.prefijo != "Limite_Comunal":
                result.incidencias.append(
                    AuditIssue(
                        issue_id=f"NAMING_SUFFIX_{shp_data.nombre_shape}",
                        capa=shp_data.nombre_shape,
                        prefijo_plantilla=tmpl.prefijo,
                        categoria="NOMENCLATURA_ARCHIVO",
                        severidad="ADVERTENCIA",
                        campo_o_regla="Nombre de Archivo",
                        valor_esperado=f"{tmpl.prefijo}_{predio_clean}.shp",
                        valor_encontrado=f"{shp_data.nombre_shape}.shp",
                        filas_afectadas="Archivo",
                        mensaje=f"El archivo '{shp_data.nombre_shape}.shp' no incluye el sufijo del predio '_{predio_clean}'.",
                        recomendacion=f"Renombrar el shapefile a '{tmpl.prefijo}_{predio_clean}.shp' para cumplir la convención de entrega.",
                    )
                )

            layer_summary = validate_single_shapefile(
                shp_data=shp_data,
                plantilla=tmpl,
                epsg_esperado=epsg_esperado,
                log_fn=log_fn,
            )
            result.shapes_auditados.append(shp_data.nombre_shape)
            result.resumen_capas.append(layer_summary)
            result.filas_tabla_r.extend(layer_summary.filas_resultado_r)
            result.incidencias.extend(layer_summary.incidencias)

        # Ejecutar las 7 validaciones espaciales y entre capas
        comp_path = Path(ruta_shp_area_comp) if ruta_shp_area_comp else None
        spatial_results, spatial_issues, nom_ref = run_all_spatial_validations(
            by_prefix=by_prefix,
            all_shps=all_loaded_shps,
            ruta_shp_area_comp=comp_path,
            log_fn=log_fn,
        )
        result.resultados_espaciales = spatial_results
        result.incidencias.extend(spatial_issues)
        result.nom_predio_referencia_atrib = nom_ref

        return result


# ==============================================================================
# SECCIÓN 6: EXPORTADOR DE REPORTE EXCEL MULTI-HOJA (PAS151ExcelExporter)
# ==============================================================================
"""
Exportador de Reporte Excel Multi-Hoja para la Auditoría PAS 151 CONAF.
Genera el archivo `00_Reporte_revision_campos_<PREDIO>.xlsx` con 6 hojas:
1. `Revision_Campos_R`: Tabla idéntica a `resultados_df` del script R (`write_xlsx`).
2. `Resumen_Ejecutivo`: KPIs generales del predio, archivos y estado global.
3. `Validacion_Capas`: Matriz de cumplimiento por shapefile (CRS, 2D, Orden, Dominios, Sup_ha, Coord_XY, Texto, Geometría).
4. `Validaciones_Espaciales`: Resultados de los 7 controles espaciales e inter-capas.
5. `Incidencias_Detalladas`: Listado completo de errores y advertencias con filas afectadas y recomendación.
6. `Plantilla_CONAF_PAS151`: Diccionario de datos y dominios oficiales aplicados.
"""

from pathlib import Path
from typing import Union
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter





class PAS151ExcelExporter:
    """Genera el libro Excel consolidado de auditoría PAS 151 CONAF."""

    def __init__(self) -> None:
        self.header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        self.header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

        self.sub_fill = PatternFill(start_color="2E5B88", end_color="2E5B88", fill_type="solid")
        self.sub_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

        self.title_font = Font(name="Segoe UI", size=14, bold=True, color="1B365D")
        self.normal_font = Font(name="Segoe UI", size=10)
        self.bold_font = Font(name="Segoe UI", size=10, bold=True)

        self.fill_ok = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")
        self.fill_err = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
        self.fill_warn = PatternFill(start_color="FCF3CF", end_color="FCF3CF", fill_type="solid")
        self.fill_na = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")

        thin = Side(border_style="thin", color="D5D8DC")
        self.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def export(self, result: PAS151AuditResult, output_path: Union[str, Path]) -> Path:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        self._build_sheet_tabla_r(wb, result)
        self._build_sheet_resumen(wb, result)
        self._build_sheet_validacion_capas(wb, result)
        self._build_sheet_espaciales(wb, result)
        self._build_sheet_incidencias(wb, result)
        self._build_sheet_plantilla(wb, result)

        wb.save(out_file)
        return out_file

    def _build_sheet_tabla_r(self, wb: openpyxl.Workbook, result: PAS151AuditResult) -> None:
        """
        Hoja 1: Replica exactamente la estructura de columnas que genera el script R:
        shape | nombre | Tipo_Plant | Tipo_Shp | Existe_Campo | Coin_Tipo_Campo |
        Long_Plant | Long_Shp | Long_Correcta | Dec_Plant | Dec_Shp | Dec_Correcto | orden_correcto
        """
        ws = wb.create_sheet(title="Revision_Campos_R")
        headers = [
            "shape",
            "nombre",
            "Tipo_Plant",
            "Tipo_Shp",
            "Existe_Campo",
            "Coin_Tipo_Campo",
            "Long_Plant",
            "Long_Shp",
            "Long_Correcta",
            "Dec_Plant",
            "Dec_Shp",
            "Dec_Correcto",
            "orden_correcto",
        ]
        ws.append(headers)
        self._style_header(ws)

        for r in result.filas_tabla_r:
            row_vals = [
                r.shape,
                r.nombre,
                r.Tipo_Plant,
                r.Tipo_Shp if r.Tipo_Shp is not None else "NA",
                r.Existe_Campo,
                r.Coin_Tipo_Campo if r.Coin_Tipo_Campo is not None else "NA",
                r.Long_Plant if r.Long_Plant is not None else "NA",
                r.Long_Shp if r.Long_Shp is not None else "NA",
                r.Long_Correcta if r.Long_Correcta is not None else "NA",
                r.Dec_Plant if r.Dec_Plant is not None else "NA",
                r.Dec_Shp if r.Dec_Shp is not None else "NA",
                r.Dec_Correcto if r.Dec_Correcto is not None else "NA",
                r.orden_correcto,
            ]
            ws.append(row_vals)
            row_idx = ws.max_row
            for col_idx in range(1, len(headers) + 1):
                c = ws.cell(row=row_idx, column=col_idx)
                c.font = self.normal_font
                c.border = self.border

            # Colorear celdas booleanas / de estado
            if r.Existe_Campo == "No":
                ws.cell(row=row_idx, column=5).fill = self.fill_err
            else:
                ws.cell(row=row_idx, column=5).fill = self.fill_ok

            for col_i, val_b in [(6, r.Coin_Tipo_Campo), (9, r.Long_Correcta), (12, r.Dec_Correcto), (13, r.orden_correcto)]:
                cell = ws.cell(row=row_idx, column=col_i)
                if val_b is True:
                    cell.fill = self.fill_ok
                elif val_b is False:
                    cell.fill = self.fill_err
                else:
                    cell.fill = self.fill_na

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        self._autofit(ws)

    def _build_sheet_resumen(self, wb: openpyxl.Workbook, result: PAS151AuditResult) -> None:
        ws = wb.create_sheet(title="Resumen_Ejecutivo")
        ws.cell(row=1, column=1, value="CONTROL DE CALIDAD CARTOGRÁFICO PAS 151 - PROTOCOLO CONAF").font = self.title_font
        ws.cell(row=2, column=1, value=f"Fecha de Ejecución: {result.timestamp.strftime('%Y-%m-%d %H:%M:%S')}").font = self.normal_font
        ws.cell(row=3, column=1, value=f"Directorio Auditado: {result.directorio}").font = self.normal_font
        ws.cell(row=4, column=1, value=f"Predio (Sufijo Archivos): {result.nombre_predio}").font = self.bold_font
        ws.cell(
            row=5,
            column=1,
            value=f"Nom_Predio Referencia (Limite_Predial): {result.nom_predio_referencia_atrib or 'No disponible'}",
        ).font = self.bold_font
        ws.cell(row=6, column=1, value=f"Compilado Áreas de Corta: {result.ruta_shp_area_comp or 'No especificado'}").font = self.normal_font

        ws.cell(row=8, column=1, value="Indicador General").font = self.sub_font
        ws.cell(row=8, column=1).fill = self.sub_fill
        ws.cell(row=8, column=2, value="Valor").font = self.sub_font
        ws.cell(row=8, column=2).fill = self.sub_fill

        kpis = [
            ("Total Archivos .shp en Carpeta", len(result.shapes_en_carpeta)),
            ("Total Capas Asociadas y Auditadas", len(result.shapes_auditados)),
            ("Capas Conformes (Sin Errores ni Advertencias)", result.total_capas_conformes),
            ("Capas con Errores", result.total_capas_con_error),
            ("Capas Adicionales (Fuera de Plantilla)", len(result.capas_adicionales)),
            ("Total Incidencias de ERROR", result.total_errores),
            ("Total Incidencias de ADVERTENCIA", result.total_advertencias),
        ]

        for idx, (label, val) in enumerate(kpis, start=9):
            c1 = ws.cell(row=idx, column=1, value=label)
            c2 = ws.cell(row=idx, column=2, value=val)
            c1.font = self.normal_font
            c2.font = self.bold_font
            c1.border = self.border
            c2.border = self.border
            if "Error" in label and val > 0:
                c2.fill = self.fill_err
            elif "Conformes" in label and val > 0:
                c2.fill = self.fill_ok
            elif ("Adicionales" in label or "ADVERTENCIA" in label) and val > 0:
                c2.fill = self.fill_warn

        self._autofit(ws)

    def _build_sheet_validacion_capas(self, wb: openpyxl.Workbook, result: PAS151AuditResult) -> None:
        ws = wb.create_sheet(title="Validacion_Capas")
        headers = [
            "Shapefile",
            "Plantilla CONAF",
            "Registros",
            "CRS (EPSG:32719)",
            "Geometría 2D (Sin Z/M)",
            "Estructura Campos",
            "Orden Campos",
            "Dominios CONAF",
            "Sin Vacíos/NA",
            "Sup_ha vs Geometría",
            "Coord_X / Coord_Y",
            "Texto Sin Acentos",
            "Texto Sin Doble Espacio",
            "Validez Topológica",
            "Estado Global",
        ]
        ws.append(headers)
        self._style_header(ws)

        def fmt_bool(b):
            if b is True:
                return "OK"
            if b is False:
                return "ERROR"
            return "N/A"

        for s in result.resumen_capas:
            row = [
                s.nombre_shape,
                s.prefijo_plantilla,
                s.num_registros,
                "OK" if s.crs_ok else f"ERROR ({s.crs_detalle})",
                "OK" if s.es_2d_ok else f"ERROR ({s.geometria_detalle})",
                fmt_bool(s.estructura_campos_ok),
                fmt_bool(s.orden_campos_ok),
                fmt_bool(s.dominios_ok),
                fmt_bool(s.campos_sin_vacios_ok),
                fmt_bool(s.sup_ha_ok),
                fmt_bool(s.coord_xy_ok),
                fmt_bool(s.texto_caracteres_ok),
                fmt_bool(s.texto_doble_espacio_ok),
                "OK" if s.geometrias_validas_ok else f"ERROR ({s.num_geometrias_invalidas} inv.)",
                s.estado_global,
            ]
            ws.append(row)
            r_idx = ws.max_row
            for col_i in range(1, len(headers) + 1):
                cell = ws.cell(row=r_idx, column=col_i)
                cell.font = self.normal_font
                cell.border = self.border
                val_s = str(cell.value or "")
                if col_i >= 4:
                    if val_s.startswith("OK") or val_s == "CONFORME":
                        cell.fill = self.fill_ok
                    elif "ERROR" in val_s:
                        cell.fill = self.fill_err
                    elif "ADVERTENCIA" in val_s:
                        cell.fill = self.fill_warn
                    elif val_s == "N/A":
                        cell.fill = self.fill_na

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        self._autofit(ws)

    def _build_sheet_espaciales(self, wb: openpyxl.Workbook, result: PAS151AuditResult) -> None:
        ws = wb.create_sheet(title="Validaciones_Espaciales")
        headers = [
            "ID Control",
            "Grupo de Validación Espacial",
            "Capas Evaluadas",
            "Estado",
            "Resultado",
            "Detalle Técnico",
        ]
        ws.append(headers)
        self._style_header(ws)

        for sp in result.resultados_espaciales:
            ws.append([
                sp.control_id,
                sp.grupo,
                sp.capas_evaluadas,
                sp.estado,
                sp.resumen_resultado,
                sp.detalle_tecnico,
            ])
            r_idx = ws.max_row
            for col_i in range(1, len(headers) + 1):
                c = ws.cell(row=r_idx, column=col_i)
                c.font = self.normal_font
                c.border = self.border
            est_cell = ws.cell(row=r_idx, column=4)
            if sp.estado == "CONFORME":
                est_cell.fill = self.fill_ok
            elif sp.estado == "ERROR":
                est_cell.fill = self.fill_err
            elif sp.estado == "ADVERTENCIA":
                est_cell.fill = self.fill_warn
            else:
                est_cell.fill = self.fill_na

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        self._autofit(ws)

    def _build_sheet_incidencias(self, wb: openpyxl.Workbook, result: PAS151AuditResult) -> None:
        ws = wb.create_sheet(title="Incidencias_Detalladas")
        headers = [
            "ID Incidencia",
            "Shapefile",
            "Plantilla",
            "Categoría",
            "Severidad",
            "Campo / Regla",
            "Valor Esperado",
            "Valor Encontrado",
            "Filas Afectadas",
            "Descripción",
            "Recomendación",
        ]
        ws.append(headers)
        self._style_header(ws)

        for iss in result.incidencias:
            ws.append([
                iss.issue_id,
                iss.capa,
                iss.prefijo_plantilla,
                iss.categoria,
                iss.severidad,
                iss.campo_o_regla,
                iss.valor_esperado,
                iss.valor_encontrado,
                iss.filas_afectadas,
                iss.mensaje,
                iss.recomendacion,
            ])
            r_idx = ws.max_row
            for col_i in range(1, len(headers) + 1):
                c = ws.cell(row=r_idx, column=col_i)
                c.font = self.normal_font
                c.border = self.border
            sev_cell = ws.cell(row=r_idx, column=5)
            if iss.severidad == "ERROR":
                sev_cell.fill = self.fill_err
            elif iss.severidad == "ADVERTENCIA":
                sev_cell.fill = self.fill_warn

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        self._autofit(ws)

    def _build_sheet_plantilla(self, wb: openpyxl.Workbook, result: PAS151AuditResult) -> None:
        ws = wb.create_sheet(title="Plantilla_CONAF_PAS151")
        headers = ["shp", "nombre", "tipo", "longitud", "decimales", "orden_esperado", "dominio_conaf"]
        ws.append(headers)
        self._style_header(ws)

        for tmpl in result.plantillas_usadas.values():
            for c in tmpl.campos:
                dom_vals = DOMINIOS_CONAF.get(c.nombre)
                dom_str = ", ".join(str(x) for x in dom_vals) if dom_vals else ""
                ws.append([
                    tmpl.prefijo,
                    c.nombre,
                    c.tipo,
                    c.longitud if c.longitud is not None else "NA",
                    c.decimales if c.decimales is not None else "NA",
                    c.orden,
                    dom_str,
                ])
                r_idx = ws.max_row
                for col_i in range(1, len(headers) + 1):
                    cell = ws.cell(row=r_idx, column=col_i)
                    cell.font = self.normal_font
                    cell.border = self.border

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        self._autofit(ws)

    def _style_header(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
        for cell in ws[1]:
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _autofit(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 55)


# ==============================================================================
# SECCIÓN 7: DEFINICIÓN DE LA CAJA DE HERRAMIENTAS ARCGIS PRO (.pyt)
# ==============================================================================

MODO_FUSIONAR = "FUSIONAR CON PLANTILLA OFICIAL CONAF (Agrega o actualiza capas sobre PAS 151)"
MODO_SOLO_EXCEL = "USAR SOLO EL EXCEL PERSONALIZADO (Audita únicamente las capas del Excel)"


class Toolbox(object):
    def __init__(self):
        """Define la caja de herramientas (.pyt) autónoma para ArcGIS Pro."""
        self.label = "Auditor PAS 151 CONAF (Control de Calidad Cartográfico)"
        self.alias = "auditor_pas151_conaf"
        self.tools = [AuditarPAS151CompletoTool, ExportarPlantillaPAS151Tool]


class AuditarPAS151CompletoTool(object):
    def __init__(self):
        """Herramienta principal que replica y expande Revision_Campos_PAS151_Rev11 2.R."""
        self.label = "1. Revisión Integral de Campos y Topología PAS 151 (CONAF)"
        self.description = (
            "Ejecuta el control de calidad cartográfico y de bases de datos para entregas PAS 151 "
            "conforme al protocolo de CONAF (Descripcion_general.docx y Revision_Campos_PAS151_Rev11 2.R):\n"
            "- Inventario de capas esperadas, faltantes y adicionales (<Prefijo>_<Predio>.shp).\n"
            "- Sistema de coordenadas WGS84 / UTM Zona 19S (EPSG:32719).\n"
            "- Verificación de geometría 2D (ausencia de dimensiones Z/M).\n"
            "- Estructura de tabla de atributos: existencia, tipo, longitud, decimales (Regla NA) y orden exacto.\n"
            "- Dominios CONAF y ausencia de valores vacíos o NA.\n"
            "- Consistencia de superficie declarada (Sup_ha) vs área geométrica recalculada (tolerancia 1%).\n"
            "- Coherencia de coordenadas puntuales (Coord_X, Coord_Y) vs centroide recalculado.\n"
            "- Revisión de formato de texto (acentos, apóstrofes, caracteres especiales y dobles espacios).\n"
            "- Validez geométrica/topológica (autointersecciones, geometrías nulas).\n"
            "- Validaciones espaciales entre capas: IDs duplicados, traslape con compilado de áreas de corta, "
            "Parcelas dentro de Rodales y coincidencia de N_Rodal, contención de Area/Suelos/Rangos_pend en "
            "Limite_Predial y Area_Proyecto, traslape de Predios_vecinos, Infra_aislada sobre Area_Proyecto, "
            "y consistencia transversal de Nom_Predio."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        param_dir = arcpy.Parameter(
            displayName="Carpeta de Shapefiles a Revisar (Directorio SHP)",
            name="directorio_shp",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input",
        )

        param_predio = arcpy.Parameter(
            displayName="Nombre del Predio (ej. PARTE_TORRES_PAEZ, o 'AUTO' para autodetectar)",
            name="nombre_predio",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
        )
        param_predio.value = "AUTO"

        param_compilado = arcpy.Parameter(
            displayName="Shapefile Compilado de Áreas de Corta (Opcional - para validar traslapes)",
            name="ruta_shp_area_comp",
            datatype="DEFile",
            parameterType="Optional",
            direction="Input",
        )
        param_compilado.filter.list = ["shp"]

        param_excel_std = arcpy.Parameter(
            displayName="Archivo Excel con Estándar de Datos Personalizado (Opcional)",
            name="excel_estandar",
            datatype="DEFile",
            parameterType="Optional",
            direction="Input",
        )
        param_excel_std.filter.list = ["xlsx", "xls"]

        param_modo_excel = arcpy.Parameter(
            displayName="Modo de Aplicación del Excel Personalizado",
            name="modo_estandar_excel",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
        )
        param_modo_excel.filter.type = "ValueList"
        param_modo_excel.filter.list = [MODO_FUSIONAR, MODO_SOLO_EXCEL]
        param_modo_excel.value = MODO_FUSIONAR

        param_epsg = arcpy.Parameter(
            displayName="Código EPSG Esperado (WGS84 / UTM 19S = 32719)",
            name="epsg_esperado",
            datatype="GPLong",
            parameterType="Optional",
            direction="Input",
        )
        param_epsg.value = 32719

        param_subfolders = arcpy.Parameter(
            displayName="Buscar Shapefiles recursivamente en subcarpetas",
            name="buscar_subcarpetas",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input",
        )
        param_subfolders.value = False

        param_out_xlsx = arcpy.Parameter(
            displayName="Reporte Excel de Salida (00_Reporte_revision_campos_<PREDIO>.xlsx)",
            name="reporte_salida_xlsx",
            datatype="DEFile",
            parameterType="Required",
            direction="Output",
        )
        param_out_xlsx.filter.list = ["xlsx"]

        return [
            param_dir,
            param_predio,
            param_compilado,
            param_excel_std,
            param_modo_excel,
            param_epsg,
            param_subfolders,
            param_out_xlsx,
        ]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        """Autocompleta la ruta del reporte Excel de salida al seleccionar la carpeta."""
        if parameters[0].valueAsText and not parameters[7].altered:
            folder = Path(parameters[0].valueAsText)
            predio_txt = (parameters[1].valueAsText or "AUTO").strip()
            if predio_txt.upper() == "AUTO" or not predio_txt:
                try:
                    shps = [
                        p for p in folder.iterdir()
                        if p.is_file() and p.suffix.lower() == ".shp" and not p.name.lower().endswith(".shp.xml")
                    ] if folder.exists() else []
                    predio_txt = infer_nombre_predio(shps, get_builtin_templates())
                except Exception:
                    predio_txt = "PREDIO"
            out_name = f"00_Reporte_revision_campos_{predio_txt}.xlsx"
            parameters[7].value = str(folder / out_name)
        return

    def updateMessages(self, parameters):
        if parameters[7].valueAsText and not parameters[7].valueAsText.lower().endswith(".xlsx"):
            parameters[7].setWarningMessage("Se agregará automáticamente la extensión .xlsx al archivo de salida.")
        return

    def execute(self, parameters, messages):
        directorio = parameters[0].valueAsText
        nombre_predio = parameters[1].valueAsText
        ruta_compilado = parameters[2].valueAsText
        excel_estandar = parameters[3].valueAsText
        modo_excel = parameters[4].valueAsText or MODO_FUSIONAR
        epsg_esperado = int(parameters[5].value) if parameters[5].value is not None else 32719
        buscar_subcarpetas = bool(parameters[6].value) if parameters[6].value is not None else False
        reporte_salida = parameters[7].valueAsText

        solo_excel = (modo_excel == MODO_SOLO_EXCEL)

        if not reporte_salida.lower().endswith(".xlsx"):
            reporte_salida += ".xlsx"

        def arcgis_logger(msg: str, level: str = "INFO") -> None:
            if level in ("ERROR", "WARN"):
                arcpy.AddWarning(msg)
            else:
                arcpy.AddMessage(msg)

        arcpy.AddMessage("======================================================================")
        arcpy.AddMessage("AUDITORÍA CARTOGRÁFICA Y ESPACIAL PAS 151 - PROTOCOLO CONAF (Rev 11)")
        arcpy.AddMessage("======================================================================")
        arcpy.AddMessage(f"Directorio de trabajo : {directorio}")
        arcpy.AddMessage(f"Predio indicado       : {nombre_predio or 'AUTO'}")
        arcpy.AddMessage(f"Compilado Áreas Corta : {ruta_compilado or 'No especificado (opcional)'}")
        arcpy.AddMessage(f"Estándar de atributos : {excel_estandar or 'Plantilla Interna PAS 151 Rev11'}")
        if excel_estandar:
            arcpy.AddMessage(f"Modo estándar Excel   : {modo_excel}")
        arcpy.AddMessage(f"EPSG esperado         : {epsg_esperado}")

        engine = PAS151AuditEngine()
        result = engine.run(
            directorio=directorio,
            nombre_predio=nombre_predio,
            ruta_shp_area_comp=ruta_compilado,
            excel_estandar_path=excel_estandar,
            solo_excel=solo_excel,
            epsg_esperado=epsg_esperado,
            buscar_subcarpetas=buscar_subcarpetas,
            log_fn=arcgis_logger,
        )

        exporter = PAS151ExcelExporter()
        out_path = exporter.export(result, reporte_salida)

        arcpy.AddMessage("\n======================================================================")
        arcpy.AddMessage("RESUMEN FINAL DE LA AUDITORÍA PAS 151")
        arcpy.AddMessage("======================================================================")
        arcpy.AddMessage(f"Predio auditado (Sufijo)        : {result.nombre_predio}")
        arcpy.AddMessage(f"Nom_Predio ref (Limite_Predial) : {result.nom_predio_referencia_atrib or 'N/A'}")
        arcpy.AddMessage(f"Total Shapefiles en carpeta     : {len(result.shapes_en_carpeta)}")
        arcpy.AddMessage(f"Total Capas auditadas           : {len(result.shapes_auditados)}")
        arcpy.AddMessage(f"Capas 100% conformes            : {result.total_capas_conformes}")
        arcpy.AddMessage(f"Capas con errores               : {result.total_capas_con_error}")
        arcpy.AddMessage(f"Total incidencias ERROR         : {result.total_errores}")
        arcpy.AddMessage(f"Total incidencias ADVERTENCIA   : {result.total_advertencias}")
        arcpy.AddMessage(f"Reporte Excel generado en       : {out_path}")
        return


class ExportarPlantillaPAS151Tool(object):
    def __init__(self):
        """Exporta el diccionario de datos oficial PAS 151 Rev11 a un archivo Excel editable."""
        self.label = "2. Exportar Plantilla Oficial CONAF PAS 151 a Excel (.xlsx)"
        self.description = (
            "Exporta todas las plantillas de capas, tipos, longitudes, decimales y dominios "
            "del protocolo CONAF PAS 151 (Rev 11) a un archivo Excel (.xlsx) editable para crear "
            "o personalizar nuevos estándares de datos."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        param_out = arcpy.Parameter(
            displayName="Archivo Excel de Salida (Diccionario_PAS151_Rev11.xlsx)",
            name="output_dictionary_xlsx",
            datatype="DEFile",
            parameterType="Required",
            direction="Output",
        )
        param_out.filter.list = ["xlsx"]
        return [param_out]

    def isLicensed(self):
        return True

    def execute(self, parameters, messages):
        out_path = Path(parameters[0].valueAsText)
        if out_path.suffix.lower() != ".xlsx":
            out_path = out_path.with_suffix(".xlsx")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Diccionario de datos"
        ws.append(["shp", "nombre", "tipo", "longitud", "decimales", "dominio_conaf"])

        templates = get_builtin_templates()
        for tmpl in templates.values():
            for c in tmpl.campos:
                dom = DOMINIOS_CONAF.get(c.nombre)
                ws.append([
                    tmpl.prefijo,
                    c.nombre,
                    c.tipo,
                    c.longitud if c.longitud is not None else "NA",
                    c.decimales if c.decimales is not None else "NA",
                    ", ".join(str(x) for x in dom) if dom else "",
                ])

        out_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out_path)
        arcpy.AddMessage(f"Diccionario oficial PAS 151 exportado exitosamente en: {out_path}")
        return
