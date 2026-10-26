# -*- coding: utf-8 -*-
"""
Caja de Herramientas Python (.pyt) para ArcGIS Pro:
Auditor Cartográfico y Espacial PAS 151 - Protocolo CONAF (Rev 11).

Implementa íntegramente las validaciones especificadas en:
- Descripcion_general.docx
- Revision_Campos_PAS151_Rev11 2.R
"""

import os
from pathlib import Path
import sys
import arcpy


def _ensure_package_path() -> None:
    toolbox_dir = os.path.dirname(os.path.abspath(__file__))
    if toolbox_dir not in sys.path:
        sys.path.insert(0, toolbox_dir)


class Toolbox(object):
    def __init__(self):
        """Define la caja de herramientas (.pyt) para ArcGIS Pro."""
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
            displayName="Archivo Excel con Diccionario Personalizado (Opcional - usa Plantilla CONAF Rev11 por defecto)",
            name="excel_estandar",
            datatype="DEFile",
            parameterType="Optional",
            direction="Input",
        )
        param_excel_std.filter.list = ["xlsx", "xls"]

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
            param_epsg,
            param_subfolders,
            param_out_xlsx,
        ]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        """Autocompleta la ruta del reporte Excel de salida al seleccionar la carpeta."""
        _ensure_package_path()
        if parameters[0].valueAsText and not parameters[6].altered:
            folder = Path(parameters[0].valueAsText)
            predio_txt = (parameters[1].valueAsText or "AUTO").strip()
            if predio_txt.upper() == "AUTO" or not predio_txt:
                try:
                    from pas151_auditor.engine import infer_nombre_predio
                    from pas151_auditor.templates_conaf import get_builtin_templates

                    shps = [
                        p for p in folder.iterdir()
                        if p.is_file() and p.suffix.lower() == ".shp" and not p.name.lower().endswith(".shp.xml")
                    ] if folder.exists() else []
                    predio_txt = infer_nombre_predio(shps, get_builtin_templates())
                except Exception:
                    predio_txt = "PREDIO"
            out_name = f"00_Reporte_revision_campos_{predio_txt}.xlsx"
            parameters[6].value = str(folder / out_name)
        return

    def updateMessages(self, parameters):
        if parameters[6].valueAsText and not parameters[6].valueAsText.lower().endswith(".xlsx"):
            parameters[6].setWarningMessage("Se agregará automáticamente la extensión .xlsx al archivo de salida.")
        return

    def execute(self, parameters, messages):
        _ensure_package_path()
        from pas151_auditor.engine import PAS151AuditEngine
        from pas151_auditor.excel_exporter import PAS151ExcelExporter

        directorio = parameters[0].valueAsText
        nombre_predio = parameters[1].valueAsText
        ruta_compilado = parameters[2].valueAsText
        excel_estandar = parameters[3].valueAsText
        epsg_esperado = int(parameters[4].value) if parameters[4].value is not None else 32719
        buscar_subcarpetas = bool(parameters[5].value) if parameters[5].value is not None else False
        reporte_salida = parameters[6].valueAsText

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
        arcpy.AddMessage(f"EPSG esperado         : {epsg_esperado}")

        engine = PAS151AuditEngine()
        result = engine.run(
            directorio=directorio,
            nombre_predio=nombre_predio,
            ruta_shp_area_comp=ruta_compilado,
            excel_estandar_path=excel_estandar,
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
        arcpy.AddMessage(f"Total Capas CONAF auditadas     : {len(result.shapes_auditados)}")
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
            "del protocolo CONAF PAS 151 (Rev 11) a un archivo Excel (.xlsx) compatible con la herramienta 1."
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
        _ensure_package_path()
        import openpyxl
        from pas151_auditor.templates_conaf import DOMINIOS_CONAF, get_builtin_templates

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
