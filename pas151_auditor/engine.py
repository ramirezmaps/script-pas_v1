# -*- coding: utf-8 -*-
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

from pas151_auditor.attribute_validators import (
    AuditIssue,
    LayerValidationSummary,
    RFieldRow,
    validate_single_shapefile,
)
from pas151_auditor.feature_reader import ShapefileData, read_shapefile_complete
from pas151_auditor.spatial_validators import SpatialCheckResult, run_all_spatial_validations
from pas151_auditor.templates_conaf import LayerTemplate, load_templates, normalize_layer_key


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
        epsg_esperado: int = 32719,
        buscar_subcarpetas: bool = False,
        log_fn: Optional[Callable[[str, str], None]] = None,
    ) -> PAS151AuditResult:
        def emit(msg: str, level: str = "INFO") -> None:
            if log_fn:
                log_fn(msg, level)

        dir_path = Path(directorio)
        templates = load_templates(excel_estandar_path)

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
