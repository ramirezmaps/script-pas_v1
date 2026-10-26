# -*- coding: utf-8 -*-
"""
Validaciones Espaciales y de Consistencia entre Capas (Sección 2 de Revision_Campos_PAS151_Rev11 2.R,
líneas 573-883, y Descripcion_general.docx).
"""

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from pas151_auditor.attribute_validators import AuditIssue, _fmt_rows
from pas151_auditor.feature_reader import ShapefileData, read_shapefile_complete
from pas151_auditor.templates_conaf import (
    CAPAS_CONTENCION_PREDIO,
    CAPAS_EXCLUIDAS_NOM_PREDIO,
    CATEGORIAS_INFRA_EROSION,
    normalize_layer_key,
)


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
