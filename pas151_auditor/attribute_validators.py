# -*- coding: utf-8 -*-
"""
Validadores individuales por Shapefile (equivalentes al ciclo principal líneas 342-563
de Revision_Campos_PAS151_Rev11 2.R y Descripcion_general.docx).
"""

from dataclasses import dataclass, field
import re
from typing import Any, Callable, Dict, List, Optional

from pas151_auditor.feature_reader import ShapefileData
from pas151_auditor.templates_conaf import DOMINIOS_CONAF, LayerTemplate


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
