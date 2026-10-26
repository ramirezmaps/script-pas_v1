# -*- coding: utf-8 -*-
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
