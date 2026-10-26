# Auditor Cartográfico y Espacial PAS 151 — CONAF (ArcGIS Pro `.pyt`)

Caja de herramientas en Python (**Python Toolbox `.pyt`**) para **ArcGIS Pro** que automatiza el control de calidad cartográfico, de tablas de atributos, dominios y topología espacial para entregas digitales del **Permiso Ambiental Sectorial 151 (PAS 151)** conforme al protocolo oficial de **CONAF (Rev. 11)**.

---

## Estructura del Repositorio

```text
├── Auditor_PAS151_CONAF.pyt        # Caja de Herramientas Python para ArcGIS Pro
├── Diccionario_PAS151_Rev11.xlsx   # Plantilla oficial CONAF PAS 151 (43 capas y dominios)
├── pas151_auditor/                 # Motor modular de validación cartográfica y espacial
│   ├── __init__.py
│   ├── templates_conaf.py          # 43 plantillas de capas CONAF, regla NA y 31 dominios cerrados
│   ├── feature_reader.py           # Lector binario DBF/SHP/PRJ + soporte híbrido Shapely / ArcPy
│   ├── attribute_validators.py     # Validación CRS, 2D, orden de campos, dominios, Sup_ha, Coord_XY y texto
│   ├── spatial_validators.py       # 7 controles espaciales e inter-capas
│   ├── engine.py                   # Orquestador PAS151AuditEngine y detección automática de predio
│   └── excel_exporter.py           # Exportador de reporte Excel multi-hoja (6 pestañas)
└── requirements.txt                # Dependencias opcionales para ejecución fuera de ArcGIS Pro
```

---

## Controles y Validaciones Implementadas

### 1. Inventario de Archivos Shapefile
- Autodetección del sufijo del predio (`<Prefijo>_<Predio>.shp` o `<Prefijo>.shp`).
- Detección de capas faltantes respecto a la plantilla oficial CONAF y archivos `.shp` adicionales fuera de norma.

### 2. Validación Individual por Capa (43 Plantillas CONAF)
- **Sistema de Referencia de Coordenadas (CRS)**: Verifica proyección `WGS 84 / UTM zone 19S` (`EPSG:32719`).
- **Geometría 2D**: Verifica ausencia de dimensiones `Z` o `M`.
- **Estructura de Campos y Regla `NA`**: Verifica existencia, tipo de dato (`String`, `Integer`, `Real`, `Date`), longitud y decimales directamente desde el encabezado binario `.dbf` (ignorando longitud o decimales cuando la plantilla especifica `NA`).
- **Orden Exacto de Columnas**: Compara la secuencia de campos del `.dbf` contra el orden oficial de la plantilla CONAF.
- **Dominios Oficiales CONAF y Campos Vacíos**: Valida valores permitidos en los 31 campos con dominio cerrado y detecta valores vacíos o `NA`.
- **Superficie (`Sup_ha`)**: Recalcula el área geométrica en hectáreas (`round(area / 10000, 2)`) y verifica que la diferencia porcentual no supere el **1%**.
- **Coordenadas Puntuales (`Coord_X`, `Coord_Y`)**: Recalcula el centroide redondeado al entero y alerta diferencias mayores a 1 m.
- **Formato de Texto**: Detecta caracteres especiales no permitidos (tildes, apóstrofes, acentos graves/agudos) y dobles espacios.
- **Validez Geométrica**: Detecta autointersecciones, anillos inválidos y geometrías vacías (`st_is_valid`).

### 3. Validaciones Espaciales y de Consistencia entre Capas
1. **IDs Duplicados**: Verifica unicidad de `N_Area` (`Area`), `Nom_Pto` (`Referencia`), `N_Parc` (`Parcela`) y la clave compuesta `Nom_Predio + N_Rodal` (`Rodales`).
2. **Traslape con Compilado de Áreas de Corta**: Evalúa intersecciones espaciales (`> 0.01 m²`) entre `Area` y el shapefile compilado de áreas de corta.
3. **Consistencia `Parcela` vs `Rodales`**: Verifica que cada punto de `Parcela` esté contenido espacialmente dentro de un polígono de `Rodales` y que ambos compartan el mismo `N_Rodal`.
4. **Contención en `Limite_Predial` y `Area_Proyecto`**: Verifica que `Area`, `Suelos`, `Rangos_pend` y `Pendientes_caminos` no sobresalgan de `Limite_Predial` ni de `Area_Proyecto` (tolerancia `0.005 m²`).
5. **Traslape `Predios_vecinos` vs `Limite_Predial`**: Detecta superposiciones mayores a `0.01 m²`.
6. **`Infra_aislada` (Control de Erosión) sobre `Area_Proyecto`**: Verifica que las obras de control de erosión no se ubiquen dentro de `Area_Proyecto`.
7. **Consistencia Transversal de `Nom_Predio`**: Verifica que todas las capas contengan exactamente el mismo `Nom_Predio` oficial registrado en `Limite_Predial`.

---

## Uso en ArcGIS Pro

1. Clona o descarga este repositorio en tu equipo.
2. En **ArcGIS Pro**, abre el panel **Catalog** (*Catálogo*).
3. Haz clic derecho sobre **Toolboxes** → **Add Toolbox** (*Agregar caja de herramientas*) y selecciona `Auditor_PAS151_CONAF.pyt`.
4. Ejecuta la herramienta **`1. Revisión Integral de Campos y Topología PAS 151 (CONAF)`**:
   - **Carpeta de Shapefiles a Revisar**: Carpeta que contiene los `.shp` del predio.
   - **Nombre del Predio**: Déjalo en `AUTO` para autodetección o ingresa el sufijo del predio (ej. `KIMAL`).
   - **Shapefile Compilado de Áreas de Corta (Opcional)**: Selecciona el `.shp` compilado si deseas verificar traslapes.
   - **Archivo Excel con Diccionario Personalizado (Opcional)**: Si se deja vacío, utiliza la plantilla oficial CONAF Rev. 11 integrada.
   - **Reporte Excel de Salida**: Ruta del archivo `.xlsx` generado (por defecto `00_Reporte_revision_campos_<PREDIO>.xlsx`).
