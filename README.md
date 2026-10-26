# Auditor Cartográfico y Espacial PAS 151 — CONAF (ArcGIS Pro `.pyt` Autónomo)

Caja de herramientas en Python (**Python Toolbox `.pyt` 100% autónoma en un único archivo**) para **ArcGIS Pro** que automatiza el control de calidad cartográfico, de tablas de atributos, dominios y topología espacial para entregas digitales del **Permiso Ambiental Sectorial 151 (PAS 151)** conforme al protocolo oficial de **CONAF (Rev. 11)**.

> **Archivo único autónomo**: Para utilizar la herramienta en cualquier equipo con ArcGIS Pro **solo necesitas el archivo `Auditor_PAS151_CONAF.pyt`** (no requiere carpetas adicionales).

---

## Estructura del Repositorio

```text
├── Auditor_PAS151_CONAF.pyt        # Caja de Herramientas Python 100% autónoma para ArcGIS Pro
├── Diccionario_PAS151_Rev11.xlsx   # Plantilla editable en Excel (referencia para personalizar estándares)
└── requirements.txt                # Dependencias opcionales fuera de ArcGIS Pro
```

---

## ¿Cómo Adicionar o Personalizar un Nuevo Estándar de Datos?

Tienes **2 formas** de agregar nuevas capas, modificar campos (`tipo`, `longitud`, `decimales`) o definir nuevos dominios de valores permitidos:

### Opción A: Sin tocar código (mediante un archivo Excel desde ArcGIS Pro)

1. **Exportar la plantilla base**:
   En ArcGIS Pro, abre `Auditor_PAS151_CONAF.pyt` y ejecuta la herramienta **`2. Exportar Plantilla Oficial CONAF PAS 151 a Excel (.xlsx)`** (o utiliza directamente `Diccionario_PAS151_Rev11.xlsx`).
2. **Editar o crear tu estándar en Excel**:
   Abre el archivo `.xlsx`. Encontrarás la hoja `Diccionario de datos` con 6 columnas:
   | Columna | Descripción | Ejemplo |
   | :--- | :--- | :--- |
   | `shp` | Prefijo del archivo Shapefile (sin el sufijo `_<PREDIO>`) | `Area` o `Nueva_Capa` |
   | `nombre` | Nombre exacto del campo en la tabla `.dbf` | `Tipo_Bos` |
   | `tipo` | Tipo de dato (`String`, `Integer`, `Real`, `Date`) | `String` |
   | `longitud` | Ancho del campo (escribe `NA` si no aplica restricción de ancho) | `50` o `NA` |
   | `decimales` | Cantidad de decimales para campos `Real` (`NA` si no aplica) | `2` o `NA` |
   | `dominio_conaf` | Valores permitidos separados por coma (dejar vacío si es texto/número libre) | `Permanente, Temporal` |

   *Nota*: El orden en que coloques las filas de cada `shp` en el Excel define el **orden obligatorio de las columnas** que exigirá el validador.
3. **Cargar el nuevo estándar al auditar**:
   En la herramienta **`1. Revisión Integral de Campos y Topología PAS 151 (CONAF)`**:
   - Selecciona tu archivo Excel en **`Archivo Excel con Estándar de Datos Personalizado (Opcional)`**.
   - En **`Modo de Aplicación del Excel Personalizado`**, elige:
     - **`FUSIONAR CON PLANTILLA OFICIAL CONAF`**: Mantiene las 43 capas de CONAF y actualiza/agrega las capas o dominios definidos en tu Excel.
     - **`USAR SOLO EL EXCEL PERSONALIZADO`**: Ignora el estándar interno y audita **únicamente** las capas y reglas definidas en tu archivo Excel (ideal para auditar otros permisos o proyectos distintos al PAS 151).

---

### Opción B: Dejar el nuevo estándar integrado permanentemente dentro del `.pyt`

Si deseas modificar el estándar interno por defecto para no tener que seleccionar un Excel externo:
1. Abre `Auditor_PAS151_CONAF.pyt` en cualquier editor de texto (VS Code, Notepad++, IDLE).
2. Ubica la **`SECCIÓN 1: PLANTILLAS CONAF PAS 151 Y DOMINIOS OFICIALES`** al inicio del archivo:
   - En el diccionario **`DOMINIOS_CONAF`**, agrega o edita la lista de valores permitidos para el nombre del campo:
     ```python
     "Nuevo_Campo": ["Valor 1", "Valor 2"],
     ```
   - En el diccionario **`_RAW_TEMPLATES`**, agrega el prefijo de la nueva capa y su lista de tuplas `(nombre, tipo, longitud, decimales)` usando `None` cuando aplique `NA`:
     ```python
     "Mi_Nueva_Capa": [
         ("ID_Elem", "Integer", 10, None),
         ("Nom_Predio", "String", 100, None),
         ("Sup_ha", "Real", None, 2),
     ],
     ```
3. Guarda el archivo `.pyt` y haz clic derecho -> **Refresh** (*Actualizar*) sobre la caja de herramientas en ArcGIS Pro.
