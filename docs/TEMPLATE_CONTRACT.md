# Contrato de la Plantilla Maestra

**Regla fundamental:** La plantilla maestra (`template_path`) **nunca** debe modificarse en cuanto a su estilo/tipografía. Solo se reemplazan sus shortcodes `(GRADE)`, `(LEVEL)`, etc.

## ¿Por qué?

El usuario provee una plantilla .docx con su propio diseño (fuente, colores, márgenes, encabezado). El trabajo de GESA es **insertar** los exámenes procesados respetando ese diseño al 100%.

## ¿Cómo se garantiza?

### Pipeline de merge

1. **Pre-procesamiento** (por cada subdocumento):
   - Se abre el subdocumento
   - Se aplica `apply_formatting_to_document()` → **Century Gothic 11pt** a todos sus runs
   - Se guarda como `.tmp.docx`

2. **Preparación de plantilla**:
   - Se abre la plantilla
   - Solo se reemplazan shortcodes `(GRADE)` → `2°`, etc.
   - **NO** se aplica `apply_formatting_to_document()` a la plantilla

3. **Merge principal** (`merge_docx_with_guaranteed_header`):
   - **Intento 1: python-docx** (desde v5.2d)
     - Copia los elementos del subdocumento (ya con Century Gothic 11pt) al final del body de la plantilla
     - Preserva el XML original → la fuente de la plantilla **no se toca**
   - **Intento 2: Word COM** (fallback si python-docx falla)
     - Word COM puede eliminar la fuente explícita durante `InsertFile`
     - En este caso el post-processing aplica Century Gothic a **todo** (incluyendo plantilla)
     - Es un fallback — la pérdida del estilo de plantilla es aceptable ante un documento sin la fuente correcta

4. **Post-processing**:
   - **Footers**: Siempre Century Gothic 11pt para paginación
   - **Headers**: Siempre Century Gothic 11pt
   - **Body**: Solo si se usó Word COM (fallback)

## Advertencias para futuros cambios

| Cambio | ¿Permitido? | Notas |
|--------|------------|-------|
| Agregar shortcode nuevo | ✅ | Solo `replace_in_all()` |
| Cambiar fuente de subdocumentos | ✅ | En `apply_formatting_to_document()` |
| Agregar formato a headers | ❌ | Rompe el diseño de la plantilla |
| Cambiar `_force_run_font()` | ⚠️ | Solo si se entiende el contrato |
| Usar Word COM como primary | ❌ | Destruye la fuente explícita |
| Aplicar `apply_formatting_to_document()` a la plantilla | ❌ | `merge_docx_with_guaranteed_header` línea ~1432 |

## Archivos clave

- `Code.py` → `merge_docx_with_guaranteed_header()`: merge y post-processing
- `Code.py` → `apply_formatting_to_document()`: formateo de subdocumentos
- `Code.py` → `_force_run_font()`: función que fuerza fuente/tamaño en runs
