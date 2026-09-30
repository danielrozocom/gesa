# GESA — Gestor de Evaluaciones de Suficiencia Académica

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![PyQt6](https://img.shields.io/badge/UI-PyQt6-green.svg)](https://pypi.org/project/PyQt6/)
[![License](https://img.shields.io/badge/License-MIT-purple.svg)](LICENSE)

**GESA** es una aplicación de escritorio diseñada para automatizar la estructuración, combinación y generación masiva de **Evaluaciones de Suficiencia Académica (E.S.A.)** a partir de documentos Microsoft Word (`.docx`) y plantillas institucionales.

---

## 📌 Características

- **Gestión Jerárquica de Exámenes**: Organización estructurada por sesiones, subsesiones y archivos `.docx` con soporte de reordenamiento por arrastrar y soltar (*Drag & Drop*).
- **Motor de Plantillas Dinámicas**: Personalización del nombre de archivo y título del documento mediante variables configurables (`{grade}`, `{period}`, `{session}`, `{year}`, `{level}`, `{day}`, `{month}`).
- **Vista Previa en Tiempo Real**: Inspección instantánea de los metadatos de salida generados antes del procesamiento.
- **Interfaz Moderna y Adaptativa**: Soporte para temas Claro, Oscuro y Sincronización con el sistema.
- **Gestión de Estado Robusta**: Historial deshacer/rehacer (`Ctrl+Z` / `Ctrl+Y`), importación y exportación de configuraciones en formato JSON.
- **Automatización de Documentos**: Combinación de archivos Word preservando el formato y encabezados institucionales.

---

## ⚙️ Procesamiento y Correcciones Aplicadas

Al generar una evaluación, el motor (`Code.py`) aplica automáticamente las siguientes transformaciones y correcciones sobre cada documento fuente y sobre el documento final combinado:

### 📄 Estructura y combinación de documentos

- **Combinación vía Microsoft Word (COM)**: fusiona los subdocumentos dentro de la plantilla preservando el encabezado institucional, con respaldo de fusión nativa de `python-docx` (incluye transferencia de numeración entre documentos).
- **Marcador de límite invisible** (`GESABOUNDARY`, texto blanco de 1 pt) que delimita dónde inicia el contenido insertado; se elimina automáticamente tras la fusión.
- **Eliminación de encabezados y pies de página de los subdocumentos**: heredan únicamente los de la plantilla.
- **Eliminación de saltos de sección internos** y de párrafos/saltos vacíos al inicio y al final de cada documento.
- **Normalización de márgenes y configuración de página**: márgenes uniformes (1 cm), una sola columna y alineación vertical superior en todas las secciones.

### 🔢 Numeración y opciones de respuesta

- **Renumeración continua de preguntas**: las preguntas se renumeran de forma correlativa entre todos los archivos combinados, respetando el número inicial configurable (*offset*).
- **Actualización de referencias cruzadas**: los rangos internos del tipo *"de la pregunta 5 a la 8"* se recalculan según la nueva numeración.
- **Conversión de numeración automática a texto fijo** y normalización a listas nativas de Word (números en negrita para preguntas, letras mayúsculas `A.–E.` para opciones, viñetas `•` para listas).
- **Corrección de secuencias de letras**: si las opciones de una pregunta empiezan en `B.` o se saltan letras, se reinician y corrigen a `A., B., C., D., E.`.
- **Soporte de múltiples formatos de prefijo**: `1)`, `(1`, `A)`, `a.`, `(A)`, etc., todos se normalizan a `1.` y `A.` en negrita.

### 🏷️ Metadatos académicos (Competencia / Componente / Habilidad)

- **Extracción desde tablas**: los bloques de *Competencia*/*Componente* atrapados dentro de tablas se convierten en párrafos normales.
- **Separación en líneas**: si *Competencia* y *Componente* vienen pegados en la misma línea (en cualquier orden), se dividen en dos párrafos con *Competencia* siempre primero.
- **Normalización de idioma y formato**: etiqueta en **negrita**, valor en *Sentence case* y consistencia de idioma (español/inglés) por pareja.
- **Reordenamiento**: los bloques de *Competencia*/*Componente* siempre se ubican **antes** del enunciado de la pregunta.
- **Deduplicación**: si el mismo par *Competencia + Componente* se repite en preguntas consecutivas, se elimina la repetición.
- **Consolidación de Habilidades**: cuando hay varias habilidades (en viñetas o separadas por comas/`;`/`y` en la misma línea), se fusionan en una sola línea con etiqueta en plural:
  - `Habilidades: Focalizar, completar, seleccionar.`
  - Siguiendo la norma RAE para listas inline, **solo la primera palabra lleva mayúscula inicial**.
  - Con una sola habilidad, la etiqueta queda en singular: `Habilidad: Comparar.`
  - Se reconocen variantes como `Habilidad(es)`, `Habilidad (es)` y `Habilidades`, normalizándolas a `Habilidad` o `Habilidades` según corresponda.
  - La línea de habilidades siempre se posiciona **encima del enunciado de la pregunta, sin línea en blanco entre ellas**.

### 🎨 Formato tipográfico uniforme

- **Fuente global**: Century Gothic 11 pt en todo el cuerpo, encabezados, pies y glifos de numeración (se preservan fuentes de símbolos como Wingdings en viñetas).
- **Color de texto negro** (`#000000`) forzado en todos los párrafos, encabezados, pies y numeración, eliminando colores de tema heredados.
- **Cero sangría agregada**: se elimina toda indentación (izquierda, derecha, primera línea y colgante) en párrafos normales, preguntas, opciones y viñetas; también las definiciones de lista se fijan en sangría `0` para que Word no herede indentación alguna. Todo el texto queda alineado al margen izquierdo.
- **Alineación**: justificada para enunciados, opciones y viñetas; izquierda para *Competencia*, *Componente* y *Habilidades*.
- **Interlineado sencillo (1.0)** y espaciado antes/después de 0 pt en todos los párrafos.
- **Párrafos vacíos separadores** reducidos a 2 pt para ahorrar papel.

### 🧹 Limpieza de contenido

- **Eliminación de encabezados redundantes de origen**: nombres de estudiante, docente, grado, fechas, códigos y títulos de evaluación propios de cada archivo fuente.
- **Limpieza de tabulaciones y tabuladores XML**, tanto en el texto como en las propiedades de párrafo.
- **Punto final automático** en opciones y viñetas que no terminen en signo de puntuación.
- **Corrección de errores tipográficos comunes** (p. ej. *Compontencia* → *Competencia*).
- **Espaciado entre preguntas**: se eliminan líneas en blanco superfluas entre encabezados y enunciados, garantizando exactamente una línea en blanco entre el final de una pregunta y el inicio del siguiente bloque.
- **Saneamiento XML**: eliminación de caracteres inválidos para XML 1.0, atributos `rsid` residuales, relaciones de encabezado/pie huérfanas y reconstrucción del ZIP para prevenir corrupción.
- **Cumplimiento estricto del esquema OpenXML (ECMA-376)**: reordenamiento de elementos `pPr`, `rPr`, `sectPr`, `tblPr`, etc., para garantizar que Word abra el documento sin advertencias de "contenido ilegible".

### 🖨️ Salida final

- **Pie de página con paginación dinámica** *"Página X de Y"* en todas las secciones.
- **Metadatos del documento** actualizados (título, categoría, autor *GESA by Daniel Rozo*, fechas).
- **Nombre de archivo y título** generados a partir de la plantilla de variables configurable.

---

## 🛠️ Requisitos del Sistema

- **Sistema Operativo:** Windows 10 / Windows 11 (64-bit).
- **Microsoft Word:** Requerido para la combinación de documentos `.docx`.
- **Python:** 3.10 o superior *(configurado automáticamente por el ejecutable si no está presente)*.

---

## 🚀 Instalación y Uso

### Opción 1: Usuario Final

1. Descarga el archivo ejecutable o el paquete `.zip` del repositorio.
2. Descomprime en una carpeta local.
3. Ejecuta **`GESA.exe`**.

### Opción 2: Desarrollo

1. Clona el repositorio:
   ```bash
   git clone https://github.com/danielrozocom/gesa.git
   cd gesa
   ```
2. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3. Ejecuta la aplicación:
   ```bash
   python desktop_app.py
   ```

---

## ⌨️ Atajos de Teclado

| Atajo | Descripción |
| :--- | :--- |
| **`Ctrl + Z`** | Deshacer la última acción |
| **`Ctrl + Y`** / **`Ctrl + Shift + Z`** | Rehacer la acción deshecha |

---

## 📂 Estructura del Proyecto

```text
GESA/
├── GESA.exe            # Lanzador ejecutable para Windows
├── desktop_app.py      # Interfaz de usuario (PyQt6)
├── Code.py             # Motor de procesamiento de documentos Word
├── start.bat           # Script de inicialización de entorno
├── requirements.txt    # Dependencias del proyecto
└── README.md           # Documentación
```

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT.
