d2ff9a553536808f5a2b1c15ce1548f264646ec9# Reglas Generales para el Proyecto GESA

## Motor de Fusión de Documentos (Merge)
**NUNCA ELIMINES NI REEMPLACES** la lógica de ensamblaje basada en automatización COM (`win32com.client`). 
- Aunque la librería `python-docx` es mucho más rápida, **rompe las imágenes, figuras anidadas y encabezados** al intentar copiar entre distintos archivos (`No se puede mostrar la imagen`).
- El método `InsertFile` de Microsoft Word ejecutado en segundo plano vía COM es OBLIGATORIO para garantizar que todo el material gráfico, el diseño del template y las estructuras complejas se peguen con fidelidad al 100%.

## Formateo y Espaciado de Párrafos
- **Interlineado para Ahorro de Papel**: Para los párrafos vacíos que actúan como separadores entre preguntas u opciones, el tamaño de fuente debe forzarse siempre a **2pt** (`<w:sz w:val="4"/>`). No los infles a 11pt, ya que el objetivo es comprimir visualmente la página.
- Todo el texto estándar debe ser forzado a fuente **Century Gothic 11pt**.
- Los encabezados de bloques (Competencia, Componente, PART) siempre deben reordenarse para ubicarse obligatoriamente **antes** de la pregunta asociada.
