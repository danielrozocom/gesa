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
