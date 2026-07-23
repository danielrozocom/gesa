# 🎓 GESA — Gestor de Evaluaciones de Suficiencia Académica

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![PyQt6](https://img.shields.io/badge/UI-PyQt6-green.svg)](https://pypi.org/project/PyQt6/)
[![License](https://img.shields.io/badge/License-MIT-purple.svg)](LICENSE)

**GESA** es una aplicación de escritorio profesional construida con **PyQt6** y **Python** diseñada para automatizar la combinación, formateo y generación masiva de **Evaluaciones de Suficiencia Académica (E.S.A.)** a partir de archivos Microsoft Word (`.docx`) y plantillas institucionales.

---

## ✨ Características Principales

- 🎨 **Interfaz Moderna & Adaptativa (Light / Dark / Sistema)**
  - Diseño inspirado en componentes Shadcn UI con paleta cromática profesional.
  - Sincronización automática con el tema del sistema operativo (Windows DWM Title Bar Sync).
  - Modo Claro pulido con fondos de tarjeta blancos limpios y acentos de color.

- 📅 **Selector de Fechas Inteligente**
  - Calendario emergente nativo traducido al español (`Aceptar` / `Cancelar`).
  - Formato de fecha limpio y regular (no en negrita).
  - Días de fin de semana (sábado y domingo) resaltados en rojo institucional.

- 📁 **Organizador Jerárquico de Exámenes (Sesiones, Subsesiones y Archivos)**
  - Estructura flexible de **Sesiones** ➔ **Subsesiones** ➔ **Archivos `.docx`**.
  - Reordenamiento por arrastrar y soltar (*Drag & Drop*) o botones de subir/bajar.
  - Reubicación automática de archivos si cambias de carpeta o equipo.

- 👁️ **Vista Previa de Salida en Tiempo Real**
  - Muestra cómo quedará el **nombre del archivo** (`.docx`) y el **título del documento** expandido.
  - Actualización instantánea al pasar el cursor (*hover*) o seleccionar cada subsesión.
  - Soporte de etiquetas dinámicas: `{grade}`, `{period}`, `{session}`, `{year}`, `{level}`, `{day}`, `{month}`.

- ↩️ **Historial Deshacer/Rehacer (`Ctrl+Z` / `Ctrl+Y`) & Confirmaciones**
  - Deshaz (`Ctrl+Z`) y reház (`Ctrl+Y` / `Ctrl+Shift+Z`) cualquier cambio en sesiones, subsesiones, archivos o formularios.
  - Diálogos de confirmación antes de eliminar elementos o limpiar la configuración.

- 🧹 **Importar, Exportar y Limpiar**
  - Guarda y carga configuraciones completas en formato `.json`.
  - Botón de **Limpiar** para restablecer la aplicación a un estado por defecto limpio en 1 clic.

- 🚀 **Ejecución Automatizada para Windows (`GESA.exe`)**
  - Incluye `GESA.exe` que detecta e instala Python, instala las dependencias necesarias y ejecuta la aplicación de forma transparente.

---

## 🛠️ Requisitos del Sistema

- **Sistema Operativo:** Windows 10 / Windows 11 (64-bit).
- **Microsoft Word:** Requerido para la combinación e integración de encabezados y numeración en los documentos `.docx`.
- **Python:** Versión 3.10 o superior (`GESA.exe` lo instala automáticamente si no existe).

---

## 🚀 Instalación y Uso Rápido

### 🟢 Opción 1: Para Usuarios Finales (Sin Consola / Fácil)

1. En la página de GitHub ([https://github.com/danielrozocom/gesa](https://github.com/danielrozocom/gesa)), haz clic en el botón verde **`< > Code`** y selecciona **`Download ZIP`**.
2. Descomprime el archivo `.zip` descargado en cualquier carpeta de tu computadora.
3. Haz **doble clic en `GESA.exe`**.
4. El programa verificará e instalará automáticamente Python y todas las librerías necesarias sin que tengas que configurar nada.

---

### 💻 Opción 2: Usando Git / CMD (Desarrolladores)

1. Abre la consola (CMD o PowerShell) y clona el repositorio:
   ```bash
   git clone https://github.com/danielrozocom/gesa.git
   cd gesa
   ```
2. Ejecuta `GESA.exe` o arranca la aplicación directamente:
   ```bash
   python desktop_app.py
   ```

---

## ⌨️ Atajos de Teclado

| Atajo | Acción |
| :--- | :--- |
| **`Ctrl + Z`** | Deshacer la última acción |
| **`Ctrl + Y`** o **`Ctrl + Shift + Z`** | Rehacer la acción deshecha |
| **`Clic en ⚙️`** | Alternar Tema (Oscuro ➔ Claro ➔ Sistema) |

---

## 📂 Estructura del Proyecto

```text
GESA/
├── GESA.exe            # Lanzador ejecutable nativo de Windows con ícono
├── desktop_app.py      # Interfaz de usuario gráfica principal (PyQt6)
├── Code.py             # Motor de combinación de Word (python-docx / win32com)
├── start.bat           # Script de arranque y auto-instalación para Windows
├── requirements.txt    # Dependencias de Python
├── README.md           # Documentación principal
└── .gitignore          # Archivos ignorados por Git
```

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT.
