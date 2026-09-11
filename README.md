# Exportador de estructura de archivos carpetas
Exporta estructura de archivos y carpetas a distintos formatos (txt, csv, html, docx, md)

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![PySide6](https://img.shields.io/badge/UI-PySide6%20(Qt6)-green)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)
![License](https://img.shields.io/badge/license-MIT-yellow)

---

📁 Lista y exporta la estructura de archivos y carpetas
Un potente listador de estructuras de carpetas con interfaz gráfica, filtros avanzados y múltiples formatos de exportación.

✨ Características
🖥️ Interfaz gráfica intuitiva con PySide6

📂 Selección de carpeta con exploración visual

🔍 Filtrado por extensiones mediante casillas de verificación

🚫 Exclusión de archivos y carpetas específicos

🔎 Búsqueda en tiempo real dentro de los resultados

📊 Barra de progreso para operaciones largas

💾 Exportación a múltiples formatos:

TXT (texto plano)

CSV (con delimitador configurable)

Markdown (.md)

HTML (página web)

DOCX (documento Word)

📋 Menú contextual para excluir elementos rápidamente

🎨 Resaltado de sintaxis en la visualización

🌍 Soporte para Unicode (UTF-8)

🚀 Instalación
Requisitos previos
bash
Python 3.7 o superior
Instalación de dependencias
bash
pip install PySide6 python-docx
Nota: python-docx es opcional, solo necesario para exportar a DOCX. Si no está instalado, la opción aparecerá deshabilitada.

Clonar el repositorio
bash
git clone https://github.com/tu-usuario/folder-structure-lister.git
cd folder-structure-lister
💻 Uso
Ejecutar la aplicación
bash
python crear_estructura_archivos.py
Pasos para usar
Seleccionar carpeta - Haz clic en "Examinar..." o escribe la ruta

Configurar opciones:

Activar/desactivar subcarpetas

Seleccionar extensiones a listar

Especificar archivos/carpetas a excluir

Listar - Haz clic en "Listar Estructura"

Buscar - Usa el campo de búsqueda para filtrar resultados

Exportar - Guarda en cualquiera de los formatos disponibles

Menú contextual (clic derecho)
En el área de resultados puedes:

📄 Excluir archivo seleccionado

📁 Excluir carpeta seleccionada

🔤 Excluir por extensión

↩️ Quitar exclusión

📋 Copiar nombre al portapapeles

📦 Formatos de exportación
Formato	Extensión	Descripción
TXT	.txt	Texto plano, compatible con cualquier editor
CSV	.csv	Dos variantes: separado por coma o punto y coma
Markdown	.md	Formato ligero para documentación
HTML	.html	Página web con estilos integrados
DOCX	.docx	Documento Word (requiere python-docx)
🛠️ Estructura del código
text
FolderStructureLister/
├── crear_estructura_archivos.py   # Aplicación principal
├── README.md                       # Este archivo
├── LICENSE                         # Licencia MIT
└── requirements.txt               # Dependencias
Clases principales
Clase	Descripción
MainWindow	Interfaz principal y lógica de control
WorkerThread	Hilo para procesamiento en segundo plano
ExportDialog	Diálogo de opciones de exportación
🔧 Configuración avanzada
Extensiones predeterminadas
La aplicación incluye extensiones comunes preconfiguradas. Puedes detectar automáticamente las extensiones de tu carpeta haciendo clic en "Detectar extensiones".

Exclusiones
Carpetas: Ejemplo cache, temporal, __pycache__

Archivos: Ejemplo temp.log, imagen.png, datos.csv

📝 Ejemplo de salida
text
Estructura de: /proyecto
======================================================================
Fecha: 2026-09-06 15:30:25

[CARPETA] src
    - main.py
    - utils.py
    [CARPETA] data
        - config.json
        - database.sqlite
[CARPETA] tests
    - test_main.py
    - test_utils.py
🤝 Contribuciones
Las contribuciones son bienvenidas. Por favor:

Haz un Fork del proyecto

Crea tu rama de características (git checkout -b feature/AmazingFeature)

Haz commit de tus cambios (git commit -m 'Add some AmazingFeature')

Push a la rama (git push origin feature/AmazingFeature)

Abre un Pull Request

📄 Licencia
Este proyecto está bajo la Licencia MIT. Ver el archivo LICENSE para más detalles.

📞 Contacto
Autor: Luis Miguel Ramos 

GitHub: Zontrox01

Email: zontrox@gmail.com

🙏 Agradecimientos
PySide6 - Framework de interfaz gráfica

python-docx - Para exportación DOCX

⭐ Si te ha sido útil, no olvides darle una estrella al repositorio.
