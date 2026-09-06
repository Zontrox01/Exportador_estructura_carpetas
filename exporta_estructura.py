# -*- coding: utf-8 -*-

"""
======================================================================
listar_estructura_gui.py - INTERFAZ GRÁFICA PARA LISTADO DE CARPETAS
======================================================================

Aplicación GUI con PySide6 para listar estructuras de carpetas con:
- Selección de carpeta origen
- Control de profundidad (carpetas y subcarpetas)
- Filtrado por tipos de archivo mediante casillas de verificación
- Búsqueda en resultados
- Exclusión de archivos específicos
- Exportación a múltiples formatos: TXT, CSV, Markdown, DOCX, HTML
"""

import os
import re
import sys
import csv
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QCheckBox,
    QGroupBox, QTextEdit, QProgressBar, QMessageBox, QScrollArea,
    QFrame, QGridLayout, QSplitter, QComboBox, QToolBar, QStatusBar,
    QDialog, QDialogButtonBox, QFormLayout, QSpinBox,
    QMenu, QRadioButton  # <--- AÑADIR ESTO
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QFont, QIcon, QAction, QTextCursor, QKeySequence

# Intentar importar python-docx para exportar a DOCX
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


class WorkerThread(QThread):
    """Hilo para ejecutar el listado sin bloquear la interfaz"""
    progress = Signal(int)
    status = Signal(str)
    finished = Signal(str)
    error = Signal(str)
    
    def __init__(self, root_path, incluir_subcarpetas, extensiones_seleccionadas, 
                 carpetas_excluidas=None,
                 archivos_excluidos=None):
        super().__init__()
        self.root_path = root_path
        self.incluir_subcarpetas = incluir_subcarpetas
        self.extensiones_seleccionadas = extensiones_seleccionadas
        self.carpetas_excluidas = carpetas_excluidas or []
        self.archivos_excluidos = archivos_excluidos or []
        self._is_running = True
        
    def stop(self):
        self._is_running = False
        
    def run(self):
        try:
            if not os.path.isdir(self.root_path):
                self.error.emit(f"No existe la carpeta: {self.root_path}")
                return
                
            # Construir conjunto de extensiones a ignorar
            extensiones_ignoradas = set()
            
            # Extensiones a listar (las que NO están en las ignoradas)
            extensiones_a_listar = set(self.extensiones_seleccionadas) - extensiones_ignoradas
            
            # Crear conjunto de archivos a excluir (nombres exactos)
            archivos_excluidos_set = set(self.archivos_excluidos)
            
            lineas = []
            lineas.append(f"Estructura de: {self.root_path}")
            lineas.append("=" * 70)
            lineas.append(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            lineas.append("")
            
            total_items = self._contar_items(self.root_path)
            items_procesados = 0
            
            self._listar(self.root_path, "", lineas, extensiones_ignoradas, 
                        extensiones_a_listar, archivos_excluidos_set, 
                        items_procesados, total_items)
            
            resultado = "\n".join(lineas)
            self.finished.emit(resultado)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def _contar_items(self, ruta):
        """Cuenta aproximadamente el número de items para la barra de progreso"""
        try:
            count = 0
            for root, dirs, files in os.walk(ruta):
                if not self._is_running:
                    break
                count += len(files) + len(dirs)
            return max(count, 1)  # Evitar división por cero
        except:
            return 100  # Valor por defecto si hay error
    
    def _listar(self, ruta, prefijo, lineas, extensiones_ignoradas, 
                extensiones_a_listar, archivos_excluidos_set,
                items_procesados, total_items):
        """Función recursiva para listar la estructura"""
        if not self._is_running:
            return items_procesados
        
        try:
            entradas = sorted(os.listdir(ruta))
        except PermissionError:
            lineas.append(f"{prefijo}[SIN PERMISO]")
            return items_procesados + 1
        
        carpetas = []
        ficheros = []
        
        for nombre in entradas:
            if not self._is_running:
                return items_procesados
                
            ruta_completa = os.path.join(ruta, nombre)
            if os.path.isdir(ruta_completa):
                # Verificar si la carpeta debe ser excluida
                if nombre not in self.carpetas_excluidas:
                    carpetas.append(nombre)
            else:
                # Verificar si el archivo debe ser excluido por nombre
                if nombre in archivos_excluidos_set:
                    items_procesados += 1
                    if total_items > 0:
                        progress_value = int((items_procesados / total_items) * 100)
                        self.progress.emit(min(progress_value, 100))
                    continue
                    
                ext = os.path.splitext(nombre)[1].lower()
                # Verificar si la extensión debe ser mostrada
                if ext not in extensiones_ignoradas and ext in extensiones_a_listar:
                    ficheros.append(nombre)
                
                # Actualizar progreso para archivos
                items_procesados += 1
                if total_items > 0:
                    progress_value = int((items_procesados / total_items) * 100)
                    self.progress.emit(min(progress_value, 100))
        
        # Mostrar archivos
        for nombre in ficheros:
            lineas.append(f"{prefijo}- {nombre}")
            items_procesados += 1
            if total_items > 0:
                progress_value = int((items_procesados / total_items) * 100)
                self.progress.emit(min(progress_value, 100))
        
        # Procesar carpetas
        if self.incluir_subcarpetas:
            for nombre in carpetas:
                lineas.append(f"{prefijo}[CARPETA] {nombre}")
                items_procesados = self._listar(
                    os.path.join(ruta, nombre), 
                    prefijo + "    ", 
                    lineas, 
                    extensiones_ignoradas, 
                    extensiones_a_listar,
                    archivos_excluidos_set,
                    items_procesados, 
                    total_items
                )
        
        return items_procesados


class ExportDialog(QDialog):
    """Diálogo para opciones de exportación"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Exportar Resultado")
        self.setModal(True)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Formulario
        form_layout = QFormLayout()

        # Formato
        self.format_combo = QComboBox()
        self.format_combo.addItems([
            "TXT - Texto plano",
            "CSV - Separado por comas",
            "CSV - Separado por punto y coma",
            "Markdown (.md)",
            "HTML - Página web",
            "DOCX - Documento Word"
        ])

        # Deshabilitar DOCX si no está disponible
        if not DOCX_AVAILABLE:
            index = self.format_combo.findText("DOCX - Documento Word")
            if index >= 0:
                self.format_combo.setItemText(
                    index,
                    "DOCX - Documento Word "
                    "(NO DISPONIBLE - instalar python-docx)"
                )
                self.format_combo.setItemData(index, False, Qt.UserRole)

        form_layout.addRow("Formato:", self.format_combo)

        # Opciones específicas según formato
        self.options_group = QGroupBox("Opciones adicionales")
        options_layout = QVBoxLayout()

        # Incluir encabezados (para CSV)
        self.include_headers_check = QCheckBox("Incluir encabezados")
        self.include_headers_check.setChecked(True)
        options_layout.addWidget(self.include_headers_check)

        # Incluir fecha
        self.include_date_check = QCheckBox("Incluir fecha en el nombre")
        self.include_date_check.setChecked(False)
        options_layout.addWidget(self.include_date_check)

        # Contenido para Markdown / HTML / DOCX
        self.content_group = QGroupBox("Contenido a exportar")
        content_layout = QVBoxLayout()

        self.content_structure_radio = QRadioButton("Estructura")
        self.content_detail_radio = QRadioButton("Detalle estructurado")
        self.content_both_radio = QRadioButton("Ambos")

        # Mantener el comportamiento actual como opción predeterminada
        self.content_both_radio.setChecked(True)

        content_layout.addWidget(self.content_structure_radio)
        content_layout.addWidget(self.content_detail_radio)
        content_layout.addWidget(self.content_both_radio)

        self.content_group.setLayout(content_layout)
        options_layout.addWidget(self.content_group)

        self.options_group.setLayout(options_layout)
        form_layout.addRow(self.options_group)

        layout.addLayout(form_layout)

        # Mostrar/ocultar opciones según el formato
        self.format_combo.currentIndexChanged.connect(
            self.actualizar_opciones_formato
        )

        # Estado inicial
        self.actualizar_opciones_formato()

        # Botones
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def actualizar_opciones_formato(self):
        """Actualizar las opciones disponibles según el formato."""

        format_type, _ = self.get_format_info()

        # El selector de contenido solo tiene sentido para
        # Markdown, HTML y DOCX.
        es_formato_estructurado = format_type in (
            "markdown",
            "html",
            "docx"
        )

        self.content_group.setVisible(es_formato_estructurado)

        # Los encabezados solo tienen sentido para CSV.
        es_csv = format_type in (
            "csv_comma",
            "csv_semicolon"
        )

        self.include_headers_check.setVisible(es_csv)

    def get_content_type(self):
        """Obtener el tipo de contenido seleccionado."""
        if self.content_structure_radio.isChecked():
            return "estructura"
        elif self.content_detail_radio.isChecked():
            return "detalle"
        else:
            return "ambos"
        
    def get_format_info(self):
        """Obtener información del formato seleccionado"""
        format_text = self.format_combo.currentText()
        if "TXT" in format_text:
            return "txt", ".txt"
        elif "CSV - Separado por comas" in format_text:
            return "csv_comma", ".csv"
        elif "CSV - Separado por punto y coma" in format_text:
            return "csv_semicolon", ".csv"
        elif "Markdown" in format_text:
            return "markdown", ".md"
        elif "HTML" in format_text:
            return "html", ".html"
        elif "DOCX" in format_text:
            return "docx", ".docx"
        return "txt", ".txt"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.extensiones_disponibles = set()
        self.resultado_completo = ""
        self.resultado_marcado = "" 
        self.estructura_parseada = []  # Para exportación estructurada
        self.elemento_seleccionado = None  # <--- AÑADIR ESTA
        self.tipo_elemento_seleccionado = None  # <--- AÑADIR ESTA
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Listador de Estructura de Carpetas")
        self.setGeometry(100, 100, 1200, 800)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        
        # ========== SECCIÓN DE SELECCIÓN DE CARPETA ==========
        folder_layout = QHBoxLayout()
        folder_layout.setSpacing(5)
        
        self.folder_label = QLabel("Carpeta:")
        self.folder_label.setFixedWidth(60)
        folder_layout.addWidget(self.folder_label)
        
        self.folder_path = QLineEdit()
        self.folder_path.setPlaceholderText("Selecciona una carpeta...")
        self.folder_path.textChanged.connect(self.on_folder_changed)
        folder_layout.addWidget(self.folder_path)
        
        self.browse_btn = QPushButton("Examinar...")
        self.browse_btn.clicked.connect(self.browse_folder)
        folder_layout.addWidget(self.browse_btn)
        
        main_layout.addLayout(folder_layout)
        
        # ========== OPCIONES GENERALES ==========
        options_group = QGroupBox("Opciones Generales")
        options_layout = QHBoxLayout()
        
        self.subfolders_check = QCheckBox("Incluir subcarpetas")
        self.subfolders_check.setChecked(True)
        options_layout.addWidget(self.subfolders_check)
        
        options_layout.addSpacing(20)
                
        options_layout.addStretch()
        
        # Botón para detectar extensiones
        self.detect_ext_btn = QPushButton("Detectar extensiones")
        self.detect_ext_btn.clicked.connect(self.detectar_extensiones)
        options_layout.addWidget(self.detect_ext_btn)
        
        options_group.setLayout(options_layout)
        main_layout.addWidget(options_group)
        
        # ========== FILTRO DE EXTENSIONES ==========
        ext_group = QGroupBox("Filtrar por tipo de archivo")
        ext_layout = QVBoxLayout()
        
        # Scroll area para las casillas de verificación
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setMaximumHeight(150)
        
        ext_widget = QWidget()
        self.ext_layout_grid = QGridLayout(ext_widget)
        self.ext_layout_grid.setSpacing(5)
        scroll_area.setWidget(ext_widget)
        
        ext_layout.addWidget(scroll_area)
        
        # Botones para seleccionar/deseleccionar todos
        btn_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("Seleccionar todos")
        self.select_all_btn.clicked.connect(lambda: self.select_all_extensions(True))
        btn_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("Deseleccionar todos")
        self.deselect_all_btn.clicked.connect(lambda: self.select_all_extensions(False))
        btn_layout.addWidget(self.deselect_all_btn)
        
        btn_layout.addStretch()
        
        self.estado_ext_label = QLabel("Extensiones detectadas: 0")
        btn_layout.addWidget(self.estado_ext_label)
        
        ext_layout.addLayout(btn_layout)
        ext_group.setLayout(ext_layout)
        main_layout.addWidget(ext_group)
        
        # ========== CARPETAS Y ARCHIVOS A EXCLUIR ==========
        excl_group = QGroupBox("Elementos a excluir")
        excl_layout = QVBoxLayout()
        
        # Carpetas
        carpetas_layout = QHBoxLayout()

        carpetas_label = QLabel("Carpetas (separadas por comas):")
        carpetas_layout.addWidget(carpetas_label)

        self.excluir_carpetas_input = QLineEdit()
        self.excluir_carpetas_input.setPlaceholderText("Ej: cache, temporal, backup")
        carpetas_layout.addWidget(self.excluir_carpetas_input)

        self.clear_carpetas_btn = QPushButton("Limpiar")
        self.clear_carpetas_btn.clicked.connect(self.excluir_carpetas_input.clear)
        carpetas_layout.addWidget(self.clear_carpetas_btn)

        excl_layout.addLayout(carpetas_layout)
        
        # Archivos
        archivos_layout = QHBoxLayout()

        archivos_label = QLabel("Archivos (separados por comas):")
        archivos_layout.addWidget(archivos_label)

        self.excluir_archivos_input = QLineEdit()
        self.excluir_archivos_input.setPlaceholderText(
            "Ej: listado.txt, imagen.png, temp.log"
        )
        archivos_layout.addWidget(self.excluir_archivos_input)

        self.clear_archivos_btn = QPushButton("Limpiar")
        self.clear_archivos_btn.clicked.connect(self.excluir_archivos_input.clear)
        archivos_layout.addWidget(self.clear_archivos_btn)

        excl_layout.addLayout(archivos_layout)
        
        excl_group.setLayout(excl_layout)
        main_layout.addWidget(excl_group)
        
        # ========== BOTONES DE ACCIÓN Y PROGRESO ==========
        action_layout = QHBoxLayout()
        
        self.listar_btn = QPushButton("Listar Estructura")
        self.listar_btn.clicked.connect(self.listar_estructura)
        self.listar_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        action_layout.addWidget(self.listar_btn)
        
        self.cancelar_btn = QPushButton("Cancelar")
        self.cancelar_btn.clicked.connect(self.cancelar_proceso)
        self.cancelar_btn.setEnabled(False)
        self.cancelar_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        action_layout.addWidget(self.cancelar_btn)
        
        self.guardar_btn = QPushButton("Guardar Resultado...")
        self.guardar_btn.clicked.connect(self.guardar_resultado)
        self.guardar_btn.setEnabled(False)
        self.guardar_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        action_layout.addWidget(self.guardar_btn)
        
        action_layout.addStretch()
        
        main_layout.addLayout(action_layout)
        
        # ========== BARRA DE PROGRESO ==========
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # ========== ÁREA DE RESULTADOS CON BÚSQUEDA ==========
        result_group = QGroupBox("Resultado")
        result_layout = QVBoxLayout()
        
        # Barra de búsqueda
        search_layout = QHBoxLayout()
        search_label = QLabel("Buscar:")
        search_layout.addWidget(search_label)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Escribe texto para buscar en los resultados...")
        self.search_input.textChanged.connect(self.buscar_en_resultados)
        search_layout.addWidget(self.search_input)
        
        self.search_count_label = QLabel("0 coincidencias")
        search_layout.addWidget(self.search_count_label)
        
        self.clear_search_btn = QPushButton("Limpiar")
        self.clear_search_btn.clicked.connect(self.limpiar_busqueda)
        search_layout.addWidget(self.clear_search_btn)
        
        result_layout.addLayout(search_layout)
        
        # ===== IMPORTANTE: Crear self.result_text AQUÍ =====
        self.result_text = QTextEdit()
        self.result_text.setFont(QFont("Courier New", 10))
        self.result_text.setReadOnly(True)

        self.result_text.mouseDoubleClickEvent = self.doble_click_resultado

        result_layout.addWidget(self.result_text)
        
        result_group.setLayout(result_layout)
        main_layout.addWidget(result_group, 1)  # Stretch factor
        
        # Estado inicial - cargar extensiones comunes
        self.cargar_extensiones_comunes()
        
        # ========== CONFIGURAR MENÚ CONTEXTUAL (AQUÍ AL FINAL) ==========
        self.setup_context_menu()
        
        # Barra de estado
        self.statusBar().showMessage("Listo")
        
    def browse_folder(self):
        """Abrir diálogo para seleccionar carpeta"""
        folder = QFileDialog.getExistingDirectory(
            self, 
            "Seleccionar Carpeta", 
            os.path.expanduser("~"),
            QFileDialog.ShowDirsOnly
        )
        if folder:
            self.folder_path.setText(folder)
            # Detectar extensiones automáticamente (sin mensaje)
            self.detectar_extensiones_silencioso()
    
    def on_folder_changed(self, text):
        """Cuando cambia el texto de la carpeta, detectar extensiones automáticamente"""
        if text and os.path.isdir(text):
            # Detectar extensiones silenciosamente cuando cambia la carpeta
            self.detectar_extensiones_silencioso()
    
    def cargar_extensiones_comunes(self):
        """Cargar algunas extensiones comunes por defecto"""
        extensiones_comunes = [".txt", ".py", ".js", ".html", ".css", ".json", 
                              ".xml", ".csv", ".md", ".yml", ".yaml", ".ini",
                              ".cfg", ".conf", ".log", ".sql", ".sh", ".bat",
                              ".ps1", ".rb", ".php", ".java", ".c", ".cpp", 
                              ".h", ".go", ".rs", ".swift", ".kt"]
        
        self.extensiones_disponibles = set(extensiones_comunes)
        self.actualizar_checkboxes_extensiones()
        
    def detectar_extensiones(self):
        """Detectar extensiones en la carpeta seleccionada (con mensaje)"""
        root_path = self.folder_path.text()
        if not root_path or not os.path.isdir(root_path):
            QMessageBox.warning(self, "Advertencia", "Por favor, selecciona una carpeta válida primero.")
            return
        
        self.detect_ext_btn.setEnabled(False)
        self.detect_ext_btn.setText("Detectando...")
        
        # Usar un QTimer para no bloquear la UI
        QTimer.singleShot(100, lambda: self._detectar_extensiones_async(root_path, True))
    
    def detectar_extensiones_silencioso(self):
        """Detectar extensiones en la carpeta seleccionada (sin mensaje)"""
        root_path = self.folder_path.text()
        if not root_path or not os.path.isdir(root_path):
            return
        
        # Usar un QTimer para no bloquear la UI
        QTimer.singleShot(100, lambda: self._detectar_extensiones_async(root_path, False))
        
    def _detectar_extensiones_async(self, root_path, mostrar_mensaje):
        """Detectar extensiones en segundo plano"""
        try:
            extensiones = set()
            contador_archivos = 0
            for root, dirs, files in os.walk(root_path):
                for file in files:
                    ext = os.path.splitext(file)[1].lower()
                    if ext:
                        extensiones.add(ext)
                    contador_archivos += 1
                # Limitar el recorrido para no tardar demasiado
                if len(extensiones) > 100:
                    break
            
            if extensiones:
                self.extensiones_disponibles = extensiones
                self.actualizar_checkboxes_extensiones()
                # Solo mostrar mensaje si se solicitó explícitamente
                if mostrar_mensaje:
                    QMessageBox.information(self, "Extensiones detectadas", 
                                           f"Se detectaron {len(extensiones)} extensiones diferentes.")
        except Exception as e:
            if mostrar_mensaje:
                QMessageBox.warning(self, "Error", f"Error al detectar extensiones: {str(e)}")
        finally:
            self.detect_ext_btn.setEnabled(True)
            self.detect_ext_btn.setText("Detectar extensiones")
    
    def actualizar_checkboxes_extensiones(self):
        """Actualizar las casillas de verificación de extensiones"""
        # Limpiar layout
        while self.ext_layout_grid.count():
            item = self.ext_layout_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Crear casillas de verificación para cada extensión
        extensiones_ordenadas = sorted(self.extensiones_disponibles)
        col_count = 8
        
        for i, ext in enumerate(extensiones_ordenadas):
            row = i // col_count
            col = i % col_count
            checkbox = QCheckBox(ext)
            checkbox.setChecked(True)  # Todas seleccionadas por defecto
            self.ext_layout_grid.addWidget(checkbox, row, col)
        
        self.estado_ext_label.setText(f"Extensiones disponibles: {len(extensiones_ordenadas)}")
    
    def select_all_extensions(self, selected):
        """Seleccionar o deseleccionar todas las extensiones"""
        for i in range(self.ext_layout_grid.count()):
            widget = self.ext_layout_grid.itemAt(i).widget()
            if isinstance(widget, QCheckBox):
                widget.setChecked(selected)
    
    def get_extensiones_seleccionadas(self):
        """Obtener lista de extensiones seleccionadas"""
        extensiones = []
        for i in range(self.ext_layout_grid.count()):
            widget = self.ext_layout_grid.itemAt(i).widget()
            if isinstance(widget, QCheckBox) and widget.isChecked():
                extensiones.append(widget.text())
        return extensiones
    
    def get_carpetas_excluidas(self):
        """Obtener lista de carpetas a excluir - AÑADIR manejo de espacios"""
        texto = self.excluir_carpetas_input.text().strip()
        if not texto:
            return []
        # Separar por comas y limpiar espacios, ignorando elementos vacíos
        carpetas = [c.strip() for c in texto.split(",") if c.strip()]
        return carpetas
    
    def get_archivos_excluidos(self):
        """Obtener lista de archivos a excluir - AÑADIR manejo de espacios"""
        texto = self.excluir_archivos_input.text().strip()
        if not texto:
            return []
        # Separar por comas y limpiar espacios, ignorando elementos vacíos
        archivos = [a.strip() for a in texto.split(",") if a.strip()]
        return archivos
    
    def listar_estructura(self):
        """Iniciar el proceso de listado"""
        root_path = self.folder_path.text().strip()
        if not root_path:
            QMessageBox.warning(self, "Advertencia", "Por favor, selecciona una carpeta.")
            return
        
        if not os.path.isdir(root_path):
            QMessageBox.warning(self, "Advertencia", "La carpeta seleccionada no existe.")
            return
        
        extensiones = self.get_extensiones_seleccionadas()
        if not extensiones:
            QMessageBox.warning(self, "Advertencia", "Selecciona al menos un tipo de archivo para listar.")
            return
        
        carpetas_excluidas = self.get_carpetas_excluidas()
        archivos_excluidos = self.get_archivos_excluidos()
        
        # Deshabilitar controles
        self.listar_btn.setEnabled(False)
        self.cancelar_btn.setEnabled(True)
        self.guardar_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        self.result_text.clear()
        self.search_input.clear()
        self.search_count_label.setText("0 coincidencias")
        self.estructura_parseada = []
        
        # Crear y ejecutar hilo worker - CORREGIDO con todos los argumentos
        self.worker = WorkerThread(
            root_path,
            self.subfolders_check.isChecked(),
            extensiones,
            carpetas_excluidas,
            archivos_excluidos
        )
        
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.status.connect(self.update_status)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        
        self.worker.start()
    
    def update_status(self, mensaje):
        """Actualizar estado en la barra de estado"""
        self.statusBar().showMessage(mensaje)
    
    def on_finished(self, resultado):
        """Manejar finalización del proceso - MODIFICADO para soportar marcado de exclusión"""
        self.resultado_completo = resultado
        
        # Marcar elementos ya excluidos en el resultado
        todas_exclusiones = self.get_archivos_excluidos() + self.get_carpetas_excluidas()
        
        lineas = resultado.split('\n')
        lineas_marcadas = []
        
        for linea in lineas:
            linea_marcada = linea
            for excluido in todas_exclusiones:
                if excluido in linea and ('[CARPETA]' in linea or '- ' in linea):
                    if ' <-- EXCLUIDO' not in linea:
                        linea_marcada = linea + ' <-- EXCLUIDO'
                    break
            lineas_marcadas.append(linea_marcada)
        
        self.resultado_marcado = '\n'.join(lineas_marcadas)
        
        # Mostrar con colores
        self.actualizar_vista_resultado()
        
        # Parsear la estructura para exportación
        self.estructura_parseada = self.parsed_estructura(self.resultado_marcado)
        
        # Habilitar controles
        self.listar_btn.setEnabled(True)
        self.cancelar_btn.setEnabled(False)
        self.guardar_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage("Listado completado")
        
        # Ajustar el cursor al inicio
        cursor = self.result_text.textCursor()
        cursor.movePosition(QTextCursor.Start)
        self.result_text.setTextCursor(cursor)
    
    def on_error(self, error_msg):
        """Manejar errores"""
        QMessageBox.critical(self, "Error", f"Error al listar: {error_msg}")
        self.result_text.setText(f"ERROR: {error_msg}")
        
        self.listar_btn.setEnabled(True)
        self.cancelar_btn.setEnabled(False)
        self.guardar_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage("Error en el listado")
    
    def cancelar_proceso(self):
        """Cancelar el proceso en curso"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
            self.statusBar().showMessage("Proceso cancelado")
            self.listar_btn.setEnabled(True)
            self.cancelar_btn.setEnabled(False)
            self.guardar_btn.setEnabled(False)
            self.progress_bar.setVisible(False)
    
    def parsed_estructura(self, texto):
        """Parsear el texto de la estructura para exportación estructurada"""
        lineas = texto.split('\n')
        estructura = []
        for linea in lineas:
            if linea.strip():
                # Detectar nivel de indentación
                indent = len(linea) - len(linea.lstrip())
                # Detectar si es carpeta o archivo
                is_folder = '[CARPETA]' in linea
                is_file = '- ' in linea and not is_folder
                # Limpiar el nombre
                nombre = linea.strip()
                if is_folder:
                    nombre = nombre.replace('[CARPETA]', '').strip()
                elif is_file:
                    nombre = nombre.replace('-', '').strip()
                
                estructura.append({
                    'indent': indent,
                    'nombre': nombre,
                    'tipo': 'carpeta' if is_folder else 'archivo' if is_file else 'encabezado',
                    'linea': linea
                })
        return estructura
    
    def guardar_resultado(self):
        """Abrir diálogo de exportación y guardar en el formato seleccionado"""
        if not self.resultado_completo:
            QMessageBox.warning(self, "Advertencia", "No hay resultado para guardar.")
            return
        
        # Mostrar diálogo de exportación
        dialog = ExportDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return
        
        format_type, extension = dialog.get_format_info()

        # Usar directamente el nombre de la carpeta raíz
        nombre_base = os.path.basename(
            os.path.normpath(self.folder_path.text())
        ) or "estructura"

        # Añadir fecha si se solicita
        if dialog.include_date_check.isChecked():
            fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_base = f"{nombre_base}_{fecha}"
        
        # Diálogo para guardar archivo
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            f"Guardar como {format_type.upper()}",
            f"{nombre_base}{extension}",
            f"{format_type.upper()} (*{extension});;Todos los archivos (*)"
        )
        
        if not file_path:
            return
        
        try:
            self.exportar_formato(
                file_path,
                format_type,
                dialog.include_headers_check.isChecked(),
                dialog.get_content_type()
)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo:\n{str(e)}")
    
    def exportar_formato(
        self,
        file_path,
        format_type,
        include_headers,
        content_type="ambos"
    ):
        """Exportar en el formato seleccionado."""

        if format_type == "txt":
            self.exportar_txt(file_path)

        elif format_type == "csv_comma":
            self.exportar_csv(
                file_path,
                ',',
                include_headers
            )

        elif format_type == "csv_semicolon":
            self.exportar_csv(
                file_path,
                ';',
                include_headers
            )

        elif format_type == "markdown":
            self.exportar_markdown(
                file_path,
                content_type
            )

        elif format_type == "html":
            self.exportar_html(
                file_path,
                content_type
            )

        elif format_type == "docx":
            self.exportar_docx(
                file_path,
                content_type
            )
    
    def exportar_txt(self, file_path):
        """Exportar a TXT (texto plano)"""
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(self.resultado_completo)
    
    def exportar_csv(self, file_path, delimiter, include_headers):
        """Exportar a CSV con el delimitador especificado"""
        
        with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f, delimiter=delimiter)
            
            if include_headers:
                writer.writerow(['Nivel', 'Tipo', 'Nombre', 'Ruta'])
            
            for item in self.estructura_parseada:
                if item['tipo'] in ['archivo', 'carpeta']:
                    nivel = item['indent'] // 4
                    ruta = item['linea']
                    
                    # Solución: Agregar un tabulador al inicio
                    # Excel interpreta esto como texto, no como fórmula
                    if ruta.startswith("- "):
                        ruta = "\t" + ruta
                    
                    writer.writerow([
                        nivel,
                        item['tipo'],
                        item['nombre'],
                        ruta
                    ])
    
    def exportar_markdown(self, file_path, content_type="ambos"):
        """Exportar a Markdown según el contenido seleccionado."""

        with open(file_path, 'w', encoding='utf-8') as f:

            # Título
            if content_type == "estructura":
                f.write("# Estructura de carpetas\n\n")
            elif content_type == "detalle":
                f.write("# Detalle estructurado\n\n")
            else:
                f.write("# Estructura de carpetas\n\n")

            # Información general
            f.write(
                f"**Fecha:** "
                f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
            )
            f.write(
                f"**Carpeta:** {self.folder_path.text()}\n\n"
            )

            # ==========================================================
            # ESTRUCTURA
            # ==========================================================
            if content_type in ("estructura", "ambos"):

                if content_type == "ambos":
                    f.write("## Estructura de carpetas\n\n")

                f.write("```\n")
                f.write(self.resultado_completo)
                f.write("\n```\n\n")

            # ==========================================================
            # DETALLE ESTRUCTURADO
            # ==========================================================
            if content_type in ("detalle", "ambos"):

                if content_type == "ambos":
                    f.write("## Detalle estructurado\n\n")

                for item in self.estructura_parseada:
                    if item['tipo'] in ['archivo', 'carpeta']:
                        indent = "  " * (item['indent'] // 4)
                        icono = (
                            "📁"
                            if item['tipo'] == 'carpeta'
                            else "📄"
                        )

                        f.write(
                            f"{indent}- {icono} {item['nombre']}\n"
                        )
    
    def exportar_html(self, file_path, content_type="ambos"):
        """Exportar a HTML según el contenido seleccionado."""

        with open(file_path, 'w', encoding='utf-8') as f:

            f.write("""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Estructura de Carpetas</title>
    <style>
        body {
            font-family: 'Courier New', monospace;
            background-color: #f5f5f5;
            padding: 20px;
            margin: 0;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        h1 {
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
            font-family: 'Segoe UI', Arial, sans-serif;
        }

        .info {
            color: #7f8c8d;
            font-size: 14px;
            margin-bottom: 20px;
            font-family: 'Segoe UI', Arial, sans-serif;
        }

        .info span {
            background-color: #ecf0f1;
            padding: 3px 10px;
            border-radius: 4px;
        }

        .estructura {
            background-color: #2c3e50;
            color: #ecf0f1;
            padding: 20px;
            border-radius: 8px;
            overflow-x: auto;
            white-space: pre;
            font-family: 'Courier New', monospace;
            font-size: 13px;
            line-height: 1.6;
            margin: 20px 0;
        }

        .folder {
            color: #3498db;
        }

        .file {
            color: #2ecc71;
        }

        .header {
            color: #f1c40f;
        }

        .permalink {
            color: #e74c3c;
        }

        .detalle {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 2px solid #ecf0f1;
        }

        .detalle h2 {
            color: #2c3e50;
            font-family: 'Segoe UI', Arial, sans-serif;
        }

        .tree-item {
            font-family: 'Courier New', monospace;
            padding: 2px 0;
        }

        .tree-item.folder {
            color: #2980b9;
            font-weight: bold;
        }

        .tree-item.file {
            color: #27ae60;
        }

        .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #ecf0f1;
            color: #7f8c8d;
            font-size: 12px;
            font-family: 'Segoe UI', Arial, sans-serif;
            text-align: center;
        }
    </style>
</head>
<body>
    <div class="container">
""")

            # ==========================================================
            # TÍTULO
            # ==========================================================

            if content_type == "estructura":
                titulo = "📂 Estructura de Carpetas"
            elif content_type == "detalle":
                titulo = "📋 Detalle Estructurado"
            else:
                titulo = "📂 Estructura de Carpetas"

            f.write(f"        <h1>{titulo}</h1>\n")

            # ==========================================================
            # INFORMACIÓN
            # ==========================================================

            f.write("""        <div class="info">
""")

            f.write(
                "            <span>📅 Fecha: "
                + datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                + "</span>\n"
            )

            f.write(
                '            <span style="margin-left: 10px;">📁 '
                + self.folder_path.text()
                + "</span>\n"
            )

            f.write("""        </div>
""")

            # ==========================================================
            # ESTRUCTURA
            # ==========================================================

            if content_type in ("estructura", "ambos"):

                f.write("""
        <div class="estructura">
""")

                # Escribir el contenido con colores
                for linea in self.resultado_completo.split('\n'):

                    if '[CARPETA]' in linea:

                        linea_coloreada = linea.replace(
                            '[CARPETA]',
                            '<span class="folder">[CARPETA]</span>'
                        )

                        f.write(
                            f'<span class="folder">'
                            f'{linea_coloreada}'
                            f'</span>\n'
                        )

                    elif '- ' in linea and '[CARPETA]' not in linea:

                        linea_coloreada = linea.replace(
                            '-',
                            '<span class="file">-</span>',
                            1
                        )

                        f.write(
                            f'<span class="file">'
                            f'{linea_coloreada}'
                            f'</span>\n'
                        )

                    elif '=' in linea:

                        f.write(
                            f'<span class="header">'
                            f'{linea}'
                            f'</span>\n'
                        )

                    else:
                        f.write(f'{linea}\n')

                f.write("""        </div>
""")

            # ==========================================================
            # DETALLE ESTRUCTURADO
            # ==========================================================

            if content_type in ("detalle", "ambos"):

                f.write("""
        <div class="detalle">
            <h2>📋 Detalle estructurado</h2>
            <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px;">
""")

                for item in self.estructura_parseada:

                    if item['tipo'] in ['archivo', 'carpeta']:

                        indent = "&nbsp;" * item['indent']

                        icono = (
                            "📁"
                            if item['tipo'] == 'carpeta'
                            else "📄"
                        )

                        clase = (
                            "folder"
                            if item['tipo'] == 'carpeta'
                            else "file"
                        )

                        f.write(
                            f'<div class="tree-item {clase}">'
                            f'{indent}{icono} {item["nombre"]}'
                            f'</div>\n'
                        )

                f.write("""            </div>
        </div>
""")

            # ==========================================================
            # PIE
            # ==========================================================

            f.write("""
        <div class="footer">
            Generado con Listador de Estructura de Carpetas v2.0
        </div>
    </div>
</body>
</html>
""")
    
    def exportar_docx(self, file_path, content_type="ambos"):
        """Exportar a DOCX según el contenido seleccionado."""

        if not DOCX_AVAILABLE:
            raise Exception(
                "python-docx no está instalado. "
                "Ejecuta: pip install python-docx"
            )

        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = Document()

        # ==========================================================
        # TÍTULO
        # ==========================================================

        if content_type == "estructura":
            titulo = "Estructura de Carpetas"
        elif content_type == "detalle":
            titulo = "Detalle Estructurado"
        else:
            titulo = "Estructura de Carpetas"

        title = doc.add_heading(titulo, 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # ==========================================================
        # INFORMACIÓN
        # ==========================================================

        doc.add_paragraph(
            f'Fecha: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
        )

        doc.add_paragraph(
            f'Carpeta: {self.folder_path.text()}'
        )

        doc.add_paragraph()

        # ==========================================================
        # ESTILO PARA LA ESTRUCTURA
        # ==========================================================

        style = doc.styles.add_style('CodeStyle', 1)
        style.font.name = 'Courier New'
        style.font.size = Pt(9)
        style.font.color.rgb = RGBColor(44, 62, 80)

        # ==========================================================
        # ESTRUCTURA
        # ==========================================================

        if content_type in ("estructura", "ambos"):

            if content_type == "ambos":
                doc.add_heading(
                    'Estructura de Carpetas',
                    level=1
                )

            for linea in self.resultado_completo.split('\n'):

                p = doc.add_paragraph(
                    linea,
                    style='CodeStyle'
                )

                # Resaltar carpetas en azul
                if '[CARPETA]' in linea:

                    for run in p.runs:
                        run.font.color.rgb = RGBColor(
                            52, 152, 219
                        )
                        run.font.bold = True

                # Resaltar archivos en verde
                elif '- ' in linea and '[CARPETA]' not in linea:

                    for run in p.runs:
                        run.font.color.rgb = RGBColor(
                            46, 204, 113
                        )

        # ==========================================================
        # DETALLE ESTRUCTURADO
        # ==========================================================

        if content_type in ("detalle", "ambos"):

            # Solo hacemos salto de página si también hemos
            # exportado previamente la estructura.
            if content_type == "ambos":
                doc.add_page_break()

            doc.add_heading(
                'Detalle Estructurado',
                level=1
            )

            # Tabla de detalle
            table = doc.add_table(
                rows=1,
                cols=3
            )

            table.style = 'Light Grid Accent 1'

            hdr_cells = table.rows[0].cells

            hdr_cells[0].text = 'Nivel'
            hdr_cells[1].text = 'Tipo'
            hdr_cells[2].text = 'Nombre'

            for item in self.estructura_parseada:

                if item['tipo'] in ['archivo', 'carpeta']:

                    nivel = item['indent'] // 4

                    row_cells = table.add_row().cells

                    row_cells[0].text = str(nivel)

                    row_cells[1].text = (
                        '📁 Carpeta'
                        if item['tipo'] == 'carpeta'
                        else '📄 Archivo'
                    )

                    row_cells[2].text = item['nombre']

        # ==========================================================
        # GUARDAR
        # ==========================================================

        doc.save(file_path)

    def buscar_en_resultados(self):
        """Buscar texto en los resultados - MODIFICADO para manejar HTML correctamente"""
        texto_buscar = self.search_input.text().strip()
        if not texto_buscar:
            self.actualizar_vista_resultado()
            self.search_count_label.setText("0 coincidencias")
            return
        
        # Obtener el texto completo sin formato HTML
        texto_completo = self.resultado_marcado if hasattr(self, 'resultado_marcado') else self.resultado_completo
        
        # Buscar todas las ocurrencias
        import re
        patron = re.compile(re.escape(texto_buscar), re.IGNORECASE)
        ocurrencias = list(patron.finditer(texto_completo))
        
        if not ocurrencias:
            self.actualizar_vista_resultado()
            self.search_count_label.setText("0 coincidencias")
            return
        
        # Construir texto con resaltado
        html_resultado = "<pre>"
        last_end = 0
        
        for match in ocurrencias:
            start = match.start()
            end = match.end()
            html_resultado += texto_completo[last_end:start]
            html_resultado += f'<span style="background-color: #ffff00; font-weight: bold;">{texto_completo[start:end]}</span>'
            last_end = end
        
        html_resultado += texto_completo[last_end:]
        html_resultado += "</pre>"
        
        self.result_text.setHtml(html_resultado)
        self.search_count_label.setText(f"{len(ocurrencias)} coincidencias")
        
        # Ir a la primera coincidencia
        cursor = self.result_text.textCursor()
        cursor.movePosition(QTextCursor.Start)
        self.result_text.setTextCursor(cursor)
    
    def limpiar_busqueda(self):
        """Limpiar el campo de búsqueda - MODIFICADO para restaurar vista con colores"""
        self.search_input.clear()
        self.actualizar_vista_resultado()
        self.search_count_label.setText("0 coincidencias")

    def setup_context_menu(self):
        """Configurar el menú contextual para el área de resultados"""
        self.result_text.setContextMenuPolicy(Qt.CustomContextMenu)
        self.result_text.customContextMenuRequested.connect(self.show_context_menu)
        
        # Crear el menú
        self.context_menu = QMenu(self)
        
        # Acciones del menú
        self.action_excluir_archivo = QAction("📄 Excluir este archivo", self)
        self.action_excluir_archivo.triggered.connect(lambda: self.excluir_elemento('archivo'))
        
        self.action_excluir_carpeta = QAction("📁 Excluir esta carpeta", self)
        self.action_excluir_carpeta.triggered.connect(lambda: self.excluir_elemento('carpeta'))
        
        self.action_excluir_extension = QAction("🔤 Excluir por extensión", self)
        self.action_excluir_extension.triggered.connect(self.excluir_por_extension)
        
        self.action_quitar_exclusion = QAction("↩️ Quitar exclusión", self)
        self.action_quitar_exclusion.triggered.connect(self.quitar_exclusion)
        
        # Separador
        self.context_menu.addSeparator()
        
        self.action_copiar_nombre = QAction("📋 Copiar nombre", self)
        self.action_copiar_nombre.triggered.connect(self.copiar_nombre_seleccionado)
        
        # Añadir acciones al menú
        self.context_menu.addAction(self.action_excluir_archivo)
        self.context_menu.addAction(self.action_excluir_carpeta)
        self.context_menu.addAction(self.action_excluir_extension)
        self.context_menu.addSeparator()
        self.context_menu.addAction(self.action_quitar_exclusion)
        self.context_menu.addSeparator()
        self.context_menu.addAction(self.action_copiar_nombre)
        
        # Almacenar el elemento seleccionado
        self.elemento_seleccionado = None
        self.tipo_elemento_seleccionado = None

    def doble_click_resultado(self, event):
        """Al hacer doble clic sobre un archivo, añadirlo a las exclusiones."""
        cursor = self.result_text.cursorForPosition(event.position().toPoint())
        cursor.select(QTextCursor.LineUnderCursor)
        linea = cursor.selectedText()

        if not linea.strip():
            return

        # Solo actuar sobre archivos
        if '- ' not in linea or '[CARPETA]' in linea:
            return

        nombre = self.limpiar_nombre(linea)

        # Comprobar si ya está excluido
        if self.esta_excluido(nombre):
            self.statusBar().showMessage(
                f"'{nombre}' ya está excluido", 2000
            )
            return

        # Añadir al campo de archivos excluidos
        archivos_actuales = self.excluir_archivos_input.text().strip()

        if archivos_actuales:
            nuevo_texto = f"{archivos_actuales}, {nombre}"
        else:
            nuevo_texto = nombre

        self.excluir_archivos_input.setText(nuevo_texto)

        # Marcarlo visualmente como excluido
        self.marcar_excluido_en_resultado(nombre)

        self.statusBar().showMessage(
            f"🚫 Archivo excluido: {nombre}", 3000
        )

    def show_context_menu(self, position):
        """Mostrar menú contextual en la posición del cursor"""
        cursor = self.result_text.cursorForPosition(position)
        cursor.select(QTextCursor.LineUnderCursor)
        linea_seleccionada = cursor.selectedText()
        
        if not linea_seleccionada.strip():
            return
        
        # Determinar qué tipo de elemento es
        self.elemento_seleccionado = linea_seleccionada.strip()
        self.tipo_elemento_seleccionado = 'desconocido'
        
        # Limpiar formato para obtener el nombre limpio
        nombre_limpio = self.limpiar_nombre(linea_seleccionada)
        
        if '[CARPETA]' in linea_seleccionada:
            self.tipo_elemento_seleccionado = 'carpeta'
            self.action_excluir_carpeta.setEnabled(True)
            self.action_excluir_archivo.setEnabled(False)
            self.action_excluir_archivo.setText("📄 Excluir este archivo")
        elif '- ' in linea_seleccionada and '[CARPETA]' not in linea_seleccionada:
            self.tipo_elemento_seleccionado = 'archivo'
            self.action_excluir_archivo.setEnabled(True)
            self.action_excluir_carpeta.setEnabled(False)
            self.action_excluir_carpeta.setText("📁 Excluir esta carpeta")
        else:
            self.action_excluir_archivo.setEnabled(False)
            self.action_excluir_carpeta.setEnabled(False)
        
        # Verificar si ya está excluido
        if self.esta_excluido(nombre_limpio):
            self.action_quitar_exclusion.setEnabled(True)
            self.action_excluir_archivo.setEnabled(False)
            self.action_excluir_carpeta.setEnabled(False)
        else:
            self.action_quitar_exclusion.setEnabled(False)
        
        # Habilitar/deshabilitar opción de extensión
        if self.tipo_elemento_seleccionado == 'archivo':
            self.action_excluir_extension.setEnabled(True)
        else:
            self.action_excluir_extension.setEnabled(False)
        
        # Mostrar el menú
        self.context_menu.exec_(self.result_text.mapToGlobal(position))

    def limpiar_nombre(self, linea):
        """Limpiar el nombre de un elemento para obtener solo el nombre del archivo/carpeta"""
        nombre = linea.strip()
        
        # Quitar prefijos
        if '[CARPETA]' in nombre:
            nombre = nombre.replace('[CARPETA]', '').strip()
        elif '- ' in nombre:
            nombre = nombre.replace('-', '').strip()
        
        # Quitar posibles anotaciones de exclusión
        if ' <-- EXCLUIDO' in nombre:
            nombre = nombre.replace(' <-- EXCLUIDO', '').strip()
        
        return nombre

    def esta_excluido(self, nombre):
        """Verificar si un nombre ya está en la lista de exclusiones"""
        exclusiones_actuales = self.get_archivos_excluidos() + self.get_carpetas_excluidas()
        return nombre in exclusiones_actuales

    def excluir_elemento(self, tipo):
        """Excluir el elemento seleccionado (archivo o carpeta)"""
        if not self.elemento_seleccionado:
            return
        
        nombre_limpio = self.limpiar_nombre(self.elemento_seleccionado)
        
        # Verificar si ya está excluido
        if self.esta_excluido(nombre_limpio):
            QMessageBox.information(self, "Información", f"'{nombre_limpio}' ya está excluido.")
            return
        
        # Añadir al campo correspondiente
        if tipo == 'archivo':
            archivos_actuales = self.excluir_archivos_input.text().strip()
            if archivos_actuales:
                nuevos_archivos = f"{archivos_actuales}, {nombre_limpio}"
            else:
                nuevos_archivos = nombre_limpio
            self.excluir_archivos_input.setText(nuevos_archivos)
        else:  # carpeta
            carpetas_actuales = self.excluir_carpetas_input.text().strip()
            if carpetas_actuales:
                nuevas_carpetas = f"{carpetas_actuales}, {nombre_limpio}"
            else:
                nuevas_carpetas = nombre_limpio
            self.excluir_carpetas_input.setText(nuevas_carpetas)
        
        # Marcar en el resultado
        self.marcar_excluido_en_resultado(nombre_limpio)
        
        # Actualizar el contador de búsqueda
        self.statusBar().showMessage(f"✅ Excluido: {nombre_limpio}", 3000)

    def excluir_por_extension(self):
        """Excluir todos los archivos con la misma extensión"""
        if not self.elemento_seleccionado or self.tipo_elemento_seleccionado != 'archivo':
            return
        
        nombre = self.limpiar_nombre(self.elemento_seleccionado)
        extension = os.path.splitext(nombre)[1]
        
        if not extension:
            QMessageBox.warning(self, "Advertencia", "Este archivo no tiene extensión.")
            return
        
        # Verificar si la extensión ya está en las extensiones seleccionadas
        # Desmarcar la extensión en las casillas de verificación
        for i in range(self.ext_layout_grid.count()):
            widget = self.ext_layout_grid.itemAt(i).widget()
            if isinstance(widget, QCheckBox) and widget.text() == extension:
                widget.setChecked(False)
                self.statusBar().showMessage(f"🔤 Extensión '{extension}' excluida", 3000)
                return
        
        # Si no está en la lista, preguntar si quiere añadirla
        reply = QMessageBox.question(
            self, 
            "Excluir extensión",
            f"La extensión '{extension}' no está en la lista de filtros.\n"
            f"¿Quieres agregarla a las extensiones excluidas?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Añadir como extensión a excluir permanentemente
            # Nota: Esto requeriría modificar la lógica del WorkerThread
            self.statusBar().showMessage(f"🔤 Extensión '{extension}' marcada para exclusión", 3000)

    def quitar_exclusion(self):
        """Quitar la exclusión del elemento seleccionado"""
        if not self.elemento_seleccionado:
            return
        
        nombre_limpio = self.limpiar_nombre(self.elemento_seleccionado)
        
        # Quitar del campo de archivos excluidos
        archivos_actuales = self.excluir_archivos_input.text().strip()
        if archivos_actuales:
            archivos_lista = [a.strip() for a in archivos_actuales.split(',')]
            if nombre_limpio in archivos_lista:
                archivos_lista.remove(nombre_limpio)
                self.excluir_archivos_input.setText(', '.join(archivos_lista))
                self.statusBar().showMessage(f"↩️ Exclusión removida: {nombre_limpio}", 3000)
                self.actualizar_vista_resultado()
                return
        
        # Quitar del campo de carpetas excluidas
        carpetas_actuales = self.excluir_carpetas_input.text().strip()
        if carpetas_actuales:
            carpetas_lista = [c.strip() for c in carpetas_actuales.split(',')]
            if nombre_limpio in carpetas_lista:
                carpetas_lista.remove(nombre_limpio)
                self.excluir_carpetas_input.setText(', '.join(carpetas_lista))
                self.statusBar().showMessage(f"↩️ Exclusión removida: {nombre_limpio}", 3000)
                self.actualizar_vista_resultado()

    def marcar_excluido_en_resultado(self, nombre):
        """Marcar un elemento como excluido en el resultado (en rojo)"""
        texto_actual = self.result_text.toPlainText()
        
        # Buscar líneas que contengan el nombre
        lineas = texto_actual.split('\n')
        lineas_modificadas = []
        
        for linea in lineas:
            if nombre in linea and ('[CARPETA]' in linea or '- ' in linea):
                # Añadir marcador de exclusión si no lo tiene ya
                if ' <-- EXCLUIDO' not in linea:
                    linea += ' <-- EXCLUIDO'
                lineas_modificadas.append(linea)
            else:
                lineas_modificadas.append(linea)
        
        # Actualizar el texto con colores (HTML)
        html_resultado = "<pre>"
        for linea in lineas_modificadas:
            if ' <-- EXCLUIDO' in linea:
                # Resaltar en rojo
                linea_limpia = linea.replace(' <-- EXCLUIDO', '')
                html_resultado += f'<span style="color: #ff0000; text-decoration: line-through;">{linea_limpia}</span> <span style="color: #ff0000; font-weight: bold;">🚫 EXCLUIDO</span>\n'
            elif '[CARPETA]' in linea:
                html_resultado += f'<span style="color: #2980b9;">{linea}</span>\n'
            elif '- ' in linea:
                html_resultado += f'<span style="color: #27ae60;">{linea}</span>\n'
            else:
                html_resultado += f'{linea}\n'
        html_resultado += "</pre>"
        
        self.result_text.setHtml(html_resultado)
        
        # Guardar el texto sin formato para futuras operaciones
        self.resultado_completo = '\n'.join(lineas_modificadas)

    def actualizar_vista_resultado(self):
        """Actualizar la vista del resultado después de cambios en exclusiones"""
        if not self.resultado_completo:
            return
        
        # Volver a mostrar el resultado con los colores apropiados
        lineas = self.resultado_completo.split('\n')
        html_resultado = "<pre>"
        
        for linea in lineas:
            if ' <-- EXCLUIDO' in linea:
                linea_limpia = linea.replace(' <-- EXCLUIDO', '')
                html_resultado += f'<span style="color: #ff0000; text-decoration: line-through;">{linea_limpia}</span> <span style="color: #ff0000; font-weight: bold;">🚫 EXCLUIDO</span>\n'
            elif '[CARPETA]' in linea:
                html_resultado += f'<span style="color: #2980b9;">{linea}</span>\n'
            elif '- ' in linea:
                html_resultado += f'<span style="color: #27ae60;">{linea}</span>\n'
            else:
                html_resultado += f'{linea}\n'
        html_resultado += "</pre>"
        
        self.result_text.setHtml(html_resultado)

    def copiar_nombre_seleccionado(self):
        """Copiar el nombre del elemento seleccionado al portapapeles"""
        if not self.elemento_seleccionado:
            return
        
        nombre = self.limpiar_nombre(self.elemento_seleccionado)
        clipboard = QApplication.clipboard()
        clipboard.setText(nombre)
        self.statusBar().showMessage(f"📋 Copiado: {nombre}", 2000)

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()