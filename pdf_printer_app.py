import sys
import os
import json
import fitz  # PyMuPDF
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QLabel, QScrollArea, QListWidget,
                             QListWidgetItem, QCheckBox, QGridLayout, QStyledItemDelegate, QLineEdit,
                             QProgressDialog, QDialog, QMessageBox, QSizePolicy, QGroupBox)
from PyQt5.QtGui import QPixmap, QImage, QDragEnterEvent, QDropEvent, QPainter, QIcon, QFontMetrics
from PyQt5.QtCore import Qt, PYQT_VERSION_STR, QTimer, pyqtSlot, QSize, QEvent
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
import time

try:
    from updater import check_for_updates, CURRENT_VERSION
    UPDATER_AVAILABLE = True
except ImportError:
    UPDATER_AVAILABLE = False
    CURRENT_VERSION = "0.0.1"
    print("Update checker not available - some features will be disabled")

# Custom widget for PDF list items
class PDFListItem(QWidget):
    def __init__(self, filename, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 2, 5, 2)
        layout.setSpacing(10)
        
        self.checkbox = QCheckBox()
        self.checkbox.setFixedSize(20, 20)
        
        self.label = QLabel(os.path.basename(filename))
        self.label.setFixedWidth(450)  # Adjusted width
        self.label.setToolTip(filename)  # Show full path on hover
        
        fm = QFontMetrics(self.label.font())
        elided_text = fm.elidedText(os.path.basename(filename), Qt.ElideMiddle, self.label.width())
        self.label.setText(elided_text)
        
        layout.addWidget(self.checkbox)
        layout.addWidget(self.label)
        
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.checkbox.setStyleSheet("""
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                background-color: #3a3a3a;
                border: 2px solid #555555;
                border-radius: 2px;
            }
            QCheckBox::indicator:checked {
                background-color: #4CAF50;
                border-color: #4CAF50;
            }
        """)

    def sizeHint(self):
        return QSize(480, 30)  # Adjusted width

# Main application class
class PDFPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        print("PDFPrinterApp.__init__ started")
        self.setWindowTitle("PDF Print Station")
        self.setGeometry(100, 100, 1200, 800)
        self.setAcceptDrops(True)  # Enable drag and drop for the main window

        # Set application icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'app_icon.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            # Set taskbar icon for Windows
            if sys.platform == 'win32':
                import ctypes
                myappid = 'sandeepaulakh.pdfprintstation.1.0'  # Updated app ID
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QHBoxLayout(self.central_widget)

        self.pdf_files = []
        self.preview_update_timer = QTimer()
        self.preview_update_timer.setSingleShot(True)
        self.preview_update_timer.timeout.connect(self.update_preview)

        self.pdf_previews = {}  # Dictionary to store previews

        # Set up temp folder for previews
        self.setup_temp_folder()

        # Set up collections directory BEFORE UI initialization
        self.collections_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'collections')
        if not os.path.exists(self.collections_dir):
            os.makedirs(self.collections_dir)

        # Initialize basic UI components
        self.init_ui()
        
        # Load saved PDFs and update collections list
        self.load_pdf_list()
        self.update_collections_list()  # Make sure collections are loaded
        self.apply_dark_theme()
        print("PDFPrinterApp.__init__ completed")

        self.installEventFilter(self)

        # Add a timer for debouncing resize events
        self.resize_timer = QTimer()
        self.resize_timer.setSingleShot(True)
        self.resize_timer.setInterval(200)  # 200ms delay
        self.resize_timer.timeout.connect(self.update_preview)

        # Check for updates on startup
        if UPDATER_AVAILABLE:
            QTimer.singleShot(2000, lambda: check_for_updates(self))

    def eventFilter(self, obj, event):
        if event.type() == QEvent.DragEnter:
            self.dragEnterEvent(event)
            return True
        elif event.type() == QEvent.Drop:
            self.dropEvent(event)
            return True
        elif event.type() == QEvent.KeyPress and event.key() == Qt.Key_Delete:
            # Handle Delete key press
            if obj == self.all_files_list and self.all_files_list.hasFocus():
                self.remove_pdf()
                return True
            elif obj == self.selected_files_list and self.selected_files_list.hasFocus():
                self.remove_from_selection()
                return True
            elif obj == self.collections_list and self.collections_list.hasFocus():
                self.delete_collection()
                return True
            
        # Handle other events (keep existing event handling)
        elif obj == self.scroll_area.viewport() and event.type() == QEvent.Resize:
            self.resize_timer.start()
            return True
        
        return super().eventFilter(obj, event)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            print("Drag enter event accepted")
            event.accept()
        else:
            print("Drag enter event ignored")
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile().lower().endswith('.pdf')]
        print(f"Dropped files: {files}")
        
        if files:
            temp_label = QLabel("Processing dropped files...", self)
            temp_label.setStyleSheet("background-color: #3a3a3a; color: white; padding: 10px;")
            temp_label.move(self.rect().center() - temp_label.rect().center())
            temp_label.show()
            
            QTimer.singleShot(100, lambda: self.process_dropped_files(files, temp_label))

    def process_dropped_files(self, files, temp_label):
        self.add_pdf(files)
        temp_label.deleteLater()

    # Initialize the user interface
    def init_ui(self):
        main_layout = QHBoxLayout()
        main_layout.setSpacing(20)  # Add some space between left and right widgets

        # Left side: Lists and buttons
        left_widget = QWidget()
        left_widget.setFixedWidth(520)  # Set a fixed width for the entire left panel
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(10)  # Increase spacing between widgets

        # Set a fixed width for all elements in the left column
        fixed_width = 500  # Width for internal widgets

        # Selected files list (top)
        selected_header = QHBoxLayout()
        selected_header.setContentsMargins(0, 0, 0, 0)
        
        selected_label = QLabel("Selected Files")
        selected_label.setStyleSheet("font-weight: bold;")
        selected_header.addWidget(selected_label)
        selected_header.addStretch()
        left_layout.addLayout(selected_header)

        self.selected_files_list = QListWidget()
        self.selected_files_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.selected_files_list.setFixedWidth(fixed_width)
        self.selected_files_list.setMinimumHeight(200)  # Set consistent height
        left_layout.addWidget(self.selected_files_list)

        # Print button
        self.print_button = QPushButton("Print Selected")
        self.print_button.clicked.connect(self.print_pdf)
        self.print_button.setFixedWidth(fixed_width)
        self.print_button.setFixedHeight(32)  # Fixed height for consistency
        left_layout.addWidget(self.print_button)

        # Selection buttons
        selection_button_layout = QHBoxLayout()
        self.add_to_selection_button = QPushButton("Add to Print Selection ↑")
        self.add_to_selection_button.clicked.connect(self.add_to_selection)
        self.add_to_selection_button.setFixedWidth(fixed_width // 2 - 5)
        self.add_to_selection_button.setFixedHeight(32)  # Fixed height
        selection_button_layout.addWidget(self.add_to_selection_button)

        self.remove_from_selection_button = QPushButton("Remove from Selection ↓")
        self.remove_from_selection_button.clicked.connect(self.remove_from_selection)
        self.remove_from_selection_button.setFixedWidth(fixed_width // 2 - 5)
        self.remove_from_selection_button.setFixedHeight(32)  # Fixed height
        selection_button_layout.addWidget(self.remove_from_selection_button)
        left_layout.addLayout(selection_button_layout)

        # All files list (bottom)
        all_files_container = QWidget()
        all_files_layout = QVBoxLayout(all_files_container)
        all_files_layout.setSpacing(5)
        all_files_layout.setContentsMargins(0, 10, 0, 0)

        # Header with label and buttons
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        
        all_files_label = QLabel("All PDF Files")
        all_files_label.setStyleSheet("font-weight: bold;")
        header_layout.addWidget(all_files_label)
        
        header_layout.addStretch()
        
        # Add PDF button (Green hover)
        add_pdf_button = QPushButton("+")
        add_pdf_button.setFixedSize(28, 28)
        add_pdf_button.clicked.connect(lambda: self.add_pdf())
        add_pdf_button.setToolTip("Add PDF files")
        add_pdf_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0px;
                font-size: 18px;
                color: #ffffff;
            }
            QPushButton:hover {
                color: #4CAF50;
            }
        """)
        header_layout.addWidget(add_pdf_button)
        
        # Remove PDF button (Red hover)
        remove_pdf_button = QPushButton("-")
        remove_pdf_button.setFixedSize(28, 28)
        remove_pdf_button.clicked.connect(self.remove_pdf)
        remove_pdf_button.setToolTip("Remove selected PDF files")
        remove_pdf_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0px;
                font-size: 18px;
                color: #ffffff;
            }
            QPushButton:hover {
                color: #cc4444;
            }
        """)
        header_layout.addWidget(remove_pdf_button)
        
        all_files_layout.addLayout(header_layout)

        # Search bar
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search files...")
        self.search_bar.textChanged.connect(self.filter_files)
        self.search_bar.setFixedWidth(fixed_width)
        self.search_bar.setFixedHeight(32)  # Match height with other elements
        self.search_bar.setStyleSheet("""
            QLineEdit {
                padding: 5px 10px;
                background-color: #3a3a3a;
                border: 1px solid #555555;
                border-radius: 3px;
            }
        """)
        all_files_layout.addWidget(self.search_bar)

        # Files list with sort buttons overlay
        list_container = QWidget()
        list_layout = QGridLayout(list_container)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(0)

        # Files list
        self.all_files_list = QListWidget(self)
        self.all_files_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.all_files_list.setFixedWidth(fixed_width)
        self.all_files_list.setMinimumHeight(200)  # Set consistent height
        self.all_files_list.itemDoubleClicked.connect(self.on_pdf_double_click)
        list_layout.addWidget(self.all_files_list, 0, 0, 1, 1)

        # Sort buttons in bottom-right corner
        sort_buttons_widget = QWidget()
        sort_buttons_widget.setFixedSize(125, 30)
        sort_buttons_widget.setStyleSheet("QWidget { background-color: transparent; }")
        sort_buttons_layout = QHBoxLayout(sort_buttons_widget)
        sort_buttons_layout.setContentsMargins(0, 0, 15, 0)
        sort_buttons_layout.setSpacing(1)

        self.sort_asc_button = QPushButton("↑ A-Z")
        self.sort_asc_button.clicked.connect(lambda: self.sort_files(ascending=True))
        self.sort_asc_button.setFixedWidth(60)
        self.sort_asc_button.setToolTip("Sort files A to Z")
        self.sort_asc_button.setStyleSheet("""
            QPushButton {
                background-color: #3a3a3a;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 3px;
                font-size: 10px;
                color: #ffffff;
            }
            QPushButton:hover {
                background-color: #4a4a4a;
                border-color: #666666;
            }
        """)
        
        self.sort_desc_button = QPushButton("↓ Z-A")
        self.sort_desc_button.clicked.connect(lambda: self.sort_files(ascending=False))
        self.sort_desc_button.setFixedWidth(60)
        self.sort_desc_button.setToolTip("Sort files Z to A")
        self.sort_desc_button.setStyleSheet(self.sort_asc_button.styleSheet())

        sort_buttons_layout.addWidget(self.sort_asc_button)
        sort_buttons_layout.addWidget(self.sort_desc_button)
        
        # Position sort buttons in bottom-right corner
        list_layout.addWidget(sort_buttons_widget, 0, 0, 1, 1, Qt.AlignBottom | Qt.AlignRight)

        all_files_layout.addWidget(list_container, 1)  # Add stretch factor
        left_layout.addWidget(all_files_container, 2)  # Increased stretch factor

        # Add Collections section with action buttons
        collections_container = QWidget()
        collections_layout = QVBoxLayout(collections_container)
        collections_layout.setSpacing(5)
        collections_layout.setContentsMargins(0, 10, 0, 0)  # Added top margin for spacing
        
        # Header with label and buttons
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        
        collections_label = QLabel("Collections")
        collections_label.setStyleSheet("font-weight: bold;")
        header_layout.addWidget(collections_label)
        
        header_layout.addStretch()  # Push buttons to the right
        
        # Save button (Green hover)
        save_collection_button = QPushButton("+")
        save_collection_button.setFixedSize(28, 28)
        save_collection_button.clicked.connect(self.save_collection)
        save_collection_button.setToolTip("Save current list as a collection")
        save_collection_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0px;
                font-size: 18px;
                color: #ffffff;
            }
            QPushButton:hover {
                color: #4CAF50;
            }
        """)
        header_layout.addWidget(save_collection_button)
        
        # Load button (Orange hover)
        load_collection_button = QPushButton("↑")
        load_collection_button.setFixedSize(28, 28)
        load_collection_button.setToolTip("Load selected collection")
        load_collection_button.clicked.connect(lambda: self.load_collection(
            os.path.join(self.collections_dir, self.collections_list.currentItem().data(Qt.UserRole))
            if self.collections_list.currentItem() else None
        ))
        load_collection_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0px;
                font-size: 18px;
                color: #ffffff;
            }
            QPushButton:hover {
                color: #FFA500;
            }
        """)
        header_layout.addWidget(load_collection_button)
        
        # Delete button (Red hover)
        delete_collection_button = QPushButton("-")
        delete_collection_button.setFixedSize(28, 28)
        delete_collection_button.setToolTip("Delete selected collection")
        delete_collection_button.clicked.connect(self.delete_collection)
        delete_collection_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0px;
                font-size: 18px;
                color: #ffffff;
            }
            QPushButton:hover {
                color: #cc4444;
            }
        """)
        header_layout.addWidget(delete_collection_button)
        
        collections_layout.addLayout(header_layout)

        # Collections list with double-click support
        self.collections_list = QListWidget()
        self.collections_list.setFixedWidth(fixed_width)
        self.collections_list.setFixedHeight(100)  # Fixed height for collections
        self.collections_list.itemDoubleClicked.connect(self.on_collection_double_click)
        self.update_collections_list()
        collections_layout.addWidget(self.collections_list)
        
        left_layout.addWidget(collections_container)

        # Add Settings and Close buttons side by side
        bottom_buttons_container = QWidget()
        bottom_buttons_layout = QHBoxLayout(bottom_buttons_container)
        bottom_buttons_layout.setContentsMargins(0, 0, 0, 0)
        bottom_buttons_layout.setSpacing(5)
        
        # Settings button (Blue hover)
        settings_button = QPushButton("⚙")
        settings_button.setFixedSize(32, 32)
        settings_button.setToolTip("Open settings")
        settings_button.clicked.connect(self.show_settings)
        settings_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0px;
                font-size: 22px;
                color: #ffffff;
            }
            QPushButton:hover {
                color: #3498db;
            }
        """)
        bottom_buttons_layout.addWidget(settings_button)
        
        # Close app button (Red hover)
        close_button = QPushButton("×")
        close_button.setFixedSize(32, 32)
        close_button.setToolTip("Close application")
        close_button.clicked.connect(self.close)
        close_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0px;
                font-size: 22px;
                color: #ffffff;
            }
            QPushButton:hover {
                color: #cc4444;
            }
        """)
        bottom_buttons_layout.addWidget(close_button)
        
        # Add stretch to push buttons to the right
        bottom_buttons_layout.addStretch()
        
        left_layout.addWidget(bottom_buttons_container)

        # Right side: Preview
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        self.preview_label = QLabel("PDF Preview")
        right_layout.addWidget(self.preview_label)

        # Create scroll area for previews
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        right_layout.addWidget(self.scroll_area)

        # Create preview widget and layout
        self.preview_widget = QWidget()
        self.preview_layout = QGridLayout(self.preview_widget)
        self.preview_layout.setSpacing(7)
        self.preview_layout.setContentsMargins(7, 1, 7, 7)
        self.preview_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.scroll_area.setWidget(self.preview_widget)

        # Connect resize event
        self.scroll_area.resizeEvent = self.on_scroll_area_resize

        # Add widgets to main layout
        main_layout.addWidget(left_widget)
        main_layout.addWidget(right_widget)

        # Set the main layout
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Add event filters for keyboard shortcuts
        self.all_files_list.installEventFilter(self)
        self.selected_files_list.installEventFilter(self)
        self.collections_list.installEventFilter(self)

        # About button (Blue hover)
        about_button = QPushButton("ℹ")  # Info symbol
        about_button.setFixedSize(32, 32)
        about_button.clicked.connect(self.show_about)
        about_button.setToolTip("About PDF Print Station")
        about_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0px;
                font-size: 22px;
                color: #ffffff;
            }
            QPushButton:hover {
                color: #3498db;
            }
        """)
        right_layout.addWidget(about_button, alignment=Qt.AlignRight | Qt.AlignTop)

    def on_scroll_area_resize(self, event):
        # Instead of updating directly, start/restart the timer
        self.resize_timer.start()
        event.accept()

    def show_settings(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Settings")
        dialog.setFixedSize(400, 400)  # Reduced height since removing collections
        
        layout = QVBoxLayout(dialog)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Print Settings Group
        print_group = QGroupBox("Print Settings")
        print_layout = QVBoxLayout(print_group)
        print_layout.setSpacing(10)
        print_layout.setContentsMargins(10, 20, 10, 10)
        
        # Add double-sided printing checkbox
        self.double_sided_cb = QCheckBox("Add blank pages for double-sided printing")
        self.double_sided_cb.setChecked(getattr(self, 'add_blank_pages', True))
        self.double_sided_cb.stateChanged.connect(self.update_print_settings)
        print_layout.addWidget(self.double_sided_cb)
        
        layout.addWidget(print_group)
        
        # Cache Management Group
        cache_group = QGroupBox("Cache Management")
        cache_layout = QVBoxLayout(cache_group)
        cache_layout.setSpacing(10)
        cache_layout.setContentsMargins(10, 20, 10, 10)
        
        # Cache info
        cache_size = self.get_cache_size()
        cache_info = QLabel(f"Cache Size: {cache_size:.2f} MB")
        cache_layout.addWidget(cache_info)
        
        # Clear cache button
        clear_cache_button = QPushButton("Clear Preview Cache")
        clear_cache_button.setFixedHeight(32)
        clear_cache_button.clicked.connect(self.clear_cache)
        clear_cache_button.clicked.connect(lambda: cache_info.setText(f"Cache Size: {self.get_cache_size():.2f} MB"))
        cache_layout.addWidget(clear_cache_button)
        
        layout.addWidget(cache_group)
        
        # Add stretch to push everything to the top
        layout.addStretch()
        
        # Close button
        close_button = QPushButton("Close")
        close_button.setFixedHeight(32)
        close_button.clicked.connect(dialog.accept)
        layout.addWidget(close_button)
        
        # Add attribution label at the bottom
        attribution_label = QLabel("Made with ❤️ by @SandeepSAulakh")
        attribution_label.setAlignment(Qt.AlignRight)
        attribution_label.setStyleSheet("""
            QLabel {
                color: #888888;
                font-size: 9pt;
                padding: 5px;
                font-weight: normal;
            }
        """)
        layout.addWidget(attribution_label)
        
        dialog.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QPushButton {
                background-color: #3a3a3a;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 5px;
                color: #ffffff;
                font-size: 11pt;
            }
            QPushButton:hover {
                background-color: #4a4a4a;
                border-color: #666666;
            }
            QLabel {
                color: #ffffff;
                font-size: 11pt;
            }
            QGroupBox {
                border: 1px solid #555555;
                margin-top: 1ex;
                color: #ffffff;
                padding-top: 10px;
                font-size: 11pt;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
            }
        """)
        
        dialog.exec_()

    def get_cache_size(self):
        """Get the size of the preview cache in MB"""
        total_size = 0
        try:
            for file in os.listdir(self.temp_dir):
                file_path = os.path.join(self.temp_dir, file)
                if os.path.isfile(file_path):
                    total_size += os.path.getsize(file_path)
        except Exception as e:
            print(f"Error calculating cache size: {e}")
        return total_size / (1024 * 1024)  # Convert to MB

    def clear_cache(self):
        try:
            for file in os.listdir(self.temp_dir):
                try:
                    os.remove(os.path.join(self.temp_dir, file))
                except Exception as e:
                    print(f"Error removing temp file {file}: {e}")
            self.update_preview()  # Refresh the preview after clearing cache
        except Exception as e:
            print(f"Error clearing cache: {e}")

    def clear_old_previews(self, max_age_days=7):
        """Clear previews older than max_age_days"""
        try:
            current_time = time.time()
            for file in os.listdir(self.temp_dir):
                file_path = os.path.join(self.temp_dir, file)
                if os.path.isfile(file_path):
                    file_age = current_time - os.path.getmtime(file_path)
                    if file_age > (max_age_days * 24 * 60 * 60):
                        try:
                            os.remove(file_path)
                        except Exception as e:
                            print(f"Error removing old preview {file}: {e}")
        except Exception as e:
            print(f"Error clearing old previews: {e}")

    # Add selected PDFs to the print selection
    def add_to_selection(self):
        selected_items = self.all_files_list.selectedItems()  # Get all selected items
        if not selected_items:
            return

        # Show loading indicator
        loading_label = QLabel("Adding files...", self)
        loading_label.setStyleSheet("background-color: #3a3a3a; padding: 10px; border-radius: 5px;")
        loading_label.move(self.rect().center() - loading_label.rect().center())
        loading_label.show()
        QApplication.processEvents()

        try:
            # Get list of files already in selection
            existing_items = [self.selected_files_list.item(i).data(Qt.UserRole) 
                             for i in range(self.selected_files_list.count())]
            
            # Add new items without updating preview for each one
            for item in selected_items:
                file_path = item.data(Qt.UserRole)
                if file_path not in existing_items:
                    new_item = QListWidgetItem(item.text())
                    new_item.setData(Qt.UserRole, file_path)
                    self.selected_files_list.addItem(new_item)
        finally:
            loading_label.hide()
            loading_label.deleteLater()
        
        # Update preview once after all items are added
        self.update_preview()

    # Remove PDFs from the print selection
    def remove_from_selection(self):
        for item in self.selected_files_list.selectedItems():
            self.selected_files_list.takeItem(self.selected_files_list.row(item))
        self.update_preview()

    # Add new PDFs to the application
    def add_pdf(self, file_names=None):
        if file_names is None or isinstance(file_names, bool):
            file_names, _ = QFileDialog.getOpenFileNames(self, "Open PDF File", "", "PDF Files (*.pdf)")
        
        if not file_names:
            return
        
        progress = QProgressDialog("Processing PDF files...", "Cancel", 0, len(file_names), self)
        progress.setWindowModality(Qt.WindowModal)
        
        try:
            for i, file_name in enumerate(file_names):
                if progress.wasCanceled():
                    break
                
                progress.setValue(i)
                if os.path.exists(file_name) and file_name.lower().endswith('.pdf'):
                    if self.generate_preview(file_name):
                        item = QListWidgetItem(os.path.basename(file_name))
                        item.setData(Qt.UserRole, file_name)
                        self.all_files_list.addItem(item)
        finally:
            progress.setValue(len(file_names))
            self.save_pdf_list()

    # Remove PDFs from the application
    def remove_pdf(self):
        for item in self.all_files_list.selectedItems():
            self.all_files_list.takeItem(self.all_files_list.row(item))
        self.save_pdf_list()

    # Update the PDF preview
    def update_preview(self):
        try:
            # Clear previous previews
            for i in reversed(range(self.preview_layout.count())): 
                widget = self.preview_layout.itemAt(i).widget()
                if widget:
                    widget.setParent(None)
                    widget.deleteLater()
            
            import gc
            gc.collect()

            # Fixed sizes and spacing
            preview_size = 160
            spacing = 7
            side_margin = 7
            top_margin = 1
            
            # Calculate number of columns based on viewport width
            viewport_width = self.scroll_area.viewport().width()
            item_width = preview_size + spacing
            columns = max(1, (viewport_width - 2 * side_margin) // item_width)
            
            # Configure grid layout
            self.preview_layout.setSpacing(spacing)
            self.preview_layout.setContentsMargins(side_margin, top_margin, side_margin, side_margin)
            self.preview_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)
            
            # Force the preview widget to use at least the viewport width
            self.preview_widget.setMinimumWidth(viewport_width - 20)
            
            # Add previews to grid
            row = col = 0
            for i in range(self.selected_files_list.count()):
                file_name = self.selected_files_list.item(i).data(Qt.UserRole)
                preview_path = os.path.join(self.temp_dir, f"preview_{os.path.basename(file_name)}.png")
                
                if os.path.exists(preview_path):
                    label = QLabel()
                    label.setFixedSize(preview_size, preview_size)
                    pixmap = QPixmap(preview_path)
                    scaled_pixmap = pixmap.scaled(preview_size, preview_size,
                                                Qt.KeepAspectRatio, 
                                                Qt.SmoothTransformation)
                    
                    label.setPixmap(scaled_pixmap)
                    label.setAlignment(Qt.AlignCenter)
                    label.setStyleSheet("QLabel { background-color: #2b2b2b; }")
                    
                    self.preview_layout.addWidget(label, row, col)
                    
                    col += 1
                    if col >= columns:
                        col = 0
                        row += 1
                
        except Exception as e:
            print(f"Error updating preview: {e}")

    # Filter PDF files based on search text
    def filter_files(self, text):
        for i in range(self.all_files_list.count()):
            item = self.all_files_list.item(i)
            file_name = item.data(Qt.UserRole)
            item.setHidden(text.lower() not in os.path.basename(file_name).lower())

    # Save the list of PDFs to a JSON file
    def save_pdf_list(self):
        try:
            pdf_files = [self.all_files_list.item(i).data(Qt.UserRole) 
                         for i in range(self.all_files_list.count())]
            save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pdf_list.json')
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(pdf_files, f, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving PDF list: {e}")

    # Load the list of PDFs from a JSON file
    def load_pdf_list(self):
        try:
            pdf_list_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pdf_list.json')
            if not os.path.exists(pdf_list_path):
                return
                
            with open(pdf_list_path, 'r', encoding='utf-8') as f:
                pdf_files = json.load(f)
                
            # Show loading indicator
            loading_label = QLabel("Loading previews...", self)
            loading_label.setStyleSheet("background-color: #3a3a3a; padding: 10px; border-radius: 5px;")
            loading_label.move(self.rect().center() - loading_label.rect().center())
            loading_label.show()
            QApplication.processEvents()
            
            try:
                for file_path in pdf_files:
                    if not os.path.exists(file_path):
                        continue
                        
                    if self.generate_preview(file_path):
                        item = QListWidgetItem(os.path.basename(file_path))
                        item.setData(Qt.UserRole, file_path)
                        self.all_files_list.addItem(item)
            finally:
                loading_label.hide()
                loading_label.deleteLater()
                
        except FileNotFoundError:
            print("No saved PDF list found. Starting with an empty list.")
        except json.JSONDecodeError:
            print("Error reading the saved file list. Starting with an empty list.")
        except Exception as e:
            print(f"An error occurred while loading the PDF list: {e}")

    # Print selected PDFs
    def print_pdf(self):
        selected_files = [self.selected_files_list.item(i).data(Qt.UserRole) 
                          for i in range(self.selected_files_list.count())]
        if not selected_files:
            self.show_error_dialog("No Files Selected", "Please select PDF files to print.")
            return
        
        try:
            # Create progress dialog
            progress = QProgressDialog("Preparing PDFs for printing...", "Cancel", 0, len(selected_files), self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setWindowTitle("Processing PDFs")
            progress.setMinimumDuration(0)  # Show immediately
            
            # Create a PDF document containing all selected PDFs
            combined_pdf = fitz.open()
            failed_files = []
            
            for i, pdf_file in enumerate(selected_files):
                if progress.wasCanceled():
                    return
                
                progress.setValue(i)
                progress.setLabelText(f"Processing: {os.path.basename(pdf_file)}")
                
                if not os.path.exists(pdf_file):
                    failed_files.append(f"{os.path.basename(pdf_file)} (file not found)")
                    continue
                
                try:
                    doc = fitz.open(pdf_file)
                    # Insert the PDF
                    combined_pdf.insert_pdf(doc)
                    
                    # Add blank page if enabled and document has odd number of pages
                    if getattr(self, 'add_blank_pages', True) and doc.page_count % 2 != 0:
                        combined_pdf.new_page(-1,  # Insert at end
                                               width=doc[0].rect.width,  # Match first page dimensions
                                               height=doc[0].rect.height)
                    
                except Exception as e:
                    failed_files.append(f"{os.path.basename(pdf_file)} ({str(e)})")
                    continue
                finally:
                    doc.close()
            
            progress.setValue(len(selected_files))
            
            if failed_files:
                self.show_error_dialog("Print Errors", 
                    "The following files had errors:\n" + "\n".join(failed_files))
                if not combined_pdf.page_count:
                    return
            
            # Save the combined PDF to a temporary file
            progress.setLabelText("Creating combined PDF file...")
            temp_pdf_path = os.path.join(os.path.dirname(selected_files[0]), "temp_combined.pdf")
            combined_pdf.save(temp_pdf_path)
            combined_pdf.close()
            
            # Platform-specific print handling
            if sys.platform == "darwin":  # macOS
                os.system(f"open -a 'Preview' '{temp_pdf_path}'")
            elif sys.platform == "win32":  # Windows
                # Use native Windows print dialog
                printer = QPrinter(QPrinter.HighResolution)
                print_dialog = QPrintDialog(printer, self)
                if print_dialog.exec_() == QDialog.Accepted:
                    progress.setLabelText("Printing...")
                    # Open and print the PDF using PyMuPDF
                    doc = fitz.open(temp_pdf_path)
                    for page_num in range(doc.page_count):
                        if progress.wasCanceled():
                            break
                        progress.setValue(page_num)
                        progress.setLabelText(f"Printing page {page_num + 1} of {doc.page_count}")
                        
                        page = doc[page_num]
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                        if page_num == 0:
                            printer.setPageSize(QPageSize(QSizeF(page.rect.width, page.rect.height), QPageSize.Point))
                        
                        painter = QPainter(printer)
                        painter.drawImage(printer.pageRect(), img)
                        if page_num < doc.page_count - 1:
                            printer.newPage()
                        painter.end()
                    doc.close()
            else:  # Linux
                os.system(f"xdg-open '{temp_pdf_path}'")
            
            # Clean up the temporary file after a delay
            QTimer.singleShot(30000, lambda: os.remove(temp_pdf_path) if os.path.exists(temp_pdf_path) else None)
            
        except Exception as e:
            self.show_error_dialog("Print Error", f"An error occurred while printing: {str(e)}")
            print(f"Error printing PDFs: {e}")
        finally:
            progress.close()

    # Apply dark theme to the application
    def apply_dark_theme(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QLabel {
                font-weight: bold;
            }
            QPushButton {
                background-color: #3a3a3a;
                border: 1px solid #555555;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #4a4a4a;
            }
            QLineEdit, QListWidget {
                background-color: #3a3a3a;
                border: 1px solid #555555;
            }
            QScrollBar:vertical {
                border: 1px solid #555555;
                background: #3a3a3a;
                width: 10px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #555555;
                min-height: 20px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: 1px solid #555555;
                background: #3a3a3a;
                height: 0px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }
        """)

    def cleanup_resources(self):
        # Clear temp preview files that aren't in use
        if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
            current_files = set()
            # Collect currently used preview files
            for i in range(self.all_files_list.count()):
                file_name = self.all_files_list.item(i).data(Qt.UserRole)
                preview_name = f"preview_{os.path.basename(file_name)}.png"
                current_files.add(preview_name)
            
            # Remove only unused preview files
            for file in os.listdir(self.temp_dir):
                if file not in current_files:
                    try:
                        os.remove(os.path.join(self.temp_dir, file))
                    except Exception as e:
                        print(f"Error removing temp file: {e}")

    # Call this method when closing the application
    def closeEvent(self, event):
        self.cleanup_resources()
        self.save_pdf_list()
        event.accept()

    def generate_preview(self, file_path):
        preview_filename = f"preview_{os.path.basename(file_path)}.png"
        preview_path = os.path.join(self.temp_dir, preview_filename)
        
        if not os.path.exists(preview_path):
            doc = None
            try:
                doc = fitz.open(file_path)
                if doc.page_count > 0:
                    page = doc[0]
                    # Reduced matrix size for better performance
                    matrix = fitz.Matrix(0.2, 0.2)
                    # Disable alpha and use RGB colorspace for smaller files
                    pix = page.get_pixmap(matrix=matrix, alpha=False, colorspace="rgb")
                    pix.save(preview_path, output="png", jpg_quality=85)
                else:
                    print(f"Warning: {file_path} has no pages")
                    return False
            except Exception as e:
                print(f"Error generating preview for {file_path}: {e}")
                return False
            finally:
                if doc:
                    doc.close()
        return True

    def setup_temp_folder(self):
        # Create temp folder in the same directory as the script
        self.temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_previews')
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)

    def show_error_dialog(self, title, message):
        QMessageBox.critical(self, title, message)

    # Add this new method to handle sorting
    def sort_files(self, ascending=True):
        # Get all items data from the list
        items_data = []
        for i in range(self.all_files_list.count()):
            item = self.all_files_list.item(i)
            items_data.append({
                'text': item.text(),
                'data': item.data(Qt.UserRole)
            })
        
        # Sort items based on filename
        items_data.sort(key=lambda x: os.path.basename(x['data']).lower(), 
                       reverse=not ascending)
        
        # Clear and re-add items in sorted order
        self.all_files_list.clear()
        for item_data in items_data:
            new_item = QListWidgetItem(item_data['text'])
            new_item.setData(Qt.UserRole, item_data['data'])
            self.all_files_list.addItem(new_item)
        
        # Save the sorted list
        self.save_pdf_list()

    # Add these new methods for collection management
    def save_collection(self):
        try:
            # Get file name for saving
            file_name, _ = QFileDialog.getSaveFileName(
                self,
                "Save Collection",
                os.path.join(self.collections_dir, "New Collection"),  # Default name
                "PDF Collection (*.pdfcol)"  # Removed "All Files" option
            )
            
            if not file_name:
                return
                
            # Add .pdfcol extension if not present
            if not file_name.endswith('.pdfcol'):
                file_name += '.pdfcol'
            
            # Save collection
            collection_data = {
                'files': [
                    {
                        'path': self.all_files_list.item(i).data(Qt.UserRole),
                        'name': self.all_files_list.item(i).text()
                    }
                    for i in range(self.all_files_list.count())
                ],
                'date_saved': time.strftime('%Y-%m-%d %H:%M:%S'),
                'version': '1.0'
            }
            
            with open(file_name, 'w', encoding='utf-8') as f:
                json.dump(collection_data, f, ensure_ascii=False, indent=2)
                
            self.update_collections_list()
            QMessageBox.information(self, "Success", "Collection saved successfully!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save collection: {str(e)}")

    def load_collection(self, selected_item=None):
        try:
            if selected_item is None and self.collections_list.currentItem():
                # Get the full filename from the item's data
                filename = self.collections_list.currentItem().data(Qt.UserRole)
                file_path = os.path.join(self.collections_dir, filename)  # Removed extra .pdfcol
            else:
                file_path = selected_item
                
            if not file_path:
                return
                
            # Load collection data
            with open(file_path, 'r', encoding='utf-8') as f:
                collection_data = json.load(f)
                
            # Verify version compatibility
            if 'version' not in collection_data:
                raise ValueError("Invalid collection file format")
                
            # Ask user about loading behavior
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setText("How would you like to load the collection?")
            msg_box.setWindowTitle("Load Collection")
            replace_button = msg_box.addButton("Replace Current", QMessageBox.ActionRole)
            append_button = msg_box.addButton("Append to Current", QMessageBox.ActionRole)
            cancel_button = msg_box.addButton("Cancel", QMessageBox.RejectRole)
            
            msg_box.exec_()
            clicked_button = msg_box.clickedButton()
            
            if clicked_button == cancel_button:
                return
                
            # Clear list if replacing
            if clicked_button == replace_button:
                self.all_files_list.clear()
            
            # Add files from collection
            missing_files = []
            for file_data in collection_data['files']:
                if os.path.exists(file_data['path']):
                    item = QListWidgetItem(file_data['name'])
                    item.setData(Qt.UserRole, file_data['path'])
                    self.all_files_list.addItem(item)
                    self.generate_preview(file_data['path'])
                else:
                    missing_files.append(file_data['name'])
            
            # Report any missing files
            if missing_files:
                QMessageBox.warning(
                    self,
                    "Missing Files",
                    "The following files were not found:\n" + "\n".join(missing_files)
                )
            
            self.save_pdf_list()  # Save the updated list
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load collection: {str(e)}")

    def update_print_settings(self, state):
        self.add_blank_pages = bool(state)

    def update_collections_list(self):
        """Update the collections list in main window"""
        if hasattr(self, 'collections_list'):
            self.collections_list.clear()
            if hasattr(self, 'collections_dir') and os.path.exists(self.collections_dir):
                # Get all collection files and sort them by name
                collection_files = [f for f in os.listdir(self.collections_dir) if f.endswith('.pdfcol')]
                collection_files.sort(key=lambda x: x.lower())  # Sort case-insensitive
                
                for file in collection_files:
                    # Remove the .pdfcol extension for display
                    display_name = os.path.splitext(file)[0]
                    item = QListWidgetItem(display_name)
                    # Store full filename as data
                    item.setData(Qt.UserRole, file)
                    self.collections_list.addItem(item)

    def delete_collection(self):
        """Delete selected collection"""
        current_item = self.collections_list.currentItem()
        if not current_item:
            return
        
        reply = QMessageBox.question(
            self, 'Delete Collection',
            f"Are you sure you want to delete '{current_item.text()}'?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                # Get the full filename from the item's data
                filename = current_item.data(Qt.UserRole)
                file_path = os.path.join(self.collections_dir, filename)  # Removed extra .pdfcol
                os.remove(file_path)
                self.update_collections_list()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete collection: {str(e)}")

    def show_collections_settings(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Collections Management")
        dialog.setFixedSize(400, 200)
        
        layout = QVBoxLayout(dialog)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("Save Collection")
        save_button.clicked.connect(self.save_collection)
        buttons_layout.addWidget(save_button)
        
        load_button = QPushButton("Load Collection")
        load_button.clicked.connect(lambda: self.load_collection(
            os.path.join(self.collections_dir, self.collections_list.currentItem().text())
            if self.collections_list.currentItem() else None
        ))
        buttons_layout.addWidget(load_button)
        
        delete_button = QPushButton("Delete Collection")
        delete_button.clicked.connect(self.delete_collection)
        buttons_layout.addWidget(delete_button)
        
        layout.addLayout(buttons_layout)
        
        # Close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(dialog.accept)
        layout.addWidget(close_button)
        
        dialog.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QPushButton {
                background-color: #3a3a3a;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 5px;
                color: #ffffff;
            }
            QPushButton:hover {
                background-color: #4a4a4a;
                border-color: #666666;
            }
        """)
        
        dialog.exec_()

    def on_collection_double_click(self, item):
        """Handle double-click on collection item"""
        try:
            # Get the full filename from the item's data
            filename = item.data(Qt.UserRole)
            file_path = os.path.join(self.collections_dir, filename)
            self.load_collection(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load collection: {str(e)}")

    def on_pdf_double_click(self, item):
        """Handle double-click on PDF item to add it to selection"""
        try:
            # Check if item is already in selected list
            existing_items = [self.selected_files_list.item(i).data(Qt.UserRole) 
                         for i in range(self.selected_files_list.count())]
            
            file_path = item.data(Qt.UserRole)
            if file_path not in existing_items:
                new_item = QListWidgetItem(item.text())
                new_item.setData(Qt.UserRole, file_path)
                self.selected_files_list.addItem(new_item)
                self.update_preview()
        except Exception as e:
            print(f"Error adding file to selection: {e}")

    def show_about(self):
        """Show about dialog"""
        about_text = f"""
        <h2>PDF Print Station v{CURRENT_VERSION}</h2>
        <p>A modern, user-friendly PDF management and printing application.</p>
        <p>Made with ❤️ by <a href='https://github.com/SandeepSAulakh'>Sandeep S Aulakh</a></p>
        <p>Copyright © 2024 Sandeep S Aulakh</p>
        <p><a href='https://github.com/SandeepSAulakh/PDF-Print-Station'>GitHub Repository</a></p>
        """
        QMessageBox.about(self, "About PDF Print Station", about_text)

# Main function to run the application
def main():
    import os
    os.environ['QT_ACCESSIBILITY'] = '0'
    os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.accessibility.core=false'
    
    print(f"Python version: {sys.version}")
    print(f"PyQt5 version: {PYQT_VERSION_STR}")
    print(f"PyMuPDF version: {fitz.version[0]}")
    print(f"Current working directory: {os.getcwd()}")
    
    app = QApplication(sys.argv)
    app.setApplicationName("PDF Print Station")
    app.setApplicationDisplayName("PDF Print Station")
    app.setOrganizationName("Sandeep S Aulakh")
    app.setOrganizationDomain("sandeepaulakh.com")
    
    try:
        print("Initializing PDFPrinterApp...")
        # Set application icon for all windows
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'app_icon.png')
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
            
        window = PDFPrinterApp()
        print("PDFPrinterApp initialized successfully.")
        
        print("Showing window...")
        window.show()
        print("Window shown successfully.")
        
        print("Entering main event loop...")
        return app.exec_()
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
