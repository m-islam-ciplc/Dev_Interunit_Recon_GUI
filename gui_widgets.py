"""
GUI Widgets for Interunit Loan Matcher
Modular UI components for better maintainability
"""

import os
import time
from pathlib import Path
from typing import Dict, Any

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QProgressBar, 
    QTextEdit, QFileDialog, QMessageBox, QApplication
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent


class FileSelectionWidget(QWidget):
    """Widget for file selection with drag and drop support"""
    
    files_selected = Signal(str, str)  # file1_path, file2_path
    
    def __init__(self):
        super().__init__()
        self.file1_path = ""
        self.file2_path = ""
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(5)
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setProperty("class", "section-container")
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(15, 15, 15, 15)  # Add padding inside container
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Select Interunit Loan Ledgers")
        title.setProperty("class", "heading")
        title.setStyleSheet("margin-bottom: 8px;")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(title)
        
        # Browse button
        self.browse_button = QPushButton("Browse Ledgers")
        self.browse_button.setProperty("class", "browse-button")
        self.browse_button.setMinimumSize(80, 20)
        self.browse_button.clicked.connect(self.select_both_files)
        section_layout.addWidget(self.browse_button)
        
        # Selected files section
        files_label = QLabel("Selected Ledgers")
        files_label.setProperty("class", "heading")
        files_label.setStyleSheet("margin-bottom: 8px;")
        files_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(files_label)
        
        # File list container - static height for 2 files
        self.files_container = QWidget()
        self.files_container.setProperty("class", "files-list")
        self.files_container.setFixedHeight(80)  # Increased height for better visibility
        files_layout = QVBoxLayout()
        files_layout.setContentsMargins(8, 8, 8, 8)
        files_layout.setSpacing(6)  # Increased spacing between files
        self.files_container.setLayout(files_layout)
        section_layout.addWidget(self.files_container)
        
        # Clear button
        self.clear_files_button = QPushButton("Clear Ledgers")
        self.clear_files_button.setProperty("class", "clear-button")
        self.clear_files_button.setMinimumSize(70, 20)
        self.clear_files_button.clicked.connect(self.clear_files)
        section_layout.addWidget(self.clear_files_button)
        
        # Run Match button
        self.run_match_button = QPushButton("Run Match")
        self.run_match_button.setProperty("class", "run-button")
        self.run_match_button.setMinimumSize(70, 24)
        self.run_match_button.clicked.connect(self.run_matching)
        self.run_match_button.setEnabled(False)
        section_layout.addWidget(self.run_match_button)
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
        self.setAcceptDrops(True)
    
    def select_both_files(self):
        """Open file dialog to select both files at once"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Both Excel Files",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if len(files) >= 2:
            # Set first file as File 1 (Pole Book)
            self.set_file(1, files[0])
            # Set second file as File 2 (Steel Book)
            self.set_file(2, files[1])
            
            # If more than 2 files selected, show warning
            if len(files) > 2:
                QMessageBox.information(
                    self, 
                    "Multiple Files Selected", 
                    f"Selected {len(files)} files. Using first two files:\n\n"
                    f"File 1: {os.path.basename(files[0])}\n"
                    f"File 2: {os.path.basename(files[1])}"
                )
        elif len(files) == 1:
            QMessageBox.warning(
                self, 
                "Insufficient Files", 
                "Please select at least 2 Excel files for matching."
            )
    
    def clear_files(self):
        """Clear both file selections"""
        self.file1_path = ""
        self.file2_path = ""
        self.update_file_display()
            
    def set_file(self, file_num: int, file_path: str):
        """Set the selected file path"""
        if file_num == 1:
            self.file1_path = file_path
        else:
            self.file2_path = file_path
        
        self.update_file_display()
        self.validate_files()
    
    def update_file_display(self):
        """Update the file display list"""
        # Clear existing file items
        layout = self.files_container.layout()
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        # Add current files with tick icons and actual filenames
        if self.file1_path:
            file1_name = os.path.basename(self.file1_path)
            file1_item = QLabel(f"✓ {file1_name}")
            file1_item.setProperty("class", "file-item")
            file1_item.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            layout.addWidget(file1_item)
        
        if self.file2_path:
            file2_name = os.path.basename(self.file2_path)
            file2_item = QLabel(f"✓ {file2_name}")
            file2_item.setProperty("class", "file-item")
            file2_item.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            layout.addWidget(file2_item)
        
        # Enable/disable Run Match button
        self.run_match_button.setEnabled(bool(self.file1_path and self.file2_path))
    
    def run_matching(self):
        """Trigger matching process"""
        if self.file1_path and self.file2_path:
            self.files_selected.emit(self.file1_path, self.file2_path)
            # Emit a signal to start matching
            if hasattr(self.parent(), 'start_matching'):
                self.parent().start_matching()
        
    def validate_files(self):
        """Validate that both files are selected and are valid Excel files"""
        if self.file1_path and self.file2_path:
            # Check if files exist and are Excel files
            if (os.path.exists(self.file1_path) and os.path.exists(self.file2_path) and
                (self.file1_path.endswith('.xlsx') or self.file1_path.endswith('.xls')) and
                (self.file2_path.endswith('.xlsx') or self.file2_path.endswith('.xls'))):
                
                self.files_selected.emit(self.file1_path, self.file2_path)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter event"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        """Handle drop event"""
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        
        # Filter for Excel files only
        excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
        
        if len(excel_files) >= 2:
            self.set_file(1, excel_files[0])
            self.set_file(2, excel_files[1])
            
            # Show info if more than 2 Excel files were dropped
            if len(excel_files) > 2:
                QMessageBox.information(
                    self,
                    "Multiple Files Dropped",
                    f"Dropped {len(excel_files)} Excel files. Using first two:\n\n"
                    f"File 1: {os.path.basename(excel_files[0])}\n"
                    f"File 2: {os.path.basename(excel_files[1])}"
                )
        elif len(excel_files) == 1:
            QMessageBox.warning(
                self,
                "Insufficient Files",
                f"Only 1 Excel file dropped. Please drop at least 2 Excel files.\n\n"
                f"Dropped: {os.path.basename(excel_files[0])}"
            )
        elif len(files) > 0:
            QMessageBox.warning(
                self,
                "Invalid Files",
                f"No Excel files found in dropped files.\n\n"
                f"Please drop .xlsx or .xls files only."
            )
        
        event.acceptProposedAction()


class ProcessingWidget(QWidget):
    """Widget for displaying processing progress and status"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(5)
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setProperty("class", "section-container")
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(15, 15, 15, 15)  # Add padding inside container
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Match Progress")
        title.setProperty("class", "heading")
        title.setStyleSheet("margin-bottom: 8px;")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(title)
        
        # Step progress
        self.step_labels = {}
        self.step_progresses = {}
        
        steps = [
            "Narration Matches",
            "LC Matches",
            "One to One PO Matches",
            "Interunit Matches",
            "USD Matches"
        ]
        
        for step in steps:
            step_layout = QHBoxLayout()
            step_layout.setSpacing(5)
            
            # Step name
            step_label = QLabel(f"{step}")
            step_label.setFixedWidth(180)  # Fixed width for all labels
            step_label.setFont(QFont("Segoe UI", 10))
            step_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            
            step_layout.addWidget(step_label)
            
            # Step progress bar
            step_progress = QProgressBar()
            step_progress.setRange(0, 100)
            step_progress.setValue(0)
            step_progress.setFixedHeight(20)  # Fixed height for all progress bars
            step_progress.setProperty("class", "step-progress")
            step_layout.addWidget(step_progress)
            
            section_layout.addLayout(step_layout)
            
            self.step_labels[step] = step_label
            self.step_progresses[step] = step_progress
        
        # Add some spacing before overall progress
        section_layout.addSpacing(10)
        
        # Overall progress section
        overall_label = QLabel("Overall Progress")
        overall_label.setProperty("class", "heading")
        overall_label.setStyleSheet("margin-bottom: 8px;")
        overall_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(overall_label)
        
        self.overall_progress = QProgressBar()
        self.overall_progress.setRange(0, 100)
        self.overall_progress.setValue(0)
        self.overall_progress.setProperty("class", "overall-progress")
        section_layout.addWidget(self.overall_progress)
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
    
    def complete_step(self, step_name: str, matches_found: int):
        """Mark a step as completed"""
        # Map full step names to display names
        step_mapping = {
            "Narration Matching": "Narration Matches",
            "LC Matching": "LC Matches",
            "PO Matching": "One to One PO Matches",
            "Interunit Matching": "Interunit Matches",
            "USD Matching": "USD Matches"
        }
        
        short_name = step_mapping.get(step_name, step_name)
        
        if short_name in self.step_progresses:
            progress_bar = self.step_progresses[short_name]
            progress_bar.setValue(100)
            
            # Also update the step label to show completion
            if short_name in self.step_labels:
                self.step_labels[short_name].setProperty("class", "step-completed")
    
    def reset_progress(self):
        """Reset all progress indicators"""
        for step_name, progress_bar in self.step_progresses.items():
            progress_bar.setValue(0)
            
            # Reset step label styling
            if step_name in self.step_labels:
                self.step_labels[step_name].setProperty("class", "step-pending")
    
    def set_processing_state(self, is_processing: bool):
        """Enable/disable processing controls"""
        pass


class ResultsWidget(QWidget):
    """Widget for displaying matching results and statistics"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(5)
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setProperty("class", "section-container")
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(15, 15, 15, 15)  # Add padding inside container
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Match Summary")
        title.setProperty("class", "heading")
        title.setStyleSheet("margin-bottom: 8px;")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(title)
        
        # Results summary - all match types and total in one row
        single_row_layout = QHBoxLayout()
        single_row_layout.setSpacing(5)
        
        # Narration Matches - bold text style
        self.narration_matches_label = QLabel("Narration: 0")
        self.narration_matches_label.setProperty("class", "match-summary-text")
        single_row_layout.addWidget(self.narration_matches_label)
        
        # LC Matches - bold text style
        self.lc_matches_label = QLabel("LC: 0")
        self.lc_matches_label.setProperty("class", "match-summary-text")
        single_row_layout.addWidget(self.lc_matches_label)
        
        # PO Matches - bold text style
        self.po_matches_label = QLabel("PO: 0")
        self.po_matches_label.setProperty("class", "match-summary-text")
        single_row_layout.addWidget(self.po_matches_label)
        
        # Interunit Matches - bold text style
        self.interunit_matches_label = QLabel("Interunit: 0")
        self.interunit_matches_label.setProperty("class", "match-summary-text")
        single_row_layout.addWidget(self.interunit_matches_label)
        
        # USD Matches - bold text style
        self.usd_matches_label = QLabel("USD: 0")
        self.usd_matches_label.setProperty("class", "match-summary-text")
        single_row_layout.addWidget(self.usd_matches_label)
        
        # Add some spacing before total
        single_row_layout.addStretch(1)
        
        # Total Matches box - smaller height but keeps current style
        self.total_matches_label = QLabel("Total Matches: 0")
        self.total_matches_label.setProperty("class", "total-match-summary")
        self.total_matches_label.setMinimumSize(150, 30)
        single_row_layout.addWidget(self.total_matches_label)
        
        section_layout.addLayout(single_row_layout)
        
        # Button row for opening files/folders
        button_row = QHBoxLayout()
        button_row.setSpacing(5)
        
        # Open Folder button
        self.open_folder_button = QPushButton("Open Output Folder")
        self.open_folder_button.setProperty("class", "open-folder-button")
        self.open_folder_button.setMinimumSize(120, 30)
        # Note: Connection will be set up in MainWindow
        self.open_folder_button.setEnabled(False)  # Disabled until files are processed
        button_row.addWidget(self.open_folder_button)
        
        # Open Output Files button
        self.open_files_button = QPushButton("Open Output Files")
        self.open_files_button.setProperty("class", "open-files-button")
        self.open_files_button.setMinimumSize(120, 30)
        # Note: Connection will be set up in MainWindow
        self.open_files_button.setEnabled(False)  # Disabled until files are processed
        button_row.addWidget(self.open_files_button)
        
        section_layout.addLayout(button_row)
        
        # Add stretch to push content to top (consistent with other sections)
        section_layout.addStretch()
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
    
    def update_results(self, statistics: Dict[str, Any]):
        """Update results display with statistics"""
        self.narration_matches_label.setText(f"Narration: {statistics.get('narration_matches', 0)}")
        self.lc_matches_label.setText(f"LC: {statistics.get('lc_matches', 0)}")
        self.po_matches_label.setText(f"PO: {statistics.get('po_matches', 0)}")
        self.interunit_matches_label.setText(f"Interunit: {statistics.get('interunit_matches', 0)}")
        self.usd_matches_label.setText(f"USD: {statistics.get('usd_matches', 0)}")
        
        # Calculate total matches
        total = (statistics.get('narration_matches', 0) + 
                statistics.get('lc_matches', 0) + 
                statistics.get('po_matches', 0) + 
                statistics.get('interunit_matches', 0) + 
                statistics.get('usd_matches', 0))
        self.total_matches_label.setText(f"Total Matches: {total}")
        
        # Enable the Open buttons when results are available
        if total > 0:
            self.open_folder_button.setEnabled(True)
            self.open_files_button.setEnabled(True)
    
    def reset_results(self):
        """Reset results display"""
        self.narration_matches_label.setText("Narration: 0")
        self.lc_matches_label.setText("LC: 0")
        self.po_matches_label.setText("PO: 0")
        self.interunit_matches_label.setText("Interunit: 0")
        self.usd_matches_label.setText("USD: 0")
        self.total_matches_label.setText("Total Matches: 0")
        self.open_folder_button.setEnabled(False)
        self.open_files_button.setEnabled(False)


class LogWidget(QWidget):
    """Widget for displaying processing logs"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(5)
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setProperty("class", "section-container")
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(15, 15, 15, 15)  # Add padding inside container
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Process Log")
        title.setProperty("class", "heading")
        title.setStyleSheet("margin-bottom: 8px;")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(title)
        
        # Log text area
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        self.log_text.setProperty("class", "log-text")
        section_layout.addWidget(self.log_text)
        
        # Add stretch to push content to top (consistent with other sections)
        section_layout.addStretch()
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
    
    def add_log(self, message: str):
        """Add a message to the log with real-time timestamp"""
        # Get timestamp when the message is actually displayed
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        # Auto-scroll to bottom
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        # Force immediate GUI update
        self.log_text.repaint()
        QApplication.processEvents()
    
    def clear_log(self):
        """Clear the log"""
        self.log_text.clear()
