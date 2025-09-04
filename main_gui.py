"""
Interunit Loan Matcher - GUI Application
Modern PyQt6 interface for automated Excel transaction matching
"""

import sys
import os
import threading
import time
from pathlib import Path
from typing import Optional, Dict, Any

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QLabel, QPushButton, QProgressBar, QTextEdit,
    QFileDialog, QMessageBox, QGroupBox, QFrame, QSplitter,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QStatusBar, QMenuBar, QMenu
)
from PySide6.QtCore import (
    Qt, QThread, Signal, QTimer, QSize, QPropertyAnimation,
    QEasingCurve, QRect
)
from PySide6.QtGui import (
    QFont, QPixmap, QIcon, QPalette, QColor, QDragEnterEvent,
    QDropEvent, QPainter, QPen, QAction
)

# Import existing matching logic
from interunit_loan_matcher import ExcelTransactionMatcher
from config import INPUT_FILE1_PATH, INPUT_FILE2_PATH, OUTPUT_FOLDER


class MatchingThread(QThread):
    """Background thread for running the matching process"""
    
    # Signals for communication with main thread
    progress_updated = Signal(int, str, int)  # step, status, matches_found
    step_completed = Signal(str, int)  # step_name, matches_found
    matching_finished = Signal(list, dict)  # matches, statistics
    error_occurred = Signal(str)  # error_message
    
    def __init__(self, file1_path: str, file2_path: str):
        super().__init__()
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.is_cancelled = False
        
    def run(self):
        """Run the matching process in background thread"""
        try:
            # Initialize matcher
            self.progress_updated.emit(0, "Initializing matcher...", 0)
            matcher = ExcelTransactionMatcher(self.file1_path, self.file2_path)
            
            if self.is_cancelled:
                return
                
            # Process files
            self.progress_updated.emit(10, "Loading and processing files...", 0)
            transactions1, transactions2, blocks1, blocks2, lc_numbers1, lc_numbers2, po_numbers1, po_numbers2, interunit_accounts1, interunit_accounts2, usd_amounts1, usd_amounts2 = matcher.process_files()
            
            if self.is_cancelled:
                return
                
            # Step 1: Narration Matching
            self.progress_updated.emit(15, "Finding narration matches...", 0)
            narration_matches = matcher.narration_matching_logic.find_potential_matches(
                transactions1, transactions2, {}, None
            )
            self.step_completed.emit("Narration Matching", len(narration_matches))
            
            if self.is_cancelled:
                return
                
            # Step 2: LC Matching
            self.progress_updated.emit(30, "Finding LC matches...", 0)
            # Create masks for unmatched records (after Narration matching)
            narration_matched_indices1 = set()
            narration_matched_indices2 = set()
            
            for match in narration_matches:
                narration_matched_indices1.add(match['File1_Index'])
                narration_matched_indices2.add(match['File2_Index'])
            
            # Filter LC numbers to only unmatched records
            lc_numbers1_unmatched = lc_numbers1.copy()
            lc_numbers2_unmatched = lc_numbers2.copy()
            
            for idx in narration_matched_indices1:
                if idx < len(lc_numbers1_unmatched):
                    lc_numbers1_unmatched.iloc[idx] = None
            
            for idx in narration_matched_indices2:
                if idx < len(lc_numbers2_unmatched):
                    lc_numbers2_unmatched.iloc[idx] = None
            
            lc_matches = matcher.lc_matching_logic.find_potential_matches(
                transactions1, transactions2, lc_numbers1_unmatched, lc_numbers2_unmatched,
                {}, None
            )
            self.step_completed.emit("LC Matching", len(lc_matches))
            
            if self.is_cancelled:
                return
                
            # Step 3: PO Matching
            self.progress_updated.emit(45, "Finding PO matches...", 0)
            # Create masks for unmatched records (after Narration and LC matching)
            narration_lc_matched_indices1 = set()
            narration_lc_matched_indices2 = set()
            
            for match in narration_matches + lc_matches:
                narration_lc_matched_indices1.add(match['File1_Index'])
                narration_lc_matched_indices2.add(match['File2_Index'])
            
            # Filter PO numbers to only unmatched records
            po_numbers1_unmatched = po_numbers1.copy()
            po_numbers2_unmatched = po_numbers2.copy()
            
            for idx in narration_lc_matched_indices1:
                if idx < len(po_numbers1_unmatched):
                    po_numbers1_unmatched.iloc[idx] = None
            
            for idx in narration_lc_matched_indices2:
                if idx < len(po_numbers2_unmatched):
                    po_numbers2_unmatched.iloc[idx] = None
            
            po_matches = matcher.po_matching_logic.find_potential_matches(
                transactions1, transactions2, po_numbers1_unmatched, po_numbers2_unmatched,
                {}, None
            )
            self.step_completed.emit("PO Matching", len(po_matches))
            
            if self.is_cancelled:
                return
                
            # Step 4: Interunit Matching
            self.progress_updated.emit(60, "Finding interunit matches...", 0)
            # Create masks for unmatched records (after Narration, LC, and PO matching)
            narration_lc_po_matched_indices1 = set()
            narration_lc_po_matched_indices2 = set()
            
            for match in narration_matches + lc_matches + po_matches:
                narration_lc_po_matched_indices1.add(match['File1_Index'])
                narration_lc_po_matched_indices2.add(match['File2_Index'])
            
            # Filter interunit accounts to only unmatched records
            interunit_accounts1_unmatched = interunit_accounts1.copy()
            interunit_accounts2_unmatched = interunit_accounts2.copy()
            
            for idx in narration_lc_po_matched_indices1:
                if idx < len(interunit_accounts1_unmatched):
                    interunit_accounts1_unmatched.iloc[idx] = None
            
            for idx in narration_lc_po_matched_indices2:
                if idx < len(interunit_accounts2_unmatched):
                    interunit_accounts2_unmatched.iloc[idx] = None
            
            interunit_matches = matcher.interunit_loan_matcher.find_potential_matches(
                transactions1, transactions2, interunit_accounts1_unmatched, interunit_accounts2_unmatched,
                self.file1_path, self.file2_path, {}, None
            )
            self.step_completed.emit("Interunit Matching", len(interunit_matches))
            
            if self.is_cancelled:
                return
                
            # Step 5: USD Matching
            self.progress_updated.emit(75, "Finding USD matches...", 0)
            # Create masks for unmatched records (after Narration, LC, PO, and Interunit matching)
            narration_lc_po_interunit_matched_indices1 = set()
            narration_lc_po_interunit_matched_indices2 = set()
            
            for match in narration_matches + lc_matches + po_matches + interunit_matches:
                narration_lc_po_interunit_matched_indices1.add(match['File1_Index'])
                narration_lc_po_interunit_matched_indices2.add(match['File2_Index'])
            
            # Filter USD amounts to only unmatched records
            usd_amounts1_unmatched = usd_amounts1.copy()
            usd_amounts2_unmatched = usd_amounts2.copy()
            
            for idx in narration_lc_po_interunit_matched_indices1:
                if idx < len(usd_amounts1_unmatched):
                    usd_amounts1_unmatched.iloc[idx] = None
            
            for idx in narration_lc_po_interunit_matched_indices2:
                if idx < len(usd_amounts2_unmatched):
                    usd_amounts2_unmatched.iloc[idx] = None
            
            usd_matches = matcher.usd_matching_logic.find_potential_matches(
                transactions1, transactions2, usd_amounts1_unmatched, usd_amounts2_unmatched,
                {}, None
            )
            self.step_completed.emit("USD Matching", len(usd_matches))
            
            if self.is_cancelled:
                return
                
            # Step 6: Aggregated PO Matching
            self.progress_updated.emit(90, "Finding aggregated PO matches...", 0)
            # Create masks for unmatched records (after all previous matching)
            narration_lc_po_interunit_usd_matched_indices1 = set()
            narration_lc_po_interunit_usd_matched_indices2 = set()
            
            for match in narration_matches + lc_matches + po_matches + interunit_matches + usd_matches:
                narration_lc_po_interunit_usd_matched_indices1.add(match['File1_Index'])
                narration_lc_po_interunit_usd_matched_indices2.add(match['File2_Index'])
            
            # Filter PO numbers to only unmatched records
            po_numbers1_unmatched_for_aggregated = po_numbers1.copy()
            po_numbers2_unmatched_for_aggregated = po_numbers2.copy()
            
            for idx in narration_lc_po_interunit_usd_matched_indices1:
                if idx < len(po_numbers1_unmatched_for_aggregated):
                    po_numbers1_unmatched_for_aggregated.iloc[idx] = None
            
            for idx in narration_lc_po_interunit_usd_matched_indices2:
                if idx < len(po_numbers2_unmatched_for_aggregated):
                    po_numbers2_unmatched_for_aggregated.iloc[idx] = None
            
            aggregated_po_matches = matcher.aggregated_po_matching_logic.find_potential_matches(
                transactions1, transactions2, po_numbers1_unmatched_for_aggregated, po_numbers2_unmatched_for_aggregated,
                {}, None
            )
            self.step_completed.emit("Aggregated PO Matching", len(aggregated_po_matches))
            
            if self.is_cancelled:
                return
                
            # Combine all matches
            all_matches = narration_matches + lc_matches + po_matches + interunit_matches + usd_matches + aggregated_po_matches
            
            # Assign sequential Match IDs
            match_counter = 1
            for match in all_matches:
                match_id = f"M{match_counter:03d}"
                match['match_id'] = match_id
                match_counter += 1
            
            # Sort matches by the newly assigned sequential Match IDs
            all_matches.sort(key=lambda x: x['match_id'])
            
            # Create statistics
            stats = {
                'total_matches': len(all_matches),
                'narration_matches': len(narration_matches),
                'lc_matches': len(lc_matches),
                'po_matches': len(po_matches),
                'interunit_matches': len(interunit_matches),
                'usd_matches': len(usd_matches),
                'aggregated_po_matches': len(aggregated_po_matches)
            }
            
            self.progress_updated.emit(100, "Matching completed successfully!", stats['total_matches'])
            self.matching_finished.emit(all_matches, stats)
            
        except Exception as e:
            self.error_occurred.emit(str(e))
    
    def cancel(self):
        """Cancel the matching process"""
        self.is_cancelled = True


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
        
        # Title
        title = QLabel("Select Excel Files for Matching")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # File selection grid
        grid = QGridLayout()
        
        # File 1
        self.file1_label = QLabel("File 1 (Pole Book):")
        self.file1_display = QLabel("No file selected")
        self.file1_display.setStyleSheet("""
            QLabel {
                border: 2px dashed #ccc;
                padding: 20px;
                background-color: #f9f9f9;
                border-radius: 5px;
            }
        """)
        self.file1_button = QPushButton("Browse...")
        self.file1_button.clicked.connect(lambda: self.select_file(1))
        
        grid.addWidget(self.file1_label, 0, 0)
        grid.addWidget(self.file1_display, 1, 0)
        grid.addWidget(self.file1_button, 1, 1)
        
        # File 2
        self.file2_label = QLabel("File 2 (Steel Book):")
        self.file2_display = QLabel("No file selected")
        self.file2_display.setStyleSheet("""
            QLabel {
                border: 2px dashed #ccc;
                padding: 20px;
                background-color: #f9f9f9;
                border-radius: 5px;
            }
        """)
        self.file2_button = QPushButton("Browse...")
        self.file2_button.clicked.connect(lambda: self.select_file(2))
        
        grid.addWidget(self.file2_label, 2, 0)
        grid.addWidget(self.file2_display, 3, 0)
        grid.addWidget(self.file2_button, 3, 1)
        
        layout.addLayout(grid)
        
        # Validation status
        self.validation_label = QLabel("")
        self.validation_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.validation_label)
        
        self.setLayout(layout)
        self.setAcceptDrops(True)
        
    def select_file(self, file_num: int):
        """Open file dialog to select Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            f"Select File {file_num}",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if file_path:
            self.set_file(file_num, file_path)
            
    def set_file(self, file_num: int, file_path: str):
        """Set the selected file path"""
        if file_num == 1:
            self.file1_path = file_path
            self.file1_display.setText(os.path.basename(file_path))
            self.file1_display.setStyleSheet("""
                QLabel {
                    border: 2px solid #4CAF50;
                    padding: 20px;
                    background-color: #e8f5e8;
                    border-radius: 5px;
                }
            """)
        else:
            self.file2_path = file_path
            self.file2_display.setText(os.path.basename(file_path))
            self.file2_display.setStyleSheet("""
                QLabel {
                    border: 2px solid #4CAF50;
                    padding: 20px;
                    background-color: #e8f5e8;
                    border-radius: 5px;
                }
            """)
        
        self.validate_files()
        
    def validate_files(self):
        """Validate that both files are selected and are valid Excel files"""
        if self.file1_path and self.file2_path:
            # Check if files exist and are Excel files
            if (os.path.exists(self.file1_path) and os.path.exists(self.file2_path) and
                (self.file1_path.endswith('.xlsx') or self.file1_path.endswith('.xls')) and
                (self.file2_path.endswith('.xlsx') or self.file2_path.endswith('.xls'))):
                
                self.validation_label.setText("✓ Files validated successfully")
                self.validation_label.setStyleSheet("color: green; font-weight: bold;")
                self.files_selected.emit(self.file1_path, self.file2_path)
            else:
                self.validation_label.setText("✗ Invalid files selected")
                self.validation_label.setStyleSheet("color: red; font-weight: bold;")
        else:
            self.validation_label.setText("Please select both files")
            self.validation_label.setStyleSheet("color: orange; font-weight: bold;")
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter event"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        """Handle drop event"""
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if len(files) >= 2:
            self.set_file(1, files[0])
            self.set_file(2, files[1])
        event.acceptProposedAction()


class ProcessingWidget(QWidget):
    """Widget for displaying processing progress and status"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Processing Status")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # Overall progress
        self.overall_progress = QProgressBar()
        self.overall_progress.setRange(0, 100)
        self.overall_progress.setValue(0)
        layout.addWidget(QLabel("Overall Progress:"))
        layout.addWidget(self.overall_progress)
        
        # Current status
        self.status_label = QLabel("Ready to start matching...")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        
        # Step progress
        self.steps_group = QGroupBox("Matching Steps")
        steps_layout = QVBoxLayout()
        
        self.step_labels = {}
        self.step_progresses = {}
        
        steps = [
            "Narration Matching",
            "LC Matching", 
            "PO Matching",
            "Interunit Matching",
            "USD Matching",
            "Aggregated PO Matching"
        ]
        
        for step in steps:
            step_layout = QHBoxLayout()
            
            # Step name
            step_label = QLabel(f"{step}:")
            step_label.setMinimumWidth(150)
            step_layout.addWidget(step_label)
            
            # Step progress bar
            step_progress = QProgressBar()
            step_progress.setRange(0, 100)
            step_progress.setValue(0)
            step_progress.setMaximumHeight(20)
            step_layout.addWidget(step_progress)
            
            # Step status
            step_status = QLabel("Pending")
            step_status.setMinimumWidth(80)
            step_layout.addWidget(step_status)
            
            steps_layout.addLayout(step_layout)
            
            self.step_labels[step] = step_label
            self.step_progresses[step] = (step_progress, step_status)
        
        self.steps_group.setLayout(steps_layout)
        layout.addWidget(self.steps_group)
        
        # Control buttons
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("Start Matching")
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setEnabled(False)
        self.cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def update_progress(self, step: int, status: str, matches_found: int):
        """Update progress display"""
        self.overall_progress.setValue(step)
        self.status_label.setText(f"{status} ({matches_found} matches found)")
        
    def complete_step(self, step_name: str, matches_found: int):
        """Mark a step as completed"""
        if step_name in self.step_progresses:
            progress_bar, status_label = self.step_progresses[step_name]
            progress_bar.setValue(100)
            status_label.setText(f"✓ {matches_found}")
            status_label.setStyleSheet("color: green; font-weight: bold;")
            
            # Also update the step label to show completion
            if step_name in self.step_labels:
                self.step_labels[step_name].setStyleSheet("color: green; font-weight: bold;")
    
    def reset_progress(self):
        """Reset all progress indicators"""
        self.overall_progress.setValue(0)
        self.status_label.setText("Ready to start matching...")
        
        for step_name, (progress_bar, status_label) in self.step_progresses.items():
            progress_bar.setValue(0)
            status_label.setText("Pending")
            status_label.setStyleSheet("")
            
            # Reset step label styling
            if step_name in self.step_labels:
                self.step_labels[step_name].setStyleSheet("")
    
    def set_processing_state(self, is_processing: bool):
        """Enable/disable processing controls"""
        self.start_button.setEnabled(not is_processing)
        self.cancel_button.setEnabled(is_processing)


class ResultsWidget(QWidget):
    """Widget for displaying matching results and statistics"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Matching Results")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # Results summary
        self.summary_group = QGroupBox("Match Summary")
        summary_layout = QGridLayout()
        
        # Create summary labels
        self.total_matches_label = QLabel("Total Matches: 0")
        self.narration_matches_label = QLabel("Narration: 0")
        self.lc_matches_label = QLabel("LC: 0")
        self.po_matches_label = QLabel("PO: 0")
        self.interunit_matches_label = QLabel("Interunit: 0")
        self.usd_matches_label = QLabel("USD: 0")
        self.aggregated_po_matches_label = QLabel("Aggregated PO: 0")
        
        # Style the labels
        for label in [self.total_matches_label, self.narration_matches_label, 
                     self.lc_matches_label, self.po_matches_label,
                     self.interunit_matches_label, self.usd_matches_label,
                     self.aggregated_po_matches_label]:
            label.setFont(QFont("Arial", 10))
            label.setStyleSheet("padding: 5px;")
        
        # Add to grid
        summary_layout.addWidget(self.total_matches_label, 0, 0)
        summary_layout.addWidget(self.narration_matches_label, 1, 0)
        summary_layout.addWidget(self.lc_matches_label, 1, 1)
        summary_layout.addWidget(self.po_matches_label, 2, 0)
        summary_layout.addWidget(self.interunit_matches_label, 2, 1)
        summary_layout.addWidget(self.usd_matches_label, 3, 0)
        summary_layout.addWidget(self.aggregated_po_matches_label, 3, 1)
        
        self.summary_group.setLayout(summary_layout)
        layout.addWidget(self.summary_group)
        
        # Action buttons
        button_layout = QHBoxLayout()
        
        self.export_button = QPushButton("Export Excel Files")
        self.export_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        
        self.open_folder_button = QPushButton("Open Output Folder")
        self.open_folder_button.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        """)
        
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.open_folder_button)
        layout.addLayout(button_layout)
        
        # Initially disable export button
        self.export_button.setEnabled(False)
        
        self.setLayout(layout)
    
    def update_results(self, statistics: Dict[str, Any]):
        """Update results display with statistics"""
        self.total_matches_label.setText(f"Total Matches: {statistics['total_matches']}")
        self.narration_matches_label.setText(f"Narration: {statistics['narration_matches']}")
        self.lc_matches_label.setText(f"LC: {statistics['lc_matches']}")
        self.po_matches_label.setText(f"PO: {statistics['po_matches']}")
        self.interunit_matches_label.setText(f"Interunit: {statistics['interunit_matches']}")
        self.usd_matches_label.setText(f"USD: {statistics['usd_matches']}")
        self.aggregated_po_matches_label.setText(f"Aggregated PO: {statistics['aggregated_po_matches']}")
        
        # Enable export button if matches were found
        self.export_button.setEnabled(statistics['total_matches'] > 0)
    
    def reset_results(self):
        """Reset results display"""
        self.update_results({
            'total_matches': 0,
            'narration_matches': 0,
            'lc_matches': 0,
            'po_matches': 0,
            'interunit_matches': 0,
            'usd_matches': 0,
            'aggregated_po_matches': 0
        })


class LogWidget(QWidget):
    """Widget for displaying processing logs"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Processing Log")
        title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # Log text area
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 5px;
                font-family: 'Courier New', monospace;
                font-size: 9pt;
            }
        """)
        layout.addWidget(self.log_text)
        
        # Clear button
        clear_button = QPushButton("Clear Log")
        clear_button.clicked.connect(self.clear_log)
        layout.addWidget(clear_button)
        
        self.setLayout(layout)
    
    def add_log(self, message: str):
        """Add a message to the log"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        # Auto-scroll to bottom
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def clear_log(self):
        """Clear the log"""
        self.log_text.clear()


class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.matching_thread = None
        self.current_file1 = ""
        self.current_file2 = ""
        self.current_matches = []
        self.init_ui()
        
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Interunit Loan Matcher - GUI")
        self.setGeometry(100, 100, 1000, 700)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create main layout
        main_layout = QVBoxLayout()
        
        # Create file selection widget
        self.file_selection = FileSelectionWidget()
        self.file_selection.files_selected.connect(self.on_files_selected)
        main_layout.addWidget(self.file_selection)
        
        # Create splitter for processing and results
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Processing widget
        self.processing_widget = ProcessingWidget()
        self.processing_widget.start_button.clicked.connect(self.start_matching)
        self.processing_widget.cancel_button.clicked.connect(self.cancel_matching)
        splitter.addWidget(self.processing_widget)
        
        # Results widget
        self.results_widget = ResultsWidget()
        self.results_widget.export_button.clicked.connect(self.export_files)
        self.results_widget.open_folder_button.clicked.connect(self.open_output_folder)
        splitter.addWidget(self.results_widget)
        
        # Set splitter proportions
        splitter.setSizes([400, 300])
        main_layout.addWidget(splitter)
        
        # Create log widget
        self.log_widget = LogWidget()
        main_layout.addWidget(self.log_widget)
        
        central_widget.setLayout(main_layout)
        
        # Create menu bar
        self.create_menu_bar()
        
        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")
        
        # Apply styling
        self.apply_styling()
        
        # Add initial log message
        self.log_widget.add_log("Application started. Please select Excel files to begin.")
    
    def create_menu_bar(self):
        """Create the menu bar"""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu('File')
        
        open_action = QAction('Open Files...', self)
        open_action.setShortcut('Ctrl+O')
        open_action.triggered.connect(self.open_files_dialog)
        file_menu.addAction(open_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction('Exit', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Help menu
        help_menu = menubar.addMenu('Help')
        
        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def apply_styling(self):
        """Apply modern styling to the application"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QPushButton {
                background-color: #e0e0e0;
                border: 1px solid #b0b0b0;
                border-radius: 3px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #d0d0d0;
            }
            QPushButton:pressed {
                background-color: #c0c0c0;
            }
            QProgressBar {
                border: 2px solid #b0b0b0;
                border-radius: 5px;
                text-align: center;
                background-color: #f0f0f0;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 3px;
            }
        """)
    
    def on_files_selected(self, file1_path: str, file2_path: str):
        """Handle file selection"""
        self.current_file1 = file1_path
        self.current_file2 = file2_path
        self.log_widget.add_log(f"Files selected: {os.path.basename(file1_path)} and {os.path.basename(file2_path)}")
        self.status_bar.showMessage(f"Files ready: {os.path.basename(file1_path)} & {os.path.basename(file2_path)}")
    
    def open_files_dialog(self):
        """Open file selection dialog"""
        # This would trigger the file selection widget's browse buttons
        pass
    
    def start_matching(self):
        """Start the matching process"""
        if not self.current_file1 or not self.current_file2:
            QMessageBox.warning(self, "No Files", "Please select both Excel files first.")
            return
        
        self.log_widget.add_log("Starting matching process...")
        self.processing_widget.set_processing_state(True)
        self.processing_widget.reset_progress()
        self.results_widget.reset_results()
        
        # Create and start matching thread
        self.matching_thread = MatchingThread(self.current_file1, self.current_file2)
        self.matching_thread.progress_updated.connect(self.processing_widget.update_progress)
        self.matching_thread.step_completed.connect(self.processing_widget.complete_step)
        self.matching_thread.matching_finished.connect(self.on_matching_finished)
        self.matching_thread.error_occurred.connect(self.on_matching_error)
        self.matching_thread.start()
        
        self.status_bar.showMessage("Processing...")
    
    def cancel_matching(self):
        """Cancel the matching process"""
        if self.matching_thread and self.matching_thread.isRunning():
            self.matching_thread.cancel()
            self.matching_thread.wait()
            self.log_widget.add_log("Matching process cancelled by user.")
            self.processing_widget.set_processing_state(False)
            self.status_bar.showMessage("Cancelled")
    
    def on_matching_finished(self, matches: list, statistics: dict):
        """Handle matching completion"""
        self.current_matches = matches
        self.results_widget.update_results(statistics)
        self.processing_widget.set_processing_state(False)
        
        self.log_widget.add_log(f"Matching completed successfully! Found {statistics['total_matches']} matches.")
        self.status_bar.showMessage(f"Completed - {statistics['total_matches']} matches found")
    
    def on_matching_error(self, error_message: str):
        """Handle matching errors"""
        self.processing_widget.set_processing_state(False)
        self.log_widget.add_log(f"Error: {error_message}")
        QMessageBox.critical(self, "Matching Error", f"An error occurred during matching:\n\n{error_message}")
        self.status_bar.showMessage("Error occurred")
    
    def export_files(self):
        """Export matched files"""
        if not self.current_matches:
            QMessageBox.warning(self, "No Matches", "No matches to export. Please run matching first.")
            return
        
        try:
            # Create matcher instance and generate output files
            matcher = ExcelTransactionMatcher(self.current_file1, self.current_file2)
            matcher.transactions1, matcher.transactions2, _, _, _, _, _, _, _, _, _, _ = matcher.process_files()
            matcher.create_matched_files(self.current_matches, matcher.transactions1, matcher.transactions2)
            
            self.log_widget.add_log("Excel files exported successfully!")
            QMessageBox.information(self, "Export Complete", "Matched Excel files have been exported to the Output folder.")
            
        except Exception as e:
            self.log_widget.add_log(f"Export error: {str(e)}")
            QMessageBox.critical(self, "Export Error", f"Failed to export files:\n\n{str(e)}")
    
    def open_output_folder(self):
        """Open the output folder in file explorer"""
        output_path = Path(OUTPUT_FOLDER)
        if output_path.exists():
            os.startfile(str(output_path))
        else:
            QMessageBox.warning(self, "Folder Not Found", f"Output folder not found: {OUTPUT_FOLDER}")
    
    def show_about(self):
        """Show about dialog"""
        QMessageBox.about(self, "About Interunit Loan Matcher", 
                         "Interunit Loan Matcher GUI v1.0\n\n"
                         "Automated Excel transaction matching tool\n"
                         "Built with PyQt6")
    
    def closeEvent(self, event):
        """Handle application close event"""
        if self.matching_thread and self.matching_thread.isRunning():
            reply = QMessageBox.question(self, "Exit Confirmation",
                                       "Matching is in progress. Are you sure you want to exit?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.matching_thread.cancel()
                self.matching_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    app.setApplicationName("Interunit Loan Matcher")
    app.setApplicationVersion("1.0")
    
    # Set application style
    app.setStyle('Fusion')
    
    # Create and show main window
    window = MainWindow()
    window.show()
    
    # Start event loop
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
