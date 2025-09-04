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
        layout.setSpacing(10)
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setProperty("class", "section-container")
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(20, 20, 20, 20)  # Add padding inside container
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Select Interunit Loan Ledgers:")
        title.setProperty("class", "heading")
        title.setStyleSheet("margin-bottom: 10px;")
        section_layout.addWidget(title)
        
        # Browse button
        self.browse_button = QPushButton("Browse Ledgers")
        self.browse_button.setProperty("class", "browse-button")
        self.browse_button.setMinimumSize(100, 24)
        self.browse_button.clicked.connect(self.select_both_files)
        section_layout.addWidget(self.browse_button)
        
        # Selected files section
        files_label = QLabel("Selected Ledgers:")
        files_label.setProperty("class", "heading")
        files_label.setStyleSheet("margin-top: 15px; margin-bottom: 5px;")
        section_layout.addWidget(files_label)
        
        # File list container
        self.files_container = QWidget()
        self.files_container.setProperty("class", "files-list")
        files_layout = QVBoxLayout()
        files_layout.setContentsMargins(0, 0, 0, 0)
        files_layout.setSpacing(2)
        self.files_container.setLayout(files_layout)
        section_layout.addWidget(self.files_container)
        
        # Clear button
        self.clear_files_button = QPushButton("Clear Ledgers")
        self.clear_files_button.setProperty("class", "clear-button")
        self.clear_files_button.setMinimumSize(90, 24)
        self.clear_files_button.clicked.connect(self.clear_files)
        section_layout.addWidget(self.clear_files_button)
        
        # Run Match button
        self.run_match_button = QPushButton("Run Match")
        self.run_match_button.setProperty("class", "run-button")
        self.run_match_button.setMinimumSize(90, 28)
        self.run_match_button.clicked.connect(self.run_matching)
        self.run_match_button.setEnabled(False)
        section_layout.addWidget(self.run_match_button)
        
        # Add stretch to push everything to top
        section_layout.addStretch()
        
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
        self.validation_label.setText("Please select both files")
        self.validation_label.setProperty("class", "status-warning")
            
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
        
        # Add current files with actual filenames
        if self.file1_path:
            file1_name = os.path.basename(self.file1_path)
            file1_item = QLabel(file1_name)
            file1_item.setProperty("class", "file-item")
            layout.addWidget(file1_item)
        
        if self.file2_path:
            file2_name = os.path.basename(self.file2_path)
            file2_item = QLabel(file2_name)
            file2_item.setProperty("class", "file-item")
            layout.addWidget(file2_item)
        
        # Enable/disable Run Match button
        self.run_match_button.setEnabled(bool(self.file1_path and self.file2_path))
    
    def on_drop_zone_clicked(self, event):
        """Handle click on drop zone"""
        self.select_both_files()
    
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
                
                self.validation_label.setText("✓ Files validated successfully")
                self.validation_label.setProperty("class", "status-success")
                self.files_selected.emit(self.file1_path, self.file2_path)
            else:
                self.validation_label.setText("✗ Invalid files selected")
                self.validation_label.setProperty("class", "status-error")
        else:
            self.validation_label.setText("Please select both files")
            self.validation_label.setProperty("class", "status-warning")
    
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
        layout.setSpacing(10)
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setProperty("class", "section-container")
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(20, 20, 20, 20)  # Add padding inside container
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Match Steps")
        title.setProperty("class", "heading")
        title.setStyleSheet("margin-bottom: 10px;")
        section_layout.addWidget(title)
        
        # Step progress
        self.step_labels = {}
        self.step_progresses = {}
        
        steps = [
            "Narration Match",
            "LC Match"
        ]
        
        for step in steps:
            step_layout = QHBoxLayout()
            step_layout.setSpacing(10)
            
            # Step name
            step_label = QLabel(f"{step}")
            step_label.setMinimumWidth(120)
            step_label.setFont(QFont("Arial", 10))
            step_layout.addWidget(step_label)
            
            # Step progress bar
            step_progress = QProgressBar()
            step_progress.setRange(0, 100)
            step_progress.setValue(0)
            step_progress.setMaximumHeight(20)
            step_progress.setProperty("class", "step-progress")
            step_layout.addWidget(step_progress)
            
            # Step status
            step_status = QLabel("0%")
            step_status.setMinimumWidth(40)
            step_status.setFont(QFont("Arial", 10))
            step_layout.addWidget(step_status)
            
            section_layout.addLayout(step_layout)
            
            self.step_labels[step] = step_label
            self.step_progresses[step] = (step_progress, step_status)
        
        # Add stretch to push everything to top
        section_layout.addStretch()
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
    
    def update_progress(self, step: int, status: str, matches_found: int):
        """Update progress display"""
        # This method is called from ProcessingWidget but we need to update overall progress
        pass
        
    def complete_step(self, step_name: str, matches_found: int):
        """Mark a step as completed"""
        # Map full step names to short names
        step_mapping = {
            "Narration Matching": "Narration Match",
            "LC Matching": "LC Match"
        }
        
        short_name = step_mapping.get(step_name, step_name)
        
        if short_name in self.step_progresses:
            progress_bar, status_label = self.step_progresses[short_name]
            progress_bar.setValue(100)
            status_label.setText("100%")
            status_label.setProperty("class", "status-success")
            
            # Also update the step label to show completion
            if short_name in self.step_labels:
                self.step_labels[short_name].setProperty("class", "step-completed")
    
    def reset_progress(self):
        """Reset all progress indicators"""
        for step_name, (progress_bar, status_label) in self.step_progresses.items():
            progress_bar.setValue(0)
            status_label.setText("0%")
            status_label.setProperty("class", "step-pending")
            
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
        layout.setSpacing(10)
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setProperty("class", "section-container")
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(20, 20, 20, 20)  # Add padding inside container
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Match Summary")
        title.setProperty("class", "heading")
        title.setStyleSheet("margin-bottom: 10px;")
        section_layout.addWidget(title)
        
        # Results summary - matching the image layout
        summary_layout = QHBoxLayout()
        summary_layout.setSpacing(15)
        
        # LC Matches box
        self.lc_matches_label = QLabel("LC Matches: 10")
        self.lc_matches_label.setProperty("class", "match-box")
        self.lc_matches_label.setMinimumSize(100, 40)
        summary_layout.addWidget(self.lc_matches_label)
        
        # PO Matches box
        self.po_matches_label = QLabel("PO Matches: 10")
        self.po_matches_label.setProperty("class", "match-box")
        self.po_matches_label.setMinimumSize(100, 40)
        summary_layout.addWidget(self.po_matches_label)
        
        section_layout.addLayout(summary_layout)
        
        # Add stretch to push everything to top
        section_layout.addStretch()
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
    
    def update_results(self, statistics: Dict[str, Any]):
        """Update results display with statistics"""
        self.lc_matches_label.setText(f"LC Matches: {statistics.get('lc_matches', 0)}")
        self.po_matches_label.setText(f"PO Matches: {statistics.get('po_matches', 0)}")
    
    def reset_results(self):
        """Reset results display"""
        self.lc_matches_label.setText("LC Matches: 0")
        self.po_matches_label.setText("PO Matches: 0")


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
        section_layout.setContentsMargins(20, 20, 20, 20)  # Add padding inside container
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Process Log:")
        title.setProperty("class", "heading")
        title.setStyleSheet("margin-bottom: 5px;")
        section_layout.addWidget(title)
        
        # Log text area
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setProperty("class", "log-text")
        section_layout.addWidget(self.log_text)
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
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
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Top row: File selection and Match steps side by side
        top_row = QHBoxLayout()
        top_row.setSpacing(20)
        
        # File selection widget (left) - equal width
        self.file_selection = FileSelectionWidget()
        self.file_selection.files_selected.connect(self.on_files_selected)
        self.file_selection.run_match_button.clicked.connect(self.start_matching)
        top_row.addWidget(self.file_selection, 1)  # Equal stretch factor
        
        # Processing widget (right) - equal width
        self.processing_widget = ProcessingWidget()
        top_row.addWidget(self.processing_widget, 1)  # Equal stretch factor
        
        main_layout.addLayout(top_row)
        
        # Overall progress bar
        overall_label = QLabel("Overall Progress:")
        overall_label.setProperty("class", "heading")
        main_layout.addWidget(overall_label)
        
        self.overall_progress = QProgressBar()
        self.overall_progress.setRange(0, 100)
        self.overall_progress.setValue(0)
        self.overall_progress.setProperty("class", "overall-progress")
        main_layout.addWidget(self.overall_progress)
        
        # Bottom row: Match Summary and Process Log side by side
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(20)
        
        # Results widget (left) - equal width
        self.results_widget = ResultsWidget()
        bottom_row.addWidget(self.results_widget, 1)  # Equal stretch factor
        
        # Log widget (right) - equal width
        self.log_widget = LogWidget()
        bottom_row.addWidget(self.log_widget, 1)  # Equal stretch factor
        
        main_layout.addLayout(bottom_row)
        
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
        """Apply styling to match the reference image exactly"""
        self.setStyleSheet("""
            /* Main Window */
            QMainWindow {
                background-color: #faf9f8;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            
            /* Universal Heading - Segoe UI Semibold */
            QLabel[class="heading"] {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                color: #2c3e50;
                font-size: 20px;
            }
            
            /* Section Containers with Windows Dialog Style */
            QWidget[class="section-container"] {
                background-color: white;
                border: 1px solid #e1e5e9;
                border-radius: 8px;
                margin: 8px;
            }
            
            /* Primary Button - Windows Dialog Style */
            QPushButton[class="browse-button"] {
                background-color: #0078D4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: 600;
                font-size: 12px;
                min-height: 32px;
                min-width: 80px;
            }
            QPushButton[class="browse-button"]:hover {
                background-color: #106EBE;
            }
            
            /* Secondary Button - Windows Dialog Style */
            QPushButton[class="clear-button"] {
                background-color: white;
                color: #323130;
                border: 1px solid #8A8886;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: 400;
                font-size: 12px;
                min-height: 32px;
                min-width: 80px;
            }
            QPushButton[class="clear-button"]:hover {
                background-color: #F3F2F1;
                border-color: #605E5C;
            }
            
            /* Primary Action Button - Windows Dialog Style */
            QPushButton[class="run-button"] {
                background-color: #0078D4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 10px 20px;
                font-weight: 600;
                font-size: 12px;
                min-height: 36px;
                min-width: 90px;
            }
            QPushButton[class="run-button"]:hover {
                background-color: #106EBE;
            }
            QPushButton[class="run-button"]:disabled {
                background-color: #A19F9D;
                color: #C8C6C4;
            }
            
            /* Files List */
            QWidget[class="files-list"] {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 8px;
                margin: 5px 0;
            }
            
            /* File Item */
            QLabel[class="file-item"] {
                padding: 4px 8px;
                margin: 1px 0;
                background-color: transparent;
                color: #495057;
                font-style: italic;
                font-size: 11px;
            }
            
            /* Step Progress Bars - Orange */
            QProgressBar[class="step-progress"] {
                border: 1px solid #dee2e6;
                border-radius: 4px;
                text-align: center;
                background-color: #f8f9fa;
                height: 20px;
                font-weight: 500;
                color: #495057;
                font-size: 10px;
            }
            QProgressBar[class="step-progress"]::chunk {
                background-color: #fd7e14;
                border-radius: 3px;
            }
            
            /* Overall Progress Bar - Orange */
            QProgressBar[class="overall-progress"] {
                border: 1px solid #dee2e6;
                border-radius: 4px;
                text-align: center;
                background-color: #f8f9fa;
                height: 25px;
                font-weight: 600;
                color: #495057;
                font-size: 12px;
            }
            QProgressBar[class="overall-progress"]::chunk {
                background-color: #fd7e14;
                border-radius: 3px;
            }
            
            /* Match Boxes */
            QLabel[class="match-box"] {
                border: 1px solid #007bff;
                border-radius: 4px;
                background-color: #f8f9fa;
                padding: 8px 12px;
                color: #495057;
                font-weight: 500;
                font-size: 11px;
                text-align: center;
            }
            
            /* Log Text */
            QTextEdit[class="log-text"] {
                border: 1px solid #dee2e6;
                border-radius: 4px;
                background-color: white;
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 10px;
                padding: 8px;
            }
            
            /* Status Labels */
            QLabel[class="status-success"] {
                color: #28a745;
                font-weight: 600;
            }
            QLabel[class="status-error"] {
                color: #dc3545;
                font-weight: 600;
            }
            QLabel[class="status-warning"] {
                color: #ffc107;
                font-weight: 600;
            }
            
            /* Step Labels */
            QLabel[class="step-completed"] {
                color: #28a745;
                font-weight: 600;
            }
            QLabel[class="step-pending"] {
                color: #6c757d;
            }
            
            /* Menu Bar */
            QMenuBar {
                background-color: #343a40;
                color: white;
                border: none;
                padding: 4px;
            }
            QMenuBar::item {
                background-color: transparent;
                padding: 8px 12px;
                border-radius: 4px;
            }
            QMenuBar::item:selected {
                background-color: #495057;
            }
            QMenu {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 4px;
            }
            QMenu::item {
                padding: 8px 16px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #e9ecef;
            }
            
            /* Status Bar */
            QStatusBar {
                background-color: #f8f9fa;
                border-top: 1px solid #dee2e6;
                color: #495057;
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
        self.overall_progress.setValue(0)
        
        # Create and start matching thread
        self.matching_thread = MatchingThread(self.current_file1, self.current_file2)
        self.matching_thread.progress_updated.connect(self.update_overall_progress)
        self.matching_thread.step_completed.connect(self.processing_widget.complete_step)
        self.matching_thread.matching_finished.connect(self.on_matching_finished)
        self.matching_thread.error_occurred.connect(self.on_matching_error)
        self.matching_thread.start()
        
        self.status_bar.showMessage("Processing...")
    
    def update_overall_progress(self, step: int, status: str, matches_found: int):
        """Update overall progress bar"""
        self.overall_progress.setValue(step)
        self.log_widget.add_log(f"{status} ({matches_found} matches found)")
    
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
            self.log_widget.add_log("Starting export process...")
            
            # Create matcher instance and load only the transaction data (without matching)
            matcher = ExcelTransactionMatcher(self.current_file1, self.current_file2)
            
            # Load transaction data without running matching logic
            self.log_widget.add_log("Loading transaction data...")
            matcher.metadata1, matcher.transactions1 = matcher.read_complex_excel(self.current_file1)
            matcher.metadata2, matcher.transactions2 = matcher.read_complex_excel(self.current_file2)
            
            # Create matched files using existing matches
            self.log_widget.add_log("Creating matched Excel files...")
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
