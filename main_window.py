"""
Main Window for Interunit Loan Matcher
Central window class for the GUI application
"""

import os
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QMessageBox
)
from PySide6.QtCore import Qt

from gui_widgets import FileSelectionWidget, ProcessingWidget, ResultsWidget, LogWidget
from matching_thread import MatchingThread
from gui_styles import get_main_stylesheet
from interunit_loan_matcher import ExcelTransactionMatcher


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
        main_layout.setSpacing(5)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Top row: File selection and Match steps side by side
        top_row = QHBoxLayout()
        top_row.setSpacing(5)
        
        # File selection widget (left) - equal width
        self.file_selection = FileSelectionWidget()
        self.file_selection.files_selected.connect(self.on_files_selected)
        self.file_selection.run_match_button.clicked.connect(self.start_matching)
        top_row.addWidget(self.file_selection, 1)  # Equal stretch factor
        
        # Processing widget (right) - equal width
        self.processing_widget = ProcessingWidget()
        top_row.addWidget(self.processing_widget, 1)  # Equal stretch factor
        
        # Get reference to overall progress bar from processing widget
        self.overall_progress = self.processing_widget.overall_progress
        
        main_layout.addLayout(top_row)
        
        # Match Summary section - full width
        self.results_widget = ResultsWidget()
        # Connect the open buttons to the main window's methods
        self.results_widget.open_folder_button.clicked.connect(self.open_output_folder)
        self.results_widget.open_files_button.clicked.connect(self.open_output_files)
        main_layout.addWidget(self.results_widget)
        
        # Process Log section - full width
        self.log_widget = LogWidget()
        main_layout.addWidget(self.log_widget)
        
        central_widget.setLayout(main_layout)
        
        # Apply styling
        self.apply_styling()
        
        # Add initial log message
        self.log_widget.add_log("Application started. Please select Excel files to begin.")
    
    def apply_styling(self):
        """Apply styling to match the reference image exactly"""
        self.setStyleSheet(get_main_stylesheet())
    
    def on_files_selected(self, file1_path: str, file2_path: str):
        """Handle file selection"""
        self.current_file1 = file1_path
        self.current_file2 = file2_path
        # Also set the attributes that open_output_folder expects
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.log_widget.add_log(f"Files selected: {os.path.basename(file1_path)} and {os.path.basename(file2_path)}")
    
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
        self.matching_thread.log_message.connect(self.log_widget.add_log)
        self.matching_thread.start()
    
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
    
    def on_matching_finished(self, matches: list, statistics: dict):
        """Handle matching completion"""
        self.current_matches = matches
        self.results_widget.update_results(statistics)
        
        self.log_widget.add_log(f"Matching completed successfully! Found {statistics['total_matches']} matches.")
        
        # Automatically create output files if matches were found
        if matches and len(matches) > 0:
            self.create_output_files_with_progress(matches)
        else:
            self.log_widget.add_log("No matches found. No output files created.")
            self.processing_widget.set_processing_state(False)
    
    def create_output_files_with_progress(self, matches):
        """Create output files with accurate progress tracking"""
        try:
            # Keep processing state active during file creation
            self.processing_widget.set_processing_state(True)
            
            # PHASE 1: PREPARATION (65-70%)
            self.overall_progress.setValue(67)
            self.log_widget.add_log("📁 Preparing file creation...")
            self.log_widget.add_log(f"   - Processing {len(matches)} matches for output files")
            
            from interunit_loan_matcher import ExcelTransactionMatcher
            matcher = ExcelTransactionMatcher(self.current_file1, self.current_file2)
            
            # PHASE 2: LOAD TRANSACTION DATA (70-85%) - This takes significant time
            self.overall_progress.setValue(72)
            self.log_widget.add_log("📖 Loading first transaction file...")
            self.log_widget.add_log(f"   - Reading: {self.current_file1}")
            self.log_widget.add_log("   - Processing Excel file structure...")
            matcher.metadata1, matcher.transactions1 = matcher.read_complex_excel(self.current_file1)
            self.log_widget.add_log(f"   ✅ Loaded {len(matcher.transactions1)} transactions from File 1")
            
            self.overall_progress.setValue(79)
            self.log_widget.add_log("📖 Loading second transaction file...")
            self.log_widget.add_log(f"   - Reading: {self.current_file2}")
            self.log_widget.add_log("   - Processing Excel file structure...")
            matcher.metadata2, matcher.transactions2 = matcher.read_complex_excel(self.current_file2)
            self.log_widget.add_log(f"   ✅ Loaded {len(matcher.transactions2)} transactions from File 2")
            
            # PHASE 3: CREATE MATCHED FILES (85-100%) - This is the longest part
            self.overall_progress.setValue(85)
            self.log_widget.add_log("📝 Creating matched Excel files...")
            self.log_widget.add_log("   - Generating output files with matched transactions...")
            self.log_widget.add_log("   - This may take 30+ seconds for large files...")
            self.log_widget.add_log("   - Creating Excel workbooks and formatting...")
            
            # This is where most of the time is spent - creating and formatting Excel files
            matcher.create_matched_files(matches, matcher.transactions1, matcher.transactions2)
            self.log_widget.add_log("   - Excel file generation completed!")
            
            # PHASE 4: COMPLETE (100%)
            self.overall_progress.setValue(100)
            self.log_widget.add_log("🎉 Excel files exported successfully!")
            self.log_widget.add_log("   - Matched files saved to the same folder as input files")
            QMessageBox.information(self, "Export Complete", "Matched Excel files have been exported to the same folder as the input files.")
            
        except Exception as e:
            self.log_widget.add_log(f"Export error: {str(e)}")
            QMessageBox.critical(self, "Export Error", f"Failed to export files:\n\n{str(e)}")
        finally:
            # Always reset processing state when done
            self.processing_widget.set_processing_state(False)
    
    def on_matching_error(self, error_message: str):
        """Handle matching errors"""
        self.processing_widget.set_processing_state(False)
        self.log_widget.add_log(f"Error: {error_message}")
        QMessageBox.critical(self, "Matching Error", f"An error occurred during matching:\n\n{error_message}")
    
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
            QMessageBox.information(self, "Export Complete", "Matched Excel files have been exported to the same folder as the input files.")
            
        except Exception as e:
            self.log_widget.add_log(f"Export error: {str(e)}")
            QMessageBox.critical(self, "Export Error", f"Failed to export files:\n\n{str(e)}")
    
    def open_output_folder(self):
        """Open the input file folder in file explorer"""
        if hasattr(self, 'file1_path') and self.file1_path:
            input_dir = Path(self.file1_path).parent
            if input_dir.exists():
                try:
                    os.startfile(str(input_dir))
                except Exception as e:
                    QMessageBox.critical(self, "Error Opening Folder", f"Could not open folder:\n\n{str(e)}")
            else:
                QMessageBox.warning(self, "Folder Not Found", f"Input folder not found: {input_dir}")
        else:
            QMessageBox.information(self, "No Files Selected", "Please select input files first to see the output location.")
    
    def open_output_files(self):
        """Open the output Excel files directly"""
        if hasattr(self, 'file1_path') and self.file1_path and hasattr(self, 'file2_path') and self.file2_path:
            try:
                # Get the directory of the input files
                input_dir1 = Path(self.file1_path).parent
                input_dir2 = Path(self.file2_path).parent
                
                # Construct output file paths
                base_name1 = os.path.splitext(os.path.basename(self.file1_path))[0]
                base_name2 = os.path.splitext(os.path.basename(self.file2_path))[0]
                
                output_file1 = input_dir1 / f"{base_name1}_MATCHED.xlsx"
                output_file2 = input_dir2 / f"{base_name2}_MATCHED.xlsx"
                
                # Check if files exist and open them
                files_opened = 0
                if output_file1.exists():
                    os.startfile(str(output_file1))
                    files_opened += 1
                if output_file2.exists():
                    os.startfile(str(output_file2))
                    files_opened += 1
                
                if files_opened == 0:
                    QMessageBox.warning(self, "No Output Files", "No output files found. Please run the matching process first.")
                elif files_opened == 1:
                    QMessageBox.information(self, "Partial Success", "Only one output file was found and opened.")
                # If files_opened == 2, both files were opened successfully (no message needed)
                
            except Exception as e:
                QMessageBox.critical(self, "Error Opening Files", f"Could not open output files:\n\n{str(e)}")
        else:
            QMessageBox.information(self, "No Files Selected", "Please select input files first.")
    
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
