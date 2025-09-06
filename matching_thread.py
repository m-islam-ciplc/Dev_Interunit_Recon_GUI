"""
Matching Thread for Interunit Loan Matcher
Background processing logic for GUI application
"""

import sys
import io
import time
from contextlib import redirect_stdout
from PySide6.QtCore import QThread, Signal
from interunit_loan_matcher import ExcelTransactionMatcher


class MatchingThread(QThread):
    """Background thread for running the matching process"""
    
    # Signals for communication with main thread
    progress_updated = Signal(int, str, int)  # step, status, matches_found
    step_completed = Signal(str, int)  # step_name, matches_found
    matching_finished = Signal(list, dict)  # matches, statistics
    error_occurred = Signal(str)  # error_message
    log_message = Signal(str)  # detailed log message
    
    def __init__(self, file1_path: str, file2_path: str):
        super().__init__()
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.is_cancelled = False
        self._original_stdout = None
        self._captured_output = None
    
    def _capture_print_output(self):
        """Start capturing print output"""
        self._original_stdout = sys.stdout
        self._captured_output = io.StringIO()
        sys.stdout = self._captured_output
    
    def _release_print_output(self):
        """Stop capturing and emit captured output"""
        if self._original_stdout and self._captured_output:
            sys.stdout = self._original_stdout
            captured_text = self._captured_output.getvalue()
            if captured_text.strip():
                # Split by lines and emit each line as a log message
                for line in captured_text.strip().split('\n'):
                    if line.strip():
                        self.log_message.emit(line.strip())
            self._captured_output.close()
            self._captured_output = None
            self._original_stdout = None
    
    def _log_with_delay(self, message: str, delay: float = 0.1):
        """Emit log message with small delay to ensure real-time display"""
        self.log_message.emit(message)
        time.sleep(delay)  # Small delay to ensure GUI updates
        
    def run(self):
        """Run the matching process in background thread with accurate progress tracking"""
        try:
            # Start capturing print output
            self._capture_print_output()
            
            # PHASE 1: INITIALIZATION (0-5%)
            self.progress_updated.emit(2, "Initializing matcher...", 0)
            self.log_message.emit("🚀 Starting Interunit Loan Matcher...")
            self.log_message.emit(f"📁 File 1: {self.file1_path}")
            self.log_message.emit(f"📁 File 2: {self.file2_path}")
            
            matcher = ExcelTransactionMatcher(self.file1_path, self.file2_path)
            
            # Release any captured output so far
            self._release_print_output()
            
            if self.is_cancelled:
                return
                
            # PHASE 2: FILE PROCESSING (5-25%) - This is the heaviest part
            self.progress_updated.emit(8, "Loading Excel files...", 0)
            self.log_message.emit("📊 Processing Excel files...")
            
            # Capture print output during file processing
            self._capture_print_output()
            self.log_message.emit("   - Loading Excel files and extracting data...")
            transactions1, transactions2, blocks1, blocks2, lc_numbers1, lc_numbers2, po_numbers1, po_numbers2, interunit_accounts1, interunit_accounts2, usd_amounts1, usd_amounts2 = matcher.process_files()
            self._release_print_output()
            self.log_message.emit(f"   ✅ File processing completed - {len(transactions1)} and {len(transactions2)} transactions loaded")
            
            if self.is_cancelled:
                return
                
            # PHASE 3: MATCHING LOGIC (25-60%) - 5 steps, 7% each
            # Step 1: Narration Matching (25-32%)
            self.progress_updated.emit(25, "Finding narration matches...", 0)
            self.log_message.emit("🔍 Step 1/5: Narration Matching")
            self.log_message.emit("   - Searching for exact text matches in transaction descriptions...")
            self.log_message.emit("   - Analyzing transaction descriptions for matches...")
            
            narration_matches = matcher.narration_matching_logic.find_potential_matches(
                transactions1, transactions2, {}, None
            )
            self.log_message.emit(f"   ✅ Found {len(narration_matches)} narration matches")
            self.log_message.emit(f"   - Narration matching completed in {len(narration_matches)} matches")
            self.step_completed.emit("Narration Matching", len(narration_matches))
            
            if self.is_cancelled:
                return
                
            # Step 2: LC Matching (32-39%)
            self.progress_updated.emit(32, "Finding LC matches...", 0)
            self.log_message.emit("🔍 Step 2/5: LC Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
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
            
            self.log_message.emit("   - Searching for LC number matches...")
            self.log_message.emit("   - Analyzing LC numbers for potential matches...")
            lc_matches = matcher.lc_matching_logic.find_potential_matches(
                transactions1, transactions2, lc_numbers1_unmatched, lc_numbers2_unmatched,
                {}, None
            )
            self.log_message.emit(f"   ✅ Found {len(lc_matches)} LC matches")
            self.log_message.emit(f"   - LC matching completed in {len(lc_matches)} matches")
            self.step_completed.emit("LC Matching", len(lc_matches))
            
            if self.is_cancelled:
                return
                
            # Step 3: PO Matching (39-46%)
            self.progress_updated.emit(39, "Finding PO matches...", 0)
            self.log_message.emit("🔍 Step 3/5: PO Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
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
            
            self.log_message.emit("   - Searching for PO number matches...")
            self.log_message.emit("   - Analyzing PO numbers for potential matches...")
            po_matches = matcher.po_matching_logic.find_potential_matches(
                transactions1, transactions2, po_numbers1_unmatched, po_numbers2_unmatched,
                {}, None
            )
            self.log_message.emit(f"   ✅ Found {len(po_matches)} PO matches")
            self.log_message.emit(f"   - PO matching completed in {len(po_matches)} matches")
            self.step_completed.emit("PO Matching", len(po_matches))
            
            if self.is_cancelled:
                return
                
            # Step 4: Interunit Matching (46-53%)
            self.progress_updated.emit(46, "Finding interunit matches...", 0)
            self.log_message.emit("🔍 Step 4/5: Interunit Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
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
            
            self.log_message.emit("   - Searching for interunit account matches...")
            self.log_message.emit("   - Analyzing interunit accounts for potential matches...")
            interunit_matches = matcher.interunit_loan_matcher.find_potential_matches(
                transactions1, transactions2, interunit_accounts1_unmatched, interunit_accounts2_unmatched,
                self.file1_path, self.file2_path, {}, None
            )
            self.log_message.emit(f"   ✅ Found {len(interunit_matches)} interunit matches")
            self.log_message.emit(f"   - Interunit matching completed in {len(interunit_matches)} matches")
            self.step_completed.emit("Interunit Matching", len(interunit_matches))
            
            if self.is_cancelled:
                return
                
            # Step 5: USD Matching (53-60%)
            self.progress_updated.emit(53, "Finding USD matches...", 0)
            self.log_message.emit("🔍 Step 5/5: USD Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
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
            
            self.log_message.emit("   - Searching for USD amount matches...")
            self.log_message.emit("   - Analyzing USD amounts for potential matches...")
            usd_matches = matcher.usd_matching_logic.find_potential_matches(
                transactions1, transactions2, usd_amounts1_unmatched, usd_amounts2_unmatched,
                {}, None
            )
            self.log_message.emit(f"   ✅ Found {len(usd_matches)} USD matches")
            self.log_message.emit(f"   - USD matching completed in {len(usd_matches)} matches")
            self.step_completed.emit("USD Matching", len(usd_matches))
            
            if self.is_cancelled:
                return
                
            # PHASE 4: MATCH PROCESSING (60-65%)
            self.progress_updated.emit(60, "Processing matches...", 0)
            self.log_message.emit("📊 Processing all matches...")
            
            # Combine all matches
            self.log_message.emit("   - Combining matches from all matching types...")
            all_matches = narration_matches + lc_matches + po_matches + interunit_matches + usd_matches
            self.log_message.emit(f"   - Combined {len(all_matches)} total matches")
            
            # Assign sequential Match IDs
            self.log_message.emit("   - Assigning sequential Match IDs...")
            self.log_message.emit("   - Organizing matches for output...")
            match_counter = 1
            for match in all_matches:
                match_id = f"M{match_counter:03d}"
                match['match_id'] = match_id
                match_counter += 1
            
            # Sort matches by the newly assigned sequential Match IDs
            self.log_message.emit("   - Sorting matches by Match ID...")
            all_matches.sort(key=lambda x: x['match_id'])
            self.log_message.emit("   - Match processing completed successfully!")
            
            # Create statistics
            stats = {
                'total_matches': len(all_matches),
                'narration_matches': len(narration_matches),
                'lc_matches': len(lc_matches),
                'po_matches': len(po_matches),
                'interunit_matches': len(interunit_matches),
                'usd_matches': len(usd_matches)
            }
            
            self.log_message.emit("🎉 Matching completed successfully!")
            self.log_message.emit(f"📈 Final Results: {stats['total_matches']} total matches found")
            self.progress_updated.emit(65, "Matching completed successfully!", stats['total_matches'])
            self.matching_finished.emit(all_matches, stats)
            
        except Exception as e:
            # Make sure to release print output capture on error
            self._release_print_output()
            self.error_occurred.emit(str(e))
    
    def cancel(self):
        """Cancel the matching process"""
        self.is_cancelled = True
