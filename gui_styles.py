"""
GUI Styles for Interunit Loan Matcher
Centralized CSS styling for better maintainability
"""

def get_main_stylesheet():
    """Get the main application stylesheet"""
    return """
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
            font-size: 16px;
        }
        
        /* Section Containers with Windows Dialog Style */
        QWidget[class="section-container"] {
            background-color: white;
            border: 1px solid #e1e5e9;
            border-radius: 8px;
            margin: 0px;
        }
        
        /* Bootstrap btn-sm - Primary Button */
        QPushButton[class="browse-button"] {
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 3px;
            padding: 4px 8px;
            font-family: "Segoe UI";
            font-weight: 600;
            font-size: 14px;
            min-height: 20px;
            min-width: 60px;
        }
        QPushButton[class="browse-button"]:hover {
            background-color: #0056b3;
        }
        
        /* Bootstrap btn-sm - Secondary Button */
        QPushButton[class="clear-button"] {
            background-color: white;
            color: #6c757d;
            border: 1px solid #6c757d;
            border-radius: 3px;
            padding: 4px 8px;
            font-family: "Segoe UI";
            font-weight: 600;
            font-size: 14px;
            min-height: 20px;
            min-width: 60px;
        }
        QPushButton[class="clear-button"]:hover {
            background-color: #6c757d;
            color: white;
        }
        
        /* Bootstrap btn-sm - Primary Action Button */
        QPushButton[class="run-button"] {
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 3px;
            padding: 6px 12px;
            font-family: "Segoe UI";
            font-weight: 600;
            font-size: 14px;
            min-height: 24px;
            min-width: 70px;
        }
        QPushButton[class="run-button"]:hover {
            background-color: #0056b3;
        }
        QPushButton[class="run-button"]:disabled {
            background-color: #6c757d;
            color: #adb5bd;
        }
        
        /* Open Folder Button */
        QPushButton[class="open-folder-button"] {
            background-color: #28a745;
            color: white;
            border: none;
            border-radius: 3px;
            padding: 6px 12px;
            font-family: "Segoe UI";
            font-weight: 600;
            font-size: 14px;
            min-height: 30px;
            min-width: 120px;
        }
        QPushButton[class="open-folder-button"]:hover {
            background-color: #218838;
        }
        QPushButton[class="open-folder-button"]:disabled {
            background-color: #6c757d;
            color: #adb5bd;
        }
        
        /* Open Files Button */
        QPushButton[class="open-files-button"] {
            background-color: #17a2b8;
            color: white;
            border: none;
            border-radius: 3px;
            padding: 6px 12px;
            font-family: "Segoe UI";
            font-weight: 600;
            font-size: 14px;
            min-height: 30px;
            min-width: 120px;
        }
        QPushButton[class="open-files-button"]:hover {
            background-color: #138496;
        }
        QPushButton[class="open-files-button"]:disabled {
            background-color: #6c757d;
            color: #adb5bd;
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
            padding: 2px 4px;
            margin: 1px 0;
            background-color: transparent;
            color: #495057;
            font-family: "Segoe UI";
            font-size: 14px;
            font-weight: 400;
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
        
        /* Match Summary Text Items - bold text style */
        QLabel[class="match-summary-text"] {
            color: #495057;
            font-weight: bold;
            font-size: 13px;
            font-family: "Segoe UI";
            padding: 2px 0px;
        }
        
        /* Total Match Summary - smaller height */
        QLabel[class="total-match-summary"] {
            border: 2px solid #28a745;
            border-radius: 6px;
            background-color: #d4edda;
            padding: 4px 12px;
            color: #155724;
            font-weight: 600;
            font-size: 14px;
            text-align: center;
            font-family: "Segoe UI";
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
    """
