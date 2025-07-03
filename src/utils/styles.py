MAIN_STYLE = """
QMainWindow {
    background-color: #f8fafc;
}

QTabWidget::pane {
    border: 2px solid #e2e8f0;
    border-radius: 4px;
    background: white;
}

QTabBar::tab {
    background-color: #f1f5f9;
    padding: 8px 12px;
    border: 1px solid #e2e8f0;
    border-radius: 4px;
    margin: 2px;
}

QTabBar::tab:selected {
    background-color: #3b82f6;
    color: white;
    border: 1px solid #2563eb;
}

/* Default button style - use blue */
QPushButton {
    background-color: #3b82f6;
    color: white;
    border: 2px solid #2563eb;
    border-radius: 4px;
    padding: 5px 10px;
}

/* Select and Create buttons - blue */
QPushButton#select_button,
QPushButton#create_button {
    background-color: #3b82f6;
    border: 2px solid #2563eb;
}

/* Test email button - green color scheme */
QPushButton#test_button {
    background-color: #10b981 !important;
    border: 2px solid #059669 !important;
    color: white !important;
}

/* Send all button - yellow color scheme */
QPushButton#send_button {
    background-color: #fbbf24 !important;  /* Amber-400 */
    border: 2px solid #d97706 !important;  /* Amber-600 */
    color: #000000 !important;  /* Black text */
}

QPushButton:hover {
    background-color: #2563eb;
    border-color: #1d4ed8;
}

/* Disabled state for both special buttons */
QPushButton#send_button:disabled,
QPushButton#test_button:disabled {
    background-color: #94a3b8;
    border-color: #64748b;
}

/* Status labels with corresponding colors */
#total_label {
    color: #1e293b;
    background-color: #f1f5f9;
    border: 1px solid #cbd5e1;
    padding: 5px 10px;
    border-radius: 4px;
}

#sent_label {
    color: #166534;
    background-color: #dcfce7;
    border: 1px solid #86efac;
    padding: 5px 10px;
    border-radius: 4px;
}

#failed_label {
    color: #991b1b;
    background-color: #fee2e2;
    border: 1px solid #fca5a5;
    padding: 5px 10px;
    border-radius: 4px;
}

#skipped_label {
    color: #475569;
    background-color: #f1f5f9;
    border: 1px solid #cbd5e1;
    padding: 5px 10px;
    border-radius: 4px;
}

QLabel {
    color: #1e293b;
}

QProgressBar {
    border: 1px solid #e2e8f0;
    border-radius: 4px;
    background-color: #f1f5f9;
}

QProgressBar::chunk {
    background-color: #3b82f6;
    border-radius: 3px;
}

QLineEdit, QPlainTextEdit {
    border: 1px solid #e2e8f0;
    background-color: white;
}

QCheckBox {
    color: #1e293b;
}
"""