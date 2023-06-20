# Import PyQt5 modules
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QDialog, QCheckBox, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QSpacerItem, QSizePolicy
from PyQt5.QtCore import Qt


# Define the popup class
class ReportPopup(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        # Set the window title and size
        self.setWindowTitle("Report Optionen")
        custom_font = QFont()
        custom_font.setWeight(38)
        # Styling font size
        # self.setStyleSheet("QLabel{font-size: 11pt;} QCheckBox{font-size: 9pt;} QPushButton{font-size: 9pt;}")
        self.resize(300, 200)
        # Create the checkboxes
        self.pdf_check = QCheckBox("PDF")
        self.html_check = QCheckBox("HTML")
        self.save_check = QCheckBox("In Downloads speichern")
        # Create the buttons
        self.ok_button = QPushButton("Fortfahren")
        self.cancel_button = QPushButton("Abbrechen")
        # Connect the buttons to their functions
        self.ok_button.clicked.connect(self.ok_clicked)
        self.cancel_button.clicked.connect(self.cancel_clicked)
        # Create the labels
        self.title_label = QLabel("Wählen Sie zumindest ein Report-Format aus:")
        self.message_label = QLabel("")
        # Create the layouts
        self.main_layout = QVBoxLayout()
        self.check_layout = QVBoxLayout()
        self.button_layout = QHBoxLayout()
        # Add the widgets to the layouts
        self.check_layout.addWidget(self.html_check)
        self.check_layout.addWidget(self.pdf_check)
        self.check_layout.addSpacing(50)
        self.check_layout.addWidget(self.save_check)
        self.button_layout.addWidget(self.ok_button)
        self.button_layout.addWidget(self.cancel_button)
        self.main_layout.addWidget(self.title_label)
        self.main_layout.addLayout(self.check_layout)
        self.main_layout.addWidget(self.message_label)
        self.main_layout.addLayout(self.button_layout)
        # Set the main layout
        self.setLayout(self.main_layout)
        # Set the standard selection
        self.pdf_check.setCheckState(Qt.CheckState.Checked)
        self.save_check.setCheckState(Qt.CheckState.Checked)
        self.result = None

    def ok_clicked(self):
        # Check if at least one option is selected
        if not (self.pdf_check.isChecked() or self.html_check.isChecked()):
            # Display an error message
            self.message_label.setText("Bitte wählen Sie mindestens ein Report-Format aus.")
            return
        # Get the selected options as a dictionary
        options = {
            "html": self.html_check.isChecked(),
            "pdf": self.pdf_check.isChecked(),
            "save": self.save_check.isChecked()
        }
        # Emit a signal with the options
        self.result = options
        # Close the popup
        self.close()

    def cancel_clicked(self):
        self.result = None
        # Close the popup without emitting any signal
        self.close()
