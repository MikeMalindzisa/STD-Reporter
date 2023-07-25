import os
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QMessageBox, QVBoxLayout, QHBoxLayout, QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
import configparser


# Get the default INI file path
default_ini_file = 'stdsettings.ini'

# Check if the INI file exists or is empty
if not os.path.exists(default_ini_file) or os.stat(default_ini_file).st_size == 0:
    # Create the INI file with default values
    config = configparser.ConfigParser()
    config['Folders'] = {
        'InputFolder': '',
        'OutputFolder': '',
        'CompletedFolder': ''
    }
    with open(default_ini_file, 'w') as config_file:
        config.write(config_file)

# Create a ConfigParser instance
config = configparser.ConfigParser()

# Read the INI file
config.read(default_ini_file)

# Read the folder paths from the INI file
input_folder = config.get('Folders', 'InputFolder')
output_folder = config.get('Folders', 'OutputFolder')
completed_folder = config.get('Folders', 'CompletedFolder')


class FolderConfigWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Folder Config")
        self.setFixedSize(400, 250)

        # Set the window icon
        icon_path = "icon.png"
        self.setWindowIcon(QIcon(icon_path))
         # Set title bar background color
        self.setStyleSheet("QWidget { background-color: #fff; }")

        # Create the main layout
        main_layout = QVBoxLayout()

        # Input folder field
        input_layout = QHBoxLayout()
        input_label = QLabel("Input Folder:")
        input_layout.addWidget(input_label)
        self.input_entry = QLineEdit()
        self.input_entry.setText(input_folder)
        input_layout.addWidget(self.input_entry)
        browse_input_button = QPushButton("Browse")
        browse_input_button.clicked.connect(self.browse_input_folder)
        input_layout.addWidget(browse_input_button)
        main_layout.addLayout(input_layout)

        # Output folder field
        output_layout = QHBoxLayout()
        output_label = QLabel("Output Folder:")
        output_layout.addWidget(output_label)
        self.output_entry = QLineEdit()
        self.output_entry.setText(output_folder)
        output_layout.addWidget(self.output_entry)
        browse_output_button = QPushButton("Browse")
        browse_output_button.clicked.connect(self.browse_output_folder)
        output_layout.addWidget(browse_output_button)
        main_layout.addLayout(output_layout)

        # Completed folder field
        completed_layout = QHBoxLayout()
        completed_label = QLabel("Completed Folder:")
        completed_layout.addWidget(completed_label)
        self.completed_entry = QLineEdit()
        self.completed_entry.setText(completed_folder)
        completed_layout.addWidget(self.completed_entry)
        browse_completed_button = QPushButton("Browse")
        browse_completed_button.clicked.connect(self.browse_completed_folder)
        completed_layout.addWidget(browse_completed_button)
        main_layout.addLayout(completed_layout)

        # Save button
        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_paths)
        main_layout.addWidget(save_button)

        self.setLayout(main_layout)

    def browse_input_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Input Folder")
        if folder_path:
            self.input_entry.setText(folder_path)

    def browse_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder_path:
            self.output_entry.setText(folder_path)

    def browse_completed_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Completed Folder")
        if folder_path:
            self.completed_entry.setText(folder_path)

    def save_paths(self):
        # Update the folder paths from the entry fields
        input_folder = self.input_entry.text()
        output_folder = self.output_entry.text()
        completed_folder = self.completed_entry.text()

        # Save the changes to the INI file
        config.set('Folders', 'InputFolder', input_folder)
        config.set('Folders', 'OutputFolder', output_folder)
        config.set('Folders', 'CompletedFolder', completed_folder)
        with open(default_ini_file, 'w') as config_file:
            config.write(config_file)

        QMessageBox.information(self, "Success", "Folder paths saved successfully.")
        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FolderConfigWindow()
    window.show()
    sys.exit(app.exec_())
