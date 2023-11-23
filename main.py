from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QLineEdit, QPushButton, QSizePolicy
from PyQt5 import QtCore
from PyQt5.QtGui import QIcon
from PyQt5.QtGui import QFont
import logging
import csv
import os
import pandas as pd
import subprocess
import csv
import io
import sys
import tempfile
import shutil
import pkg_resources

app = QApplication([])
app.setStyleSheet("""
    QMainWindow {
        background-color: #2b2b2b;
    }

    QPushButton {
        color: #fff;
        background-color: #3d3d3d;
        border: none;
        padding: 10px;
        min-width: 100px;
        min-height: 40px;
        border-radius: 5px;
    }

    QPushButton:hover {
        background-color: #555;
    }

    QPushButton:pressed {
        background-color: #777;
    }

    QLineEdit {
        padding: 10px;
        border: 2px solid #555;
        border-radius: 10px;
        color: #fff;
        background-color: #3d3d3d;
    }
""")

from PyQt5.QtGui import QFont

class DashboardWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("AD-Tool")
        self.setWindowIcon(QIcon('./icons/microsoft-active-directory5035.jpg'))

        layout = QVBoxLayout()

        font = QFont("Arial", 14)  # Create a QFont object

        run_first_script_button = QPushButton('Run Get-ADComputersExport')
        run_first_script_button.clicked.connect(self.run_first_script)
        run_first_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_first_script_button.setFont(font)  # Set the font of the button
        layout.addWidget(run_first_script_button)

        run_second_script_button = QPushButton('Run Get-ADUsersExport')
        run_second_script_button.clicked.connect(self.run_second_script)
        run_second_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_second_script_button.setFont(font)  # Set the font of the button
        layout.addWidget(run_second_script_button)

        run_third_script_button = QPushButton('Run Get-UserComputer')
        run_third_script_button.clicked.connect(self.run_third_script)
        run_third_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_third_script_button.setFont(font)
        layout.addWidget(run_third_script_button)

        # Add a button for the DataSummary script
        run_data_summary_script_button = QPushButton('Run DataSummary')
        run_data_summary_script_button.clicked.connect(self.run_data_summary_script)
        run_data_summary_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_data_summary_script_button.setFont(font)
        layout.addWidget(run_data_summary_script_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    # Set up logging
    logging.basicConfig(filename='debug.txt', level=logging.DEBUG)


    import csv

    def run_first_script(self):
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-ADComputersExportToSQL.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the first script
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()
                if process.returncode != 0:
                    logging.error(f"Error running first script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")
        except Exception as e:
            logging.exception("Exception occurred: ")
            raise e

    def run_second_script(self):
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-ADUsersExportToSQLOnPremUpdate.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the second script
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()
                if process.returncode != 0:
                    logging.error(f"Error running second script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")
        except Exception as e:
            logging.exception("Exception occurred: ")
            raise e

    def run_third_script(self):
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-UserComputerInfo.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the third script
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()
                if process.returncode != 0:
                    logging.error(f"Error running third script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")
        except Exception as e:
            logging.exception("Exception occurred: ")
            raise e      
        
    def run_data_summary_script(self):
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/DataSummary.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the DataSummary script
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()
                if process.returncode != 0:
                    logging.error(f"Error running DataSummary script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")
        except Exception as e:
            logging.exception("Exception occurred: ")
            raise e    
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("AD-Tool")
        self.setWindowIcon(QIcon('./icons/microsoft-active-directory5035.jpg'))

        layout = QVBoxLayout()

        self.username_field = QLineEdit()
        self.username_field.setPlaceholderText('Username')
        self.username_field.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.username_field)

        self.password_field = QLineEdit()
        self.password_field.setPlaceholderText('Password')
        self.password_field.setEchoMode(QLineEdit.Password)
        self.password_field.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.password_field)

        login_button = QPushButton('Log In')
        login_button.clicked.connect(self.check_credentials)
        login_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(login_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def check_credentials(self):
        username = self.username_field.text()
        password = self.password_field.text()

        # Here you would check the username and password against your database
        # For simplicity, we're just checking against hardcoded values
        if username == 'admin' and password == 'password':
            self.hide()
            self.dashboard_window = DashboardWindow()
            self.dashboard_window.show()
        else:
            print('Access denied')

window = MainWindow()
window.show()

app.exec_()