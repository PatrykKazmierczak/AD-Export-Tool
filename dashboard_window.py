# Import necessary modules
from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget, QPushButton, QSizePolicy, QHBoxLayout, QPlainTextEdit
from PyQt5.QtGui import QFont, QIcon
import logging
import tempfile
import shutil
import subprocess
import pkg_resources
from PyQt5.QtWidgets import QTabWidget, QVBoxLayout
import pandas as pd
import os
from PyQt5 import QtWidgets

class DashboardWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set window title and icon
        self.setWindowTitle("AD-Tool")
        self.setWindowIcon(QIcon('./icons/microsoft-active-directory5035.jpg'))

        # Create a QFont object
        font = QFont("Arial", 14)

        # Create a QTabWidget for the ribbon
        self.ribbon = QTabWidget()  

        # Create the first button and connect it to the first script
        run_first_script_button = QPushButton('Run Get-ADComputersExport')
        run_first_script_button.clicked.connect(self.run_first_script)
        run_first_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_first_script_button.setFont(font)

        # Create the second button and connect it to the second script
        run_second_script_button = QPushButton('Run Get-ADUsersExport')
        run_second_script_button.clicked.connect(self.run_second_script)
        run_second_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_second_script_button.setFont(font)

        # Create the third button and connect it to the third script
        run_third_script_button = QPushButton('Run Get-UserComputer')
        run_third_script_button.clicked.connect(self.run_third_script)
        run_third_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_third_script_button.setFont(font)

        # Create the fourth button and connect it to the fourth script
        run_data_summary_script_button = QPushButton('Run CreateITOADashboard')
        run_data_summary_script_button.clicked.connect(self.run_data_summary_script)
        run_data_summary_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_data_summary_script_button.setFont(font)
        
        # Create the clear console button and connect it to the clear console method
        clear_console_button = QPushButton('Clear Console')
        clear_console_button.clicked.connect(self.clear_console)
        clear_console_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        clear_console_button.setFont(font)

        # Create a QVBoxLayout object for the buttons
        buttons_layout = QVBoxLayout()

        # Add the buttons to the buttons_layout
        buttons_layout.addWidget(run_first_script_button)
        buttons_layout.addWidget(run_second_script_button)
        buttons_layout.addWidget(run_third_script_button)
        buttons_layout.addWidget(run_data_summary_script_button)
        buttons_layout.addWidget(clear_console_button)
        
        # Create a QPlainTextEdit object
        self.console = QPlainTextEdit()
        self.console.setReadOnly(True)  # Make it read-only

        # Create a QHBoxLayout object for the buttons and console
        layout = QHBoxLayout()

        layout.addLayout(buttons_layout)
        layout.addWidget(self.console)

        # Create a QVBoxLayout for the main layout
        main_layout = QVBoxLayout()

        # Add the ribbon and the layout to the main layout
        main_layout.addWidget(self.ribbon)
        main_layout.addLayout(layout)

        # Create a QWidget object, set its layout, and set it as the central widget
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
        
    # Set up logging to a file named 'debug.txt' with level set to DEBUG
    logging.basicConfig(filename='debug.txt', level=logging.DEBUG)
    
    def clear_console(self):
        self.console.clear()    

    # Define the method to run the first script
    def run_first_script(self):
        self.console.appendPlainText("Running first script Get-ADComputersExportToSQL...")
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-ADComputersExportToSQL.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the first script using subprocess
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()

                # Write the output and errors to the console
                self.console.appendPlainText(output.decode('utf-8'))
                if error:
                    self.console.appendPlainText("Error: " + error.decode('utf-8'))

                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running first script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")

            # If the script runs successfully, print the success message
            self.console.appendPlainText("CSV file generated successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            raise e

    # Define the method to run the second script
    def run_second_script(self):
        self.console.appendPlainText("Running second script Get-ADUsersExportToSQLOnPremUpdate...")
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-ADUsersExportToSQLOnPremUpdate.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the second script using subprocess
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()

                # Write the output and errors to the console
                self.console.appendPlainText(output.decode('utf-8'))
                if error:
                    self.console.appendPlainText("Error: " + error.decode('utf-8'))

                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running second script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")

            # If the script runs successfully, print the success message
            self.console.appendPlainText("CSV file generated successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            raise e

    # Define the method to run the third script
    def run_third_script(self):
        self.console.appendPlainText("Running third script Get-UserComputerInfo ...")
        try:
            # Similar to the second script, create a temporary directory, extract the script, and run it
            with tempfile.TemporaryDirectory() as temp_dir:
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-UserComputerInfo.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()

                # Write the output and errors to the console
                self.console.appendPlainText(output.decode('utf-8'))
                if error:
                    self.console.appendPlainText("Error: " + error.decode('utf-8'))

                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running third script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")

            # If the script runs successfully, print the success message
            self.console.appendPlainText("CSV file generated successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            raise e           
        
    # Define the method to run the data summary script
    

    # Define the method to run the data summary script
    

    

    def run_data_summary_script(self):
        self.console.appendPlainText("Running data summary script...")
        try:
            # Get the path to the Excel file inside the executable
            excel_file_path = pkg_resources.resource_filename(__name__, 'excel/ITOA_ProblemsDashboard.xlsx')

            # Check if the Excel file already exists in C:\ITOA
            destination_path = 'C:\\ITOA\\ITOA_ProblemsDashboard.xlsx'
            if os.path.exists(destination_path):
                self.console.appendPlainText("Excel file already exists.")
                return

            # Copy the Excel file to C:\ITOA
            shutil.copy(excel_file_path, destination_path)

            # Check if the CSV files exist and are not empty
            csv_file_paths = ['C:\\ITOA\\Get-ADComputers.csv', 'C:\\ITOA\\Get-ADUsers.csv']
            for csv_file_path in csv_file_paths:
                if not os.path.exists(csv_file_path) or os.path.getsize(csv_file_path) == 0:
                    raise ValueError(f"CSV file {csv_file_path} does not exist or is empty")

            # Read the CSV files
            df1 = pd.read_csv(csv_file_paths[0])
            df2 = pd.read_csv(csv_file_paths[1])

            # Write the dataframes to the Excel file
            with pd.ExcelWriter(destination_path) as writer:
                df1.to_excel(writer, sheet_name='Sheet1')
                df2.to_excel(writer, sheet_name='Sheet2')

            # Write a success message to the console
            self.console.appendPlainText("Excel file generated successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and show an error message
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            # Show an error message to the user
            error_dialog = QtWidgets.QErrorMessage()
            error_dialog.showMessage('Oh no!')