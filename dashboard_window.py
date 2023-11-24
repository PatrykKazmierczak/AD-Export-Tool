# Import necessary modules
from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget, QPushButton, QSizePolicy, QHBoxLayout, QPlainTextEdit
from PyQt5.QtGui import QFont, QIcon
import logging
import tempfile
import shutil
import subprocess
import pkg_resources
from PyQt5.QtWidgets import QTabWidget, QToolBar, QToolButton
from PyQt5.QtWidgets import QVBoxLayout

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

        # Create a QToolBar for each tab
        home_tab = QToolBar()
        insert_tab = QToolBar()

        # Create QToolButtons for each action
        save_button = QToolButton()
        save_button.setText("Save")
        # save_button.clicked.connect(self.save)  # Connect to the appropriate slot


        # Add the toolbars to the ribbon
        self.ribbon.addTab(home_tab, "Home")
        self.ribbon.addTab(insert_tab, "Insert")

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

        # Create a QVBoxLayout object for the buttons
        buttons_layout = QVBoxLayout()

        # Add the buttons to the buttons_layout
        buttons_layout.addWidget(run_first_script_button)
        buttons_layout.addWidget(run_second_script_button)
        buttons_layout.addWidget(run_third_script_button)
        buttons_layout.addWidget(run_data_summary_script_button)
        
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
        # Create a QWidget object, set its layout, and set it as the central widget
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
        
    # Set up logging to a file named 'debug.txt' with level set to DEBUG
    logging.basicConfig(filename='debug.txt', level=logging.DEBUG)    

    # Define the method to run the first script
    def run_first_script(self):
        self.console.appendPlainText("Running first script...")
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
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            console.appendPlainText("Exception: " + str(e))
            raise e

    # Define the method to run the second script
    def run_second_script(self):
        self.console.appendPlainText("Running second script...")
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
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            console.appendPlainText("Exception: " + str(e))
            raise e

    # Define the method to run the third script
    def run_third_script(self):
        self.console.appendPlainText("Running third script...")
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
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            console.appendPlainText("Exception: " + str(e))
            raise e           
        
    # Define the method to run the data summary script
    def run_data_summary_script(self):
        self.console.appendPlainText("Running data summary script...")
        try:
            # Get the path to the Excel file inside the executable
            excel_file_path = pkg_resources.resource_filename(__name__, 'excel/ITOA_ProblemsDashboard.xlsx')

            # Copy the Excel file to C:\ITOA
            shutil.copy(excel_file_path, 'C:\\ITOA\\ITOA_ProblemsDashboard.xlsx')

            # Write a success message to the console
            self.console.appendPlainText("Excel file copied successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            raise e   


    # Set up logging
    