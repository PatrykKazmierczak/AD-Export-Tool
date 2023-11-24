# Import necessary modules
from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget, QPushButton, QSizePolicy
from PyQt5.QtGui import QFont, QIcon
import logging

# Define the DashboardWindow class
class DashboardWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set window title and icon
        self.setWindowTitle("AD-Tool")
        self.setWindowIcon(QIcon('./icons/microsoft-active-directory5035.jpg'))

        # Create a QVBoxLayout object
        layout = QVBoxLayout()

        # Create a QFont object
        font = QFont("Arial", 14)

        # Create the first button and connect it to the first script
        run_first_script_button = QPushButton('Run Get-ADComputersExport')
        run_first_script_button.clicked.connect(self.run_first_script)
        run_first_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_first_script_button.setFont(font)
        layout.addWidget(run_first_script_button)

        # Create the second button and connect it to the second script
        run_second_script_button = QPushButton('Run Get-ADUsersExport')
        run_second_script_button.clicked.connect(self.run_second_script)
        run_second_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_second_script_button.setFont(font)
        layout.addWidget(run_second_script_button)

        # Create the third button and connect it to the third script
        run_third_script_button = QPushButton('Run Get-UserComputer')
        run_third_script_button.clicked.connect(self.run_third_script)
        run_third_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_third_script_button.setFont(font)
        layout.addWidget(run_third_script_button)

        # Create the fourth button and connect it to the fourth script
        run_data_summary_script_button = QPushButton('Run CreateITOADashboard')
        run_data_summary_script_button.clicked.connect(self.run_data_summary_script)
        run_data_summary_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_data_summary_script_button.setFont(font)
        layout.addWidget(run_data_summary_script_button)

        # Create a QWidget object, set its layout, and set it as the central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
    # Set up logging to a file named 'debug.txt' with level set to DEBUG
    logging.basicConfig(filename='debug.txt', level=logging.DEBUG)    

    # Define the method to run the first script
    def run_first_script(self):
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-ADComputersExportToSQL.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the first script using subprocess
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()
                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running first script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            raise e

    # Define the method to run the second script
    def run_second_script(self):
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-ADUsersExportToSQLOnPremUpdate.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the second script using subprocess
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()
                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running second script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            raise e

    # Define the method to run the third script
    def run_third_script(self):
        try:
            # Similar to the second script, create a temporary directory, extract the script, and run it
            with tempfile.TemporaryDirectory() as temp_dir:
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-UserComputerInfo.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()
                if process.returncode != 0:
                    logging.error(f"Error running third script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")
        except Exception as e:
            logging.exception("Exception occurred: ")
            raise e           
        
    # Define the method to run the data summary script
    def run_data_summary_script(self):
        try:
            # Get the path to the Excel file inside the executable
            # pkg_resources.resource_filename is used to get the path of a resource
            # It takes two arguments: package name and resource name
            excel_file_path = pkg_resources.resource_filename(__name__, 'excel/ITOA_ProblemsDashboard.xlsx')

            # Copy the Excel file to C:\ITOA
            # shutil.copy is used to copy the file source (first argument) to the file destination (second argument)
            shutil.copy(excel_file_path, 'C:\\ITOA\\ITOA_ProblemsDashboard.xlsx')
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            # logging.exception logs a message with level ERROR on the root logger
            # The exception information is added to the logging message
            logging.exception("Exception occurred: ")
            raise e   


    # Set up logging
    