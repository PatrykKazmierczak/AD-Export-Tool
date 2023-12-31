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
import platform
import socket
import psutil
from PyQt5.QtWidgets import QLabel
import datetime
import logging
import pkg_resources
import shutil
import sqlite3
import subprocess
import tempfile
import pkg_resources
import sqlite3
import tempfile
import pkg_resources
import shutil
import subprocess
import logging
import csv
import io


class DashboardWindow(QMainWindow):
    def __init__(self, conn):
        super().__init__()
        self.conn = conn
        
        
         # Set window title and icon
        self.setWindowTitle("ITOA  ActiveDirectory  Tool")
        self.setWindowIcon(QIcon('./icons/microsoft-active-directory5035.jpg'))

        # Create a QMenuBar
        menu_bar = self.menuBar()

        # Create the menus
        file_menu = menu_bar.addMenu('File')
        edit_menu = menu_bar.addMenu('Edit')
        view_menu = menu_bar.addMenu('View')
        go_menu = menu_bar.addMenu('Go')
        tools_menu = menu_bar.addMenu('Tools')
        settings_menu = menu_bar.addMenu('Settings')
        help_menu = menu_bar.addMenu('Help')

        # Create a QFont object
        font = QFont("Arial", 14)

        # Create a QTabWidget for the ribbon
        self.ribbon = QTabWidget()
        
        # Get the server/computer data
        os_info = f"OS: {platform.system()} {platform.release()}"
        hostname = f"Hostname: {socket.gethostname()}"
        ip_address = f"IP: {socket.gethostbyname(socket.gethostname())}"
        cpu_info = f"CPU: {platform.processor()}"
        total_memory = f"Total Memory: {round(psutil.virtual_memory().total / (1024.0 ** 3), 2)} GB"
        disk_usage = psutil.disk_usage('/')
        total_disk_space = f"Total Disk Space: {round(disk_usage.total / (1024.0 ** 3), 2)} GB"
        free_disk_space = f"Free Disk Space: {round(disk_usage.free / (1024.0 ** 3), 2)} GB"
        # Drive info and adapter info would require additional code to retrieve

        # Create a QLabel for the server/computer data
        server_data_label = QLabel()
        server_data_label.setText(f"{os_info}, {hostname}, {ip_address}, {cpu_info}, {total_memory}, {total_disk_space}, {free_disk_space}")
        server_data_label.setStyleSheet("color: black;")  # Set the text color to black 

        # Add the server_data_label to the ribbon
        self.ribbon.addTab(server_data_label, "Server Data")

        # Create the first button and connect it to the first script
        run_first_script_button = QPushButton('Run Get-AD-Computers-Export')
        run_first_script_button.clicked.connect(self.run_first_script)
        run_first_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_first_script_button.setFont(font)

        # Create the second button and connect it to the second script
        run_second_script_button = QPushButton('Run Get-AD-Users-Export')
        run_second_script_button.clicked.connect(self.run_second_script)
        run_second_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_second_script_button.setFont(font)

        # Create the third button and connect it to the third script
        run_third_script_button = QPushButton('Run Get-User-Computer')
        run_third_script_button.clicked.connect(self.run_third_script)
        run_third_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_third_script_button.setFont(font)

        # Create the fourth button and connect it to the fourth script
        run_data_summary_script_button = QPushButton('Run Create-ITOA-Dashboard')
        run_data_summary_script_button.clicked.connect(self.run_data_summary_script)
        run_data_summary_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_data_summary_script_button.setFont(font)
        
        # Create the Send-AD-Status button and connect it to the appropriate method
        send_ad_status_button = QPushButton('Run Send-AD-Status')
        send_ad_status_button.clicked.connect(self.send_ad_status)
        send_ad_status_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        send_ad_status_button.setFont(font)
        
        # Create the clear console button and connect it to the clear console method
        clear_console_button = QPushButton('Run Clear-Console')
        clear_console_button.clicked.connect(self.clear_console)
        clear_console_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        clear_console_button.setFont(font)
        
        extract_users_from_azure_ad_button = QPushButton('Run Extract-Users-From-AzureAD')
        extract_users_from_azure_ad_button.clicked.connect(self.run_extract_users_from_azure_ad_script)
        extract_users_from_azure_ad_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        extract_users_from_azure_ad_button.setFont(font)

        # Create a QVBoxLayout object for the buttons
        buttons_layout = QVBoxLayout()

        # Add the buttons to the buttons_layout
        buttons_layout.addWidget(run_first_script_button)
        buttons_layout.addWidget(run_second_script_button)
        buttons_layout.addWidget(run_third_script_button)
        buttons_layout.addWidget(run_data_summary_script_button)
        buttons_layout.addWidget(clear_console_button)
        buttons_layout.addWidget(send_ad_status_button)
        buttons_layout.addWidget(extract_users_from_azure_ad_button)
        
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
                self.console.appendPlainText(output.decode('utf-8', errors='replace'))
                if error:
                    self.console.appendPlainText("Error: " + error.decode('utf-8', errors='replace'))

                # Print the output before parsing
                print("Output:", output)

                # Create a connection to the database
                conn = self.conn

                # Create the table if it doesn't exist
                self.create_table_if_not_exists1(conn)

                # Parse the output data
                data = self.parse_output1(output)

                # Insert the data into the database
                self.insert_data_into_database1(conn, data)

                # Print the data after parsing
                print("Parsed data:", data)

                # Query the database after inserting
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM ad_computers")
                rows = cursor.fetchall()
                print("Database rows:", rows)

                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running first script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")

            # If the script runs successfully, print the success message
            self.console.appendPlainText("Data saved in SQLite database successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            raise e
        
    

    def parse_output1(self, output):
        # Decode the output
        try:
            decoded_output = output.decode('utf-8')
        except UnicodeDecodeError:
            decoded_output = output.decode('cp1252')

        # Use csv.reader to split the data into columns
        reader = csv.reader(io.StringIO(decoded_output), delimiter=';')
        data = list(reader)

        # Fill missing values with None
        for row in data:
            while len(row) < 18:
                row.append(None)

        return data

    def insert_data_into_database1(self, conn, data):
        cursor = conn.cursor()
        for row in data:
            cursor.execute("""
                INSERT INTO ad_computers (Name, DNSHostName, Description, Enabled, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, Location, UserAccountControl, PasswordLastSet, WhenCreated, WhenChanged, LastLogonTimestampDT, ManagedBy, Owner, CanonicalName, DistinguishedName, AdditionalColumn) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, row)
        conn.commit()

    def create_table_if_not_exists1(self, conn):
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ad_computers (
                Name TEXT,
                DNSHostName TEXT,
                Description TEXT,
                Enabled TEXT,
                OperatingSystem TEXT,
                OperatingSystemServicePack TEXT,
                OperatingSystemVersion TEXT,
                Location TEXT,
                UserAccountControl TEXT,
                PasswordLastSet TEXT,
                WhenCreated TEXT,
                WhenChanged TEXT,
                LastLogonTimestampDT TEXT,
                ManagedBy TEXT,
                Owner TEXT,
                CanonicalName TEXT,
                DistinguishedName TEXT,
                AdditionalColumn TEXT
            )
        """)
        conn.commit()
        
    ###########################################################################################################################################################    
        
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
                self.console.appendPlainText(output.decode('utf-8', errors='replace'))
                if error:
                    self.console.appendPlainText("Error: " + error.decode('utf-8', errors='replace'))

                # Print the output before parsing
                print("Output:", output)

                # Create a connection to the database
                conn = self.conn

                # Create the table if it doesn't exist
                self.create_table_if_not_exists2(conn)

                # Parse the output data
                data = self.parse_output2(output)

                # Insert the data into the database
                self.insert_data_into_database2(conn, data)

                # Print the data after parsing
                print("Parsed data:", data)

                # Query the database after inserting
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM ad_users")
                rows = cursor.fetchall()
                print("Database rows:", rows)

                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running second script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")

            # If the script runs successfully, print the success message
            self.console.appendPlainText("Data saved in SQLite database successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            raise e

    def parse_output2(self, output):
        # Decode the output and split it into a list of values
        data = output.decode('utf-8').strip().split(',')
        
        # The number of columns in the ad_users table
        num_columns = 39

        # Fill missing values with None
        while len(data) < num_columns:
            data.append(None)
        
        return data 

    def insert_data_into_database2(self, conn, data):
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO ad_users (UserName, SamAccountName, UserPrincipalName, DisplayName, Name, Surname, Title, Description, Department, Company, Office, ManagerName, ManagerUPN, StreetAddress, City, State, PostalCode, CountryCode, Country, EmailAddress, OfficePhone, HomePhone, Mobile, FAX, EmployeeID, EmployeeNumber, HomeDirectory, HomeDrive, WhenCreated, WhenChanged, LastBadPasswordAttempt, LastLogonDate, PasswordLastSet, PasswordExpired, PasswordNeverExpires, Enabled, DistinguishedName, CanonicalName, AccountExpirationDate) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, data)
        conn.commit()

    def create_table_if_not_exists2(self, conn):
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ad_users (
                UserName TEXT,
                SamAccountName TEXT,
                UserPrincipalName TEXT,
                DisplayName TEXT,
                Name TEXT,
                Surname TEXT,
                Title TEXT,
                Description TEXT,
                Department TEXT,
                Company TEXT,
                Office TEXT,
                ManagerName TEXT,
                ManagerUPN TEXT,
                StreetAddress TEXT,
                City TEXT,
                State TEXT,
                PostalCode TEXT,
                CountryCode TEXT,
                Country TEXT,
                EmailAddress TEXT,
                OfficePhone TEXT,
                HomePhone TEXT,
                Mobile TEXT,
                FAX TEXT,
                EmployeeID TEXT,
                EmployeeNumber TEXT,
                HomeDirectory TEXT,
                HomeDrive TEXT,
                WhenCreated TEXT,
                WhenChanged TEXT,
                LastBadPasswordAttempt TEXT,
                LastLogonDate TEXT,
                PasswordLastSet TEXT,
                PasswordExpired TEXT,
                PasswordNeverExpires TEXT,
                Enabled TEXT,
                DistinguishedName TEXT,
                CanonicalName TEXT,
                AccountExpirationDate TEXT
            )
        """)
        conn.commit()
    # Define the method to run the third script
    
    ###########################################################################################################################################################

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
                self.console.appendPlainText(output.decode('utf-8', errors='replace'))
                if error:
                    self.console.appendPlainText("Error: " + error.decode('utf-8', errors='replace'))

                # Print the output before parsing
                print("Output:", output)

                # Connect to the SQLite database
                conn = self.conn

                # Create the table if it doesn't exist
                self.create_table_if_not_exists(conn)
                
                # Parse the output data
                data = self.parse_output(output)
                
                # Insert the data into the database
                self.insert_data_into_database(conn, data)

                # Print the data after parsing
                print("Parsed data:", data)

                # Query the database after inserting
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM my_table")
                rows = cursor.fetchall()
                print("Database rows:", rows)

                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running third script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")

                # If the script runs successfully, print the success message
                self.console.appendPlainText("Data saved in SQLite database successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            raise e

    def parse_output(self, output):
        # Decode the output and split it into a list of values
        data = output.decode('utf-8').strip().split(',')
        
        # Fill missing values with None
        while len(data) < 6:
            data.append(None)
        
        return data 

    def insert_data_into_database(self, conn, data):
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO my_table (Username, OS, CPU, TotalMemoryGB, DriveInfo, AdapterInfo) 
            VALUES (?, ?, ?, ?, ?, ?)
        """, data)
        conn.commit()

    def create_table_if_not_exists(self, conn):
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS my_table (
                Username TEXT,
                OS TEXT,
                CPU TEXT,
                TotalMemoryGB TEXT,
                DriveInfo TEXT,
                AdapterInfo TEXT
            )
        """)
        conn.commit()       
        
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
            
    def run_extract_users_from_azure_ad_script(self):
        self.console.appendPlainText("Running script Extract-Users-From-AzureAD...")
        try:
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                # Extract the PowerShell script to the temporary directory
                script_path = pkg_resources.resource_filename(__name__, 'scripts/Get-Extract-Users-From-AzureAD.ps1')
                temp_script_path = shutil.copy(script_path, temp_dir)

                # Run the script using subprocess
                process = subprocess.Popen(['powershell.exe', temp_script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output, error = process.communicate()

                # Write the output and errors to the console
                self.console.appendPlainText(output.decode('utf-8'))
                if error:
                    self.console.appendPlainText("Error: " + error.decode('utf-8'))

                # If the script returns a non-zero exit status, log the error and raise an exception
                if process.returncode != 0:
                    logging.error(f"Error running script: {error.decode('utf-8')}")
                    raise Exception(f"Script returned non-zero exit status {process.returncode}")

            # If the script runs successfully, print the success message
            self.console.appendPlainText("CSV file generated successfully.")
        except Exception as e:
            # If an exception occurs, log the exception and re-raise it
            logging.exception("Exception occurred: ")
            self.console.appendPlainText("Exception: " + str(e))
            raise e     
     
    def send_ad_status(self):
        self.console.appendPlainText("Sending AD status...")
        # Add your code to send the AD status here        