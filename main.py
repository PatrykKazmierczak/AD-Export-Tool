from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem
from PyQt5 import QtCore
from PyQt5.QtWidgets import QSizePolicy
from PyQt5.QtGui import QIcon
from PyQt5.QtGui import QFont
import csv
import os
import pandas as pd
import subprocess
import csv
import io

app = QApplication([])
app.setStyleSheet("""
    QMainWindow {
        background-color: #333;
    }

    QPushButton {
        color: #fff;
        background-color: #555;
        border: none;
        padding: 10px;
        min-width: 100px;
        min-height: 40px;
    }

    QPushButton:hover {
        background-color: #777;
    }

    QPushButton:pressed {
        background-color: #999;
    }

    QLineEdit {
        padding: 10px;
        border: 2px solid #555;
        border-radius: 10px;
    }
""")

class OutputWindow(QMainWindow):
    def __init__(self, output):
        super().__init__()

        self.setWindowTitle("Script Output")

        self.table = QTableWidget()
        self.setCentralWidget(self.table)

        self.show_output(output)

    def show_output(self, output):
        data = list(csv.reader(io.StringIO(output)))

        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(data[0]))

        for i, row in enumerate(data):
            for j, value in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(value))

        self.table.resizeColumnsToContents()  # Add this line

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

        export_first_script_button = QPushButton('Export First Script Results')
        export_first_script_button.clicked.connect(self.export_first_script)
        export_first_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        export_first_script_button.setFont(font)
        layout.addWidget(export_first_script_button)

        run_second_script_button = QPushButton('Run Get-ADUsersExport')
        run_second_script_button.clicked.connect(self.run_second_script)
        run_second_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_second_script_button.setFont(font)  # Set the font of the button
        layout.addWidget(run_second_script_button)

        export_second_script_button = QPushButton('Export Second Script Results')
        export_second_script_button.clicked.connect(self.export_second_script)
        export_second_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        export_second_script_button.setFont(font)
        layout.addWidget(export_second_script_button)
        
        run_third_script_button = QPushButton('Run Get-UserComputer')
        run_third_script_button.clicked.connect(self.run_third_script)
        run_third_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        run_third_script_button.setFont(font)
        layout.addWidget(run_third_script_button)

        export_third_script_button = QPushButton('Export Third Script Results')
        export_third_script_button.clicked.connect(self.export_third_script)
        export_third_script_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        export_third_script_button.setFont(font)
        layout.addWidget(export_third_script_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)


    def run_first_script(self):
            # Run the first script and get the output
            output = subprocess.check_output(['powershell.exe', './scripts/Get-ADComputersExportToSQL.ps1'])
            # Show the output in a new window
            self.output_window = OutputWindow(output.decode('utf-8'))
            self.output_window.show()
            return output.decode('utf-8')

    def run_second_script(self):
            # Run the second script and get the output
            output = subprocess.check_output(['powershell.exe', './scripts/Get-AD-UsersExportToSQLOnPremUpdate.ps1'])
            # Show the output in a new window
            self.output_window = OutputWindow(output.decode('utf-8'))
            self.output_window.show()
            return output.decode('utf-8')    

    def run_third_script(self):
            # Run the third script and get the output
            output = subprocess.check_output(['powershell.exe', './scripts/Get-UserComputerInfo.ps1'])
            # Show the output in a new window
            self.output_window = OutputWindow(output.decode('utf-8'))
            self.output_window.show()
            return output.decode('utf-8')    

    def export_first_script(self):
        try:
            if not os.path.exists('C:/test/'):
                os.makedirs('C:/test/')
            data = self.run_first_script()  # Get the data from the first script
            data = list(csv.reader(data.splitlines()))  # Parse the CSV string into a list of lists
            df = pd.DataFrame(data)  # Convert the list of lists to a pandas DataFrame
            df.to_csv('C:/test/first_script_results.csv')  # Write the DataFrame to a CSV file
        except Exception as e:
            print(f"An error occurred: {e}")

    def export_second_script(self):
        try:
            if not os.path.exists('C:/test/'):
                os.makedirs('C:/test/')
            data = self.run_second_script()  # Get the data from the second script
            data = list(csv.reader(data.splitlines()))  # Parse the CSV string into a list of lists
            df = pd.DataFrame(data)  # Convert the list of lists to a pandas DataFrame
            df.to_csv('C:/test/second_script_results.csv')  # Write the DataFrame to a CSV file
        except Exception as e:
            print(f"An error occurred: {e}")

    def export_third_script(self):
        try:
            if not os.path.exists('C:/test/'):
                os.makedirs('C:/test/')
            data = self.run_third_script()  # Get the data from the third script
            data = list(csv.reader(data.splitlines()))  # Parse the CSV string into a list of lists
            df = pd.DataFrame(data)  # Convert the list of lists to a pandas DataFrame
            df.to_csv('C:/test/third_script_results.csv')  # Write the DataFrame to a CSV file
        except Exception as e:
            print(f"An error occurred: {e}")     
        
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