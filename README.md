# AD-Export-Tool

This tool is a PyQt5 application that allows users to run several PowerShell scripts related to Active Directory. The scripts are bundled with the application and are run in a temporary directory.

## Features

- Login form with hardcoded credentials.
- Dashboard with buttons to run each script.
- Error handling and logging for each script.

## Usage

1. Run the application.
2. Enter the username and password in the login form.
3. Click on the buttons in the dashboard to run the scripts.

## Requirements

- Python 3
- PyQt5
- PowerShell

## Installation

1. Clone the repository.
2. Install the requirements with `pip install -r requirements.txt`.
3. Run `python main.py` to start the application.

pyinstaller --onefile --add-data 'scripts/Get-ADComputersExportToSQL.ps1;scripts/' --add-data 'scripts/Get-ADUsersExportToSQLOnPremUpdate.ps1;scripts/' --add-data 'scripts/Get-Extract-Users-From-AzureAD.ps1;scripts/' --add-data 'scripts/Get-UserComputerInfo.ps1;scripts/' --add-data 'scripts/Send-Email.ps1;scripts/' main.py
pyinstaller --onefile --add-data "scripts/*;scripts/" --add-data "styles/*;styles/" main.py