import sys
import os
import sqlite3

# Import necessary modules from PyQt5
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont
# Import the MainWindow class from the main_window module
from main_window import MainWindow

# Create a QApplication instance
app = QApplication([])

# Get the directory where the script (or executable) is located
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the path to the CSS file
css_path = os.path.join(base_dir, 'styles', 'styles.css')

# Open the CSS file and set it as the application style sheet
with open(css_path, 'r') as f:
    stylesheet = f.read()

# Append the text-align style to the style sheet
stylesheet += 'QPushButton { text-align: left; }'

# Set the style sheet
app.setStyleSheet(stylesheet)

# Define the directory where you want to save the database
db_dir = 'C:/ITOA'

# Ensure the directory exists
os.makedirs(db_dir, exist_ok=True)

# Define the full path to the database file
db_path = os.path.join(db_dir, 'database.db')

# Create a connection to the database
conn = sqlite3.connect(db_path)

# Create a MainWindow instance and show it
window = MainWindow(conn)
window.show()

# Start the application's event loop
app.exec_()