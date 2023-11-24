# Import necessary modules from PyQt5
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont
# Import the MainWindow class from the main_window module
from main_window import MainWindow

# Create a QApplication instance
app = QApplication([])

# Open the CSS file and set it as the application style sheet
with open('styles/styles.css', 'r') as f:
    app.setStyleSheet(f.read())

# Create a MainWindow instance and show it
window = MainWindow()
window.show()

# Start the application's event loop
app.exec_()