from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget, QLineEdit, QPushButton, QSizePolicy
from PyQt5.QtGui import QIcon
from dashboard_window import DashboardWindow

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set window properties
        self.setWindowTitle("✨ AD-Tool ✨")
        self.setWindowIcon(QIcon('./icons/microsoft-active-directory5035.jpg'))

        # Set up layout and widgets
        layout = QVBoxLayout()
        self.username_field = QLineEdit()
        self.username_field.setPlaceholderText('Username')
        layout.addWidget(self.username_field)

        self.password_field = QLineEdit()
        self.password_field.setPlaceholderText('Password')
        self.password_field.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.password_field)

        login_button = QPushButton('Log In')
        login_button.clicked.connect(self.check_credentials)
        layout.addWidget(login_button)

        # Set layout in a container and set it as central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    # Check credentials and open dashboard if they're correct
    def check_credentials(self):
        username = self.username_field.text()
        password = self.password_field.text()

        if username == 'admin' and password == 'password':
            self.hide()
            self.dashboard_window = DashboardWindow()
            self.dashboard_window.show()
        else:
            print('Access denied')