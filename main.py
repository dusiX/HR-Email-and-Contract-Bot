from PyQt6.QtWidgets import QDialog, QFormLayout, QPushButton, QApplication
import sys
from email_main import email_main_form
from contract_main import contract_main


class main_form(QDialog):
    """
    Main menu dialog for the application.
    This window serves as the entry point and allows the user to choose
    between the email automation module and the contract creation module.
    """

    def __init__(self):
        super().__init__()

        # Main layout for the menu window (simple vertical form layout)
        layout = QFormLayout()

        # --------------------------------------------------------
        # EMAIL MODULE BUTTON
        # --------------------------------------------------------
        self.email_main = QPushButton("Email templates")
        # When clicked, open the email templates window
        self.email_main.clicked.connect(self.email_templates)
        layout.addRow(self.email_main)

        # --------------------------------------------------------
        # CONTRACT MODULE BUTTON
        # --------------------------------------------------------
        self.contract_main = QPushButton("Create contract")
        # When clicked, open the contract creation window
        self.contract_main.clicked.connect(self.create_contract)
        layout.addRow(self.contract_main)

        # Apply layout to the main window
        self.setLayout(layout)

    # ------------------------------------------------------------
    # BUTTON ACTIONS
    # ------------------------------------------------------------

    def email_templates(self):
        """
        Called when the user selects the email templates option.
        Closes the main menu and opens the email module window.
        """
        self.close()
        form = email_main_form()
        form.exec()

    def create_contract(self):
        """
        Called when the user selects the contract creation option.
        Closes the main menu and opens the contract module window.
        """
        self.close()
        form = contract_main()
        form.exec()


# ------------------------------------------------------------
# APPLICATION ENTRY POINT
# ------------------------------------------------------------

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Create and show the main application menu
    window = main_form()
    window.show()

    # Start the PyQt event loop
    app.exec()
