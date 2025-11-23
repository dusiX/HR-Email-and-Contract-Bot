from PyQt6.QtWidgets import QFormLayout, QPushButton, QApplication, QDialog
import sys
from CN_availability import CN_availability_draft_email
from NIN_chaser import NIN_chaser_draft_email
from policies_chaser import policies_chaser_draft_email
from case_study import case_study_draft_email
from Onboarding_CN_chaser import Onboarding_CN_chaser_draft_email
from Onboarding_HM_chaser import Onboarding_HM_chaser_draft_email
from iForm_chaser import iForm_chaser_draft_email
from form import basic_form, HM_HRBP_chaser, CN_availability, case_study

class email_main_form(QDialog):
    """
    Main GUI dialog to select an email template and trigger the corresponding email draft.
    Each button opens a form to collect data and then generates an Outlook draft email.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Choose email template")

        # Create the form layout
        layout = QFormLayout()

        # Button for Candidate's availability request email
        self.CN_availability_button = QPushButton("Candidate's availability request")
        self.CN_availability_button.clicked.connect(self.open_CN_availability)
        layout.addRow(self.CN_availability_button)

        # Button for Case Study email
        self.case_study_button = QPushButton("Case Study")
        self.case_study_button.clicked.connect(self.open_case_study)
        layout.addRow(self.case_study_button)

        # Button for iForm chaser email
        self.iForm_chaser_button = QPushButton("iForm chaser")
        self.iForm_chaser_button.clicked.connect(self.open_iForm_chaser)
        layout.addRow(self.iForm_chaser_button)

        # Button for National Insurance Number chaser email
        self.NIN_chaser_button = QPushButton("National Insurance Number chaser")
        self.NIN_chaser_button.clicked.connect(self.open_NIN_chaser)
        layout.addRow(self.NIN_chaser_button)

        # Button for Onboarding Candidate's tasks chaser email
        self.new_hire_CN_chaser_button = QPushButton("Onboarding Candidate's tasks chaser")
        self.new_hire_CN_chaser_button.clicked.connect(self.open_new_hire_CN_chaser)
        layout.addRow(self.new_hire_CN_chaser_button)

        # Button for Onboarding Hiring Manager's tasks chaser email
        self.new_hire_HM_chaser_button = QPushButton("Onboarding Hiring Manager's tasks chaser")
        self.new_hire_HM_chaser_button.clicked.connect(self.open_new_hire_HM_chaser)
        layout.addRow(self.new_hire_HM_chaser_button)

        # Button for UK Policies chaser email
        self.policies_chaser_button = QPushButton("UK Policies chaser")
        self.policies_chaser_button.clicked.connect(self.open_policies_chaser)
        layout.addRow(self.policies_chaser_button)

        # Set the layout for the dialog
        self.setLayout(layout)

    # Each of the following methods handles opening the respective form,
    # collecting input, and creating a draft email via the corresponding function.

    def open_CN_availability(self):
        self.close()
        form = CN_availability()
        form.show()
        form.exec() 
        CN_availability_draft_email(form)

    def open_NIN_chaser(self):
        self.close()
        form = basic_form(caller_module="NIN_chaser")
        form.show()
        form.exec() 
        NIN_chaser_draft_email(form)

    def open_policies_chaser(self):
        self.close()
        form = HM_HRBP_chaser(caller_module="policies_chaser")
        form.show()
        form.exec() 
        policies_chaser_draft_email(form)

    def open_case_study(self):
        self.close()
        form = case_study()
        form.show()
        form.exec() 
        case_study_draft_email(form)

    def open_new_hire_CN_chaser(self):
        self.close()
        form = basic_form(caller_module="Onboarding_CN_chaser")
        form.show()
        form.exec() 
        Onboarding_CN_chaser_draft_email(form)

    def open_new_hire_HM_chaser(self):
        self.close()
        form = HM_HRBP_chaser(caller_module="Onboarding_HM_chaser")
        form.show()
        form.exec() 
        Onboarding_HM_chaser_draft_email(form)

    def open_iForm_chaser(self):
        self.close()
        form = basic_form(caller_module="iForm_chaser")
        form.show()
        form.exec() 
        iForm_chaser_draft_email(form)


if __name__ == "__main__":
    # Initialize the PyQt application
    app = QApplication(sys.argv)
    
    # Create and show the main email template selection form
    window = email_main_form()
    window.show()
    
    # Run the PyQt event loop
    app.exec()