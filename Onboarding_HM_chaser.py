from PyQt6.QtWidgets import QApplication
import win32com.client as win32
from form import HM_HRBP_chaser
import sys

def Onboarding_HM_chaser_draft_email(window):
    """
    Create a draft email in Outlook for a Hiring Manager (HM) to complete 
    onboarding tasks for a new hire, with HRBP copied.
    """
    # Initialize Outlook COM object
    outlook = win32.Dispatch('outlook.application')
    
    # Create a new email item
    mail = outlook.CreateItem(0)
    
    # Set primary recipient as the Hiring Manager
    mail.To = window.HM_email
    
    # CC the HR Business Partner for visibility
    mail.CC = window.HRBP_email
    
    # Set email subject including candidate's name
    mail.Subject = " ".join(['ACTION NEEDED:', window.CN_name, 'onboarding process'])
    
    # Set plain text body (required, though HTMLBody will display)
    mail.Body = 'Message body'
    
    # Build HTML formatted email body
    mail.HTMLBody = (
        '<p>Hi ' + (window.HM_name).split()[0] + ',</p>'  # Greet HM by first name
        '<p>Hope you are well!</p>'
        '<p>I kindly request that you complete all the tasks assigned to you in Onboarding platform, considering our new hire starting on ' + window.start_date + '.</p>'
        '<p>This is essential for a seamless and positive candidate onboarding experience.</p>'
        '<p>Your prompt attention to these tasks will greatly facilitate our onboarding process.</p>'
        '<p>Kind regards,</p>'
    )

    # Display the draft email so user can review before sending
    mail.Display()


if __name__ == "__main__":
    # Initialize the PyQt application
    app = QApplication(sys.argv)
    
    # Open the HM_HRBP_chaser form to collect Hiring Manager, HRBP, and candidate info
    window = HM_HRBP_chaser(caller_module="Onboarding_HM_chaser")
    window.show()
    
    # Run the PyQt event loop
    app.exec()
    
    # After form is closed, generate the draft email for HM
    Onboarding_HM_chaser_draft_email(window)
