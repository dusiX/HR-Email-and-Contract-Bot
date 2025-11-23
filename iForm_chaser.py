from PyQt6.QtWidgets import QApplication
import win32com.client as win32
from form import basic_form
import sys

def iForm_chaser_draft_email(window):
    """
    Create a draft Outlook email to remind a candidate to complete the mandatory iForm.
    """
    # Initialize Outlook COM object
    outlook = win32.Dispatch('outlook.application')
    
    # Create a new email item
    mail = outlook.CreateItem(0)
    
    # Set the recipient as the candidate's email
    mail.To = window.email
    
    # Set email subject
    mail.Subject = 'Action Required: Please Complete Your Pre-Employment Form'
    
    # Set plain text body (fallback for non-HTML email clients)
    mail.Body = 'Message body'
    
    # Build HTML formatted email body with personalized greeting
    mail.HTMLBody = (
        '<p>Hi ' + (window.name).split()[0] + ',</p>'  # Use first name for personalization
        '<p>Hope you are well!</p>'
        '<p>We have recently sent you an email containing an iForm that requires your attention. As a new member of our Company, it is mandatory for you to fill out this form.</p>'
        '<p>Kindly check your inbox for the subject line "Action Required: Pre-Employment Form" to locate the email.</p>'
        '<p>Thank you for your prompt attention to this matter.</p>'
        '<p>Kind regards,</p>'
    )

    # Display the draft email so the user can review and send manually
    mail.Display()


if __name__ == "__main__":
    # Initialize the PyQt application
    app = QApplication(sys.argv)
    
    # Open the basic_form to collect candidate information
    window = basic_form(caller_module="iForm_chaser")
    window.show()
    
    # Run the PyQt event loop
    app.exec()
    
    # After the form is closed, generate the draft email for the candidate
    iForm_chaser_draft_email(window)
