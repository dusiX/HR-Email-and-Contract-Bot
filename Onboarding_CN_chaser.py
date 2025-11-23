from PyQt6.QtWidgets import QApplication
import win32com.client as win32
from form import basic_form
import sys

def Onboarding_CN_chaser_draft_email(window):
    """
    Create a draft email in Outlook for a candidate to complete 
    onboarding tasks in the company's onboarding platform.
    """
    # Initialize Outlook COM object
    outlook = win32.Dispatch('outlook.application')
    
    # Create a new mail item
    mail = outlook.CreateItem(0)
    
    # Set recipient
    mail.To = window.email
    
    # Set email subject
    mail.Subject = 'ACTION NEEDED: Your onboarding process'
    
    # Set plain text body (required by Outlook, but HTMLBody will override display)
    mail.Body = 'Message body'
    
    # Build HTML formatted email body
    mail.HTMLBody = (
        '<p>Hi ' + (window.name).split()[0] + ',</p>'  # Greet using first name
        '<p>We kindly request your immediate attention to the following tasks within our Onboarding platform, as outlined in our previous email:</p>'
        '<ol>'
        '<li>Upload your Photo.</li>'
        '<li>Submit Required Documents.</li>'
        '<li>Other required tasks.</li>'
        '</ol>'
        '<p>Timely completion of these tasks is essential for the successful processing of your hiring. Your prompt action in this matter is greatly appreciated.</p>'
        '<p>Feel free to reach out to me in case of any problems.</p>'
        '<p>Kind regards,</p>'
    )

    # Display the draft email to allow user review before sending
    mail.Display()


if __name__ == "__main__":
    # Standard PyQt application launcher
    app = QApplication(sys.argv)
    
    # Open the basic_form dialog to collect candidate name and email
    window = basic_form(caller_module="Onboarding_CN_chaser")
    window.show()
    
    # Execute the PyQt event loop
    app.exec()
    
    # After form is closed, generate the draft onboarding email
    Onboarding_CN_chaser_draft_email(window)
