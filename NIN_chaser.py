from PyQt6.QtWidgets import QApplication
import win32com.client as win32
from form import basic_form
import sys

def NIN_chaser_draft_email(window):
    """
    Create a draft Outlook email to notify a candidate about an incorrect or missing 
    National Insurance Number (NIN) on their submitted iForm.
    """
    # Initialize Outlook COM object
    outlook = win32.Dispatch('outlook.application')
    
    # Create a new email item
    mail = outlook.CreateItem(0)
    
    # Set recipient to candidate's email
    mail.To = window.email
    
    # Set the email subject
    mail.Subject = 'Information Missing'
    
    # Set plain text fallback body
    mail.Body = 'Message body'
    
    # Build HTML formatted email with personalized greeting
    mail.HTMLBody = (
        '<p class="editor-paragraph">Hi ' + (window.name).split()[0] + ','  # First name only
        '<br><br>I hope you are well!'
        '<br><br>I’ve received the iForm that you filled and I noticed that the National Insurance Number is filled incorrectly.'
        '<br><br>Could you please share with me the correct NI Number?'
        '<br><br>Kind regards,</p>'
    )

    # Display the draft email for review before sending
    mail.Display()


if __name__ == "__main__":
    # Initialize PyQt application
    app = QApplication(sys.argv)
    
    # Open the basic_form to collect candidate information
    window = basic_form(caller_module="NIN_chaser")
    window.show()
    
    # Run the PyQt event loop
    app.exec()
    
    # After the form is closed, generate the NIN chaser draft email
    NIN_chaser_draft_email(window)