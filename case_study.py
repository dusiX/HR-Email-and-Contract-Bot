from PyQt6.QtWidgets import QApplication
import win32com.client as win32
from form import case_study
import sys

def case_study_draft_email(window):
    # Create an Outlook application instance through COM
    outlook = win32.Dispatch('outlook.application')

    # Create a new email item
    mail = outlook.CreateItem(0)

    # Set the recipient using the email typed in the GUI form
    mail.To = window.email

    # Build the subject using ID + Job Title
    mail.Subject = " ".join(['Case Study: ', window.id, window.job_title])

    # Plain text body (Outlook will ignore this if HTMLBody is set)
    mail.Body = 'Message body'

    # HTML content of the email
    mail.HTMLBody = (
        '<p>Hi ' + (window.name).split()[0] + ',</p>'          # Use only the first name
        '<p>Hope you’re well!</p>'
        '<p>I’m sending you the case study file attached.</p>'
        '<p>You can already start doing it as you’ll receive the feedback about it during the interview.</p>'
        '<p>Please let me know if the file opens correctly.</p>'
        '<p>Kind regards,</p>'
    )

    # Display the email window (does not send it automatically)
    mail.Display()


if __name__ == "__main__":
    # Standard PyQt application setup
    app = QApplication(sys.argv)

    # Open the case study form and wait for user input
    window = case_study()
    window.show()

    # Run the event loop and block until the form is closed
    app.exec()

    # After closing the form, generate the Outlook draft
    case_study_draft_email(window)
