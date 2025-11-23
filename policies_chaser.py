from PyQt6.QtWidgets import QApplication
import win32com.client as win32
from form import HM_HRBP_chaser
import sys

def policies_chaser_draft_email(window):
    """
    Create a draft Outlook email to the Hiring Manager (and CC HRBP) reminding them
    to ensure the new hire completes mandatory onboarding policy acknowledgements.
    """
    # Initialize Outlook COM object
    outlook = win32.Dispatch('outlook.application')
    
    # Create a new email item
    mail = outlook.CreateItem(0)
    
    # Set recipient (Hiring Manager) and CC (HRBP)
    mail.To = window.HM_email
    mail.CC = window.HRBP_email
    
    # Set the subject of the email
    mail.Subject = 'Urgent Action Required: Acknowledgement Onboarding Policies'
    
    # Plain text fallback for the email
    mail.Body = 'Message body'
    
    # Build HTML formatted email body with personalization
    mail.HTMLBody = (
        '<p class="editor-paragraph">Hi ' + (window.HM_name).split()[0] + ','  # First name only
        '<br><br>I trust this message finds you well.'
        '<br><br>All external hires are requested to acknowledge various policies during their onboarding. Once signed, these documents are then saved electronically for legal compliance.'
        '<br><br>Despite numerous reminders, your new hire ' + window.CN_name + ', who commenced on ' + window.start_date + ', has still not completed the mandatory tasks within their onboarding journey.'
        '<br><br>Please can I kindly request that you encourage them to re-visit their onboarding case and ensure that all tasks are completed <strong>as soon as possible</strong>?'
        '<br><br>If there are any challenges or queries, please reach out for further assistance.'
        '<br><br>Thank you for your prompt attention to this matter.'
        '<br><br>Kind regards,</p>'
    )

    # Display the draft email for review before sending
    mail.Display()


if __name__ == "__main__":
    # Initialize PyQt application
    app = QApplication(sys.argv)
    
    # Open the HM_HRBP_chaser form to collect Hiring Manager and candidate information
    window = HM_HRBP_chaser(caller_module="policies_chaser")
    window.show()
    
    # Run the PyQt event loop
    app.exec()
    
    # After the form is closed, generate the policies chaser draft email
    policies_chaser_draft_email(window)
