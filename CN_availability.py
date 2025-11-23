from PyQt6.QtWidgets import QApplication
import win32com.client as win32
from form import CN_availability
import sys

def CN_availability_draft_email(window):
    # Create Outlook application instance via COM automation
    outlook = win32.Dispatch('outlook.application')

    # Create a new email item
    mail = outlook.CreateItem(0)

    # Set recipient from the GUI form
    mail.To = window.email

    # Email subject containing candidate name
    mail.Subject = " ".join(['Your Interview at Example Company -', window.name])

    # Plain-text fallback body
    mail.Body = 'Message body'

    # Check interview format selected by the user
    # If the interview is face-to-face:
    if window.ivformat == 'F2F':

        # Build HTML email for face-to-face interview
        mail.HTMLBody = (
            '<p class="editor-paragraph">Hi ' + (window.name).split()[0] + ','
            '<br><br>I am delighted to inform you that you have been selected to take part in an interview for our role '
            + window.id + ' ' + window.job_title + '.'  
            '<br><br>The interview will take ' + window.duration + '.'
            '<br><br>Format: Face to face'
            '<br><br>Address: ' + window.address +
            '<br><br>I am in the process of scheduling your interview. Could you please share your availability for '
            + window.slot + '.'
            '<br><br>Once you reply I will send an official invite.'
            '<br><br>Kind regards,</p>'
        )

    else:
        # Build HTML email for online interview (Teams meeting)
        mail.HTMLBody = (
            '<p class="editor-paragraph">Hi ' + (window.name).split()[0] + ','
            '<br><br>I am delighted to inform you that you have been selected to take part in an interview for our role '
            + window.id + ' ' + window.job_title + '.'  
            '<br><br>The interview will take ' + window.duration + '.'
            '<br><br>Format: Teams Meeting'
            '<br><br>I am in the process of scheduling your interview. Could you please share your availability for '
            + window.slot + '.'
            '<br><br>Once you reply I will send an official invite.'
            '<br><br>Kind regards,</p>'
        )

    # Open email window for review (does not send automatically)
    mail.Display()


if __name__ == "__main__":
    # PyQt application setup
    app = QApplication(sys.argv)

    # Launch the availability input form
    window = CN_availability()
    window.show()

    # Execute event loop (waiting for user input)
    app.exec()

    # After the form closes, generate Outlook draft email
    CN_availability_draft_email(window)
