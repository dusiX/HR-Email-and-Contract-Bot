# HR-Email-and-Contract-Bot (General Copy)

This repository contains a generalized version of Python scripts designed to automate onboarding and HR processes, including email drafting, tracker updates, and contract creation. **This code is a generalized copy of the original internal scripts and does not contain any proprietary or confidential company information.**

---

## Overview

The repository provides automation for several HR workflows:

1. **Email Templates & Automation:**

   * Candidate availability requests
   * Case study invitations
   * iForm completion reminders
   * National Insurance Number (NIN) chasers
   * Onboarding task reminders for candidates (CN) and hiring managers (HM)
   * Policies acknowledgment reminders

2. **Tracker Updates:**

   * Updates candidate onboarding progress in an Excel tracker
   * Marks tasks as chased, pending, or completed based on user input

3. **Contract Creation:**

   * Prepares employment contracts using Excel macros
   * Automatically fills Word documents with the provided employee data
   * Generates contract in PDF

4. **GUI Forms:**

   * PyQt6-based dialog forms collect input for email generation and contract creation

---

## File Structure

* `form.py`

  * Contains PyQt6 dialog classes (`basic_form`, `CN_availability`, `HM_HRBP_chaser`, `case_study`) for capturing user input.
  * Contains functions to update the Excel tracker based on the form input.

* `email_main.py`

  * Main GUI to select and trigger different email templates.
  * Opens the corresponding form, collects input, and drafts an Outlook email.

* Email template scripts:

  * `Onboarding_CN_chaser.py`
  * `Onboarding_HM_chaser.py`
  * `iForm_chaser.py`
  * `NIN_chaser.py`
  * `policies_chaser.py`
  * `CN_availability.py`
  * `case_study.py`
  * Each script contains a function to draft an Outlook email for a specific HR scenario.

* `contract_main.py`

  * GUI for creating employment contracts.
  * Collects contract details and runs an Excel macro to populate a Word document.

---

## Prerequisites

* Python 3.10 or higher
* Required packages:

  ```bash
  pip install pyqt6 openpyxl pywin32
  ```
* Windows OS (required for Outlook and Word automation using `win32com`)
* Excel macro-enabled template file for contract creation (update `macro_path` in `contract_main.py`).

---

## Usage

### Run the Main Email GUI

```bash
python email_main.py
```

* Select the desired email template.
* Fill in the form fields.
* A draft email will open in Outlook.

### Run the Contract Creator GUI

```bash
python contract_main.py
```

* Fill in the contract form.
* The Excel macro fills the Word contract with the provided information.

### Tracker Updates

* Automatically triggered when forms are saved.
* Excel workbook path for tracker updates must be configured in the scripts.

---

## Notes

* This is a general copy of the original scripts; sensitive or proprietary company data has been removed.
* Designed for Windows environments due to Outlook and Word automation.
* Ensure Outlook is properly configured for `win32com` to function.

---

## License

This project is provided as a reference and **should not be used with real employee data**.
