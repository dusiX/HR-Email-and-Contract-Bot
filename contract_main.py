from PyQt6.QtWidgets import QFormLayout, QPushButton, QApplication, QComboBox, QLineEdit, QDialog
import win32com.client as win32
import sys, time

# Path to the Excel macro-enabled workbook responsible for generating contracts
# NOTE: Replace {Username} with an actual user or load dynamically from config.
macro_path = r"C:\Users\{Username}\Path_to_macro_file"


class contract_main(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Prepare contract")
        layout = QFormLayout()

        # Text input fields for contract data
        self.employee_name_input = QLineEdit()
        self.job_title_input = QLineEdit()
        self.start_date_input = QLineEdit()
        self.end_date_input = QLineEdit()
        self.hours_input = QLineEdit()
        self.location_input = QLineEdit()
        self.salary_input = QLineEdit()
        self.shift_pattern_input = QLineEdit()
        self.shift_allowance_input = QLineEdit()
        self.disruption_allowance_input = QLineEdit()

        # Legal entity dropdown
        self.legal_entity_input = QComboBox()
        self.legal_entity_input.addItems(["Company 1", "Company 2", "Company 3"])
        layout.addRow("Legal Entity:", self.legal_entity_input)

        # Candidate and role details
        layout.addRow("Candidate's name:", self.employee_name_input)
        layout.addRow("Job title:", self.job_title_input)

        # Job level dropdown
        self.level_input = QComboBox()
        self.level_input.addItems([str(i) for i in range(1, 12)])
        layout.addRow("Level:", self.level_input)

        # Dates
        layout.addRow("Start date:", self.start_date_input)
        layout.addRow("End date:", self.end_date_input)

        # Part-time selection dropdown
        self.part_time_input = QComboBox()
        self.part_time_input.addItems(["No", "Yes"])
        layout.addRow("Part time:", self.part_time_input)

        # Additional job details
        layout.addRow("Weekly hours:", self.hours_input)
        layout.addRow("Location:", self.location_input)
        layout.addRow("Salary:", self.salary_input)

        # Holidays dropdown
        self.holidays_input = QComboBox()
        self.holidays_input.addItems(["25", "23"])
        layout.addRow("Holidays:", self.holidays_input)

        # Business car dropdown
        self.car_input = QComboBox()
        self.car_input.addItems(["No", "Yes"])
        layout.addRow("Business needs car:", self.car_input)

        # Field Sales dropdown
        self.FS_input = QComboBox()
        self.FS_input.addItems(["No", "Yes"])
        layout.addRow("Field Sales:", self.FS_input)

        # Shift details
        layout.addRow("Shift pattern:", self.shift_pattern_input)
        layout.addRow("Shift allowance:", self.shift_allowance_input)
        layout.addRow("Disruption allowance:", self.disruption_allowance_input)

        # Proceed button triggers contract creation
        self.proceed = QPushButton("Proceed")
        self.proceed.clicked.connect(self.create_contract)
        layout.addRow(self.proceed)

        self.setLayout(layout)

    def create_contract(self):
        # Close the current dialog before running automation
        self.close()

        # Open Excel in background using COM automation
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        # Open the macro-enabled workbook
        workbook = excel.Workbooks.Open(macro_path)
        sheet = workbook.Sheets("MAIN")

        # Populate required cells in the Excel contract template
        sheet.Range("C2").Value = self.legal_entity_input.currentText()
        sheet.Range("C3").Value = self.employee_name_input.text()
        sheet.Range("C4").Value = self.job_title_input.text()
        sheet.Range("C5").Value = self.level_input.currentText()
        sheet.Range("C6").Value = self.start_date_input.text()
        sheet.Range("C7").Value = self.start_date_input.text()
        sheet.Range("C8").Value = self.end_date_input.text() if self.end_date_input.text() else None
        sheet.Range("C9").Value = self.part_time_input.currentText() if self.part_time_input.currentText() == "Yes" else None
        sheet.Range("C10").Value = self.hours_input.text()
        sheet.Range("C11").Value = self.location_input.text()
        sheet.Range("C12").Value = int(self.salary_input.text())
        sheet.Range("C13").Value = int(self.holidays_input.currentText())
        sheet.Range("C14").Value = self.car_input.currentText() if self.car_input.currentText() == "Yes" else None
        sheet.Range("C15").Value = self.FS_input.currentText() if self.FS_input.currentText() == "Yes" else None
        sheet.Range("C16").Value = self.shift_pattern_input.text() if self.shift_pattern_input.text() else None
        sheet.Range("C17").Value = self.shift_allowance_input.text() if self.shift_allowance_input.text() else None
        sheet.Range("C18").Value = self.disruption_allowance_input.text() if self.disruption_allowance_input.text() else None

        # Macro inside Excel produces a Word document with a complete contract
        macro_name = "Run"
        excel.Application.Run(macro_name)

        # After macro finishes, copy the generated contract from Word
        self.copy_contract()

        # Cleanup
        workbook.Close()
        excel.Quit()

    def copy_contract(self):
        print("copy contract working")

        # The macro generates a Word document.
        # We must wait until Word opens and loads the document.
        max_wait = 10
        wait_time = 0
        word = None

        # Poll for Word application being ready with a document
        while wait_time < max_wait:
            try:
                word = win32.GetActiveObject("Word.Application")
                if word.Documents.Count > 0:
                    break
            except Exception:
                pass

            time.sleep(0.5)
            wait_time += 0.5

        # If no Word document was detected, stop
        if not word or word.Documents.Count == 0:
            print("Couldn't find open Word document.")
            return

        # Get the first (newly generated) Word document
        doc = word.Documents(1)

        # Copy full contract content to clipboard
        doc.Content.Copy()

        # Disable Word alerts before closing
        word.DisplayAlerts = 0

        # Close without saving (macro already produced final output)
        doc.Close(SaveChanges=False)

        # Quit Word session
        word.Quit()


if __name__ == "__main__":
    # Standard PyQt application launcher
    app = QApplication(sys.argv)
    window = contract_main()
    window.show()
    app.exec()