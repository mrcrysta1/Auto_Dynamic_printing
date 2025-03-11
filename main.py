import pandas as pd
import win32print
import win32ui

# Load Excel file
file_path = "company_data.xlsx"  # Update with your actual file path
df = pd.read_excel(file_path)

# Define the fixed letter content
letter_template = """\
Dear {company_name},

We are pleased to inform you that Anwar Hardware Store, Multan, is offering a special Eid-ul-Fitr sale on handle locks.

Thank you for being a valued customer.

Best Regards,
Anwar Hardware Store
"""

# Get default printer
printer_name = win32print.GetDefaultPrinter()
hprinter = win32print.OpenPrinter(printer_name)
printer_info = win32print.GetPrinter(hprinter, 2)
pdc = win32ui.CreateDC()
pdc.CreatePrinterDC(printer_name)

# Iterate over each company and send print jobs
for index, row in df.iterrows():
    letter_content = letter_template.format(
        company_name=row["company_name"]
    )
    
    # Start the print job
    pdc.StartDoc(f"Letter_{index+1}")
    pdc.StartPage()
    pdc.TextOut(100, 100, letter_content)  # Positioning text on the page
    pdc.EndPage()
    pdc.EndDoc()

# Cleanup
pdc.DeleteDC()
win32print.ClosePrinter(hprinter)

print("All letters have been sent to the printer.")
