import win32print
import win32ui

# Set your printer name (Use the exact name from Step 1 output)
printer_name = win32print.GetDefaultPrinter()

# Open the printer
hprinter = win32print.OpenPrinter(printer_name)
printer_info = win32print.GetPrinter(hprinter, 2)

# Create Printer Device Context
pdc = win32ui.CreateDC()
pdc.CreatePrinterDC(printer_name)

# Start the print job
pdc.StartDoc("Test_Print")
pdc.StartPage()
pdc.TextOut(100, 100, "This is a test print from Python.")
pdc.EndPage()
pdc.EndDoc()

# Cleanup
pdc.DeleteDC()
win32print.ClosePrinter(hprinter)

print("Test print sent successfully.")
