import win32print

# Get all available printers
printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)

print("Available Printers:")
for printer in printers:
    print(printer[2])  # Printer Name

# Get default printer
default_printer = win32print.GetDefaultPrinter()
print("\nDefault Printer:", default_printer)
