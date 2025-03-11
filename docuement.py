import pandas as pd
import win32print
import win32ui

# Load Excel file
file_path = "ProvisionalListofVoters2024-2026CorporateClass.xlsx"  # Update with your actual file path
df = pd.read_excel(file_path)

# Take the first company for testing
company = df.iloc[1]
bussinuss_address = df.iloc[1]
person_name = df.iloc[1]
# Define letter template
letter_template = """\
Dear {name},

our {copmany_name} and {address} We are pleased to inform you that Anwar Hardware Store, Multan, is offering a special Eid-ul-Fitr sale on handle locks.

Thank you for being a valued customer.

Best Regards,
Anwar Hardware Store
"""

letter_content = letter_template.format(
    company_name=company["NameOfCompany/Firm"],
    address=bussinuss_address["BusinessAddress"],
    name=person_name["NameOfAuthorizedRepresentative"]
)

print(letter_content)
