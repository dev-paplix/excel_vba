import re

# Read the input file (using latin-1 encoding to handle special characters)
with open('InvoiceData.txt', 'r', encoding='latin-1') as file:
    content = file.read()

# Function to add 4 to numbers beginning with 20
def add_four_to_year(match):
    year = int(match.group(0))
    return str(year + 4)

# Replace all numbers that start with 20 (years like 2018, 2019, etc.)
# This pattern matches numbers starting with 20 followed by 2 digits
updated_content = re.sub(r'\b20\d{2}\b', add_four_to_year, content)

# Write to a new file (using same encoding)
with open('InvoiceData_Updated.txt', 'w', encoding='latin-1') as file:
    file.write(updated_content)

print("File processing complete!")
print("Original file: InvoiceData.txt")
print("Updated file: InvoiceData_Updated.txt")
print("All years beginning with 20 have been increased by 4 years.")
