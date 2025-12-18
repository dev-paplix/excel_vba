import csv

# Read the input CSV file
input_file = 'pivot_table.csv'
output_file = 'pivot_table_transformed.csv'

# Define all the replacements
replacements = {
    'America': 'Malaysia',
    'Europe': 'Singapore',
    '1010US': '1010MY',
    '1020US': '1020MY',
    '1040DE': '1040SG',
    '/2020': '/2025',
    'Women type T simple white': 'White Ink',
    'Women type T simple black': 'Black Ink',
    'Women crop top black': 'Blue Ink',
    'Women basics': 'Red Ink',
    'Men basics': 'Green Ink',
    'Men type T simple white': 'Purple Ink',
    'Laptop bag black': 'Grey Ink',
    'Men dress shirt black': 'Yellow Ink',
    'Men dress shirt grey': 'Violet Ink',
    'Men shorts grey': 'Brown Ink',
    'Men shorts black': 'Orange Ink',
    'Unisex tank top white': 'White Ink',
    'Laptop bag red': 'Emerald Ink',
    'Smartphone case diamond': 'Sky Blue Ink',
    'Smartphone case simple': 'Silver Ink',
    'Men type T simple black': 'Gold Ink',
    ' GmbH': ''
}

# Read and transform the data
with open(input_file, 'r', encoding='utf-8') as infile, \
     open(output_file, 'w', encoding='utf-8', newline='') as outfile:
    
    reader = csv.reader(infile)
    writer = csv.writer(outfile)
    
    for row in reader:
        # Apply replacements to each cell in the row
        transformed_row = []
        for cell in row:
            for old_value, new_value in replacements.items():
                cell = cell.replace(old_value, new_value)
            transformed_row.append(cell)
        
        writer.writerow(transformed_row)

print(f"Transformation complete! Output saved to {output_file}")
