import os
import csv
import openpyxl
from openpyxl import Workbook
from bs4 import BeautifulSoup

# Prepare the Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Extracted Data"
ws.append(['Fiscal Year', 'Professor Name', 'Department', 'Project Name', 'Project Duration', 'Project Cost'])

# Prepare CSV output
csv_data = []
csv_data.append(['Fiscal Year', 'Professor Name', 'Department', 'Project Name', 'Project Duration', 'Project Cost'])

# Get all HTML filenames from the data folder
data_folder = 'data'
html_files = [f for f in os.listdir(data_folder) if f.endswith('.html')]

# Iterate through all HTML files in the data folder
for file_name in html_files:
    file_path = os.path.join(data_folder, file_name)
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')

    # Find all rows of the table
    rows = soup.find_all('tr', class_=['Grid_AlternatingRow', 'Grid_Row'])

    # Iterate through each row and extract the required information
    for row in rows:
        fiscal_year = row.find_all('td')[0].get_text(strip=True)
        professor_name = row.find_all('td')[1].get_text(strip=True)
        department = row.find_all('td')[2].get_text(strip=True)
        project_name = row.find('span', id=lambda x: x and 'lblAWARD_PLAN_CHI_DESCc' in x).get_text(strip=True).replace('\n', '').replace('\r', '').replace('\t', '').replace(' ', '')
        project_duration = row.find('span', id=lambda x: x and 'lblAWARD_ST_ENDc' in x).get_text(strip=True)
        project_cost = row.find('span', id=lambda x: x and 'lblAWARD_TOT_AUD_AMTc' in x).get_text(strip=True)

        # Write to Excel
        ws.append([fiscal_year, professor_name, department, project_name, project_duration, project_cost])

        # Add to CSV data
        csv_data.append([fiscal_year, professor_name, department, project_name, project_duration, project_cost])

# Set the output file name
excel_filename = "combined_extracted_data"
csv_filename = excel_filename

# Save the Excel file
wb.save(f'{excel_filename}.xlsx')

# Save the CSV file
with open(f'{csv_filename}.csv', 'w', newline='', encoding='utf-8-sig') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerows(csv_data)

print(f"Data extraction complete. Check '{excel_filename}.xlsx' and '{csv_filename}.csv' for the output.")
