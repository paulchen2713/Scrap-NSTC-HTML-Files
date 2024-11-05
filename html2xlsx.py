import os
import openpyxl
from openpyxl import Workbook
from bs4 import BeautifulSoup

# Prepare the Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Extracted Data"
ws.append(['Fiscal Year', 'Professor Name', 'Department', 'Project Name', 'Project Duration', 'Project Cost'])

# Iterate through all HTML files in the data folder
for i in range(1, 15):
    file_path = os.path.join('data', f'{i}.html')
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')

    # Find all rows of the table
    rows = soup.find_all('tr', class_=['Grid_AlternatingRow', 'Grid_Row'])

    # Iterate through each row and extract the required information
    for row in rows:
        fiscal_year = row.find_all('td')[0].get_text(strip=True)
        professor_name = row.find_all('td')[1].get_text(strip=True)
        department = row.find_all('td')[2].get_text(strip=True)
        project_name = row.find('span', id=lambda x: x and 'lblAWARD_PLAN_CHI_DESCc' in x).get_text(strip=True)
        project_duration = row.find('span', id=lambda x: x and 'lblAWARD_ST_ENDc' in x).get_text(strip=True)
        project_cost = row.find('span', id=lambda x: x and 'lblAWARD_TOT_AUD_AMTc' in x).get_text(strip=True)

        # Write to Excel
        ws.append([fiscal_year, professor_name, department, project_name, project_duration, project_cost])

# Save the Excel file
wb.save('combined_extracted_data.xlsx')

print("Data extraction complete. Check 'combined_extracted_data.xlsx' for the output.")
