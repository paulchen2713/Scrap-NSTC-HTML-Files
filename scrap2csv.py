import csv
from bs4 import BeautifulSoup

# Load the HTML file
file_name = "fy107-1"
with open(f'./data/{file_name}.html', 'r', encoding='utf-8') as file:
    soup = BeautifulSoup(file, 'html.parser')

# Find all rows of the table
rows = soup.find_all('tr', class_=['Grid_AlternatingRow', 'Grid_Row'])

# Prepare the CSV output
with open(f'{file_name}.csv', 'w', newline='', encoding='utf-8-sig') as csvfile:
    fieldnames = ['Professor Name', 'Department', 'Project Name', 'Project Duration', 'Project Cost']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()

    # Iterate through each row and extract the required information
    for row in rows:
        professor_name = row.find_all('td')[1].get_text(strip=True)
        department = row.find_all('td')[2].get_text(strip=True)
        project_name = row.find('span', id=lambda x: x and 'lblAWARD_PLAN_CHI_DESCc' in x).get_text(strip=True)
        project_duration = row.find('span', id=lambda x: x and 'lblAWARD_ST_ENDc' in x).get_text(strip=True)
        project_cost = row.find('span', id=lambda x: x and 'lblAWARD_TOT_AUD_AMTc' in x).get_text(strip=True)

        # Write to CSV
        writer.writerow({
            'Professor Name': professor_name,
            'Department': department,
            'Project Name': project_name,
            'Project Duration': project_duration,
            'Project Cost': project_cost
        })

print(f"Data extraction complete. Check '{file_name}.csv' for the output.")



