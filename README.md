# Scrap-NSTC-HTML-Files
從[國科會網站](https://wsts.nstc.gov.tw/STSWeb/Award/AwardMultiQuery.aspx) (.aspx) 找清大每位教師的[國家科學及技術委員會補助研究計畫資料](https://wsts.nstc.gov.tw/STSWeb/Award/AwardMultiQuery.aspx?year=107&code=QS01&organ=A%2CFA04%2C&name=) (.html)，抓取 107-111 年度、姓名、系所、計畫名稱、執行年限、金額 等資訊整理成一個檔案。

註: 這不是爬蟲，只是從靜態 html 網頁內容把想要的資料撈出來而已，且要自己手動 Ctrl + Shift + C 把每 1 頁 (1 頁 200 筆) 共 14 頁的 html 網頁內容存下來。


### Tag examples
```htmlembedded
<tr class="Grid_Row">
    <td align="center">110</td>
    <td align="left">鍾偉和</td>
    <td align="left">國立清華大學通訊工程研究所</td>
    <td align="left">
        <span id="wUctlAwardQueryPage_grdResult_lblAWARD_PLAN_CHI_DESCc_198">運用機器學習於巨量多天線傳輸系統之設計</span><br>
        <span id="wUctlAwardQueryPage_grdResult_lblAWARD_ST_ENDc_198">2021/08/01~2024/07/31</span><br>
        <span id="wUctlAwardQueryPage_grdResult_lblAWARD_TOT_AUD_AMTc_198">3,036,000元</span>&nbsp;
```
```htmlembedded
<tr class="Grid_AlternatingRow">
    <td align="center">107</td>
    <td align="left">鍾偉和</td>
    <td align="left">國立清華大學電機工程學系(所)</td>
    <td align="left">
        <span id="wUctlAwardQueryPage_grdResult_lblAWARD_PLAN_CHI_DESCc_191">適用於智慧型物聯人聯網中之多天線系統訊號處理</span><br>
        <span id="wUctlAwardQueryPage_grdResult_lblAWARD_ST_ENDc_191">2018/08/01~2021/10/31</span><br>
        <span id="wUctlAwardQueryPage_grdResult_lblAWARD_TOT_AUD_AMTc_191">2,598,000元</span>&nbsp;
```

### Sample code
```python
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
if os.path.isfile(f'{excel_filename}.xlsx') is False:
    wb.save(f'{excel_filename}.xlsx')

# Save the CSV file
if os.path.isfile(f'{csv_filename}.csv') is False:
    with open(f'{csv_filename}.csv', 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(csv_data)

print(f"Data extraction complete. Check '{excel_filename}.xlsx' and '{csv_filename}.csv' for the output.")
```
