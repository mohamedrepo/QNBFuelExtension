import re

filepath = r'C:\QNBFuelExtension\QNB_Fuel_3Sheets_01072023_to_30062026 (2).xls'
output = r'C:\QNBFuelExtension\QNB_Fuel_3Sheets_01072023_to_30062026_fixed.xls'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = re.sub(r'\"\}+>', '">', content)
content = re.sub(r'\}\}>', '}>', content)

with open(output, 'w', encoding='utf-8') as f:
    f.write(content)

print('File fixed and saved as QNB_Fuel_3Sheets_01072023_to_30062026_fixed.xls')