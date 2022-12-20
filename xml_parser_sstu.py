import xml.etree.ElementTree as ET
import openpyxl

file_xml = 'guid.xml'
file_xls = 'guid.xlsx'
list_xml = []

wb = openpyxl.Workbook()
wb_s = wb.active

wb_s.append(["ID", "Name"])

root_node = ET.parse('guid.xml').getroot()

for tag in root_node.findall('Department'):
    id_value = tag.get('ID')
    if not id_value:
        id_value = 'UNKNOWN DATA'
        print(id_value)
        
    name_value = tag.get('Name')
    if not name_value:
        name_value = 'UNKNOWN DATA'
        print(name_value)
    

    wb_s.append([id_value, name_value])

wb.save(file_xls)
wb.close()

print('Done')
