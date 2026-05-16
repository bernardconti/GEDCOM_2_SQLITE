#pip install excelsql

import sqlite3
import openpyxl 
import os

# Path to the Excel file
icloud = "/Users/bernardconti/Library/Mobile Documents/com~apple~CloudDocs"
geographie_excel = icloud+'/MesProgrammes/Geography_data/geo_data.xlsx'
geographie_sqlite = icloud+'/MesProgrammes/Geography_data/geo_data.db'

# Open workbook
if not os.path.isfile(geographie_excel): 
    print(geographie_excel,"not found")
    exit()

#open workbook
wb = openpyxl.load_workbook(geographie_excel)

#Create sql db and table
if os.path.isfile(geographie_sqlite): os.remove(geographie_sqlite)

connection_obj = sqlite3.connect(geographie_sqlite)
print(f"Opened SQLite database {geographie_sqlite} with version {sqlite3.sqlite_version} successfully.")
sql = connection_obj.cursor()

#iterate on sheets (one TABLE per sheet)

print(f'Processing workbook {geographie_excel}')
for wsname in wb.sheetnames:
    print(f' --> Processing {wsname}')
    ws = wb[wsname]
    le_max_col = ws.max_column
    le_max_row = ws.max_row
    if le_max_row < 2 or le_max_col < 1:
        print("not processed",wsname,le_max_row,"max_row",le_max_col,"max_col")
    else:
        
        first_row = ws[1]
        le_select_field_temp = []
        le_select_row_field_temp = []
        for cell in first_row:
            le_field = cell.value.lower()
            le_field=le_field.replace(" ","_")
            le_field=le_field.replace("(","")
            le_field=le_field.replace(")","")
            le_select_field_temp.append(f'{le_field} TEXT')
            le_select_row_field_temp.append(f'{le_field}')
        
        le_select_row_field = ",".join(le_select_row_field_temp)

        print(f'------>{le_select_row_field}')

        sql.execute(f'DROP TABLE IF EXISTS {wsname}')
        sql.execute(f'CREATE TABLE {wsname} ({",".join(le_select_field_temp)});')

        for idx,la_row in enumerate(ws.rows):
            if idx > 0 :
                le_select_row_value_temp = []
                #'INSERT INTO FAMS (indi_id,fam_id) VALUES (?,?);'
                for cell in la_row:
                    le_select_row_value_temp.append(f'"{cell.value}"')
                le_select_row_value = ",".join(le_select_row_value_temp)
                
                le_select_row = f'INSERT INTO {wsname} ({le_select_row_field}) VALUES ({le_select_row_value});'
                #print(le_select_row)
                    
                sql.execute(le_select_row)
        print(f'{idx-1} rows created')

    #import rows



# Close all
wb.close()
connection_obj.commit()
connection_obj.close()