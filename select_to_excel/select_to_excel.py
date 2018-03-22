
import glob
import fdb
import os
import datetime

from excel_creator import create_excel

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def connectar_unifiltro():
    return fdb.connect(
            host='localhost',
            database='DATABASE.FDB',
            user='sysdba', 
            password='masterkey'
    )
def select_to_list(cur_select):
    header = []
    for fieldDesc in cur_select.description:
        header.append(fieldDesc[fdb.DESCRIPTION_NAME])    

    data = []
    field_indexs = range(len(cur_select.description))
    for row in cur_select:
        row_data = []
        for i in field_indexs:        
            row_data.append(str(row[i]))        
        data.append(row_data)

    return [header] + data
    

con = connectar_unifiltro()    
cur = con.cursor()    
cur.execute(open("select_to_excel.sql").read())    

file_name='select_to_excel_{}.xlsx'.format(datetime.datetime.now().strftime('%Y.%m.%d.%H.%M'))
select_data = select_to_list(cur)
create_excel(file_name, select_data)

print("Arquivo criado: " + file_name)