import pyodbc
import csv
 
conn_string = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/katak/Desktop/Data/ConexionExcel.accdb") 
 
conn = pyodbc.connect(conn_string)
 
cursor = conn.cursor()
 
cursor.execute('select * from qryControlGestion;')

#headerList = ['Ligne', 'Code_pf', 'libelle', 'base', 'Type', 'pots_uc', 'pots', 'lots', 'UC_Palet', 'Vasos_Palets', 'Composant', 'Description', 'Familia', 'Cliente']

with open('data_gest.csv','w') as f:
    writer = csv.writer(f)
    writer.writerows([i[0] for i in cursor.description])
    writer.writerows([list(i) for i in cursor.fetchall()])
    writer.writerows(cursor) 
    #df = pd.DataFrame(rows, columns=col_headers)
    #df.to_csv("test.csv", index=False)   



cursor.close()
conn.close()



#import pyodbc
#import pandas as pd
#...
#cnxn = pyodbc.connect('DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ=' + \
                      #'{};Uid={};Pwd={};'.format(db_file, user, password)

#query = "SELECT * FROM mytable WHERE INST = '796116'"
#dataf = pd.read_sql(query, cnxn)
#cnxn.close()
#------------------------------------------------------#
#import os
#import csv
#import pyodbc

# TEXT FILE CLEAN
#with open('C:\Path\To\Raw.csv', 'r') as reader, open('C:\Path\To\Clean.csv', 'w') as writer:
    #read_csv = csv.reader(reader); write_csv = csv.writer(writer, lineterminator='\n')

    #for line in read_csv:
        #if len(line[1]) > 0:            
            #write_csv.writerow(line)

# DATABASE CONNECTION
#access_path = "C:\Path\To\Access\\DB.mdb"
#con = pyodbc.connect("DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={};" \
                     #.format(access_path))

# RUN QUERY
#strSQL = "SELECT * INTO [TableName] FROM [text;HDR=Yes;FMT=Delimited(,);" + \
         #"Database=C:\Path\To\Folder].Clean.csv;"    
#cur = con.cursor()
#cur.execute(strSQL)
#con.commit()

#con.close()                            # CLOSE CONNECTION
#os.remove('C\Path\To\Clean.csv')       # DELETE CLEAN TEMP 