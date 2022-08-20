import csv
import os
import pyodbc

# Rig db Params/Vars
# Rigs data csv file path and name.
filePath = os.getcwd() + '/'
fileName = 'export.csv'
...

# SQL to select data from the rigs table.
rigs_export_sql = "SELECT * FROM qryControlGestion;"

def export_rigs_data():
    
    # Database connection variable.
    connect = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/katak/Desktop/Data/ConexionExcel.accdb") 
    conn = pyodbc.connect(connect)
    cursor = conn.cursor()
    # Cursor to execute query.
   
    
    # Execute query.
    cursor.execute(rigs_export_sql)
    
    # Fetch the data returned.
    results = cursor.fetchall()
    
    # Extract the table headers.
    headers = [i[0] for i in cursor.description]
    
    # Open CSV file for writing.
    #csvFile = csv.writer(open(filePath + fileName, 'w', newline=''),delimiter=',', lineterminator='',quoting=csv.QUOTE_ALL, escapechar='/')
    with open('data_gest.csv','w', newline='') as f:
            writer = csv.writer(f)
    # Add the headers and data to the CSV file.
            writer.writerow(headers)
            writer.writerows(results)

if __name__ == "__main__":
    export_rigs_data()

# Database connection variable.
    #connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +
                                     #server+';DATABASE='+database+';UID='+username+';PWD=' + password)