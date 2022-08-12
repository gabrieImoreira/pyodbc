import pyodbc, datetime

# for driver in pyodbc.drivers():
#     if '.xlsx' in driver:
#         driver_odbc = driver

#it works
driver_odbc = 'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)'
conn_str = (r'DRIVER={'+ driver_odbc +'};'
            r'DBQ=C:\Presencial.xls;'
            r'ReadOnly=0') 
cnxn = pyodbc.connect(conn_str, autocommit=True)
crsr = cnxn.cursor()

# loop through all the tables
for worksheet in crsr.tables():
    table_name = worksheet[2]
    break

print(table_name)
#it works
table_name = 'Dados$A26070:T'
crsr.execute("SELECT * FROM [{}]".format(table_name))
# print each row of data.
counter = 0
# for row in crsr:
#     print(row)
#     counter += 1
#     if counter >=3:
#         break

# for row in crsr.columns(table=table_name):
#     print(row.column_name)

# sql = """
#it doesnt works
#     INSERT INTO [{}] ("PUC PR", `[wellen].[giovanella]`)
#     VALUES (?, ?)""".format(table_name)  
#it works
sql = """UPDATE [{}] SET CONTROLE = (?) WHERE DATA = ? """.format('Dados$A5:T')
print('\n\n',sql)
crsr.execute(sql, 'fabricio.ihsdfu', "08/08/2022")
print('\n\n',sql)
cnxn.commit()
cnxn.close()
