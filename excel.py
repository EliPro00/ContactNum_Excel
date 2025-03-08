from openpyxl import Workbook , load_workbook 

wb = load_workbook ('test.xlsx')
ws = wb.active
new_row = {'A' : 'sas_cust' , 'b' : '+971666666'}
ws.append(new_row)
wb.save('test.xlsx')
