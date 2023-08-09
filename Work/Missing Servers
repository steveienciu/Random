# program to check the data of two spreadsheets then outputs results into another spreadsheet

import openpyxl

client = 'Client'

customer_wb = openpyxl.load_workbook(client + 'Servers.xlsx') # open (loads) the excel spreadsheet to be manipulated with 
sentinel_wb = openpyxl.load_workbook(client + 'SentinelServers.xlsx') # open (loads) the excel spreadsheet to be manipulated with 
final_wb = openpyxl.Workbook() # Workbook() used to create new workbook, name given at the end
index = 0

customer_sheet = customer_wb['Sheet1']

for i in range(1, customer_sheet.max_row + 1):
    customer_sheet['A' + str(i)] = customer_sheet.cell(row = i, column = 1).value.partition('.')[0]
    customer_sheet['A' + str(i)] = customer_sheet.cell(row = i, column = 1).value.replace(" ", "")

customer_wb.save(filename = client + 'ServersCopy.xlsx')

customer_sheet = customer_wb['Sheet1']
sentinel_sheet = sentinel_wb['Sheet1']
final_sheet = final_wb.active

for i in range(1, customer_sheet.max_row + 1):
    customer_index1 = customer_sheet.cell(row = i, column = 1).value.lower()
    #customer_index2 = customer_sheet.cell(row = i, column = 2).value
    #customer_index3 = customer_sheet.cell(row = i, column = 3).value
    #customer_index4 = customer_sheet.cell(row = i, column = 4).value
    counter = 0
    for j in range(1, sentinel_sheet.max_row + 1):
        sentinel_index = sentinel_sheet.cell(row = j, column = 1).value.lower()
        if sentinel_index != customer_index1:
            counter += 1
        if counter == sentinel_sheet.max_row:
            index += 1
            final_sheet['A' + str(index)] = customer_index1
            #final_sheet['B' + str(index)] = customer_index2
            #final_sheet['C' + str(index)] = customer_index3
            #final_sheet['D' + str(index)] = customer_index4


final_wb.save(filename = client + 'MissingServers.xlsx')
