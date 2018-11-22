import xlrd

print("get data from excel and save to txt")

# open excel
workbook = xlrd.open_workbook('Katalogi.xls')
# get names of all excel sheets
sheetnames = workbook.sheet_names()
# create txt to write data
file = open("data.txt","w")
# array to keep all data
items = []

# iterable all excel sheets
for val in sheetnames:
	# get data from sheets
    worksheet = workbook.sheet_by_name(val)
    # iterable data from sheet
    for row in range(1, worksheet.nrows):
        # keep data from all rows
        values = []
        # if cell 3 isnt empty get data from all cell from row
        if worksheet.cell(row, 2).value != '':
        	# iterable all cells from row
            for col in range(0, worksheet.ncols):
            	# get data from cell
                wari = worksheet.cell(row, col).value
                # if cell isnt the last from row
                if(col != worksheet.ncols-1):
                	# save to file
                    file.write(str(wari) + ", ")
                else:
                    file.write(str(wari))
                # add value to array
                values.append(wari)
            # next line in text file
            file.write("\n")
            # add all rows to all data
            items.append(values)

# for item in items:
# 	print(item)

print("finish")

# close write file
file.close() 
