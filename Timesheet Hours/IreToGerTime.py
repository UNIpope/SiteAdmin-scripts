import os 
import openpyxl as xl

pathi = r"C:\Users\jduggan\Documents\A hole year ivana\year"
patho = r"C:\Users\jduggan\Documents\A hole year ivana\fake"

# edit hours and save as new file 
for f in os.listdir( pathi):
    ogfile = pathi + "\\" + f

    # read in el sheet
    workbook = xl.load_workbook(filename=ogfile)
    sheet = workbook.active
    
    pos = 16
    while sheet["S"+str(pos)].value != None:

        if sheet["S"+str(pos)].value > 48:
            diff =  sheet["S"+str(pos)].value - 48
            sheet["S"+str(pos)].value = sheet["S"+str(pos)].value - diff
            sheet["R"+str(pos)].value = sheet["R"+str(pos)].value - diff

        pos += 2

    workbook.save(patho+ "\\" + f)
