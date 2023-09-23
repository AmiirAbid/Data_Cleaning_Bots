# Import the openpyxl library for xlsx files management
import openpyxl

# Open the file where you have the data to manipulate
wb = openpyxl.load_workbook('../Data sheets/ORIGINAL DATA.xlsx')
# Open the desired worksheet in the xlsx file
sheet = wb['Sheet1']

# Create a new workbook for Lands
wb1 = openpyxl.Workbook()
ws1 = wb1.active

# Create a new workbook for Houses
wb2 = openpyxl.Workbook()
ws2 = wb2.active

# Make the header for each column
for j in range(1,8) :
    ws1.cell(row = 1, column = j).value = ws2.cell(row = 1, column = j).value = sheet.cell(row = 1, column = j).value

# Initialize a counter on the second row, skipping the table headers
i=2

# Loop through the rows
while i < sheet.max_row+1 :
    
    # Acess the cell in the i row, A column
    cell = sheet.cell(row = i, column = 1)

    # Make a wordlist for keywords that indicates that the property type is a land
    l = ['terrain','TERRAIN','Terrain','hectare','Hectare','HECTARE','lot','Lot','LOT','lotissement','Lotissement','LOTISSEMENT','أرض','ارض']
    
    # Make a wordlist for keywords that indicates that the property type is not a house
    h = ['depot','dépot','batiment','ferme','inachevé','fond','commerce','commercial','bureau','garage','parcelle','usine']
    
    # Check if the cell contains a keyword from the land wordlist
    if (cell.value != None) and any(keyword in cell.value for keyword in l) :

        # Print the accessed cell
        print(i,'LAND: ',cell.value)

        # Add this row to the lands file
        ws1.append([cell.value for cell in sheet[i]])
    
    elif (cell.value != None) and (keyword not in cell.value for keyword in h) :
        
        # Print the accessed cell
        print(i,'HOUSE: ',cell.value)

        # Add this row to the houses file
        ws2.append([cell.value for cell in sheet[i]])

    i+=1

# Save the output workbooks
print('-------------------------Saving files-------------------------')
wb1.save('../Data sheets/LANDS.xlsx')
wb2.save('../Data sheets/HOUSES.xlsx')