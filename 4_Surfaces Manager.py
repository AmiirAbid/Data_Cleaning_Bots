# Import the openpyxl library for xlsx files management
import openpyxl

# Open the file where you have the data to manipulate
wb = openpyxl.load_workbook('../Data sheets/Location/APPARTEMENTS S+2.xlsx')

# Open the desired worksheet in the xlsx file
sheet = wb['Sheet']

# Define the function that extracts the surface from the details
def surface_extractor(surface):
    # Initialize a new string
    s = ''

    # Iterate through each character in reverse order
    for char in reversed(surface):
        # Check if the character is numeric or a comma
        if char.isnumeric() or char == ',':
            # Prepend the character to the new string (to maintain the correct order)
            s = char + s
        else:
            # Stop the loop when a non-numeric character is encountered
            break

    # Return the new string
    return s

# Loop through the rows, starting from the second row (skipping headers)
for i in range(2, sheet.max_row + 1):
    print(sheet.cell(row=i, column=1).value)

    # Access the cells in the i-th row
    cellD = sheet.cell(row=i, column=4)
    cellE = sheet.cell(row=i, column=5)
    cellF = sheet.cell(row=i, column=6)

    # Check if there is no "Surface Habitable" and there is in the description
    if (cellD.value is None or "Surf" not in cellD.value) and (cellF.value is not None and 'Surf habitable' in cellF.value):
        # Split the details into 2 strings with "m² Surf habitable" as a separator
        surface = cellF.value.split("m² Surf habitable")[0]
        cellD.value = surface

    # Check if there is no "Surface Terrain" and there is in the description    
    if (cellE.value is None or "Surf" not in cellE.value) and (cellF.value is not None and 'Surf terrain' in cellF.value):
        # Split the details into 2 strings with "m² Surf terrain" as a separator
        surface = cellF.value.split("m² Surf terrain")[0]
        cellE.value = surface

    if cellD.value is not None and 'm²' in cellD.value:
        # Extract the exact surface without any alphabetic characters
        cellD.value = surface_extractor(cellD.value.split('m²')[0])
        print(cellD.value)

    if cellE.value is not None and 'm²' in cellE.value:
        # Extract the exact surface without any alphabetic characters
        cellE.value = surface_extractor(cellE.value.split('m²')[0])
        print(cellE.value)

    # Check if the "surface habitable" and "surface terrain" are swapped
    if (cellD.value is not None and cellE.value is not None) and cellD.value and cellE.value and cellE.value < cellD.value:
        # Swap the values
        cellD.value, cellE.value = cellE.value, cellD.value
        print(cellD.value,'/',cellE.value)

# Save the output workbooks
print('-------------------------Saving files-------------------------')
wb.save('../Data sheets/APPARTEMENTS S+2 CLEAN.xlsx')
