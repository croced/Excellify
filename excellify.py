import sys
import glob
import gspread
import gspread_formatting as gsf

# READ A FILE TO RETURN ITEMS AS AN ARRAY
def getItems(path):
    content = []

    # fill content array
    with open(path) as file:
        content = file.readlines()

    # strip whitespace
    content = [x.strip() for x in content]

    return content

# CREATE A SPREADSHEET FROM A GIVEN FILE
def createSheetFromFile(filePath):
    # create the spreadsheet
    gc = gspread.service_account()

    spreadsheet = gc.create(filePath)
    spreadsheet.share(ownerEmail, perm_type='user', role='writer')

    # edit the spreadsheet
    items = gc.open(filePath).sheet1

    # creating/configuring headers 
    items.update('A1', 'Item:')
    gsf.set_column_width(items, 'A', 450)

    items.update('B1', 'Status:')
    gsf.set_column_width(items, 'B', 200)

    # data validation for status column
    defaultItemStatus   = '❌ INCOMPLETE'

    validation_rule = gsf.DataValidationRule(
        gsf.BooleanCondition('ONE_OF_LIST', [defaultItemStatus, '⚠️ WIP', '✅ COMPLETE']),
        showCustomUi=True
    )

    toAdd = getItems(filePath)

    # how many cells large is the gap between title and list
    listUpperPadding    = 1

    columnACounter = listUpperPadding
    columnBCounter = listUpperPadding

    # create the list
    for i in range (0, len(toAdd)):

        columnACounter += 1
        columnBCounter += 1

        itemCell = str('A' + str(columnACounter + 1))
        statusCell = str('B' + str(columnBCounter + 1))

        items.update(itemCell, toAdd[i])
        items.update(statusCell, defaultItemStatus)

    columnBRange = str('B3:B' + str(columnBCounter + 1))
    gsf.set_data_validation_for_cell_range(items, columnBRange, validation_rule)

# Enter your
ownerEmail = ""

print("--==Excellify==---")

# create a sheet for each txt list in folder
txtFiles = glob.glob('*.txt')

print("Found files: " + str(txtFiles))
print("Creating spreadsheets (check email: '" + ownerEmail + "')")

for file in txtFiles:
    createSheetFromFile(file)
