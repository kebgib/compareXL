#! python3
print("""
Ver 0.1
This application will open and compare MS Excel docs and compare their contents.
""")
import os, openpyxl
from time import sleep

# This changes the current working dir to the directory that contains your desired excel files
os.chdir('YOUR DIRECTORY HERE')
print('\nThe current directory is: \n\t'
      +(os.getcwd()))

# Imports the excel files
wb = openpyxl.load_workbook('Example.xlsx')
wb2 = openpyxl.load_workbook('Example2.xlsx')

def sheetAssignment(book):
    """
    The sheetAssignment function queries the user for an input selection that should be from the printSheets() function
    The except clause of the try loop will catch the errors listed, throw an oops and try again. very nice.
    """
    printSheets(book)
    while True:
      try:
        sheetgrab = input("Which sheet would you like to load? (q to quit): ")
        if sheetgrab.title() == 'Q':
          print('Goodbye!')
          sleep(3)
          exit()
        sheet = book[str(sheetgrab.title())]
        print (sheet)
        return sheet
      except (ValueError, KeyError):
        print("Oops! That is not a valid sheet choice. Please try again...")

def printSheets(book):
    """
    This function prints the sheetnames within the xlsx doc. This functions is called by sheetAssignment()
    """
    print('\nThe sheets within this Excel file are:      ')
    sleep(.5)
    for sheet in book:
        print('\t' + sheet.title)
        sleep(0.15)

def nodeinventory(first_sheet, second_sheet):
    sheet1_list = []
    sheet2_list = []
    print("\tComparing contents of the workbooks, and preparing to display matches...")
    sleep(1.5)
    # Iterating through the items in first_sheet, adding them to list(sheet1_list)
		#	**Change the column values to target the groups you wish to compare**
    for x in range(2, (first_sheet.max_row)):
        sheet1_list.append(first_sheet.cell(row=x, column=6).value)
    # Iterating through the sheet2_list, adding them to list(sheet2_list)
    for i in range(2, (second_sheet.max_row)):
        sheet2_list.append(second_sheet.cell(row=i, column=2).value)
    # Compares serials from sheet1_list to the serials found in sheet2_list's doc
    count = 0
    for z in sheet1_list:
        count += 1
        if z in sheet2_list:
            count -= 1
            print ("Match found! " + z + " is in both books!" )
        if count == len(sheet1_list):
            print ("\t\n\nNo matches, dawg.")

# Running the function set for the first time
sheet1 = sheetAssignment(wb)
sheet2 = sheetAssignment(wb2)
nodeinventory(sheet1, sheet2)

# Giving the user a chance to play again, load different sheets.
try:
    while True:
        loopanswer = input("\t\tWould you like to check again? (y/n): ")
        if loopanswer.lower() == 'y':
            nodeinventory(sheet1, sheet2)
        if loopanswer.lower() == 'n':
            loopanswer2 = input("\t\tWould you like to check different sheets? (y/n)")
            if loopanswer2.lower() == 'y':
                sht1 = sheetAssignment(wb)
                sht2 = sheetAssignment(wb2)
                nodeinventory(sht1, sht2)
            if loopanswer2.lower() == 'n' or 'q':
                exit()
except(ValueError, KeyError):
    print("Oops! That is not a valid sheet choice. Please try again...")




  
  


