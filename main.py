import Leadership
import openpyxl
import Aspace

filename = str(input("Please enter the exact name of the file of the excel sheet: "))
wb = openpyxl.load_workbook(filename)
ws = wb.active
start = 0
print('Welcome to the Jimmy Stewart Composite Squadron Cadet Programs Data Base')


while start < 1:
    print('\nSelect which of these you would like to run: ')
    print('1. List cadets with completed leadership')
    print('2. List of cadets with completed aerospace')
    print('3. List of cadets needing a drill test')
    print('4. Exit the program')
    menu_start = int(input("Your Selection: "))

    if menu_start == 4:
        print("The program will now end.")
        break

    num_cadets = int(input('How many cadets are there? \n'))
    leadPass = Leadership.leadPass(num_cadets)
    aspacePass = Aspace.aspacePass(num_cadets)

    if menu_start == 1:
        for x in leadPass:
            print(ws[f'C{x}'].value + " " + ws[f'B{x}'].value)
        print("\nWould you like to continue?")
        cont = str(input("(Y/N): "))

    if menu_start == 2:
        for x in aspacePass:
            print(ws[f'C{x}'].value + " " + ws[f'B{x}'].value)
        print("\nWould you like to continue?")
        cont = str(input("(Y/N): "))

    if menu_start == 3:
        drTest = []
        for x in leadPass:
            for y in aspacePass:
                if x==y:
                    drTest.append(x)
        for z in drTest:
            print(ws[f'C{z}'].value + " " + ws[f'B{z}'].value)
        print("\nWould you like to continue?")
        cont = str(input("(Y/N): "))
        print(cont)

    if cont == 'N' or cont == 'n':
        print("The program will now end.")
        start = 1


