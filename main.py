import Leadership
import cadet_directory
import openpyxl

start = 0
print('Welcome to the Jimmy Stewart Composite Squadron Cadet Programs Data Base')

while start < 1:
    print('\nSelect which of these you would like to run: ')
    print('1. List cadets with completed leadership')
    print('2. List of cadets with completed aerospace')
    menu_start = int(input("Your Selection: "))

    if menu_start == 1:
        LeadTestStatus = Leadership.leadTest_check()
        print('\nCadets who have completed the leadership requirment are:')
        for capid in LeadTestStatus:
            print(cadet_directory.roster_name[capid])

    print("\nWould you like to continue?")
    cont = str(input("(Y/N): "))
    if cont == 'N' or 'n':
        print("The program will now end.")
        break






