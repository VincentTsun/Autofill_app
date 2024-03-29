from extract_df import find_all_contracts
from extract_info import get_all_data
from fillweb import web_login,to_booking_page,fill_in,add_po
from selenium import webdriver
import os
import pandas as pd
from DPL import dpl_setup

def part1(dir_path):

    #create an empty excel sheet with columns for user input
    input_df = pd.DataFrame(None,columns=['Contract_id','Port_num','Marks','HS_Code','Unit'])
    input_df.to_excel(os.path.join(dir_path,'Input.xlsx'),index=False)

    #stop the program for user input
    input('\n\nPlease enter the required information into Input.xlsx located at {}.\nMake sure to close the xlsx file after completed.\nPress Enter to continue...'.format(dir_path))

    #extract user input
    input_contract = pd.read_excel(os.path.join(dir_path,'Input.xlsx'),index_col=0,dtype='str')
    #transform all contract id into string type for comparison
    input_contract.index = input_contract.index.astype('str')

    #Use function find_all_contracts() with user input and returns the dictionary, with contract number for keys and the Excel sheet's dataframe for values
    all_contracts = find_all_contracts(dir_path,list(input_contract.index.astype('str')))

    #try and except to stop program from closing when error occurs
    try:
        #extract data from dataframe with get_all_data() 
        info_dict = get_all_data(all_contracts,input_contract)
        
        #Combine contract id and style id, serve as unique keys
        columns= ['Quantity','Carts','Weight','CBM','Mixed?','Marks','HS_Code','Unit']
        for i in info_dict:
            for j in info_dict[i]:
                if len(info_dict[i][j])>8:
                    print('The following caused an error:')
                    print(i,j,info_dict[i][j])
                    input('Please fix the error in the corresponding Excel and run the program again.')
        data_df = pd.DataFrame.from_dict({(i,j): info_dict[i][j] 
                                for i in info_dict.keys() 
                                for j in info_dict[i].keys()},
                            orient='index', columns = columns)
        try:
            #write data into excel file
            data_df.to_excel(os.path.join(dir_path,'Preview.xlsx'))
            #To stop console from closing from the .exe file
            input('\nExtraction has completed\n\nPlease check the Preview.xlsx file before proceeding.\n\nPress enter to continue...')
        except Exception as e:
            print(e)
            input('An error has occured while exporting to Excel. Press Enter to continue...')

    except Exception as e: 
        print(e)
        input('An error has occured while extracting data from Excel. Press Enter to continue...')

def part2(dir_path,username,password,booking_num):
    PATH = "C:\Program Files (x86)\chromedriver.exe"
    input_contract = pd.read_excel(os.path.join(dir_path,'Input.xlsx'),index_col=0,dtype='str')
    
    try:
        driver = webdriver.Chrome(PATH)
        driver.get('https://portal.damco.com/Applications/shipper/')
    except:
        print('\nPlease make sure the chrome web driver matches up with your chrome version, and is located at C:\Program Files (x86).')
        print('Make sure that in the folder C:\Program Files (x86), there is a file called chromedriver.exe')
        input('\nPlease close the program...')

    web_login(username,password,driver)
    to_booking_page(booking_num,driver)
    add_po(input_contract,driver)
    input('\n\nAdding process Complete! Press Enter to continue...')
    

def part3(dir_path,username,password,booking_num):
    PATH = "C:\Program Files (x86)\chromedriver.exe"
    
    data_df = pd.read_excel(os.path.join(dir_path,'Preview.xlsx'),index_col=0)

    advanced = ''
    while advanced == '':
        advanced = input('Advanced settings(y/n)?')
        if advanced.lower() == 'y':
            startline = input('Starting line number: ')
            endline = input('Ending line number: ')
        elif advanced.lower() == 'n':
            startline = ''
            endline = ''
        else:
            advanced = ''

    driver = webdriver.Chrome(PATH)
    driver.get('https://portal.damco.com/Applications/shipper/')
    
    web_login(username,password,driver)
    to_booking_page(booking_num,driver)
    fill_in(data_df,driver,startline,endline)


mode = int(input('Functions:\n1. Extract data from Excel\n2. Add PO to booking\n3. Fill system with data\n4. All of the above\n5. DPL\n'))
if mode == 1:
    dir_path = input('Input the directory path of the folder containing all contracts: ')
    part1(dir_path)
elif mode == 2:
    dir_path = input('Input the directory path of the folder containing all contracts: ')

    username = input('Username: ')
    password = input('Password: ')
    booking_num = input('Booking number: ')
    try:
        part2(dir_path,username,password,booking_num)
    except:
        input('Error occured when trying to add PO.\nPress Enter to close the program...')
elif mode == 3:
    dir_path = input('Input the directory path of the folder containing all contracts: ')

    username = input('Username: ')
    password = input('Password: ')
    booking_num = input('Booking number: ')
    try:
        part3(dir_path,username,password,booking_num)
    except:
        input('Error occured when trying to fill data onto the system.\nPress Enter to close the program...')
elif mode ==4:
    dir_path = input('Input the directory path of the folder containing all contracts: ')
    part1(dir_path)

    username = input('Username: ')
    password = input('Password: ')
    booking_num = input('Booking number: ')

    try:
        part2(dir_path,username,password,booking_num)
    except:
        input('Error occured when trying to add PO.\nPress Enter to close the program...')

    try:
        part3(dir_path,username,password,booking_num)
    except:
        input('Error occured when trying to fill data onto the system.\nPress Enter to close the program...')
elif mode == 5:
    dir_path = input('Input the directory path of the folder containing all contracts: ')
    try:
        dpl_setup(dir_path)
        input('DPL setup completed.\nPress Enter to close the program...')
    except:
        input('Error occured when trying to compile the DPL.\nPress Enter to close the program...')
else:
    input('Value Error. Press Enter to close the program...')


