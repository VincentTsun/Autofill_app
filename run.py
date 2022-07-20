from extract_df import find_all_contracts
from extract_info import get_all_data
import os
import pandas as pd


#Variables needed from user input
dir_path = input('Input the directory path of the folder containing all contracts: ')

input_df = pd.DataFrame(None,columns=['Contract_id','Port_num','Desc'])

input_df.to_excel(os.path.join(dir_path,'Input.xlsx'),index=False)

#input('\n\nPlease enter the required information into Input.xlsx located at {}.\nMake sure to close the xlsx file after completed.\nPress Enter to continue...'.format(dir_path))

input_contract = pd.read_excel(os.path.join(dir_path,'Input - Copy.xlsx'),index_col=0,dtype='str')

#Use function find_all_contracts() with user input
all_contracts = find_all_contracts(dir_path,list(input_contract.index))




#try and except to stop program from closing when error occurs
try:
    #extract data from dataframe
    info_dict = get_all_data(all_contracts,input_contract)
    columns= ['Quantity','Carts','Weight','CBM','Mixed?','Remark?','Remark text','Desc']
    data_df = pd.DataFrame.from_dict({(i,j): info_dict[i][j] 
                           for i in info_dict.keys() 
                           for j in info_dict[i].keys()},
                       orient='index', columns = columns)
except Exception as e: 
    print(e)
    input('An error has occured while extracting data from Excel. Press Enter to continue...')

try:
    #write data into excel file
    data_df.to_excel(os.path.join(dir_path,'Combined.xlsx'))
except Exception as e:
    print(e)
    input('An error has occured while exporting to Excel. Press Enter to continue...')

#To stop console from closing from the .exe file
input('Press enter to close...')