import pandas as pd
import os

#restrictions: 1. first 6 digits of the filename should be the contract number
#              2. Must put all files in a designated folder
#              3. Sizes must be in [XS,S,M,L,XL] or starts with a number
#              3. Changing Excel layouts may cause errors. Some caveat:
#                   a. Must have '合同号' and '款号' on the top left
#                   b. Must have destination port number on the second row
#                   c. Must have '颜色' on the top left of the main table, '每箱数量' on the top middle of the main table, and '总毛重' on the top right of the main table
#                   d. '色号' must be located to the left of the first size at the bottom chart
#                   e. Each colour must be seperated by one or more empty lines in the main table
#                   f. Colour labels must be correct!





#Reuseable functions to simplify the cleaning process
def first_char_is_num(string):
    '''Find if the first character of in the string is a number. Returns True or False.'''
    try:
        float(string[0])
        return True
    except:
        return False


def find_word_bool(df,word):
    '''Find the first location in the database with the specified keyword, returns a tuple (column,row) of boolean values as dataframe conditions.'''
    return (df.apply(lambda row: row.astype(str).str.replace(' ','').str.contains(word).any(), axis=0) ,df.apply(lambda row: row.astype(str).str.replace(' ','').str.contains(word).any(), axis=1))



def find_first_num(row):
    '''Find the first string that starts with a number, then returns the string.'''
    for i in range(1,len(row)):
        if first_char_is_num(row[i]) == True:
            num = row[i]
            return num
    



def first_six(lst):
    '''return the first 6 characters of every item in the list.'''
    lst2 = []
    for item in lst:
        lst2.append(item[:6])
    return lst2
    


#find the contracts that is in the list of contract numbers and output a dictionary with the contract numbers as keys  
#and the dataframe of the entire excel sheets as values.
def find_contracts(xls,contract_list):
    '''Locate the sheets with contract number that matches the contract_list, then adding to a dictionary. 
    Returns the dictionary, with contract number for keys and the sheet's dataframe for values.'''
    sheets = {}
    for i in range(len(xls.sheet_names)):
        df = pd.read_excel(xls,i,header=None)
        contract_row = df[find_word_bool(df,'合同号')[1]].iloc[0]
        contract_num = find_first_num(contract_row)
        if contract_num in contract_list:
            sheets[contract_num] = pd.read_excel(xls,i,header=None)
    return sheets

def find_all_contracts(dir_path,contract_list):
    '''Using function find_contracts(), loop through all files in specified folder and add to dictionary if contract number matches the contract list.
    Returns the dictionary, with contract number for keys and the sheet's dataframe for values.'''
    all_sheets = {}
    lst = first_six(contract_list)

    for filename in os.listdir(dir_path):
        f = os.path.join(dir_path,filename)
        if os.path.isfile(f):
            if filename[:6] in lst:
                xls = pd.ExcelFile(f)
                sheets = find_contracts(xls,contract_list)
                all_sheets.update(sheets)
        print(f)
    return all_sheets




#Variables needed from user input
dir_path = input('Input the directory path of the folder containing all contracts: ')

input_df = pd.DataFrame(None,columns=['Contract_id','Port_num','Desc'])

input_df.to_excel(os.path.join(dir_path,'Input.xlsx'),index=False)

input('Please enter the required information into Input.xlsx located at {}.\nMake sure to close the xlsx file after completed.\nPress Enter to continue...'.format(dir_path))

input_contract = pd.read_excel(os.path.join(dir_path,'Input.xlsx'))


#Use function find_all_contracts() with user input
all_contracts = find_all_contracts(dir_path,input_contract.iloc[:,0])



