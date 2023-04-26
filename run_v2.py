from extract_df import find_all_contracts, find_word_bool, find_first_num, first_char_is_num
from decimal import Decimal, ROUND_HALF_UP
import os
import pandas as pd

pd.options.mode.chained_assignment = None  # default='warn'

def style_code(df):
    '''extract the style code'''
    restricted_zone = df.iloc[:,:6]
    row = restricted_zone[find_word_bool(restricted_zone,'款号')[1]].iloc[0]
    return find_first_num(row)

def colour_size(df):
    '''Extract the data for the colour and size table. Returns {colours: codes}, sizes.'''
    #empty lists
    colours = []
    codes = []
    sizes = []
    

    #find the first column
    restricted_zone = df.iloc[:,:6]
    col_bool = find_word_bool(restricted_zone,'颜色')[0]
    first_col = restricted_zone[col_bool.index[col_bool]]

    #find keyword and set its index as a starting point
    row = restricted_zone[find_word_bool(restricted_zone,'颜色')[1]]
    row_start = row.index[1]
    
    #crop the top of the dataframe along with header row
    cropped_df = df.iloc[row_start+1:]
    
    #find first null value and set as end point, if there is an empty space at the end
    if cropped_df.iloc[:,first_col.columns[0]].isna().any()==True:
        row_end = cropped_df[pd.isnull(cropped_df.iloc[:,first_col.columns[0]])==True].index[0]
        #crop the bottom of the dataframe
        cropped_df = cropped_df.loc[:row_end-1]
    
    #rename the headers
    cropped_df = cropped_df.rename(columns = df.iloc[row_start])

    #append colours to the list
    for i in cropped_df.iloc[:,first_col.columns[0]]:
        colours.append(i)

    #find the column index with '色号'
    col_bool = find_word_bool(df,'色号')[0]
    col = df[col_bool.index[col_bool]]
    #append colour codes to the list 
    for i in cropped_df.iloc[:,col.columns[0]]:
        int(i)
        codes.append(i)

    #pair the colour and colour code together
    colours_dict = dict(zip(colours,codes))

    #find all the possible sizes and append to the list
    for i in range(len(cropped_df.columns)):
        if pd.notnull(cropped_df.columns[i]):
            if cropped_df.columns[i].upper() in ['OS','2XS','XS','S','M','L','XL']:
                sizes.append(cropped_df.columns[i])
            elif first_char_is_num(cropped_df.columns[i]) == True:
                sizes.append(cropped_df.columns[i])
    cropped_df = cropped_df.iloc[:,:len(sizes)+5]
    return  colours_dict, sizes, cropped_df

def main_data(df,colours_dict,sizes,contracts,contract):
    '''Clean up the dataframe for extracting number of cartons, cbm, weight, need for remark (True or False), and if it is mixed (True or False). Returns a dictionary.'''
    #find header rows and its index and combine with the sizes column
    header_rows = df.loc[find_word_bool(df,'颜色',ignore_row=[num for num in range(5)])[1]]
    header_index = header_rows.index[0]
    
    size_rows = df.loc[find_word_bool(df,sizes[0],ignore_row=[num for num in range(5)])[1]]
    size_index = size_rows.index[0]
    size_col_bool = find_word_bool(df,sizes[0],ignore_row=[num for num in range(5)])[0]
    size_col = df[size_col_bool.index[size_col_bool]]
    size_col_index = size_col.columns[0]
    
    df.iloc[header_index,size_col_index:size_col_index+len(sizes)] = df.iloc[size_index,size_col_index:size_col_index+len(sizes)]
    

    #find the end column of the main table
    end_col_bool = find_word_bool(df,'总毛重')[0]
    end_col = df[end_col_bool.index[end_col_bool]]


    #crop the top and right of the dataframe
    cropped_df = df.iloc[header_index+1:,:end_col.columns[0]+1].rename(columns=df.iloc[header_index])
    
    #create a dictionary to store data
    data = {}

    #create a dictionary to store dataframe for each colour
    data_dfs = {}

    #split the excel sheet into sections separated by colours from the colours dictionary.
    for colour in colours_dict:
        #temporary dictionary for each colour 
        temp_data = {}
        
        #calculate the starting row for each colour
        row = cropped_df.loc[find_word_bool(cropped_df,colour)[1]]
        colour_row_start = row.index[0]
        
        #find the ending row for the colour
        null_items_index = cropped_df[pd.isnull(cropped_df.loc[:,'箱号'])].index
        for i in null_items_index:
            if i>colour_row_start:
                row_end = i
                break
        #create a temporary dataframe to hold the data for the specific colour, and fill na values with 0
        temp_df = cropped_df.loc[colour_row_start:row_end-1]
        temp_df.iloc[:,0] = colours_dict[colour]
        temp_df.fillna(0,inplace=True)
        temp_df.columns = temp_df.columns.astype(str).str.replace("\n", "")
        
        #extract quantity first
        size_iter = sizes[:]
        current_box = int(temp_df['箱号'].iloc[0])
        for box_num in temp_df['箱号']:
            temp_dict = {}
            while current_box != int(box_num)+temp_df[temp_df['箱号']==box_num].loc[:,'总箱数'].sum():                
                for i in size_iter:
                    quant = temp_df[temp_df['箱号']==str(box_num)][i].sum()
                    if quant > 0:
                        style = style_code(contracts[contract])
                        name = style+'-'+colours_dict[colour]+'-'+i
                        weight = (temp_df[temp_df['箱号']==str(box_num)]['净重'].sum()/temp_df[temp_df['箱号']==str(box_num)]['每箱数量'].sum())*quant+ (temp_df[temp_df['箱号']==str(box_num)]['纸箱重量'].sum()/temp_df[temp_df['箱号']==str(box_num)]['每箱数量'].sum())*quant
                        cbm = temp_df[temp_df['箱号']==str(box_num)]['CBM'].sum() / temp_df[temp_df['箱号']==str(box_num)]['总箱数'].sum() / temp_df[temp_df['箱号']==str(box_num)]['每箱数量'].sum() * quant
                        temp_dict[name] = {'Quantity':quant,'CBM':cbm,'Weight':weight}
                box_name = "BOX"+str(current_box)
                temp_data[box_name] = temp_dict
                current_box += 1
               
        data.update(temp_data)
        data_dfs[colour] = temp_df
    return data, data_dfs

def get_all_data(contracts,input_df):
    '''Using main_data(), compile all dictionary data and store in a new dictionary with contract number as key and the old dictionary as values. Returns a dictionary.'''
    all_data = {}
    for i in contracts:
        print(i,'has begun processing')
        
        #extract data
        data = {}
        data = main_data(contracts[i],colour_size(contracts[i])[0],colour_size(contracts[i])[1],contracts,i)[0]
        while data == {}:
            print('Data for',i,'is empty, trying again...')
            data = main_data(contracts[i],colour_size(contracts[i])[0],colour_size(contracts[i])[1],contracts,i)[0]
        #insert port number into the keys
        port = input_df.loc[i,'Port_num']
        if '-' in i:
            all_data[i[:-2]+'_'+port] = (data)
        else:
            all_data[i+'_'+port] = (data)
        print(i,'has finished processing')
    return all_data




#main
def part1(dir_path):

    #create an empty excel sheet with columns for user input
    input_df = pd.DataFrame(None,columns=['Contract_id','Port_num','HSCODE','Construction','Lining','Category','Fabric','Composition','Manufacturers Details','Gender','Length'])
    input_df.to_excel(os.path.join(dir_path,'Input.xlsx'),index=False)

    #stop the program for user input
    input('\n\nPlease enter the required information into Input.xlsx located at {}.\nMake sure to close the xlsx file after completed.\nPress Enter to continue...'.format(dir_path))

    #extract user input
    input_contract = pd.read_excel(os.path.join(dir_path,'Input.xlsx'),index_col=0,dtype='str')
    #transform all contract id into string type for comparison
    input_contract.index = input_contract.index.astype('str')

    #Use function find_all_contracts() with user input and returns the dictionary, with contract number for keys and the Excel sheet's dataframe for values
    all_contracts = find_all_contracts(dir_path,list(input_contract.index.astype('str')))

    path = os.path.join(dir_path,'output')
    # Check whether the specified path exists or not
    isExist = os.path.exists(path)
    if not isExist:

        # Create a new directory because it does not exist
        os.makedirs(path)
        print("The new directory is created!")
    
    #try and except to stop program from closing when error occurs
    try:
        #extract data from dataframe with get_all_data() 
        info_dict = get_all_data(all_contracts,input_contract)
        input_contract = input_contract.fillna("N/A").reset_index()
        for key1, item1 in info_dict.items():
            new_dict = pd.DataFrame()
            for key2, item2 in item1.items():
                df = pd.DataFrame.from_dict(info_dict[key1][key2],'index').reset_index().rename(columns={"index":"SKU"})
                df.insert(0,'PurchaseOrder Number',key1)
                df.insert(1,'Carton Name',key2)
                new_dict = pd.concat([new_dict,df],ignore_index=True)
            for col in ['HSCODE','Construction','Lining','Category','Fabric','Composition','Manufacturers Details','Gender','Length']:
                new_dict[col] = input_contract[input_contract['Contract_id'].str.split("-",expand=True)[0]==key1[:-5]].loc[:,col].values[0]
            file_name = os.path.join(dir_path,'output\\'+key1+'.xlsx')
            new_dict.to_excel(file_name,index=False)
        
        input('Process completed. Please check '+path+' for the output files...')

    except Exception as e: 
        print(e)
        input('An error has occured while extracting data from Excel. Press Enter to continue...')


dir_path = input('Input the directory path of the folder containing all contracts: ')
part1(dir_path)

csv_txt = input("Enter y if you want to convert to csv file...")
dir_path = os.path.join(dir_path,'output')
if csv_txt == "y":
    for filename in os.listdir(dir_path):
        if filename[-4:] == "xlsx":
            excelfile = pd.read_excel(os.path.join(dir_path,filename),keep_default_na=False)
            excelfile.to_csv(os.path.join(dir_path,filename[:-5]+".csv"),index=False)
    input("csv file exported. Press Enter to close the program...")
else:
    input("Press Enter to end the program...") 
