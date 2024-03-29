from extract_df import find_word_bool, find_first_num, first_char_is_num
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

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
            if cropped_df.columns[i].upper() in ['2XS','XS','S','M','L','XL']:
                sizes.append(cropped_df.columns[i])
            elif first_char_is_num(cropped_df.columns[i]) == True:
                sizes.append(cropped_df.columns[i])
    cropped_df = cropped_df.iloc[:,:len(sizes)+5]
    return  colours_dict, sizes, cropped_df

def main_data(df,colours_dict,sizes):
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
        
        #Make a copy of temp_df to preserve all rows for later calculations
        temp_df1 = temp_df[:]
        
        #extract quantity first
        temp_sizes = sizes[:]
        size_iter = sizes[:]
        for i in size_iter:
            name = colours_dict[colour]+'-'+i
            quant = temp_df['总箱数'].dot(temp_df[i]).sum()
            if quant > 0:
                temp_data[name] = [quant]
            else:
                temp_sizes.remove(i)
            

        #extract number of carts, cbm, weight, need for remark (True or False), and if it is mixed (True or False)
        for i in temp_sizes:
            name = colours_dict[colour]+'-'+i
            if temp_df.empty == False:
                
                #filter for only rows that matter to each size
                size_df = temp_df[temp_df[i]!=0]
                #drop the used rows so it is easier to calculate carts, cbm, weight
                temp_df.drop(size_df.index.values,inplace=True)
                
                carts = size_df['总箱数'].sum()
                temp_data[name].append(carts)

                weight = 0
                for y in range(len(size_df['总毛重'])):
                    weight += Decimal(str(size_df['总毛重'].iloc[y])).quantize(Decimal('0.00'),rounding=ROUND_HALF_UP)
                temp_data[name].append(float(weight))
                
                cbm = 0
                for y in range(len(size_df['总毛重'])):
                    cbm += Decimal(str(size_df['CBM'].iloc[y])).quantize(Decimal('0.000'),rounding=ROUND_HALF_UP)
                temp_data[name].append(float(cbm))

                #check if the size only has one column, and assign remark and mixed accordingly
                small_df = temp_df1.loc[:,temp_sizes[0]:temp_sizes[-1]][temp_df1[i]!=0]
                if len(small_df.loc[:,(small_df != 0).any(axis=0)].columns)>1:
                    mixed = True
                    #remark = False
                else:
                    mixed = False
                    #if len(small_df.loc[:,(small_df != 0).any(axis=0)].index)>1:
                        #remark = True
                    #else:
                        #remark = False
                
                temp_data[name].append(mixed)
                
                #temp_data[name].append(remark)

                
            else:
                #mixed cartons default
                carts = 0
                temp_data[name].append(carts)

                weight = 0
                temp_data[name].append(weight)

                cbm = 0
                temp_data[name].append(cbm)

                mixed = True
                temp_data[name].append(mixed)
                
                #remark = False
                #temp_data[name].append(remark)
        #append remark texts to the dictionary
        #for key, val in temp_data.items():
        #    remark_desc = ''
        #    carts_quant_dict = {}
        #    if val[5] == True:
        #        small_df = temp_df1[temp_df1[key[3:]]!=0]
        #        for i in small_df.index.values:
        #            carts_text = small_df.loc[i,'总箱数']
        #            quant_text = small_df.loc[i,'每箱数量']
        #            if (carts_text,quant_text) in carts_quant_dict:
        #                carts_quant_dict[(carts_text,quant_text)] += 1
        #            else:
        #                carts_quant_dict[(carts_text,quant_text)] = 1
        #        counter = 0
        #        for k, v in carts_quant_dict.items():
        #            counter+=1
        #            carts_text = k[0]*v
        #            quant_text = k[1]
        #            if counter != len(carts_quant_dict):
        #                text = '{} carton x {} pcs, \n'.format(carts_text,quant_text)
        #            else:
        #                text = '{} carton x {} pcs.'.format(carts_text,quant_text)
        #            remark_desc += text
        #    temp_data[key].append(remark_desc)
        data.update(temp_data)
        data_dfs[colour] = temp_df1
    return data, data_dfs

def get_all_data(contracts,input_df):
    '''Using main_data(), compile all dictionary data and store in a new dictionary with contract number as key and the old dictionary as values. Returns a dictionary.'''
    all_data = {}
    for i in contracts:
        print(i,'has begun processing')
        
        #extract data
        data = {}
        data = main_data(contracts[i],colour_size(contracts[i])[0],colour_size(contracts[i])[1])[0]
        while data == {}:
            print('Data for',i,'is empty, trying again...')
            data = main_data(contracts[i],colour_size(contracts[i])[0],colour_size(contracts[i])[1])[0]
        #extract marks and unit from input dataframe 
        marks = input_df.loc[i,'Marks']
        hs_code = input_df.loc[i,'HS_Code']
        unit = input_df.loc[i,'Unit']
        for v in data:
            data[v].append(marks)
            data[v].append(hs_code)
            data[v].append(unit.strip().upper())
        
        style = style_code(contracts[i])
        data = {style+f"-{key}": val for key, val in data.items()}
        #insert port number into the keys
        port = input_df.loc[i,'Port_num']
        if '-' in i:
            all_data[i[:-2]+'_'+port] = (data)
        else:
            all_data[i+'_'+port] = (data)
        print(i,'has finished processing')
    return all_data




