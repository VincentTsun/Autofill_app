import xlwings as xw
import pandas as pd
import os
from extract_info import colour_size, main_data
from extract_df import find_contracts

def find_loc(sheet,label,row_end=500,col_end=40,row_start=1,col_start=1):
    '''Find the location of the specified value from the excel sheet. Returns a tuple (row,column).'''
    for row in range(row_start, row_end):
        for col in range(col_start, col_end):
            if type(sheet.range((row,col)).value) == str:
                if sheet.range((row,col)).value.strip().upper() == label:
                    location = (row,col)
                    return location

def extract_dpl_value(sheet,label,row_end=500,col_end=40,row_start=1,col_start=1,method='same_row'):
    '''extract the value based on the label from the Excel sheet'''
    location = find_loc(sheet,label,row_end,col_end,row_start,col_start)
    col = location[1]
    row = location[0]
    if method == 'same_row':
        while str(sheet.range((row,col+1)).value) == 'None':
            col+=1
        result = sheet.range((row,col+1)).value
    elif method == 'same_col':
        while str(sheet.range((row+1,col)).value) == 'None':
            row+=1
        result = sheet.range((row+1,col)).value
    return result
                

def fill_dpl(sheet,info_dict,id):
    print('Begin filling',id)
    
    #reset index for dataframes
    info_dict['colour_table'].reset_index(inplace=True)

    #fill header
    sheet.range('D4').value = info_dict['style']
    if id[:2].upper() != 'PO':
        sheet.range('G4').value = id[:6]
    else:
        sheet.range('G4').value = id[:8]
    
    #fill size headers
    data_location = find_loc(sheet,'SIZE',row_end=9,row_start=6,col_start=7,col_end=10)
    current_row = data_location[0]+1
    current_col = data_location[1]
    for num in range(len(info_dict['sizes'])):
        cell = (current_row,current_col)
        sheet.range(cell).value = info_dict['sizes'][num]
        current_col+=1
    
    #fill main data
    data_location = find_loc(sheet,'REPLEN',row_end=100)
    current_row = data_location[0]+2
    current_col = data_location[1]
    row_added = 0
    for key,table in info_dict['data_table'].items():
        table.reset_index(inplace=True)
        for row in table.index:
            sheet.range((current_row,current_col)).value = table.iloc[row][2]
            current_col += 1
            sheet.range((current_row,current_col)).value = '-'
            current_col += 1
            sheet.range((current_row,current_col)).value = table.iloc[row][4]

            current_col += 3
            sheet.range((current_row,current_col)).value = key
            current_col += 1
            sheet.range((current_row,current_col)).value = table.iloc[row][1]
            size_col = current_col
            for size_num in range(len(info_dict['sizes'])):
                size_col += 1
                if table.iloc[row][size_num+6] != 0:
                    sheet.range((current_row,size_col)).value = table.iloc[row][size_num+6]
            if len(info_dict['sizes'])<=12:
                current_col += 14
                sheet.range((current_row,current_col)).value = sheet.range((current_row,current_col)).value = '=V'+str(current_row)+'+'+'D'+str(current_row)+'*'+str(table.iloc[row]['纸箱重量'])
            else:
                current_col += 20
                sheet.range((current_row,current_col)).value = sheet.range((current_row,current_col)).value = '=AB'+str(current_row)+'+'+'D'+str(current_row)+'*'+str(table.iloc[row]['纸箱重量'])
            current_col += 2
            sheet.range((current_row,current_col)).value = info_dict['nw']
            current_col += 1
            sheet.range((current_row,current_col)).value = round(table.iloc[row]['CBM'],3)


            #add new row
            current_row += 1
            row_add = str(current_row)+':'+str(current_row)
            sheet.range(row_add).insert()
            row_added += 1
            current_col = data_location[1]
    row_add = str(current_row)+':'+str(current_row)
    sheet.range(row_add).delete()
    sheet.range(row_add).delete()
    
    #copy formula for all added rows
    row_range_start = 23
    row_range_end = 23+row_added-1
    current_col_2 = 4
    formula = sheet.range((23,current_col_2)).formula
    sheet.range((row_range_start,current_col_2),(row_range_end,current_col_2)).formula = formula
    current_col_2 += 1 
    formula = sheet.range((23,current_col_2)).formula
    sheet.range((row_range_start,current_col_2),(row_range_end,current_col_2)).formula = formula
    if len(info_dict['sizes'])<=12:
        current_col_2 += 15
    else:
        current_col_2 += 21
    formula = sheet.range((23,current_col_2)).formula
    sheet.range((row_range_start,current_col_2),(row_range_end,current_col_2)).formula = formula
    current_col_2 += 2
    formula = sheet.range((23,current_col_2)).formula
    sheet.range((row_range_start,current_col_2),(row_range_end,current_col_2)).formula = formula

    #fill summary
    row_end = sheet.used_range[-1].row
    summary_location = find_loc(sheet,'REPLEN QUANTITY',row_start=38+row_added,col_start=5,row_end=row_end)
    current_row = summary_location[0]+2
    current_col = summary_location[1]
    for c in range(len(info_dict['colour_table']['客色號'])):
        sheet.range((current_row,current_col)).value = info_dict['colour_table']['客色號'][c]
        sheet.range((current_row,current_col+1)).value = info_dict['colour_table']['色号'][c]
        for s in info_dict['sizes']:
            current_col+=1
            sheet.range((current_row,current_col+1)).value = info_dict['colour_table'][s][c]
        current_row+=1
        current_col = summary_location[1]
    print('Finish filling',id)

def dpl_extract_all(dir_path,id,filename):
    current_dict = {}
    f = xw.Book(os.path.join(dir_path,filename))
    for sheet in f.sheets:
        temp_sheet = f.sheets[sheet]
        check_value = extract_dpl_value(temp_sheet,'合同号:',4)
        if type(check_value) == float or type(check_value) == int:
            check_value = str(int(check_value))
        if check_value.upper() == id.upper():
            df = find_contracts(pd.ExcelFile(os.path.join(dir_path,filename)),[id])
            nw = extract_dpl_value(temp_sheet,'NW:',20)
            style = extract_dpl_value(temp_sheet,'款  号:',6)
            cs = colour_size(df[id])
            colours = cs[0]
            sizes = cs[1]
            colour_table = cs[2]
            data_table = main_data(df[id],colours,sizes)[1]
            for colour in colours:
                row_end = sheet.used_range[-1].row
                alt_name = extract_dpl_value(temp_sheet,colour,row_end=row_end,col_end=2,row_start=10,method='same_col')
                data_table[alt_name] = data_table.pop(colour)
            current_dict[id] = {'nw':nw,'style':style,'colours':colours,'sizes':sizes,'colour_table':colour_table,'data_table':data_table}
            break
    f.close()
    return current_dict

def dpl_setup(dir_path):
    '''Extract data from each corresponding Excel sheet and fill the dpl file.'''
    #extract contract id into a list
    input_df = pd.read_excel(os.path.join(dir_path,'input.xlsx'),converters={'Port_num':str})
    input_list = [i for i in input_df.Contract_id.astype('str')]
    input_port_num = [i for i in input_df.Port_num.astype('str')]

    #search for each corresponding Excel sheet using the input_list
    info_dict={}
    for id in input_list:
        print('Checking',id)
        for filename in os.listdir(dir_path):
            if filename[:2].upper() != 'PO' and filename[:6] == id[:6]:
                info_dict.update(dpl_extract_all(dir_path,id,filename))
                print('Completed',id)
                break
            elif filename[:2].upper() == 'PO' and filename[:8].upper() == id[:8].upper():
                info_dict.update(dpl_extract_all(dir_path,id,filename))
                print('Completed',id)
                break
            
    
    #copy the template sheet
    wb_template = xw.Book(os.path.join(dir_path,'dpl template.xlsx'))

    #create a new workbook
    dpl_wb = xw.Book()
    
    #delete file if the file exists to prevent error
    if os.path.isfile(os.path.join(dir_path,'dpl.xlsx')):
        os.remove(os.path.join(dir_path,'dpl.xlsx'))
    #save the file
    dpl_wb.save(os.path.join(dir_path,'dpl.xlsx'))

    #create sheets for each contract using the template
    for i in range(len(input_list)):
        if len(info_dict[input_list[i]]['sizes']) <= 12:
            sheet = wb_template.sheets[0]
        elif len(info_dict[input_list[i]]['sizes']) <= 18:
            sheet = wb_template.sheets[1]
        else:
            input('Too many sizes for the po. Please contact me.')
        if input_list[i][:2].upper() != 'PO':
            name = str(input_list[i][:6])+'_'+str(input_port_num[i])
        else:
            name = str(input_list[i][:8])+'_'+str(input_port_num[i])
        sheet.copy(after=dpl_wb.sheets[i],name=name)
        target_sheet = dpl_wb.sheets[name]
        fill_dpl(target_sheet,info_dict[input_list[i]],input_list[i])

    #delete the default sheet and save
    dpl_wb.sheets[0].delete()
    dpl_wb.save()
    
    #close the template excel
    wb_template.close()
    dpl_wb.save()
    dpl_wb.close()

