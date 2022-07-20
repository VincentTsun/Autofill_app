from extract_df import all_contracts, dir_path
from extract_info import get_all_data
import os
import pandas as pd

try:
    info_dict = get_all_data(all_contracts)
    columns= ['Quantity','Carts','Weight','CBM','Mixed?','Remark?','Remark text']
    data_df = pd.DataFrame.from_dict({(i,j): info_dict[i][j] 
                           for i in info_dict.keys() 
                           for j in info_dict[i].keys()},
                       orient='index', columns = columns)
except:
    input('An error has occured while extracting data from Excel. Press Enter to continue...')

try:
    data_df.to_excel(os.path.join(dir_path,'Combined.xlsx'))
except:
    input('An error has occured while exporting to Excel. Press Enter to continue...')

#To stop console from closing from the .exe file
input('Press enter to close...')