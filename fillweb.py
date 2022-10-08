from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time





def web_login(username,password,driver):
    '''Login to website'''
    username_field = driver.find_element(By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_UsernameTextBox"]')
    username_field.send_keys(username)


    password_field = driver.find_element(By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_PasswordTextBox"]')
    password_field.send_keys(password)

    login_click = driver.find_element(By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_SubmitButton"]')
    login_click.click()

def to_booking_page(booking_num,driver):
    '''load the booking page'''
    while True:
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="MyDamcoMenu_navigation"]/li/a')))
            shipper_click = driver.find_element(By.XPATH,'//*[@id="MyDamcoMenu_navigation"]/li/a')
            shipper_click.click()
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="applicationIframe"]')))
            iframe = driver.find_element(By.XPATH,'//*[@id="applicationIframe"]')
            driver.switch_to.frame(iframe)

            search_field = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,'//*[@id="quick_search_Doc_No"]')))
            search_field.send_keys(booking_num)
            break
        except:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="MyDamcoMenu_navigation"]/li/a')))
            shipper_click = driver.find_element(By.XPATH,'//*[@id="MyDamcoMenu_navigation"]/li/a')
            shipper_click.click()

    search_click = driver.find_element(By.XPATH,'//*[@id="searchDocNobtn"]')
    search_click.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="dojox_grid__View_0"]/div/div/div/div/table/tbody/tr/td[6]/a')))

    while True:
        try:
            booking_click = driver.find_element(By.XPATH,'//*[@id="dojox_grid__View_0"]/div/div/div/div/table/tbody/tr/td[6]/a')
            booking_click.click()
            break
        except:
            pass

    detail_tab_click = driver.find_element(By.XPATH,'//*[@id="tab_treetab3"]')
    detail_tab_click.click()




def save_draft(driver):
    '''click save as draft'''
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="SaveAsMenuBtnId"]')))
    while True:
        try:
            save_as_click = driver.find_element(By.XPATH,'//*[@id="SaveAsMenuBtnId"]')
            save_as_click.click()

            draft_click = driver.find_element(By.XPATH,'//*[@id="SOStatusOption"]/a[1]')
            draft_click.click()
            break
        except:
            pass

    




def add_po(input_contract,driver):
    '''Add new po'''
    contract_num_list = list(input_contract.index)[1:]
    for contract_num in contract_num_list:
        if '-' in str(contract_num):
            po_num = str(input_contract[input_contract.index == contract_num].index[0][:-2])+'_'+input_contract.loc[contract_num,:][0]
        else:
            po_num = str(input_contract[input_contract.index == contract_num].index[0])+'_'+input_contract.loc[contract_num,:][0]
        while True:
            try:
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="addButtonId"]')))
                add_click = driver.find_element(By.XPATH,'//*[@id="addButtonId"]')
                add_click.click()
                break
            except:
                pass
        while True:
            try:
                time.sleep(1)
                po_field = driver.find_element(By.XPATH,'//*[@id="srchPONoId"]')
                po_field.send_keys(po_num)
                break
            except:
                add_click = driver.find_element(By.XPATH,'//*[@id="addButtonId"]')
                add_click.click()

        search_click = driver.find_element(By.XPATH,'//*[@id="searchPObtn"]')
        search_click.click()

        WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//*[@id="grid_rowSelector_0"]')))

        box_check = driver.find_element(By.XPATH,'//*[@id="grid_rowSelector_0"]')
        box_check.click()

        add_click = driver.find_element(By.XPATH,'//*[@id="bookBtnId"]')
        add_click.click()

        print('Adding PO number:',po_num)
        while True:
            try:
                end_page_click = driver.find_element(By.XPATH,'//*[@id="mLast"]/i')
                end_page_click.click()
                break
            except:
                time.sleep(1)
        
        save_draft(driver)
        print('PO added and saved')

def fill_in(data_df,driver,startline='',endline=''):
    '''function to fill in data into the web form'''
    wrongquant = []
    wrongquantnum = []
    if startline == '':
        startline = 1
    if endline == '':
        endline = len(data_df)
    data_df_reset = data_df.reset_index()
    for i in range(int(startline)-1,int(endline)):
        num = i+1
        
        
        while True:
            try:
                po_id_web = driver.find_element(By.ID,'POID_{}'.format(i))
                sku_id_web = driver.find_element(By.ID,'SKUID_{}'.format(i))
                web_key = (po_id_web.text,sku_id_web.text)
                data_row = data_df_reset[data_df_reset['index'] == str(web_key)]

                packages_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bookedPackagesId{}"]'.format(i))))
                packages_field.click()
                break
            except:
                next_page_click = driver.find_element(By.XPATH,'//*[@id="mNext"]/i')
                next_page_click.click()
        
        print(web_key,'verified. Begin filling...')
        packages_field.send_keys(Keys.CONTROL + "a")
        packages_field.send_keys(int(data_row['Carts'].values))

        weight_field = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bookedWeightId{}"]'.format(i))))
        weight_field.click()
        weight_field.send_keys(Keys.CONTROL + "a")
        weight_field.send_keys(round(float(data_row['Weight'].values),2))


        measurement_field = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bookedMeasurementId{}"]'.format(i))))
        measurement_field.click()
        measurement_field.send_keys(Keys.CONTROL + "a")
        measurement_field.send_keys(round(float(data_row['CBM'].values),3))

        quantity_field = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bookedQtyId{}"]'.format(i))))
        if quantity_field.get_attribute('value') != str(data_row['Quantity'].item()):
            wrongquant.append('Line No. {}'.format(num))
            wrongquantnum.append(num)
            print('Found unmatched quantity: Line No. {}'.format(num))

        if data_row['Unit'].values == 'PCS':
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[2]'.format(i))
            quantity_dropdown.click()
        elif data_row['Unit'].values == 'BOX':
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[3]'.format(i))
            quantity_dropdown.click()
        elif data_row['Unit'].values == 'CAS':
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[4]'.format(i))
            quantity_dropdown.click()
        elif data_row['Unit'].values == 'CTN':
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[5]'.format(i))
            quantity_dropdown.click() 
        elif data_row['Unit'].values == 'DOZ':
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[6]'.format(i))
            quantity_dropdown.click()
        elif data_row['Unit'].values == 'EA':
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[7]'.format(i))
            quantity_dropdown.click()
        elif data_row['Unit'].values == 'PLT':
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[8]'.format(i))
            quantity_dropdown.click()
        elif data_row['Unit'].values == 'UNI':
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[9]'.format(i))
            quantity_dropdown.click()
        else:
            quantity_dropdown = driver.find_element(By.XPATH,'//*[@id="qtyUnitId{}"]/option[1]'.format(i))
            quantity_dropdown.click()

       

        marks_field = driver.find_element(By.XPATH,'//*[@id="lnMkNum{}"]'.format(i))
        marks_field.click()
        marks_field.send_keys(Keys.CONTROL + "a")
        marks_field.send_keys(Keys.BACKSPACE)
        marks_field.send_keys(data_row['Marks'].values)
        #if data_row['Remark?'].values == True:
        #    marks_field.send_keys(Keys.RETURN)
        #    marks_field.send_keys(Keys.RETURN)
        #    marks_field.send_keys(data_row['Remark text'].values)

        country_code_field = driver.find_element(By.XPATH,'//*[@id="EditSOForm_soDto_soLineDtoList_{}__soLineHtsDtoList_0__country_countryCode"]'.format(i))
        country_code_field.clear()
        country_code_field.send_keys('CN')

        if data_row['Mixed?'].bool() == True:
            try:
                mix_dropdown = driver.find_element(By.XPATH,'//*[@id="dynafield_{}_2_refValue"]/option[2]'.format(i))
                mix_dropdown.click()
            except:
                pass
        else:
            try:
                mix_dropdown = driver.find_element(By.XPATH,'//*[@id="dynafield_{}_2_refValue"]/option[1]'.format(i))
                mix_dropdown.click()
            except:
                pass
        
        if num%50 == 0:
            next_page_click = driver.find_element(By.XPATH,'//*[@id="mNext"]/i')
            next_page_click.click()
    print('PO with mismatched quantities:',wrongquant)
    save_draft(driver)
    print('\nWaiting for file to be saved...')
    while True:
        try:
            first_page_click = driver.find_element(By.XPATH,'//*[@id="mFirst"]/i')
            first_page_click.click()
            break
        except:
            time.sleep(1)
    input('\nSaving complete! \nPress Enter to proceeed...')
    
    #filling mismatched quantity on the system after the user checked for errors
    if len(wrongquantnum) != 0:
        fill_mismatch = ''
        while fill_mismatch.lower().strip() != 'y' and fill_mismatch.lower().strip() != 'n':
            fill_mismatch = input('Do you wish to replace mismatched quantities on the system(y/n)? ')
        fill_mismatch = fill_mismatch.lower().strip()
        if fill_mismatch == 'y':
            for num in wrongquantnum:
                while True:
                    try:
                        po_id_web = driver.find_element(By.ID,'POID_{}'.format(num))
                        sku_id_web = driver.find_element(By.ID,'SKUID_{}'.format(num))
                        web_key = (po_id_web.text,sku_id_web.text)
                        data_row = data_df_reset[data_df_reset['index'] == str(web_key)]

                        quantity_field = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bookedQtyId{}"]'.format(i))))
                        quantity_field.click()
                        break
                    except:
                        next_page_click = driver.find_element(By.XPATH,'//*[@id="mNext"]/i')
                        next_page_click.click()
            
                print(web_key,'verified. Begin filling...')
                quantity_field.send_keys(Keys.CONTROL + "a")
                quantity_field.send_keys(int(data_row['Quantity'].item()))
    
    print('\nData filling process complete!\nPress Enter to close the program...')


            
        



