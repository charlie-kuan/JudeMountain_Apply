# ===================== Description =====================
# Author: Charlie Cheng 
# Email: charliecheng90523@gmail.com
# Version: v 1.0
# Summary: Automatically fill in forms to apply for national park admission
# How to use: run jade_mountain.py <excel file path>
# Example: /usr/bin/python3 /Users/charlie/Desktop/jade_mountain.py /Users/charlie/Desktop/example.xlsx
# =======================================================
import time
import sys, os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

def page(file_name):
    # import data from file
    member = pd.read_excel('/Users/charlie/Desktop/example.xlsx', sheet_name = '隊員資料')
    member.dropna(subset=['姓名'], inplace = True)
    member.fillna("",inplace = True)

    infor_df = pd.read_excel(file_name, sheet_name = '隊伍資訊').fillna("")
    infor = {}
    for i in infor_df.index:
        infor[infor_df.key[i]] = infor_df.value[i]

    print('歡迎使用「玉管處請加油」～～')
    print('請使用 example.xlsx 填寫隊伍資訊及隊員資料。\n')
    print("==========  Disclaimer ============")
    print("登山活動必有風險，請注意自身安全，並遵守森林法、國家公園法及其他相關登山活動法規。")
    time.sleep(1)

    # driver initialization
    
    options = Options()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=options) 
    start_website = "https://npm.cpami.gov.tw/apply_1_2.aspx?unit=c951cdcd-b75a-46b9-8002-8ef952ec95fd"
    driver.get(start_website)
    time.sleep(2)


    # check all boxes 
    boxes = driver.find_elements(By.NAME, "chk[]")
    for box in boxes:
        if not box.is_selected():
            box.click()


    # go next page
    driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnagree").click()
    time.sleep(3)

    # fill in information
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$teams_name').send_keys(infor['隊名'])

    select_main_path = Select(driver.find_element(By.ID, "ContentPlaceHolder1_climblinemain"))
    select_main_path.select_by_index(1) # 玉山線
    time.sleep(2)
    
    select_sub_path = Select(driver.find_element(By.ID, "ContentPlaceHolder1_climbline"))
    select_sub_path.select_by_index(3) #次路線 2-5 天
    time.sleep(2)
    
    select_days = Select(driver.find_element(By.ID, "ContentPlaceHolder1_sumday"))
    select_days.select_by_index(int(infor['登山總天數'])-1) # 兩天 -> index = 1 ; 三天 -> index = 2
    time.sleep(2)
    
    select_start = Select(driver.find_element(By.ID, "ContentPlaceHolder1_applystart"))
    select_start.select_by_value(infor['入園日期'])
    time.sleep(2)

    select_start = Select(driver.find_element(By.ID, "ContentPlaceHolder1_seminar"))
    select_start.select_by_index(1) # 線上講習

    select_start = Select(driver.find_element(By.ID, "ContentPlaceHolder1_gps"))
    select_start.select_by_index(1) # GPS 是

    #driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnToStep22").click()

    os.system('clear')
    print('\n')
    print('請自行設計規劃路線')
    print('請檢查資料是否正確，確認後請按下一頁')
    print('進到資料填寫頁後，請至terminal按下enter鍵')
    input()
    os.system('clear')
    time.sleep(1)
    
    driver.find_element(By.LINK_TEXT, "申請人資料").click()
    time.sleep(1)
    
    
    # 申請人資料
    # agree button
    apply_box = driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$applycheck")
    if not apply_box.is_selected():
            apply_box.click()
    time.sleep(1)
    
    # applicant information
    applicant_data = member.iloc[0].to_dict()
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_name').send_keys(applicant_data['姓名'])
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_tel').send_keys(applicant_data['電話'])

    select_county_1 = Select(driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ddlapply_country"))
    select_county_1.select_by_visible_text(applicant_data['縣市'])
    time.sleep(0.5)
    
    select_zone_1 = Select(driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ddlapply_city"))
    select_zone_1.select_by_visible_text(applicant_data['鄉鎮市區'])

    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_addr').send_keys(applicant_data['地址'])
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_mobile').send_keys(applicant_data['手機'])
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_email').send_keys(applicant_data['電子郵件'])

    select_national_1 = Select(driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$apply_nation"))
    select_national_1.select_by_visible_text("中華民國")
    time.sleep(0.5)
    
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_sid').send_keys(applicant_data['身分證字號'])
    time.sleep(1)
    
    select_sex_1 = Select(driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_sex'))
    sex_id = 1 if applicant_data['性別'] == '男' else 2   
    select_sex_1.select_by_index(sex_id)
    time.sleep(1)
    
    birthday = str(applicant_data["生日"])
    driver.execute_script("document.getElementById('ContentPlaceHolder1_apply_birthday').value ='"+ birthday +"'")

    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_contactname').send_keys(applicant_data['緊急聯絡人姓名'])
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$apply_contacttel').send_keys(applicant_data['緊急聯絡人電話'])


    # 領隊資料
    driver.find_element(By.LINK_TEXT, "領隊資料").click()
    time.sleep(1)
    
    copy_box = driver.find_element(By.ID, "ContentPlaceHolder1_copyapply")
    if not copy_box.is_selected():
            copy_box.click()

    # 隊員資料
    driver.find_element(By.LINK_TEXT, "隊員資料").click()
    time.sleep(1)

    total_member = member.shape[0] - 1
    print("隊員人數（不含領隊）：", total_member)
    
    for count in range(1,total_member + 1):
        member_data = member.iloc[count].to_dict()
        # 新增
        driver.find_element(By.ID, "ContentPlaceHolder1_lbInsMember").click()
        time.sleep(0.5)
        driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_name').send_keys(member_data['姓名'])
        driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_tel').send_keys(member_data['電話'])

        select_county_mem = Select(driver.find_element(By.NAME, f"ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$ddlmember_country"))
        select_county_mem.select_by_visible_text(member_data['縣市'])
        time.sleep(0.5)
        
        select_zone_mem = Select(driver.find_element(By.NAME, f"ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$ddlmember_city"))
        select_zone_mem.select_by_visible_text(member_data['鄉鎮市區'])
        time.sleep(0.5)

        driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_addr').send_keys(member_data['地址'])
        driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_mobile').send_keys(member_data['手機'])
        driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_email').send_keys(member_data['電子郵件'])

        select_national_mem = Select(driver.find_element(By.NAME, f"ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_nation"))
        select_national_mem.select_by_visible_text("中華民國")
        time.sleep(0.5)
        
        driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_sid').send_keys(member_data['身分證字號'])
        time.sleep(1)
        
        select_national_mem = Select(driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_sex'))
        sex_id = 1 if member_data['性別'] == '男' else 2   
        select_national_mem.select_by_index(sex_id)

        birthday_mem = str(member_data["生日"])
        driver.execute_script("document.getElementById('ContentPlaceHolder1_lisMem_member_birthday_0').value ='"+ birthday_mem + "'")

        driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_contactname').send_keys(member_data['緊急聯絡人姓名'])
        driver.find_element(By.NAME, f'ctl00$ContentPlaceHolder1$lisMem$ctrl{count}$member_contacttel').send_keys(member_data['緊急聯絡人電話'])

        time.sleep(2)       
    
    driver.find_element(By.LINK_TEXT, "留守人資料").click()
    time.sleep(0.5)
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$stay_name').send_keys(infor['留守人姓名'])
    driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$stay_mobile').send_keys(infor['留守人電話'])

    os.system('clear')
    print("請檢查資料是否正確無誤，確認後請輸入驗證碼")
    print('本程式已結束，感謝您的使用，如有疑問請洽 charliecheng90523@gmail.com')

if __name__ == "__main__":
    page(str(sys.argv[1]))
    



    
    



