# Import Library
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

import time
import pandas as pd

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome("chrome_driver\chromedriver.exe", options=options)

# 메인 홈페이지로 이동
driver.get("https://www.koreabaseball.com/")

# 기록 페이지로 이동
css_selector = "#lnb > li:nth-child(3)"
elem = driver.find_element(By.CSS_SELECTOR, css_selector)
elem.click()

# 기록실 페이지로 이동
css_selector = "#lnb02"
elem = driver.find_element(By.CSS_SELECTOR, css_selector)
elem.click()

# 기록실 옵션(타자, 투수, 수비, 주루)
options = ['hitter', 'pitcher', 'defense', 'runner']
record_tab_xpath = '//*[@id="contents"]/div[2]/div[2]/ul/*'

for i in range(0, 4):
    # 현재 기록실 옵션 상태
    status = options[i]
    
    # 원하는 기록실 옵션 클릭
    record_tab = driver.find_elements(By.XPATH, record_tab_xpath)
    record_tab[i].click()
    
    time.sleep(3)

    # 팀 선택
    select = Select(driver.find_element(By.NAME, "ctl00$ctl00$ctl00$cphContents$cphContents$cphContents$ddlTeam$ddlTeam"))
    select.select_by_value(value='OB')

    time.sleep(3)

    # 데이터 가져오기 & 가공
    table_value = driver.find_element(By.CLASS_NAME, 'tData01').text.split('\n')
    table_head = table_value.pop(0).split(' ')

    ## 페이지 이동 및 데이터 가져오기 & 가공
    pages = driver.find_element(By.CLASS_NAME, 'paging').text.split(' ')
    if len(pages) == 1:
        pass
    else:
        for page in range(1, len(pages)):
            id_selector = 'cphContents_cphContents_cphContents_ucPager_btnNo' + str(page+1)
            elem = driver.find_element(By.ID, id_selector)
            elem.click()

            time.sleep(3)

            temp = driver.find_element(By.CLASS_NAME, 'tData01').text.split('\n')
            del temp[0]
            table_value.extend(temp)

    # 선수 기록 데이터 가공
    record_values = []

    for item in table_value:
        record_values.append(item.split(' '))
    
    if status == 'pitcher':
        for item in record_values:
            if len(item) == len(table_head):
                pass
            else:
                temp1 = item.pop(10)
                temp2 = item.pop(10)
                res_temp = temp1 + ' ' + temp2
                item.insert(10, res_temp)

    if status == 'defense':
        for item in record_values:
            if len(item) == len(table_head):
                pass
            else:
                temp1 = item.pop(6)
                temp2 = item.pop(6)
                res_temp = temp1 + ' ' + temp2
                item.insert(6, res_temp)
    
    # 데이터 -> 데이터 프레임 형태로 변환
    globals()[f'{status}_dataframe'] = pd.DataFrame(columns=table_head)
    
    for j in range(0, len(record_values)):
        globals()[f'{status}_dataframe'].loc[j] = record_values[j]

    time.sleep(2)

# 엑셀 파일로 저장
with pd.ExcelWriter('test.xlsx') as writer:
    for status in options:
        globals()[f'{status}_dataframe'].to_excel(writer, sheet_name=status, index=False)

# 창 닫기
driver.close()