# Import Library
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

import time
import pandas as pd

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome("chrome_driver\chromedriver.exe", options=options)

# 홈페이지로 이동
driver.get("https://www.koreabaseball.com/")

# 기록 페이지로 이동
css_selector = "#lnb > li:nth-child(3)"
elem = driver.find_element(By.CSS_SELECTOR, css_selector)
elem.click()

# 기록실 페이지로 이동
css_selector = "#lnb02"
elem = driver.find_element(By.CSS_SELECTOR, css_selector)
elem.click()

# 팀 선택
select = Select(driver.find_element(By.NAME, "ctl00$ctl00$ctl00$cphContents$cphContents$cphContents$ddlTeam$ddlTeam"))
select.select_by_value(value='OB')

time.sleep(3)

# 데이터 가져오기
thead_xpath = '//*[@id="cphContents_cphContents_cphContents_udpContent"]/div[3]/table/thead/tr'
tbody_xapth = '//*[@id="cphContents_cphContents_cphContents_udpContent"]/div[3]/table/tbody'

table_head = driver.find_element(By.XPATH, thead_xpath).text.split(' ')
table_value = driver.find_element(By.XPATH, tbody_xapth).text.split('\n')

# 2페이지 이동
id_selector = 'cphContents_cphContents_cphContents_ucPager_btnNo2'
elem = driver.find_element(By.ID, id_selector)
elem.click()

time.sleep(3)

# 2페이지 데이터 추가
temp = driver.find_element(By.XPATH, tbody_xapth).text.split('\n')
table_value.extend(temp)

# 데이터 리스트로 변환
record_values = []

for item in table_value:
    record_values.append(item.split(' '))

# 데이터프레임 생성
doosan_dataframe = pd.DataFrame(columns=table_head)

for i in range(0, len(record_values)):
    doosan_dataframe.loc[i] = record_values[i]

# 엑셀 파일로 저장
doosan_dataframe.to_excel('test.xlsx', index=False)

# 창 닫기
driver.close()