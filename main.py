from selenium import webdriver
import pandas as pd
from openpyxl.workbook import Workbook

PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

driver.get("https://www.barchart.com/stocks/quotes/$SPX/put-call-ratios")
print(driver.title)
driver.maximize_window()

dtes = driver.find_elements_by_xpath('//div/div[3]/div/div[2]/div/div/ng-transclude/table/tbody/tr/td[2]')
put_volume = driver.find_elements_by_xpath('//div/div[3]/div/div[2]/div/div/ng-transclude/table/tbody/tr/td[3]')
call_volume = driver.find_elements_by_xpath('//div/div[3]/div/div[2]/div/div/ng-transclude/table/tbody/tr/td[4]')
put_call_vol_ratio = driver.find_elements_by_xpath('//div/div[3]/div/div[2]/div/div/ng-transclude/table/tbody/tr/td[5]')
put_oi = driver.find_elements_by_xpath('//div/div[3]/div/div[2]/div/div/ng-transclude/table/tbody/tr/td[6]')
call_oi = driver.find_elements_by_xpath('//div/div[3]/div/div[2]/div/div/ng-transclude/table/tbody/tr/td[7]')
put_call_oi_ratio = driver.find_elements_by_xpath('//div/div[3]/div/div[2]/div/div/ng-transclude/table/tbody/tr/td[8]')
avg_ATM_volatility = driver.find_elements_by_xpath('//div/div[3]/div/div[2]/div/div/ng-transclude/table/tbody/tr/td[9]')


Table = []

for i in range(len(dtes)):
    temporary_data = {'DTE': dtes[i].text,
                      'Put Volume':put_volume[i].text,
                      'Call Volume': call_volume[i].text,
                      'Put/Call Vol Ratio': put_call_vol_ratio[i].text,
                      'Put OI': put_oi[i].text,
                      'Put/Call OI Ratio': put_call_oi_ratio[i].text,
                      'Avg ATM Volatility': avg_ATM_volatility[i].text}
    Table.append(temporary_data)

df_data = pd.DataFrame(Table)
print(df_data)

print('The table and the results from scraping the website were saved in an excel file')

df_data.to_excel('S&P_500_Index_($SPX).xlsx', index=False)


driver.quit()