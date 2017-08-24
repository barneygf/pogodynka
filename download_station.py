import selenium.webdriver as webdriver
from selenium.webdriver.common.by import By
#from selenium.webdriver.chrome import service
import pandas as pd
import time
import re

#webdriver_service = service.Service('operadriver.exe')
#webdriver_service.start()
#driver = webdriver.Remote(webdriver_service.service_url, webdriver.DesiredCapabilities.OPERA)
driver = webdriver.Chrome()  # uruchamianie przegladarki
driver.get('http://monitor.pogodynka.pl/#station/meteo/249200180')  # otwieranie strony
time.sleep(5)

raw_table = driver.find_element_by_xpath("//table[@class='table table-striped table-responsive table-bordered']").get_attribute('outerHTML')
driver.close()
table_temp = raw_table.split('</tbody', maxsplit=1)
table_text = '<table>' + table_temp[1]

header_temp = table_temp[0]
headers = re.findall('">(.*?)</', header_temp)

dataframe_table = pd.read_html(table_text)
dataframe_table[0].columns = headers
print(dataframe_table)

writer = pd.ExcelWriter('output1.xlsx')
dataframe_table[0].to_excel(writer)
writer.save()