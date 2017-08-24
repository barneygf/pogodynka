'''This program read meteo data from site pogodynka.pl for one chosen station in Poland and write it to the Excel
file.'''

import selenium.webdriver as webdriver
from selenium.webdriver.common.by import By
#from selenium.webdriver.chrome import service
import pandas as pd
import time
import re

# Attempt to run this program with Opera browser using operachromiumdriver
#webdriver_service = service.Service('operadriver.exe')
#webdriver_service.start()
#driver = webdriver.Remote(webdriver_service.service_url, webdriver.DesiredCapabilities.OPERA)

driver = webdriver.Chrome()  # Run Chrome browser
url = 'http://monitor.pogodynka.pl/#station/meteo/249200180'
driver.get(url)  # Open site with data
time.sleep(5)  # Wait a few seconds because website contains much AJAX so Selenium doen't know if the website is fully
# loaded

raw_table = driver.find_element_by_xpath("//table[@class='table table-striped table-responsive table-bordered']"
                                         ).get_attribute('outerHTML')  # Find table in the website code and scrap it
driver.close()  # Close browser
table_temp = raw_table.split('</tbody', maxsplit=1)  # Choose part with data from the scrapped table
table_text = '<table>' + table_temp[1]  # In the previous row, first <table> tag was deleted, now we add it

header_temp = table_temp[0]  # Prepare data for header
headers = re.findall('">(.*?)</', header_temp)  # Find names for column headers

dataframe_table = pd.read_html(table_text)  # Convert scrapped html code to Pandas DataFrame
dataframe_table[0].columns = headers  # Add columns headers
print(dataframe_table)

# Write table to Excel file
output_filename = 'output1.xlsx'  # Remember about extension!
writer = pd.ExcelWriter(output_filename)
dataframe_table[0].to_excel(writer)
writer.save()