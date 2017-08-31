''' Downloading meteorological data from site monitor.pogodynka.pl for three stations and save it to Excel files.
If Excel file exists, script does not delete already existed data, but append new sheet to existed file. So do not
worry about overwriting.
Python 3. Necessary packages:
    * Selenium
    * NumPy
    * Pandas
    * OpenPyXL
Version 2.0, 2017-08-30
Maciej Barnaś, e-mail: maciej.michal.barnas(at)gmail.com '''

import selenium.webdriver as webdriver
import pandas as pd
import time
import re
from openpyxl import load_workbook
from pathlib import Path

################################################################################

# Input - fill urls of websites and output Excel filenames
Siercza_url = 'http://monitor.pogodynka.pl/#station/meteo/249201010'
Siercza_output = 'Siercza1.xlsx'
Wronowice_url = 'http://monitor.pogodynka.pl/#station/meteo/249200990'
Wronowice_output = 'Wronowice1.xlsx'
Limanowa_url = 'http://monitor.pogodynka.pl/#station/meteo/249200180'
Limanowa_output = 'Limanowa1.xlsx'

################################################################################

driver = webdriver.Chrome()  # Driver must be in the same folder as a script

# Three rows for Opera browser
#webdriver_service = service.Service('operadriver.exe')
#webdriver_service.start()
#driver = webdriver.Remote(webdriver_service.service_url, webdriver.DesiredCapabilities.OPERA)

nowdate = time.strftime('%Y-%m-%d %H%M%S')  # Now date - for name of sheet in Excel files

def download_and_save(url, output_filename):
    driver.get(url)  # Open website
    time.sleep(10)  # Website is loading, script waits

    raw_table = driver.find_element_by_xpath("//table[@class='table table-striped table-responsive table-bordered']").\
        get_attribute('outerHTML')  # Find table with results

    # Prepare scrapped html code for making Pandas DataFrame from it
    table_temp = raw_table.split('</tbody', maxsplit=1)
    table_text = '<table>' + table_temp[1]
    header_temp = table_temp[0]
    headers = re.findall('">(.*?)</', header_temp)

    dataframe_table = pd.read_html(table_text)  # Make DataFrame from html code
    dataframe_table[0].columns = headers  # Fill in headers
    print(dataframe_table)

    # Check if output file exists. If no, make it and save data to it. If yes, add new sheet to it. Name of sheet
    # will be now date.
    output_path = Path(output_filename)
    if output_path.is_file():
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            writer.book = load_workbook(output_filename)
            dataframe_table[0].to_excel(writer, nowdate)
    else:
        writer = pd.ExcelWriter(output_filename)
        dataframe_table[0].to_excel(writer, sheet_name=nowdate)
        writer.save()

# Run function for every station
download_and_save(Siercza_url, Siercza_output)
download_and_save(Wronowice_url, Wronowice_output)
download_and_save(Limanowa_url, Limanowa_output)

driver.close()  # Close browser