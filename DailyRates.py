#!/usr/bin/env python
# coding: utf-8

# In[4]:


# https://medium.com/analytics-vidhya/scraping-tables-from-a-javascript-webpage-using-selenium-beautifulsoup-and-pandas-cbd305ca75fe

# https://stackoverflow.com/questions/66782145/no-tables-found-using-selenium

# run this script in VBA: https://stackoverflow.com/questions/18135551/how-to-call-python-script-on-excel-vba


def main():
    import pandas as pd
    from selenium import webdriver
    from bs4 import BeautifulSoup
    import html5lib
    import lxml
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    options = Options()
    options.headless = True

    driver = webdriver.Chrome(options=options)
    driver.get("https://www.bnz.co.nz/personal-banking/international/exchange-rates")
    table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "table"))
    )
    # driver.implicitly_wait(1000)

    soup = BeautifulSoup(driver.page_source, "lxml")

    tables = soup.find_all("table")

    dfs = pd.read_html(str(tables))
    # print(dfs[1])
    print(dfs[1])
    # driver.close()
