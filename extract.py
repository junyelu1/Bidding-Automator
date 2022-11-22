import pandas as pd
import numpy as np
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
import time


def downloadECPFile(searchTerm="协议库存"):

    # Initiate Google Chrome Driver
    service = ChromeService(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    try:
        # Open ECP2.0 Website and Manuveur to Relevant Pages
        driver.get('https://ecp.sgcc.com.cn/ecp2.0/portal/#/')
        time.sleep(5)

        # Subject to change if they changed the website
        bidSection = driver.find_element(
            By.XPATH, "/html/body/app-root/app-main/app-nav/div/div/ul/li[2]/ul/li[2]/a")
        bidSectionLink = bidSection.get_attribute("href")
        driver.get(bidSectionLink)
        time.sleep(3)

    except:
        return "ECP2.0 Website Layout Mightbe changed, Examine XPATH"

    try:
        # Search relevant keys
        searchBox = driver.find_element(By.NAME, "key")
        searchBox.send_keys()
        searchBox.submit()
        time.sleep(1)

        # If result not found, return error


def filterDownloadExcel(downloadPath: str, savePath='~/Desktop/相关清单/'):
    '''
    Filter Excel files containing multiple products downloaded from ECP 2.0
    '''

    # Input Type check
    assert ".xlsx" in downloadPath, f"Downloaded FilePath given is not Excel"

    try:
        # Import Files into Pandas DataFrame
        downloadFile = pd.read_excel(downloadPath, sheet_name=None)

        # Future Expansions possible into other products
        relatedProd = ['架空绝缘导线', '集束绝缘导线', '架空线']
        prodList = [*downloadFile]

        fileName = ""
        for key in prodList:
            if not any(prod in key for prod in relatedProd):
                file = downloadFile.pop(key)
                if not fileName:
                    # Extract fileName from Owner
                    fileName = list(file['项目单位'])[0].split(
                        '国网')[1].split('电力')[0]

        if not downloadFile:
            return 'No related subprojects present in this file.'

    except:
        return "Erros occured during extraction Process."

    # Writing Found Sheets into new sheets
    try:
        with pd.ExcelWriter(savePath + fileName + '相关清单.xlsx') as writer:
            for key in downloadFile:
                downloadFile[key].to_excel(writer, sheet_name=key, index=False)
        return f"Subprojects successully filtered. {len(downloadFile)} related subprojects found."

    except:
        return "Errors occured during writing process."
