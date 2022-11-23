from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
import time


def bidProjectSearch(searchTerm="协议库存"):

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

    # Search relevant keys
    searchBox = driver.find_element(By.NAME, "key")
    searchBox.send_keys()
    searchBox.submit(searchTerm)
    time.sleep(1)

    # If result not found, return error
    try:
        driver.find_element(By.XPATH, "//page/div")
    except:
        return f"Search Term {searchTerm} entered didn't match any results."
