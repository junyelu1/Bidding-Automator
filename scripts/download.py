from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time


def bidProjectSearch(searchTerm="协议库存"):

    # Initiate Google Chrome Driver, and allow page to stay open
    service = ChromeService(executable_path=ChromeDriverManager().install())
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    chrome_options.headless = False
    driver = webdriver.Chrome(service=service, options=chrome_options)

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
        return "ECP2.0 Website Layout might have changed, Examine XPATH"

    # Search relevant keys
    searchBox = driver.find_element(By.NAME, "key")
    searchBox.send_keys(searchTerm)
    searchBox.submit()
    time.sleep(1)

    # If result not found, return error
    try:
        driver.find_element(By.XPATH, "//page/div")
    except:
        return f"Search Term {searchTerm} entered didn't match any results."

    try:
        # Return top result, could be changed laster to accomodate other download requests
        topresult = driver.find_element(
            By.XPATH, "//app-main/app-list/div/app-list-spe/div/div/div/table/tbody/tr[1]")
        topresult.click()
        currentpage = driver.current_window_handle
        openpages = driver.window_handles

        # Switch to new page
        for page in openpages:
            if (page != currentpage):
                driver.switch_to.window(page)
        time.sleep(3)

        driver.find_element(By.XPATH, "//*[text() = '下载公告文件']").click()

    except:
        return f"Error occured during download phase."


if __name__ == "main":
    bidProjectSearch("国网湖北省电力有限公司2022年第二次配网物资协议库存招标采购")
