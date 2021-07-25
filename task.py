"""Template robot with Python."""

from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

browser = Selenium()

url = 'https://robotsparebinindustries.com/#/'
password = "thoushallnotpass"

def open_browser():
    browser.open_chrome_browser(url)
    browser.maximize_browser_window

def Log_in():
    browser.input_text("id:username", "maria")
    browser.input_text("id:password",password)
    browser.click_button("class:btn-primary")
    browser.wait_until_page_contains_element("id:firstname")

def download_excel():
    download = HTTP()
    download.download("https://robotsparebinindustries.com/SalesData.xlsx")

def excel_manipulation():
    lib = Files()
    lib.open_workbook(r"C:\Users\admin\Desktop\Robocorp\first\SalesData.xlsx")
    try:
        table = lib.read_worksheet_as_table("data",header=True)
        
       
        for row in range(10):
            browser.input_text("id:firstname", table[row][0])
            browser.input_text("id:lastname", table[row][1])
            browser.input_text("id:salesresult", table[row][3])
            browser.click_button("class:btn-primary")

    finally:
        lib.close_workbook()



if __name__ == "__main__":
    
    open_browser()
    Log_in()
    download_excel()
    excel_manipulation()
   



