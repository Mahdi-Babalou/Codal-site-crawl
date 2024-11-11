#Import pakage
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import openpyxl
import time

#Driver settings
options = webdriver.ChromeOptions()
#options.add_argument('--headless')  # Run without displaying the browser window
driver = webdriver.Chrome(options=options)

#List of symbols
symbols = ["سامان", "وکار", "وبملت", "وپارس", "وپاسار", "ونوین", "وخاور", "وبصادر", "وتجارت", 
           "وپست", "دی", "وسالت" ]

#Create a new Excel file
wb = openpyxl.Workbook()

#loop for each symbol
for symbol in symbols:
    ws = wb.create_sheet(title=symbol)
    headers = ["Date", "فروش اقساطي", "جعاله", "مرابحه", 
               "اجاره به شرط تمليک", "سلف", "مضاربه", "مشارکت مدني", 
               "خريد دين", "بدهکاران بابت اعتبارات اسنادی پرداخت شده", 
               "بدهکاران بابت ضمانت نامه های پرداخت شده", 
               "تسهيلات ارزی", "تسهيلات قرض الحسنه", "ساير تسهيلات", "جمع"]
    
    ws.append(headers)    #Adding headers

    url = f'https://codal.ir/ReportList.aspx?search&Symbol={symbol}&LetterType=58&FromDate=1400%2F09%2F30&AuditorRef=-1&PageNumber=1&Audited&NotAudited&IsNotAudited=false&Childs&Mains&Publisher=false&CompanyState=-1&Category=-1&CompanyType=-1&Consolidatable&NotConsolidatable'
    driver.get(url)

    #Find links related to financial reports
    try:
        elements = WebDriverWait(driver, 20).until(
            EC.visibility_of_all_elements_located((By.XPATH, '//a[@class="icon icon-file-eye ng-scope"]'))
        )
        print(f"Found {len(elements)} report links for {symbol}.")
    except TimeoutException:
        print("Failed to find report links. Please check the URL or XPath.")
        continue   #Continue to the next symbol if links are not found

    original_window = driver.current_window_handle

    for element in elements:
        link = element.get_attribute('href')
        driver.execute_script("window.open(arguments[0]);", link)

        #Wait for the new tab to open by checking the handle count
        original_handles_count = len(driver.window_handles)
        timeout = 0.5 # Increase wait time if needed
        start_time = time.monotonic()

        while len(driver.window_handles) <= original_handles_count:
            if time.monotonic() - start_time > timeout:
                print("Timeout while waiting for a new window to open.")
                break  #Exit loop if timeout is reached
            time.sleep(0.5)  #Sleep briefly before checking again

    for handle in driver.window_handles:
        if handle != original_window:
            driver.switch_to.window(handle)
            try:
                date_str = driver.find_element(By.XPATH, '/html/body/form/div[6]/div[3]/div[2]/div/span[5]/bdo').text
                ws.append([date_str])  #Start a new row with the date
                for i in range(1, len(headers)):
                    remain = driver.find_element(By.XPATH, f'/html/body/form/div[13]/app-root/app-statement/div/div/div/app-sheet/div[2]/div/app-table/table/tbody/tr[{i}]/td[7]').text
                    ws.cell(row=ws.max_row, column=i + 1, value=remain)
            except Exception as e:
                print(f"Error extracting data: {e}")
            driver.close()  #Close the current tab

    driver.switch_to.window(original_window)  #Return to the main window

#Delete the default sheet
if 'Sheet' in wb.sheetnames:
    del wb['Sheet']

#Save the Excel file
try:
    wb.save("Banks.xlsx")
    print("Excel file saved successfully.")
except Exception as e:
    print(f"Error saving the Excel file: {e}")

#Close the driver
driver.quit()
