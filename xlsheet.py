# importing webbrowser python module
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter

service =Service(executable_path="chromedriver.exe")
driver=webdriver.Chrome(service=service)

start = 1201
end = 1264
index=1
l=[]
l = [f"21JJ1A{i}" for i in range(start, end+1)]
  
workbook=xlsxwriter.Workbook("Marks_sheet.xlsx")
worksheet=workbook.add_worksheet("firstSheet")
worksheet.write(0,0,"S.No.")
worksheet.write(0,1,"Roll Number")
worksheet.write(0,2,"Name")
worksheet.write(0,3,"1-1")
worksheet.write(0,4,"1-2")
worksheet.write(0,5,"2-1")
worksheet.write(0,6,"2-2")
worksheet.write(0,7,"3-1")
worksheet.write(0,8,"CGPA")

for x in l:

    driver.get("https://jntuhresults.vercel.app/academicresult")
    input_element=driver.find_element(By.XPATH, "//input[contains(@placeholder,'Enter your hallticket no')]")
    input_element.clear()
    input_element.send_keys(x + Keys.ENTER)

    link=driver.find_element(By.XPATH,"//button[text()='Result']")
    link.click()


    cgpa = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '.dark\\:border-white.w-\\[50\\%\\].text-green-800'))
    )

    name = driver.find_element(By.XPATH,'//*[@class="dark:border-white"]').text
    results=driver.find_elements(By.XPATH,'//*[@class="dark:border-white w-[25%]"]')
    r_11 = results[0].text
    r_12 = results[1].text
    r_21 = results[2].text
    r_22 = results[3].text
    r_31 = results[4].text
   
    cgpa = driver.find_element(By.XPATH,'//*[@class="dark:border-white w-[50%] text-green-800"]')
    
    cgpa=cgpa.text
    worksheet.write(index,0,index)
    worksheet.write(index,1,x)
    worksheet.write(index,2,name)
    worksheet.write(index,3,r_11)
    worksheet.write(index,4,r_12)
    worksheet.write(index,5,r_21)
    worksheet.write(index,6,r_22)
    worksheet.write(index,7,r_31)
    worksheet.write(index,8,cgpa)
    index=index+1
    


workbook.close()

