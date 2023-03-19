from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import xlsxwriter
from selenium.webdriver import ActionChains
import pyautogui
from pynput.keyboard import Key, Controller
keyboard = Controller()
from tkinter import *
import tkinter as tk
global driver
def run_chrome():
    global driver

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--user-data-dir=C:\\Users\\Darakhsha\\AppData\\Local\\Google\\Chrome\\User Data")
    driver = webdriver.Chrome(executable_path="C:\chromedriver.exe", options=chrome_options)
    driver.get("https://www.linkedin.com/")
    driver.maximize_window()






def companyEmployes():

    nameSearch=Cname.get()
    outWorkbook = xlsxwriter.Workbook('data.xlsx')
    outSheet = outWorkbook.add_worksheet()
    outSheet.write("A1", "Name")
    outSheet.write("B1", "Profile link")
    driver.find_element_by_xpath("//input[@placeholder='Search']").send_keys(nameSearch)
    driver.find_element_by_xpath("//input[@placeholder='Search']").send_keys(Keys.RETURN)
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "(//a[@class='app-aware-link '])[2]"))).click()
    # driver.find_element_by_xpath("(//a[@class='app-aware-link '])[2]").click()
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "(//div[@class='mt1']/div/a)[2]/span"))).click()
    # driver.find_element_by_xpath("(//div[@class='mt1']/div/a)[2]/span").click()
    page=pageTnum.get()
    row=1
    for i in range(0,page) :
        profs = driver.find_elements_by_xpath("*//div/span[1]/span/a[@class='app-aware-link ']")
        pName = driver.find_elements_by_xpath("*//a[@class='app-aware-link ']/span/span[@aria-hidden='true']")
        print(len(pName))


        for data in  profs:
        #     print(pName[index].text," ", profs[index].get_attribute('href'))
            if('LinkedIn Member' in data.text):
                pass
            else:
                outSheet.write(row,0,data.text)
                outSheet.write(row,1,data.get_attribute('href'))
                print(profs.index(data)," ",data.text , data.get_attribute('href'))
                row+=1
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//button[@ aria-label='Next']"))).click()
        time.sleep(3)
    outWorkbook.close()


def jobColector():
    outWorkbook = xlsxwriter.Workbook('data.xlsx')
    outSheet = outWorkbook.add_worksheet()
    outSheet.write("A1", "Company Name")
    outSheet.write("B1", "Job Role ")
    outSheet.write("C1", "Job link")
    outSheet.write("D1", "Job location")
    skill= JSkill.get()
    location=Jloc.get()
    action = ActionChains(driver)
    action2 = ActionChains(driver)

    driver.find_element_by_xpath("//li-icon[@type='job']").click()


    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "*//input[@aria-label='Search by title, skill, or company']"))).send_keys(skill)
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "*//input[@aria-label='Search by title, skill, or company']"))).send_keys(Keys.TAB)
    time.sleep(3)
    elm=WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "*//input[@placeholder='City, state, or zip code']")))
    # WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "*//input[@placeholder='City, state, or zip code']"))).send_keys(Keys.RETURN)

    action.move_to_element(elm).send_keys(location).perform()
    time.sleep(3)

    # action.move_to_element(elm).send_keys(Keys.ENTER).perform()
    elm2=WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "*//input[@aria-label='Search by title, skill, or company']")))
    action2.move_to_element(elm2).send_keys(Keys.RETURN).perform()

    NumPage=pageTnum.get()
    jRow=1
    for record in range(2,NumPage):
        time.sleep(3)
        jobcompname=driver.find_elements_by_xpath('//div[@class="job-card-container__company-name"]')
        jobRole = driver.find_elements_by_xpath('//div[@class="full-width artdeco-entity-lockup__title ember-view"]/a')
        tjob=len(jobRole)
        for index in range(0,tjob):
            print(jobRole[index].text," company name ", jobcompname[index].text," link ", jobRole[index].get_attribute('href'))
            outSheet.write(jRow,0,jobcompname[index].text)
            outSheet.write(jRow,1,jobRole[index].text)
            outSheet.write(jRow,2,jobRole[index].get_attribute('href'))
            outSheet.write(jRow,3,location)
            jRow+=1
        print("Page "+str(record))
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "*//button[@aria-label="+"'Page "+str(record)+"']"))).click()
    outWorkbook.close()
emplCheck=0
def run_ALl():
    run_chrome()
    print(emplCheck)
    if(emplCheck>0):
        companyEmployes()
    else:
        jobColector()
    print("Task completed")
def sel():
    global emplCheck
    emplCheck=1
def sel2():
    global emplCheck
    emplCheck=0
root = Tk()
root.geometry("900x400")
root.title("Bot By Darakhsha_")
root['background'] = '#8C65D3'

# credential


radio = IntVar()
Cname=StringVar()
JSkill=StringVar()
Jloc = StringVar()
pageTnum=IntVar()
Label(root, text="Smart Linkedin BOT", font="Roboto,14,bold", bg="#8C65D3", fg="white").grid(row=1,column=4)
Radiobutton(root, text="Employe Collector", variable=radio, value=1,command=sel).grid(row=3,column=1)
Label(root, text="Company Name", font="Roboto,14,bold", bg="#8C65D3", fg="white").grid(row=4,column=1)
Entry(root, textvariable=Cname).grid(row=4,column=2)
Radiobutton(root, text="Job finder", variable=radio, value=0,command=sel2).grid(row=5,column=1)
Label(root, text="Job Skill", font="Roboto,14,bold", bg="#8C65D3", fg="white").grid(row=6,column=1)
Entry(root, textvariable=JSkill).grid(row=6,column=2)
Label(root, text="Job location", font="Roboto,14,bold", bg="#8C65D3", fg="white").grid(row=7,column=1)
Entry(root, textvariable=Jloc).grid(row=7,column=2)
Label(root, text="Page number", font="Roboto,14,bold", bg="#8C65D3", fg="white").grid(row=8,column=1)
Entry(root, textvariable=pageTnum).grid(row=8,column=2)

# alignment of layout






Button(text="Run Bot", command=run_ALl).grid(row=9,column=2)
root.mainloop()



