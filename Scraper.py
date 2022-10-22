import pandas as pd
from openpyxl import load_workbook
from playwright.sync_api import Playwright, sync_playwright, expect
import time

def run(playwright: Playwright, link) -> None:
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    citb = ""
    citf = ""

    page = context.new_page()
    page.goto(link, wait_until="networkidle")

    c = 1
    pc = 1
    for i in range (1, 500):
        try:
            if(pc > 20):
                page.locator('//*[@id="citnResult"]/div[1]/div[2]/span[3]/a/img').click()
                time.sleep(1)
                c+=1
                pc = 1
            a = page.query_selector('//*[@id="citnResult"]/div[3]/table/tbody/tr['+ str(pc) +']/td[6]').inner_text() #get backward, forward info
            txt = page.query_selector('//*[@id="citnResult"]/div[3]/table/tbody/tr['+ str(pc) +']/td[2]/a').inner_text() #get patent number
            if a == 'B1,F1':
                citb = citb + txt + "; "
                citf = citf + txt + "; "
            elif a == 'B1':
                citb = citb + txt + "; "
            elif a == 'F1':
                citf = citf + txt + "; " #check backward, forward
            pc += 1
        except:
            break

    context.close()
    browser.close()
    return citb, citf

xls = pd.read_excel('mod2.xlsx', engine = 'openpyxl') #open excel file with patent number and link 

data = pd.DataFrame(columns = ['Pat_No', 'Link', 'Backward', 'Forward'])    #dataframe to save result
link = "Empty"
for i in range(1305):  
    try:
        link = xls.iloc[i, 0]
        no = xls.iloc[i, 1]
        with sync_playwright() as playwright:
            cb, cf = run(playwright, link) #to function

        if len(cb) != 0:
            cb = cb + no
        if len(cf) != 0:
            cf = cf[:-2]

        df = pd.DataFrame({'Pat_No': [no],
                            'Link': [link],
                            'Backward': [cb],
                            'Forward': [cf],})  # add to dataframe
        data = pd.concat([data,df], ignore_index=True)
        link = "Empty"
    except:
        print(link)
    print(i) #cycle through list in excel file

data.to_excel('datav3.xlsx') #save dataframe as .xlsx