from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from time import sleep, time


####  these four are the inputs ####
start = 0  #where to start progress from
end = 1000  #where to end the progress
step = 20   #number of tabs to function simultaneously
batch = 'batch1.xlsx' #batch number


# reads excel file
df = pd.read_excel(batch)
df['Zip'] = df['Zip - 5 digit']

# some options to make chromedriver behave like an actual browser
options = Options()
options.add_argument("--disable-blink-features")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("start-maximized")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
dr = webdriver.Chrome(options=options)

#url to the site
url = 'https://data.hrsa.gov/tools/shortage-area/by-address'


#to store rows of excel file which faced an error
invalid = {
    'ID': [],
    'PK': [],
    'Address': [],
    'ErrorOccured': []
}

# function to dump erros
def dumpInvalid(i, ex):
    invalid['ID'].append(df.ID[i])
    invalid['PK'].append(df.PK[i])
    invalid['Address'].append(df.Address[i])
    invalid['ErrorOccured'].append(ex)


# function to scrape the site
def scrape(source, i):
    soup = BeautifulSoup(source)
    el = str(soup.findAll('div', {'class': 'rural-analyzer-info'}))
    pmc1 = el.find('In a Primary Care HPSA')
    if el[pmc1:pmc1+120].find('Yes') != -1:
        pmc = 'Y'
    elif el[pmc1:pmc1 + 120].find('No') != -1:
        pmc = 'N'
    else:
        pmc = '0'

    pmc1 = el.find('In a MUA/P')
    if el[pmc1:pmc1 + 120].find('Yes') != -1:
        mua = 'Y'
    elif el[pmc1:pmc1 + 120].find('No') != -1:
        mua = 'N'
    else:
        mua = '0'
    print("{} : pmc={} mua={}".format(i, pmc, mua))

    df['Primary Care HPSA?'].loc[i] = pmc
    df['MUA?'].loc[i] = mua



# function to fill the site
def submitForm(i):
    dr.find_element_by_xpath('//input[contains(@id,"inputAddress")]').send_keys(Keys.CONTROL+'a')
    dr.find_element_by_xpath('//input[contains(@id,"inputAddress")]').send_keys(df.Address[i])
    dr.find_element_by_xpath('//input[contains(@id,"inputCity")]').send_keys(Keys.CONTROL+'a')
    dr.find_element_by_xpath('//input[contains(@id,"inputCity")]').send_keys(df.City[i])
    dr.find_element_by_xpath('//select[contains(@id,"ddlState")]').send_keys(df.State[i])
    dr.find_element_by_xpath('//input[contains(@id,"inputZipCode")]').send_keys(Keys.CONTROL+'a')
    dr.find_element_by_xpath('//input[contains(@id,"inputZipCode")]').send_keys(str(df.Zip[i]))
    dr.find_element_by_xpath('//input[contains(@type,"button")]').click()



# function to start over the site
def nextOne():
    for x in range(3):
        try:
            elem = dr.find_element_by_xpath('//button[contains(@id, "btnStartOver")]')
            return True, elem
        except Exception as ex:
            sleep(5)
            if x == 2:
                print(ex)
                return False, ex
            else:
                continue


# main function
if __name__ == '__main__()':

    # to open multiple tabs
    for x in range(step-1):
        dr.get(url)
        dr.execute_script('window.open("{}","_blank");'.format(url))

    #function calling
    try:
        for i in range(start, end+1, step):
            
            for t in range(step):       # submits the form on one tab and then switches to other tabs to fill form
                dr.switch_to.window(dr.window_handles[t])
                submitForm(i+t)

            for t in range(step):       # extracts the tab's html, starts over that page; and then switches to other tabs
                dr.switch_to.window(dr.window_handles[t])
                (nextValid, ex) = nextOne()
                if nextValid:
                    scrape(dr.page_source, i+t)
                    ex.click()
                if not nextValid:
                    dumpInvalid(i+t, ex)

    except Exception as exp:
        print(exp)      # prints any error in case it occurs


    # to store results in excel file
    df.to_excel(batch, index=False)


    #to store error values
    invalid = pd.DataFrame(invalid, columns=['ID', 'PK', 'Address', 'ErrorOccured'])
    invalid.to_csv("inv{}".format(batch), index=False)


    # to store progress
    with open('completed.txt', 'a') as f:
        f.writelines("\nscraped {} from {} to {}\n".format(batch, start, end))
