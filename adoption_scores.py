"""
adoption_scores.py

This module retrieves adoption scores from a list of companies and stores them in an Excel sheet.
"""
import datetime
from playwright.sync_api import sync_playwright
import pandas as pd

# opens the excel sheets to obtain the name and the scores
while True:
    FILEPATH = input("Please enter the filepath to the excel sheet including the file format (.csv or .xlsx): ")
    if ".csv" in FILEPATH:
        try:
            df = pd.read_csv(FILEPATH)
        except:
            print("Incorrect filename or format. Please double check the filepath and filetype.")
            continue
    elif ".xlsx" in FILEPATH:
        try:
            df = pd.read_excel(FILEPATH)
        except:
            print("Incorrect filename or format. Please double check the filepath and filetype.")
            continue
    break

COMPANIES_SHEET = pd.read_excel(FILEPATH, sheet_name = "Report status")
SCORES_SHEET = pd.read_excel(FILEPATH, sheet_name = "Adoption score", index_col = 0)

# reports errors to excel sheet
def reportError(company, msg):
    if 'Error' not in SCORES_SHEET.columns:
        SCORES_SHEET['Error'] = ''
    SCORES_SHEET.loc[company, "Error"] = msg
    with pd.ExcelWriter(FILEPATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        if ".csv" in FILEPATH:
            SCORES_SHEET.to_csv(writer, sheet_name="Adoption score")
        else:
            SCORES_SHEET.to_excel(writer, sheet_name='Adoption score')

# goes to the client's admin centre
def goToTenant(COMPANIES_SHEET, client, company) -> bool:
    
    # goes to the customer list
    page.goto("https://partner.microsoft.com/dashboard/v2/customers/list")
        
    # fill customer name into searchbar
    page.locator("#customer-search-box").get_by_placeholder("Search").fill(COMPANIES_SHEET.loc[client]["Domain"])
    page.wait_for_timeout(5000)
    
    # go to customer profile
    page.get_by_role("radio", name = "Select row", checked = False, disabled = False).check()
    
    # go to service management
    page.get_by_text("Service management").click()
    page.wait_for_timeout(5000)
    
    # go to "Microsoft 365"
    try:
        page.locator('//*[@id="MicrosoftOffice"]').click()
        page.wait_for_timeout(5000)
    except:
        reportError(company, "No admin permissions")
        return True

    return False

# extracts scores from the adoption score page
def getScores(company, scores):
    page.goto("https://admin.microsoft.com/Adminportal/Home#/adoptionscore")
    
    # extracts "Your organization's score"
    SCORES_SHEET.loc[company, 'Your organization’s score'] = page.get_by_text("Your organization’s score: ").text_content()[-3:]
    
    # extracts "Total score"
    SCORES_SHEET.loc[company, 'Total score'] = page.get_by_text("Total score:").all_text_contents()[0][-14:-7]
    page.set_default_timeout(5000)
    
    # extracts everything else
    for link in range(len(scores)):
        if link < 6:
            try:
                SCORES_SHEET.loc[company, scores[link]] = page.locator('//html/body/div[1]/div[1]/div[1]/main/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[2]/div[' + str(link+1) + ']/div/div[1]/span[3]').text_content().strip(" points")
            except:
                SCORES_SHEET.loc[company, scores[link]] = '--'
        else:
            try:
                SCORES_SHEET.loc[company, scores[link]] = page.locator('//html/body/div[1]/div[1]/div[1]/main/div[2]/div[1]/div[2]/div[2]/div/div/div[2]/div[2]/div[' + str(link-5) + ']/div/div[1]/span[3]').text_content().strip(" points")
            except:
                SCORES_SHEET.loc[company, scores[link]] = '--'
    print(SCORES_SHEET)
    with pd.ExcelWriter(FILEPATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        if ".csv" in FILEPATH:
            SCORES_SHEET.to_csv(writer, sheet_name="Adoption score")
        else:
            SCORES_SHEET.to_excel(writer, sheet_name='Adoption score')
                                                           

# list of scores to extract
scores = ['Communication', 'Meetings', 'Content collaboration', 'Teamwork', 'Mobility', 'AI assistance', 'Endpoint analytics',	'Network connectivity', 'Microsoft 365 Apps Health']

# display program start time
print("Program start time: " + str(datetime.datetime.now().strftime("%H:%M:%S")))

for client in range(len(COMPANIES_SHEET["Domain"])): # client is a number
    
    company = COMPANIES_SHEET.loc[client, "Username"]
    
    # adds a new row to the df
    SCORES_SHEET.loc[company] = [None] * len(SCORES_SHEET.columns)

    with sync_playwright() as p:
        # connects to open browser
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        default_context = browser.contexts[0]
        page = default_context.pages[0]
        
        # goes to the client's admin centre
        if goToTenant(COMPANIES_SHEET, client, company): 
            continue
        
        # extracts the scores
        getScores(company, scores)
    
# display program end time
print("Program end time: " + str(datetime.datetime.now().strftime("%H:%M:%S")))      