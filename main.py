# Version 1.0
# Developed by Deivid Carvalho Moura

# Imports
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
import time
import configparser
from tqdm import tqdm


# standardizes input data
def standardize_data(product, banned_terms, minimum_price, maximum_price):
    # handles the values that came from the table
    product = product.lower()
    banned_terms = banned_terms.lower()
    list_banned_terms = banned_terms.split(" ")
    list_product_terms = product.split(" ")
    minimum_price = float(minimum_price)
    maximum_price = float(maximum_price)

    return product, banned_terms, minimum_price, maximum_price, list_banned_terms, list_product_terms


# checks if the search return has specified banned terms
def check_banned_terms(list_banned_terms, name):
    # name verification - if the name has any banned term
    have_banned_terms = False
    for word in list_banned_terms:
        if word in name:
            have_banned_terms = True

    return have_banned_terms


# checks if the search return has all the specified names
def check_name(list_product_terms, name):
    # check if the name has all the terms of the product name
    have_all_product_terms = True
    for word in list_product_terms:
        if word not in name:
            have_all_product_terms = False

    return have_all_product_terms


# Search for items on Google Shopping
def google_shopping_search(driver, product, banned_terms, minimum_price, maximum_price, currency_symbol, pbar):
    # enter google
    driver.get("https://www.google.com/")

    product, banned_terms, minimum_price, maximum_price, list_banned_terms, \
    list_product_terms = standardize_data(product, banned_terms, minimum_price, maximum_price)

    # search product name on Google
    xpath_google_search_bar = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input'
    driver.find_element(By.XPATH, xpath_google_search_bar).send_keys(product, Keys.ENTER)

    # click on the shopping tab
    elements = driver.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for item in elements:
        if "Shopping" in item.text:
            item.click()
            break

    # get the list of search results in google shopping
    results_list = driver.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')

    # for each result, it will check if the result matches all our conditions
    offer_list = []  # list that the function will give me as an answer
    for result in results_list:
        name = result.find_element(By.CLASS_NAME, 'Xjkr3b').text
        name = name.lower()

        # name verification - if the name has any banned term
        have_banned_terms = check_banned_terms(list_banned_terms, name)

        # check if the name has all the terms of the product name
        have_all_product_terms = check_name(list_product_terms, name)

        if not have_banned_terms and have_all_product_terms:  # checking the name
            try:
                # converting text price to float
                price = result.find_element(By.CLASS_NAME, 'a8Pemb').text
                price = float(price.replace(currency_symbol, "").replace(" ", "").replace(".", "").replace(",", "."))
                # checking if the price is within the minimum and maximum
                if minimum_price <= price <= maximum_price:
                    link_element = result.find_element(By.CLASS_NAME, 'aULzUe')
                    parent_element = link_element.find_element(By.XPATH, '..')
                    link = parent_element.get_attribute('href')
                    offer_list.append((name, price, link))
            except:
                continue
    pbar.update()

    return offer_list


# Search for items on BuscaPé
def buscape_search(driver, product, banned_terms, minimum_price, maximum_price, currency_symbol, pbar):
    # handle function values
    product, banned_terms, minimum_price, maximum_price, list_banned_terms, \
    list_product_terms = standardize_data(product, banned_terms, minimum_price, maximum_price)

    # enter BuscaPe
    driver.get("https://www.buscape.com.br/")

    # search for the product on BuscaPé
    driver.find_element(By.CLASS_NAME, 'AutoCompleteStyle_input__FInnF').send_keys(product, Keys.ENTER)

    # get the search results list from BuscaPe
    time.sleep(5)
    list_results = driver.find_elements(By.CLASS_NAME, 'Cell_Content__1630r')
    # for each result, it will check if the result matches all our conditions
    offer_list = []  # list that the function will give me as an answer
    for result in list_results:
        try:
            price = result.find_element(By.CLASS_NAME, 'CellPrice_MainValue__3s0iP').text
            name = result.get_attribute('title')
            name = name.lower()
            link = result.get_attribute('href')

            # name verification - if the name has any banned term
            have_banned_terms = check_banned_terms(list_banned_terms, name)

            # check if the name has all the terms of the product name
            have_all_product_terms = check_name(list_product_terms, name)

            if not have_banned_terms and have_all_product_terms:
                price = float(price.replace(currency_symbol, "").replace(" ", "").replace(".", "").replace(",", "."))
                if minimum_price <= price <= maximum_price:
                    offer_list.append((name, price, link))
        except:
            pass

    pbar.update()

    return offer_list


# import databases
tabel_products = pd.read_excel("search.xlsx")
df_collaborators = pd.read_excel('Send E-mails.xlsx')

# import credentials for sending emails
file_cred = configparser.RawConfigParser()
file_cred.read('Credentials.txt')
email_src = file_cred.get('LOGIN', 'email_src')
app_passw = file_cred.get('LOGIN', 'app_passw')

# create progress bar
pbar = tqdm(total=(len(tabel_products.index) * 2) + len(df_collaborators.index), position=0, leave=True)

# currency symbol that will be considered in search results according to your location
# to convert string price to float ex: R$, US$, €
currency_symbol = "R$"

# create browser instance
options = webdriver.ChromeOptions()
# options.add_argument("--headless")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# assembles pandas dataframe with Google Shopping and BuscaPe results
offer_table = pd.DataFrame()

for line in tabel_products.index:
    product = tabel_products.loc[line, "Name"]
    banned_terms = tabel_products.loc[line, "Banned Terms"]
    minimum_price = tabel_products.loc[line, "Minimum Price"]
    maximum_price = tabel_products.loc[line, "Maximum Price"]

    google_shopping_offer_list = google_shopping_search(driver, product, banned_terms, minimum_price,
                                                        maximum_price, currency_symbol, pbar)
    if google_shopping_offer_list:
        google_shopping_tabel = pd.DataFrame(google_shopping_offer_list, columns=['Product', 'Price', 'link'])
        offer_table = pd.concat([offer_table, google_shopping_tabel])
    else:
        google_shopping_tabel = None

    buscape_offer_list = buscape_search(driver, product, banned_terms, minimum_price,
                                        maximum_price, currency_symbol, pbar)
    if buscape_offer_list:
        buscape_tabel = pd.DataFrame(buscape_offer_list, columns=['Product', 'Price', 'link'])
        offer_table = pd.concat([offer_table, buscape_tabel])
    else:
        buscape_tabel = None

# export locally to excel
offer_table = offer_table.reset_index(drop=True)
offer_table.to_excel("Offers.xlsx", index=False)

# close browser instance
driver.quit()

# sends the return to each email registered in the "Send E-mails" file
for i, email in enumerate(df_collaborators['E-mail']):
    contributor = df_collaborators.loc[i, 'Contributor']

    # send the table result by email
    # checking if there is any offer within the offer table
    if len(offer_table.index) > 0:
        # I will send email
        file_attachment = 'Offers.xlsx'
        emai_body = f"""
                <p>Dear, {contributor}!</p>
                <p>We found some products on offer within the desired price range. 
                See the details in the attached spreadsheet</p>
                <p>Any doubts I am available</p>
                <p>Att.,</p>
                """

        msg = MIMEMultipart()
        msg['Subject'] = 'Products found in the desired price range'
        msg['From'] = email_src
        msg['To'] = email
        password = app_passw
        msg.attach(MIMEText(emai_body, 'html'))

        with open(file_attachment, 'rb') as attachment:
            att = MIMEBase('application', 'octet-stream')
            att.set_payload(attachment.read())
            encoders.encode_base64(att)
            att.add_header('Content-Disposition', f'attachment; filename={file_attachment}')

        msg.attach(att)

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        s.quit()

        print(f'Email sent to: {contributor}')
        pbar.update()


