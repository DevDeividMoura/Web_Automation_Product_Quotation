# Web Automation Product Quotation
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/DevDeividMoura/Web_Automation_Product_Quotation/blob/main/LICENSE) 
[![Python 3.10](https://img.shields.io/badge/python-v3.10-blue.svg)](https://www.python.org/downloads/release/python-3100/)

# About the project

- This project aims to automate web searches as a way of optimizing processes within a company, increasing team productivity and effectively presenting fast and accurate results.


### Inspiration:

- An employee who works in the purchasing area of a company needs to make a comparison of suppliers for their inputs/products. At that time, the employee will constantly search the websites of these suppliers for the products available and the price, after all, each one of them can promote at different times and with different values.

- Objective: If the value of the products is below a previously defined threshold price, the employee will discover the cheapest products and update this in a spreadsheet. It will then send an email with the list of products below its maximum purchase price to each employee in the shopping area.

- In this example, we are going to search for common products on sites like Google Shopping and Buscapé, but the idea is the same for other sites.

### What do we have available?

- Product Spreadsheet, with the names of the products, the maximum price, the minimum price (to avoid "wrong" or "too cheap to be true" products and the terms we want to avoid in our searches.
- Spreadsheet with the information of the employees of the purchasing sector with their name, email and area

### What should we do:

- Search for each product on Google Shopping and get all the results that are priced within the range and are the correct products. Same for Buscapé. Send an email to each purchasing department employee with the notification and the Spreadsheet with the items and prices found, along with the purchase link. (I'll use the email teste+(contributor's name)@gmail.com. Use your email to do the tests to see if the message is arriving)

### Solution:

- Using a simple automation script we can execute this entire process from hours to minutes, saving time and effort for the employee that can be used in other activities

# Functionalities

:heavy_check_mark: `Functionality 1:` Search thousands of products on various sites extracting data such as: name, price, link to the offer.

:heavy_check_mark: `Functionality 2:` Compare prices by filtering the results within the pre-defined range, check the product description avoiding errors and then save the spreadsheet locally with name, price and link to the offer.

:heavy_check_mark: `Functionality 3:` Forwards email with results attached in .xlsx to each person responsible in the sector.

:heavy_check_mark: `Functionality 4:` Plot progress bar tracking process enabling execution with hidden browser.

# Application



https://user-images.githubusercontent.com/116500495/199385784-e847d892-f2f3-4bea-9fb0-37c36233c4a4.mp4



# Technologies used

- Python v3.10
- Selenium
- Pandas
- Smtplib (Automatic email sending by gmail)

# Open and run the project

After downloading the project, you will need to make some adjustments:

- First, open the `search.xlsx` file and add the products you want to search for;

  • In Column `Name`, add the terms of the product name to be searched separated by a single space;
  
  • In the Column `Banned Terms`, add the terms that cannot be contained in the product name to avoid searching for the wrong product;
  
  • In the Column `Minimum Price`, add the minimum value to consider so as not to fall into scams with very low values;
      
  • In Column `Maximum Price`, add the maximum value to consider in the search;
  
    Note: The value considered in the script is in reais (BRL / R$) because it was developed in Brazil, to run in another country, just enter the value 
    according to the currency of your location and change to a line of code (currently n° 167) in the `main.py` file
  
    ```python
    # currency symbol that will be considered in search results according to your location
    # to convert string price to float ex: R$, US$, €
    currency_symbol = "R$"
    ```
  
- Soon after, it is necessary to register the names and e-mail of the collaborators who will receive the offers, in the file `Send E-mails.xlsx`;

- Then you must register your google email (which will be used as a server) and your password for applications in the `Credentials.txt` file,

   • See how to get a password here: https://support.google.com/mail/answer/185833?hl=en;

- Now open the `main.py` file:

  • Install requirements via terminal terminal:
    
     ```NET
    pip install -r requirements.txt
    ```

- Finally run the `main.py` file;
  
- After running the script a browser instance will open and you will see the search process, at the end of the process the script will save the search results in a file named `Offers.xlsx` and forward it via email to registered collaborators.

# Final remarks:

 - Depending on when you are running these Scripts, some site structures may have changed, thus needing to readjust the code to interact with the elements again;

 - The script can be executed with the browser instance hidden, so just uncomment the next line of code (line n° 172) in the `main.py` file:
    ```python
    
    ## run to the browser instance covertly
    # options.add_argument("--headless")
    
    ```


# Autor

Deivid Carvalho Moura

https://www.linkedin.com/in/devdeividmoura
