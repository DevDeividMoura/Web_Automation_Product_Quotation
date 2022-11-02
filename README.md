# Web Automation Product Quotation
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/DevDeividMoura/Automated_Web_Product_Quote/blob/main/LICENSE) 
[![Python 3.10](https://img.shields.io/badge/python-v3.10-blue.svg)](https://www.python.org/downloads/release/python-3100/)

# About the project

video do projeto rodando

### Inspiration:

- An employee who works in the purchasing area of ​​a company needs to make a comparison of suppliers for their inputs/products. At that time, the employee will constantly search the websites of these suppliers for the products available and the price, after all, each one of them can promote at different times and with different values.

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

## Aplicação

<div align="center">

![Android Emulator](https://user-images.githubusercontent.com/37356058/135944390-ec96d4ec-ee43-4db9-882f-89be66aad23a.gif)

  </div>

## Technologies used

- Python v3.10
- Selenium
- Pandas
- API Outlook (pywin32)

## Abrir e rodar o projeto

Após baixar o projeto, você precisara fazer alguns ajustes:

- Primeiramente, abra o arquivo `search.xlsx` e adicione os produtos que pretende buscar;

  • Na Coluna `Name` adicione os termos do nome do produto a buscar separados por um espaço simples;
  
  • Na Coluna `Banned Terms` adicione os termos que nao podem conter no nome do produto para evitar a busca do produto errado;
  
  • Na Coluna `Minimum Price` adicione o valor minimo a considerar para nao cair em golpes com valores muito baixos;
      
  • Na Coluna `Maximum Price`  adicione o valor maximo a considerar na busca;
  
    Observação: O valor considerano no script esta em reais (BRL / R$) pois foi desenvolvido no Brasil, para rodar em outro pais basta
    adicionar o valor de acordo com a moeda da sua localização e alterar a uma linha de codigo (atualmente n° 167) no arquivo `main.py` 
  
    ```python
    # currency symbol that will be considered in search results according to your location
    # to convert string price to float ex: R$, US$, €
    currency_symbol = "R$"
    ```
  
- Logo após é necessario cadastrar os nomes e e-mail dos colaboradores que irão receber as ofertas, no arquivo `Send E-mails.xlsx`;

- Em Seguida voce devera cadastrar seu email do google (que sera usado como servidor) e sua senha para aplicativos no arquivo `Credentials.txt`, 

   • Veja como conseguir uma senha aqui: https://support.google.com/mail/answer/185833?hl=en;

- Agora abra arquivo `main.py` (Recomenda-se a utilização da IDE PyCharm).

  • Instale os requisitos via terminal terminal:
    
     ```python
    
    pip install -r requirements.txt
    
    ```

  • Por fim rode o arquivo `main.py`;


