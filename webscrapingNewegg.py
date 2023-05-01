#This file will can only run if the Anaconda and BeautifulSoup python packages are installed on your computer. 
#The xlsxwriter Python module must also be installed on your computer for this program to compile properly.

#importing webscraping packages
import bs4 
from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup as soup #Beautiful Soup is a python package which allows us to pull data out of HTML and XML documents

#importing the excelsheet module
import xlsxwriter

my_url = 'https://www.newegg.com/p/pl?d=video+cards'

#opening up connection and grabbing the url page
uClient = uReq(my_url) #uClient downloads the webpage

page_html = uClient.read() #reads the webpage

uClient.close() #closes the webpage

#html parsing
page_soup = soup(page_html, "html.parser")

#grabs each product
containers = page_soup.findAll("div", {"class":"item-container"})

#creates the excelsheet 
workbook = xlsxwriter.Workbook("NeweggVideoCardsInfo.xlsx")
worksheet = workbook.add_worksheet("firstSheet")

data = [] #will be used to store the videocards info
count = 0 #used to keep track of video cards count

for container in containers: 

    #display product name
    productName = container.a.img["alt"]
    print("Product Name: " + productName + " ")

    #display name of brand
    brand = productName.split()[0]
    print("Brand Name: " + brand + " ")

    #find price current class
    priceCurrentList = container.findAll("li", {"class":"price-current"})
    
    noPriceCurrent = "[<li class=\"price-current\"></li>]" #webpage doesnt include price

    #check if webpage contains price or not
    if str(priceCurrentList) == noPriceCurrent: #checks if price current contains price
        price = "Must add item to cart to view price!"
    else:
        dollars = priceCurrentList[0].strong.getText() #retrieves the dollars from the strong tag
        cents = priceCurrentList[0].sup.getText() #retrieves the cents from the sup tag
        price = str(dollars) + str(cents)

    #display price
    print("Price: " + price)

    count+=1
    if count == 35: break #break once count reaches 35

    videoCard = {'Brand':str(brand), 'Product Name':str(productName), 'Price':str(price)} #create dict for videoCard
    data.append(videoCard) #add videoCard to data

#remove sponsored products
data = data[4:]

#add to first row of spreadsheet
url = 'https://www.newegg.com/p/pl?d=video+cards'
worksheet.write_url(0, 0, url, string='Link') #add link to webpage
worksheet.write(0,1,"Brand") 
worksheet.write(0,2,"Product Name")
worksheet.write(0,3,"Price")

#store data in excelsheet
for index, entry in enumerate(data):
    worksheet.write(index+1, 0, str(index+1))
    worksheet.write(index+1, 1, entry["Brand"])
    worksheet.write(index+1, 2, entry["Product Name"])
    worksheet.write(index+1, 3, entry["Price"])

bold_format = workbook.add_format({'bold': True})
worksheet.set_row(0, None, bold_format) #bold the first row
worksheet.autofit()
workbook.close()