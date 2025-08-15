from ast import Name
from http.client import PARTIAL_CONTENT
from lib2to3.pgen2 import driver
from nntplib import ArticleInfo
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import threading
import queue
import time

import openpyxl
from openpyxl import Workbook,load_workbook
import os


class TabInd:
            NAME = 0
            ARTICLE = 1
            BRAND = 2
            PRICE = 3
            SIZE = 4
            DESCRIPTION = 5
            PHOTO = 6
            LINK = 7  
            
def ParsingGoods(driver):
    
    Goods = []
    Link = driver.get("https://www.okeanobuvi.ru/%D0%B2%D1%81%D1%8F-%D0%BE%D0%B1%D1%83%D0%B2%D1%8C/397174/97-6-%D0%B6%D0%B5%D0%BB%D1%82%D1%8B%D0%B9-detail")
    time.sleep(5)
    
    # Cбор данных товаров
    Article = driver.find_element(By.XPATH,"//span[@class = 'sku']").text
    Color = Article.split('/')[-1]
    
    FName = driver.find_element(By.XPATH,"//h1").text 
    Name1 = FName + " "+ Color
    Brand = FName.split()[-1]
    
    Price = driver.find_element(By.XPATH,"//span[@class= 'price'][1]").text
    
    SizeList =[]
    Sizes = driver.find_elements(By.XPATH,"//div[@class = 'product_order']/span")
    for Siz in Sizes:
         SizeList.append(Siz.text) 
         Size = ", ".join(SizeList) + "." 
         
    DescrTable = driver.find_elements(By.XPATH,"//div[@class ='product-fields']//strong[@itemprop ='value']")
    for Index, Discr in enumerate(DescrTable):
        if Index == 2:
           Season = "Сезон:" + Discr.text 
        if Index == 4:
           UpperMaterial = "Материал верха" + Discr.text 
        if Index == 5:
           LiningMaterial = "Материал подкладок" + Discr.text
        if Index == 6:
           InsoleMaterial = "Материал стелек" + Discr.text
    Description = driver.find_element(By.XPATH,"//div[@class= 'product-description']//span").text + " " + Season + " " + UpperMaterial +" "+ LiningMaterial + " " + InsoleMaterial
    
    Picture = []
    Pictures = driver.find_elements(By.XPATH,"//div[@class ='img-container']/a")
    for Pict in Pictures:                
        Picture.append(Pict.get_attribute("href"))          
 
    # Запись данных в экземляр структуры StructureOfProducts    
    StructureOfProduct = {
         TabInd.NAME : Name1,
         TabInd.ARTICLE : Article,
         TabInd.BRAND : Brand,
         TabInd.PRICE : Price, 
         TabInd.SIZE :Size,
         TabInd.DESCRIPTION : Description,
         TabInd.PHOTO : Picture,
         TabInd.LINK : Link,
         }  
    Goods.append(StructureOfProduct)      
    print(Goods)

ChromedriverPuth = 'G:\\NRU\\SP\\Parsers\\selenium\\chromedriver\\win64\\139.0.7258.66\\chromedriver.exe'
s=Service(ChromedriverPuth)
driver = webdriver.Chrome(service=s)

ParsingGoods(driver)