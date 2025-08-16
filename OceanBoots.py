from ast import Name
from http.client import PARTIAL_CONTENT
from lib2to3.pgen2 import driver
from nntplib import ArticleInfo
from pprint import pprint
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
            
def GettingColltction(driver):
    
    driver.get("https://www.okeanobuvi.ru/%D0%BE%D0%B1%D1%83%D0%B2%D1%8C/%D0%B6%D0%B5%D0%BD%D1%81%D0%BA%D0%B0%D1%8F/%D0%B1%D0%B0%D0%BB%D0%B5%D1%82%D0%BA%D0%B8")
  
    LinksGoods = []                      
    ListLG = driver.find_elements(By.XPATH,"//h2[@class = 'product-name']/a")
    for LG in ListLG:  
        LinksGoods.append(LG.get_attribute("href"))
        ParsingGoods(driver, LinksGoods)    


def ParsingGoods(driver, LinksGoods):
    
    Goods = []  
    for Link in LinksGoods:  
        driver.get(Link)   
        try:
           OnSale = driver.find_element(By.XPATH,"//span[@class = 'in-stock']").text 
           if OnSale  == "В наличии":
          
                # Cбор данных товаров
                Article = driver.find_element(By.XPATH,"//span[@class = 'sku']").text
                Color = Article.split('/')[-1]
    
                FName = driver.find_element(By.XPATH,"//h1").text 
                Name1 = FName + " "+ Color
                Brand = FName.split()[-1]
    
                Price = driver.find_element(By.XPATH,"//div[@class = 'product-price with-discount']/span[@itemprop = 'price']").text
    
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
               
                Description = driver.find_element(By.XPATH,"//div[@class= 'product-description']//span").text + " /n" + Season + " /n" + UpperMaterial +" /n"+ LiningMaterial + " /n" + InsoleMaterial
        
                # TableSize = []
                # try: 
                #     #Модальное окно не открывается при нажатии
                #     LinkTbSz = driver.find_element(By.XPATH,"//a[@id = 'tbsize-a']").get_attribute("href")
                #     LinkTbSz.click()
                
                #     TableSizes = driver.find_elements(By.XPATH,"//div[@id= 'tab1']//td")           
                #     for TS in TableSizes:                
                #         TableSize.append(TS.text)      
                #     indsize = TableSize.index("Размер")
                #     TabSZ = ", ".join(TableSize[:indsize + 1]) + "\n" + ", ".join(map(str, TableSize[indsize + 1:]))   
                # except:
                #     print("Таблица размеров не найдена")
                # print(TabSZ)   

                Picture = []
                Pictures = driver.find_elements(By.XPATH,"//div[@class ='img-container']/a")
                if len(Pictures) > 5:
                     for index, Pict in enumerate(Pictures):                      
                        if index % 2 != 0:
                             Picture.append(Pict.get_attribute("src"))
                else:
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
        except:
            print("При сборе данных возникла ошибка")
        
        
        print(Goods)
        print()
        
ChromedriverPuth = 'G:\\NRU\\SP\\Parsers\\selenium\\chromedriver\\win64\\139.0.7258.66\\chromedriver.exe'
s=Service(ChromedriverPuth)
driver = webdriver.Chrome(service=s)

GettingColltction(driver)
