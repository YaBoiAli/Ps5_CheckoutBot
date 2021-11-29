from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from pathlib import Path
from time import sleep
import winsound
from win32com.client import Dispatch
from datetime import datetime
import time
import schedule

options = webdriver.ChromeOptions()
options.add_argument(
    r"--user-data-dir=C:\Users\muzza\AppData\Local\Google\Chrome\User Data")
options.add_argument(r'--profile-directory=Default')
driver = webdriver.Chrome(
    executable_path=r'C:\chromedriver_win32\chromedriver.exe', options=options)


# ps5 link
driver.get(
  "https://www.bestbuy.com/site/sony-playstation-5-console/6426149.p?skuId=6426149")


# test link
# driver.get("https://www.bestbuy.com/site/sony-playstation-5-dualsense-wireless-controller-white/6430163.p?skuId=6430163") #Test controller


warranty = driver.find_element_by_id("warranty-sku-2557326-warranty-selector")
warranty.click()
buyBtn = False

count = 0

while not buyBtn:
    try:
        driver.find_element_by_class_name("btn-disabled")
        warranty = driver.find_element_by_id(
            "warranty-sku-2557326-warranty-selector")
        warranty.click()
        print("OUT OF STOCK")
        count = count + 1
        print(count)

        sleep(1)

        driver.refresh()

        sleep(1)

        buyBtn = False

    except:

        addToCart = driver.find_element_by_class_name("btn-primary")
        sleep(1)
        addToCart.click()

        print("Added to cart successfully")

        buyBtn = True
        if(buyBtn == True):
            driver.get("https://www.bestbuy.com/checkout/r/fast-track")
            price = driver.find_element_by_xpath("/html/body/div[1]/main/div/div[2]/div[1]/div/div[1]/div[1]/section[2]/div/div/div[1]/div/div/div[2]")    
            
            if(price.text == "$0.00"):
                buyBtn = False
                driver.get(
                    "https://www.bestbuy.com/site/sony-playstation-5-console/6426149.p?skuId=6426149")


            else:

                code = driver.find_element_by_class_name("c-input")
                code.send_keys("") #add your card pin
                reward = driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[2]/div/div[3]/div[3]/div/div/div/div/div[2]/button")
                reward.click()

                chekout = driver.find_element_by_class_name("btn-primary")
                chekout.click();
                speak = Dispatch("SAPI.SpVoice").Speak
                speak("Item has been ordered!!")
