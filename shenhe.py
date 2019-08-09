import re
import requests
import pandas as pd
import numpy as np
from lxml import etree
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.wait import WebDriverWait

fileName = r'E:\star1\333.xlsx'
df = pd.read_excel(fileName)
link = df['电商链接']

result = []
for i in range(len(link)):
    main_url = link[i]
    if type(main_url) == float :
        print("其他销售渠道证明")
        result.append("其他销售渠道证明")
    else:
        if "jd" in str(main_url):
            options = Options()
            options.add_argument('-headless')  # 无头X参数
            driver = Firefox(firefox_options=options,executable_path = r"C:\Users\Administrator\Desktop\geckodriver.exe")  # 配了环境变量第一个参数就可以省了，不然传绝对路径
            wait = WebDriverWait(driver, timeout=3)
            driver.get(main_url)
            source = driver.page_source 
            driver.quit()
            html = etree.HTML(source)
            ziying = html.xpath(
                "//div[@class='name goodshop EDropdown']/em/text()")
            if "自营" in str(ziying):
                    name = html.xpath(
                            "//div[@class='sku-name']/text()")
                    if ("定制"in str(name))or ("防弹"in str(name))or ("射击"in str(name))or ("订制"in str(name))or ("帐篷"in str(name))or ("卫星"in str(name))or ("靶"in str(name)):
                        print("定制/专用类产品暂不通过")
                        result.append("定制/专用类产品暂不通过")
                    else:
                        wuhuo = html.xpath(
                            "//div[@id='store-prompt']/strong/text()")
                        if "无货" in str(wuhuo):
                            print("无货，请按要求提供在销渠道证明")
                            result.append("无货，请按要求提供在销渠道证明")
                        else:
                            print("通过")
                            result.append("通过")
            else:
                print("非自营，请按要求提供在销渠道证明")
                result.append("非自营，请按要求提供在销渠道证明")
                
        elif "gome" in str(main_url):
            options = Options()
            options.add_argument('-headless')  # 无头X参数
            driver = Firefox(firefox_options=options,executable_path = r"C:\Users\Administrator\Desktop\geckodriver.exe")  # 配了环境变量第一个参数就可以省了，不然传绝对路径
            wait = WebDriverWait(driver, timeout=3)
            driver.get(main_url)
            source = driver.page_source 
            driver.quit()
            html = etree.HTML(source)
            ziying = html.xpath(
                "//span[@class='identify']/text()")
            if len(ziying) == 1:
                    name = html.xpath(
                            "//div[@class='hgroup']/text()")
                    if ("定制"in str(name))or ("防弹"in str(name))or ("射击"in str(name))or ("订制"in str(name))or ("帐篷"in str(name))or ("卫星"in str(name))or ("靶"in str(name)):
                        print("定制/专用类产品暂不通过")
                        result.append("定制/专用类产品暂不通过")
                    else:
                        wuhuo = html.xpath(
                            "//div[@id='store-prompt']/strong/text()")
                        if "无货" in str(wuhuo):
                            print("无货，请按要求提供在销渠道证明")
                            result.append("无货，请按要求提供在销渠道证明")
                        else:
                            print("通过")
                            result.append("通过")
            else:
                 print("非自营，请按要求提供在销渠道证明")
                 result.append("非自营，请按要求提供在销渠道证明")
        elif "suning" in str(main_url):
            options = Options()
            options.add_argument('-headless')  # 无头X参数
            driver = Firefox(firefox_options=options,executable_path = r"C:\Users\Administrator\Desktop\geckodriver.exe")  # 配了环境变量第一个参数就可以省了，不然传绝对路径
            wait = WebDriverWait(driver, timeout=3)
            driver.get(main_url)
            source = driver.page_source 
            driver.quit()
            html = etree.HTML(source)
            ziying = html.xpath(
                "//h1[@id='itemDisplayName']//span//text()")
            if "自营" in str(ziying):
                    name = html.xpath(
                              "//h1[@id='itemDisplayName']/text()")
                    if ("定制"in str(name))or ("防弹"in str(name))or ("射击"in str(name))or ("订制"in str(name))or ("帐篷"in str(name))or ("卫星"in str(name))or ("靶"in str(name)):
                        print("定制/专用类产品暂不通过")
                        result.append("定制/专用类产品暂不通过")
                    else:
                        wuhuo = html.xpath(
                            "//div[@id='store-prompt']/strong/text()")
                        if "无货" in str(wuhuo):
                            print("无货，请按要求提供在销渠道证明")
                            result.append("无货，请按要求提供在销渠道证明")
                        else:
                            print("通过")
                            result.append("通过")
            else:
                print("非自营，请按要求提供在销渠道证明")
                result.append("非自营，请按要求提供在销渠道证明")
        else:
                 print("非自营，请按要求提供在销渠道证明")
                 result.append("非自营，请按要求提供在销渠道证明")
         
df['审核意见'] = result
df.to_excel(fileName)
