# 亚马逊爬虫

## automate_scrapy.py   
纯自动化selenium 亚马逊爬虫

##  request_amazon_data.py
selenium 和 requests 结合的亚马逊爬虫

先获取所有asin并保存。再通过读取asin文件一个个去获取asin的数据。

可能存在问题：requests请求速度过快可能会导致出现验证码

解决思路：出现验证码,使用selenium过验证码后再继续使用requests
