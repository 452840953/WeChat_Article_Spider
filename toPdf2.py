# -*- coding: UTF8 -*-
import json
import time
import random
import string
from selenium import webdriver
from selenium.webdriver.chrome.service import Service


def generate_random_filename():
    """生成随机文件名"""
    random_string = ''.join(random.choices(string.ascii_lowercase + string.digits, k=8))
    return f"{random_string}.pdf"


def save_webpage_as_pdf(url, save_directory='F:\\zwh\\WeChat_Article-master\\pdf', driver=r"C:\Program Files\Google\Chrome\Application\chromedriver.exe"):
    print("Setting up Chrome options...")
    chrome_options = webdriver.ChromeOptions()
    chrome_options.ignore_local_proxy_environment_variables()

    chrome_options.add_argument('--headless')  # 保留headless模式
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--print-to-pdf-no-header')

    # 生成随机文件名
    pdf_filename = generate_random_filename()
    pdf_path = f"{save_directory}\\{pdf_filename}"
    chrome_options.add_argument(f'--print-to-pdf={pdf_path}')  # 添加print-to-pdf选项

    print("Setting up ChromeDriver path...")
    driver_path = driver
    service = Service(driver_path)

    print("Initializing WebDriver...")
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        print("Navigating to the target URL...")
        driver.get(url)

        print("Waiting for the page to load...")
        time.sleep(5)  # 添加页面加载等待时间
    finally:
        print("Closing browser...")
        driver.quit()

    print(f"PDF saved to {pdf_path}")
    return pdf_path


# 示例调用
if __name__ == "__main__":
    url = 'http://mp.weixin.qq.com/s?__biz=MzA5MzE5MDAzOA==&mid=2664220358&idx=6&sn=a6c42566d1a38c3eca452f5396cfb132&chksm=8b59c43fbc2e4d29ff59d692d73bcdc010548dd747d7c37ff6751759896bd0d8acd6e829a38a#rd'
    save_webpage_as_pdf(url)
