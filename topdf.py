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


def save_webpage_as_pdf(url, save_directory='F:\\zwh\\WeChat_Article-master\\pdf', driver=r"C:\Program Files\Google\Chrome\Application\chromedriver.exe", title=generate_random_filename()):
    # 设置打印机的纸张大小、打印类型、保存路径等
    print("Setting up Chrome options...")
    chrome_options = webdriver.ChromeOptions()
    chrome_options.ignore_local_proxy_environment_variables()
    settings = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": ""
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "isHeaderFooterEnabled": False,
        "isCssBackgroundEnabled": True,
        "mediaSize": {
            "height_microns": 297000,
            "name": "ISO_A4",
            "width_microns": 210000,
            "custom_display_name": "A4"
        },
    }

    print("Adding Chrome arguments and preferences...")
    chrome_options.add_argument('--enable-print-browser')
    # chrome_options.add_argument('--headless')  # headless模式下，浏览器窗口不可见，可提高效率
    prefs = {
        'printing.print_preview_sticky_settings.appState': json.dumps(settings),
        'savefile.default_directory': save_directory  # 此处填写你希望文件保存的路径
    }

    chrome_options.add_argument('--kiosk-printing')  # 静默打印，无需用户点击打印页面的确定按钮
    chrome_options.add_experimental_option('prefs', prefs)

    # 设置ChromeDriver路径
    print("Setting up ChromeDriver path...")
    driver_path = driver
    service = Service(driver_path)

    print("Initializing WebDriver...")
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        print("Navigating to the target URL...")
        driver.get(url)

        print("Maximizing browser window...")
        driver.maximize_window()  # 浏览器最大化

        # 平滑滚动到底部
        print("Scrolling to the bottom of the page...")

        def smooth_scroll(driver, scroll_height):
            for i in range(0, scroll_height, 100):
                driver.execute_script(f"window.scrollTo(0, {i});")
                time.sleep(0.03)  # 调整滚动速度

        # 获取页面的总高度
        total_height = driver.execute_script("return document.body.scrollHeight")

        # 平滑滚动到页面底部
        smooth_scroll(driver, total_height)
        # 生成随机文件名
        pdf_filename = title+".pdf"

        print(f"Executing JavaScript to change the title and trigger print for {pdf_filename}...")
        driver.execute_script(
            f'document.title="{pdf_filename}";window.print();')  # 利用js修改网页的title，该title最终就是PDF文件名，利用js的window.print可以快速调出浏览器打印窗口，避免使用热键ctrl+P

        # 等待打印命令完成
        time.sleep(5)
    finally:
        # 关闭浏览器
        print("Closing browser...")
        driver.quit()

    pdf_path = f"{save_directory}\\{pdf_filename}"
    print(f"PDF saved to {pdf_path}")
    return pdf_path


# 示例调用
if __name__ == "__main__":
    url = 'https://mp.weixin.qq.com/s/FKI16e2l42OWDrdAqXfmLQ'
    save_webpage_as_pdf(url)
