token = "1872774492"
cookie = "pac_uid=0_891cbc52dc09d; iip=0; _qimei_uuid42=1821b0c1816100ee38a2eebe8c2a8a00df0bffe20a; _qimei_fingerprint=6c3def783fcdd5b5b9000359e1cbf9a1; _qimei_q36=; _qimei_h38=f8e0bf2838a2eebe8c2a8a000200000251821b; RK=5FPlrjs23/; ptcz=322cf7ba0c108b3d60dc7c746396b79c154677d9507e69322a08bb9435998340; ua_id=4K0Si4IDLsqpZZ6bAAAAAOJyBXwbpKo6VH5fIi8SKuA=; wxuin=16817929837358; mm_lang=zh_CN; ETCI=1c4314ed024d44f2ab9b18443c010a14; qq_domain_video_guid_verify=21bfe8a8d72ea95d; pgv_pvid=3377678849; _clck=3864921006|1|fnz|0; uuid=f30d0742d6c24b03d7cdd5c3d3afe76c; rand_info=CAESIPbMohfdzZWM/OvLoMc9YkpYsumJTKqviHNM1o4JwfCZ; slave_bizuin=3864921006; data_bizuin=3864921006; bizuin=3864921006; data_ticket=9F6lYcQlM0lrwe483DUGSw/ZAtXbhIKhfn6WBFrT2shxDrFakcy8C+spB07mE84t; slave_sid=ejd2enBQanhQT2lycUZzbWcyWjhidTB5ZUlPQ3g3dzFFckF5UHVEMl80ZUVHZHNCVDd5OGl5eGtUdERJT20zV18xbW9JSDNMbE1sMldzRWRqM3ZiWEhJV0F1bWRCMDdhOUdnTlZKcU9nWjQxbGhnempNcWJIdTFFdFZjNG1iTEtkQ281bVdEMFVUdWY1YTBm; slave_user=gh_4a07ff39a32d; xid=c28ae57a1df5d5060c1ffd1be58e9501; _clsk=1q9c94a|1722611439007|2|1|mp.weixin.qq.com/weheat-agent/payload/record"
Official_Account_list = ['中国信息安全', '数据要素社']
# 这个路径是pdf下载下来后存储的地方
save_url = "F:\\zwh\\WeChat_Article-master\\download\\"
driver=r"C:\Program Files\Google\Chrome\Application\chromedriver.exe"
key = ""
# 我们以这个为筛选目标关键字，要关掉，不然无法访问
# 请注意！key为"",则所有文章都会进行下载
# 刚刚页面不太好是我这边网络问题，你那边如果没有网络问题不会有这个情况
# driver应该和chrome浏览器的版本对应！
# 上一个视频演示了正常查询，这里演示筛选查询
# 由于刚刚进行了爬虫，所以理论上现在不    需要爬第二遍，我们把爬虫代码先注释掉，避免爬多了被短暂封禁
# 然后我们去找最新爬下来的数据
import Check_Offical_Account_Articl as query
import os
import pandas as pd
from topdf import save_webpage_as_pdf,generate_random_filename
from datetime import datetime

# 创建保存PDF文件的目录
now = datetime.now().strftime("%Y%m%d_%H%M%S")
# 下面代码负责爬虫，把这一行注释即可
file_path = query.query(cookie, token, Official_Account_list, now+".xlsx")
# 下面代码负责下载
directory = 'excel'
# 获取目录下的所有文件，按修改时间排序
files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')]
latest_file = max(files, key=os.path.getctime)

# 读取最新的xlsx文件
df = pd.read_excel(latest_file)

os.makedirs(save_url + now, exist_ok=True)

pdf_files = []
for index, row in df.iterrows():
    if key!="":
        title = str(row.iloc[1])
        abstract = str(row.iloc[2])
        if key and (key in title or key in abstract):
            print(f"包含关键词 '{key}' 的文章：{title}")
            # 获取第4列的文章链接
            article_link = row[3]
            print(f"Article link: {article_link}")

            # 假设pdf_path是文章链接指向的PDF文件路径
            url = article_link
            pdf_path = save_webpage_as_pdf(url, save_url + now, driver, title)
            print(f"PDF saved at: {pdf_path}")
            pdf_files.append(pdf_path)
    else:
        title = str(row.iloc[1])

        # 获取第4列的文章链接
        article_link = row[3]
        print(f"Article link: {article_link}")

        # 假设pdf_path是文章链接指向的PDF文件路径
        url = article_link
        pdf_path = save_webpage_as_pdf(url, save_url + now, driver, title)
        print(f"PDF saved at: {pdf_path}")
        pdf_files.append(pdf_path)
