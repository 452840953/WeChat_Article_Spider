import requests
import datetime
import pandas as pd
import re
import time


# 创建全局变量，用于存储取得的文章基本信息
aritcle_lists = []

# 创建全局变量，用于获取当前日期，判断文章是否新发，进而触发下一步操作
today = str(datetime.date.today())

# 创建正则的判定表达式
rule = r"(cve)|(漏洞)|(通告)|(告警)|(紧急)"

def query_info(offical_account,cookie,Query_Token):
    # 设定基础请求接口
    base_url = "https://mp.weixin.qq.com/cgi-bin/appmsg"
    # 厂商公众号的ID
    fakeid_dict = {'中国信息安全':'MzA5MzE5MDAzOA==','数据要素社':'Mzg2MDgzNDEzMw==','金融电子化':'MjM5MzA3MzAzOQ==' ,'CAICT数据要素':'MzkwNjE4ODkxNg==','数据法盟':'MzIyNjUxOTQ0MQ==','国家数据局':'MzkzMTU5MDA3OQ=='}
    # 请求参数
    params = {
        "action": "list_ex",
        "begin": 0,  # 查询启示页，默认每页5条数据
        "count": 1,  # 查询最近两天次内发布的信息（一天一次，一次可以发布多条文章）
        "fakeid": fakeid_dict[offical_account],  # 公众号的标识ID
        "type": 9,
        "query": "",
        "token": f"{Query_Token}",  # 微信公众号的查询token,每天都会变化
        "lang": "zh_CN",
        "f": "json",
        "ajax": 1
    }

    # 设置请求头，参数为微信公众平台-订阅号cookie
    headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36',
        'Cookie':f'{cookie}'
    }

    # 向指定公众号的文章查询接口发送信息，并解析为json格式
    resp = requests.get(url=base_url,headers=headers,params=params).json()

    if resp['base_resp']['err_msg'] == 'freq control':
        print("请求频率过高，需要等一段时间")
    if resp['base_resp']['err_msg'] == 'invalid csrf token':
        print("请检查请求url中的token参数,第34行~~")


    # 返回结果为该公众号最近2天次已发布文章概要信息
    lists = resp['app_msg_list']

    # 对字典内容进行解析
    for i in range(0,len(lists)):
        title = lists[i]['title']       # 文章标题
        digest = lists[i]['digest']     # 文章摘要
        link = lists[i]['link']         # 文章链接
        # 对创建时间戳解析
        dateobject = datetime.datetime.fromtimestamp(lists[i]['create_time'])
        day = str(dateobject.date())        # 文章发表时间--天
        hour = str(dateobject.hour)         # 文章发表时间--小时
        minute = str(dateobject.minute)     # 文章发表时间--分钟
        publish_time = f"{hour}:{minute}"   # 文章发表时间
        articl_list = [offical_account,title,digest,link,day,publish_time]


        # 利用正则筛选关键title信息，判断该文章是否需要关注
        match = re.search(rule,title,flags=re.IGNORECASE)
        if match:
            print(title)
            articl_list.append("True")

        # 将文章数据汇总到总表中
        aritcle_lists.append(articl_list)

def query(cookie,Query_Token,Official_Account_list, now=str(int(time.time()))+".xlsx"):
    for Official_Account in Official_Account_list:
        print(f"当前正在爬取的公众号是：{Official_Account}")
        query_info(Official_Account,cookie,Query_Token)
        # 等待三秒，防止频率过高，接口被封锁
        time.sleep(3)
    # Debugging: print aritcle_lists to check its content
    # 检查 aritcle_lists 的每一行，如果不满足 7 个字段，则填充一个空值
    for i in range(len(aritcle_lists)):
        while len(aritcle_lists[i]) < 7:
            aritcle_lists[i].append(None)
    data = pd.DataFrame(aritcle_lists,columns=['公众号',"文章标题",'文章摘要','文章链接','发布日期','发布时间','需要关注']);
    # 计算时间戳，并且作为表单名
    timestamp = int(time.time())
    file_path = 'excel/'+str(now)
    #输出表格，保存到excel文件夹
    data.to_excel(file_path,index=False)
    return file_path