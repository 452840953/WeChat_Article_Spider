from datetime import datetime
import zipfile

from fastapi import FastAPI, Form, Request, Query
from fastapi.templating import Jinja2Templates
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
import Check_Offical_Account_Articl as query
import uvicorn
import os
import pandas as pd
import numpy as np
from topdf import save_webpage_as_pdf,generate_random_filename
import time

# 创建程序，并且关闭API文档
app = FastAPI(docs_url=None, redoc_url=None)
# 加载html模板
templates = Jinja2Templates("Pages")


# 加载主页
@app.api_route("/", methods=["GET", "POST"])
def index(req: Request):
    return templates.TemplateResponse("query.html", context={"request": req})


# 获取查询结果，以表格文件返回
@app.post("/fofa")
def test(req: Request, cookie=Form(...), token=Form(...)):
    try:
        Official_Account_list = ['中国信息安全', '数据要素社', '金融电子化', 'CAICT数据要素', '数据法盟', '国家数据局']
        file_path = query.query(cookie, token ,Official_Account_list)
        if file_path is not None:
            return FileResponse(path=file_path, filename=file_path)
        else:
            return JSONResponse({'Error': '文件生成异常,请检查cookie和token。另外可能你的账户已经被限制请求接口，需要等待1天'})
    except:
        return JSONResponse({'Error': '文件生成异常,请检查cookie和token。另外可能你的账户已经被限制请求接口，需要等待1天'})


# 获取最新的xlsx文件并转换为JSON数据返回
@app.get("/latest_excel_to_json")
def latest_excel_to_json():
    try:
        directory = 'excel'
        # 获取目录下的所有文件，按修改时间排序
        files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')]
        latest_file = max(files, key=os.path.getctime)

        # 读取最新的xlsx文件
        df = pd.read_excel(latest_file)

        # 替换NaN和Infinity值
        df = df.replace({np.nan: None, np.inf: None, -np.inf: None})

        # 将DataFrame转换为JSON
        data = df.to_dict(orient='records')

        return JSONResponse(content=data)
    except Exception as e:
        return JSONResponse({'Error': str(e)})


# 提供PDF文件下载并获取特定行的文章链接
@app.get("/download_pdf")
def download_pdf(id: int = Query(...)):
    try:
        directory = 'excel'
        # 获取目录下的所有文件，按修改时间排序
        files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')]
        latest_file = max(files, key=os.path.getctime)

        # 读取最新的xlsx文件
        df = pd.read_excel(latest_file)

        # 获取指定id行的数据
        row = df.iloc[id]

        # 获取第4列的文章链接
        article_link = row[3]
        print(f"Article link: {article_link}")

        # 假设pdf_path是文章链接指向的PDF文件路径
        url = article_link
        pdf_path = save_webpage_as_pdf(url)
        print(f"PDF saved at: {pdf_path}")
        # pdf_path = "my_test_file3.pdf"
        # time.sleep(5)

        return FileResponse(path=pdf_path, filename=generate_random_filename(), media_type='application/pdf')
    except Exception as e:
        return JSONResponse({'Error': str(e)})

@app.get("/download_pdf_all")
def download_pdf_all():
    try:
        directory = 'excel'
        # 获取目录下的所有文件，按修改时间排序
        files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')]
        latest_file = max(files, key=os.path.getctime)

        # 读取最新的xlsx文件
        df = pd.read_excel(latest_file)

        # 创建保存PDF文件的目录
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        download_dir = os.path.join('download', now)
        os.makedirs(download_dir, exist_ok=True)

        pdf_files = []
        for index, row in df.iterrows():
            # 获取第4列的文章链接
            article_link = row[3]
            print(f"Article link: {article_link}")

            # 假设pdf_path是文章链接指向的PDF文件路径
            url = article_link
            pdf_path = save_webpage_as_pdf(url, "F:\\zwh\\WeChat_Article-master\\"+download_dir)
            print(f"PDF saved at: {pdf_path}")
            pdf_files.append(pdf_path)

        # 创建保存ZIP文件的目录
        zip_dir = 'zip'
        os.makedirs(zip_dir, exist_ok=True)

        # 以日期时间命名ZIP文件
        zip_filename = f"{now}.zip"
        zip_path = os.path.join(zip_dir, zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf_file in pdf_files:
                if os.path.isfile(pdf_file):  # 判断文件是否存在
                    zipf.write(pdf_file, os.path.basename(pdf_file))

        # 使用 RedirectResponse 重定向到主页
        return RedirectResponse(url="/")
    except Exception as e:
        return JSONResponse({'Error': str(e)})


if __name__ == "__main__":
    # 绑定端口为8000，并且不限定IP访问，启动程序
    uvicorn.run(app, host="0.0.0.0", port=8000)
