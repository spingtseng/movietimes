from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import schedule
import os

# 設定環境變數，不顯示密碼
my_password = os.environ.get('PASSWORD')

def job():
    page = 56
    table = {
        "頻道":[],
        "播出時間":[],
        "結束時間":[],
        "電影片名":[],
        "電影分級":[],
    }


    text_final = ""
    while True:
        text4 = ""
        url = "http://tv.atmovies.com.tw/tv/attv.cfm?action=channeltime&channel_id=CH" + str(page)
        response = requests.get(url)
        html = BeautifulSoup(response.text)
        title = html.find("table", width="90%")
        text_title = f"附件：今日電視台洋片時刻表\n{str(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))}\n\n"
        # 如果網頁沒有標題
        if not title:
            page = page + 1
            continue
        channel = title.find("font", class_="at15b")
        day = title.find("td", class_="at11")
        # print("頻道：", channel.text)
        # print("日期：", day.text)
        # print()
        text_channel = f"\n\n頻道:{channel.text}\n日期:{day.text}\n\n"
        data = html.find("table", bordercolor="#cddcaa").find_all("tr")
        # print(data)
        text = ""
        for d in data[1:-4]:
            starttime = d.find("td", class_="at9")
            # if starttime is None:
            #     continue
            endtime = starttime.find_next("tr").find("td", class_="at9")
            movie = d.find("font", class_="at11")
            # if movie is None:
            #     continue
            find_rate = d.find("font", color="#696933")
            if not find_rate:
                rate = "未分類"
                pass
            else:
                rate = find_rate.text
            # print("時間：", starttime.text, "~", endtime.text)
            # print("片名：", movie.text))
            # print("-" * 20)
            table["頻道"].append(channel.text)
            table["播出時間"].append(starttime.text)
            table["結束時間"].append(endtime.text)
            table["電影片名"].append(movie.text)
            table["電影分級"].append(rate)
            text_movie = f"時間:{starttime.text}~{endtime.text}\n片名:{movie.text}\n------------------------------\n"
            text = text + text_movie

        text_all = text_channel + text
        text_final = text_final + text_all

        print("\n\n\n")
        page = page + 1

        if page == 63:
            break
    print(text_title + text_final)

    print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    filename = str(time.strftime('%Y-%m-%d',time.localtime()))+"電視台洋片時刻表.csv"
    df = pd.DataFrame(table)
    df.to_csv(filename, encoding="utf-8", index=False)

    content = MIMEMultipart()
    content["subject"] = str(time.strftime('%m-%d',time.localtime()))+"電視台洋片時刻表" #標題
    content["from"] = "spingtseng@gmail.com"  #寄件人
    content["to"] = "spingtseng@gmail.com" #收件人
    content.attach(MIMEText(text_title + text_final)) #內容
    pdfload = MIMEApplication(open(filename,'rb').read())
    pdfload.add_header("Content-Disposition","attachment",filename = filename)
    content.attach(pdfload)

    with smtplib.SMTP(host = "smtp.gmail.com", port="587") as smtp:
        try:
            smtp.ehlo() # 驗證SMTP伺服器
            smtp.starttls() # 建立加密傳輸
            smtp.login("spingtseng@gmail.com",my_password)
            smtp.send_message(content)
            print("傳送成功")
        except Exception as e:
            print("傳送失敗,Error message:",e)

schedule.every().day.at("01:00").do(job)
# schedule.every(10).seconds.do(job)
while True:
    schedule.run_pending()
    time.sleep(10)

