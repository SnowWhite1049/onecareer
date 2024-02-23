import tkinter as tk
from tkinter import ttk
from datetime import datetime
from tkinter import Menu
import xlsxwriter

import requests
import json

def formatDateTime(date_str):   
    new_date_str = ""
    if date_str != None :   
        date_obj = datetime.fromisoformat(date_str.replace('Z', ''))    
        new_date_str = date_obj.strftime("%Y/%m/%d %H:%M")  
    return new_date_str

def hasLiveSchedule(b):       
    str = ""
    if b == True : str ="本日締切！"
    return str

root = tk.Tk()
root.title("https://gaishishukatsu.com/")

style = ttk.Style()
style.configure('Treeview', rowheight=100)
tree = ttk.Treeview(root, selectmode='browse', height=10)

tree["columns"] = (
    "NO",
    "has_live_schedule!",
    "closest_schedule_start_at",
    "closest_schedule_time_limit_at",
    "company_business_category_name",
    "company_name",
    "title"
)

tree.column("#0", 			anchor=tk.W, width=10)
tree.column("#1", 			anchor=tk.W, width=50)
tree.column("#2", 			anchor=tk.W, width=600)
tree.column("#3", 			anchor=tk.W, width=600)
tree.column("#4", 			anchor=tk.W, width=50)
tree.column("#5", 			anchor=tk.W, width=50)
tree.column("#6", 			anchor=tk.W, width=50)
# tree.column("product_category", anchor=tk.CENTER, width=100)
# tree.column("product_mana_no", 	anchor=tk.CENTER, width=100)
# tree.column("product_sale_no", 	anchor=tk.CENTER, width=100)
# tree.column("product_id", 		anchor=tk.CENTER, width=100)
# tree.column("login_id", 		anchor=tk.CENTER, width=100)
# tree.column("bid_amount", 		anchor=tk.CENTER, width=100)
# tree.column("closed_date", 		anchor=tk.CENTER, width=100)
tree.heading("#0", 				anchor=tk.W, text="")
tree.heading("#1", 				anchor=tk.W, text="No")
tree.heading("#2", 				anchor=tk.W, text="会社情報")
tree.heading("#3", 		anchor=tk.W, text="スケジュール")
tree.heading("#4", 		anchor=tk.W, text="")
tree.heading("#5", 		anchor=tk.W, text="")
tree.heading("#6", 		anchor=tk.W, text="")


# tree.heading("product_category", 	anchor=tk.CENTER, text="カテゴリ")
# tree.heading("product_mana_no", 	anchor=tk.CENTER, text="管理番号")
# tree.heading("product_sale_no", 	anchor=tk.CENTER, text="併売番号")
# tree.heading("product_id", 			anchor=tk.CENTER, text="商品ID")
# tree.heading("login_id", 			anchor=tk.CENTER, text="ログインID")
# tree.heading("bid_amount", 			anchor=tk.CENTER, text="落札金額")
# tree.heading("closed_date", 		anchor=tk.CENTER, text="終了日時")

tree.pack(fill=tk.BOTH, expand=True, pady=0)

def loadData (): 
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'X-Csrf': 'onecareer'}
    
    link = "https://api-public.gaishishukatsu.com/3.0/recruitment?page=1&order=recommend&limit=1000&allgy=false"

    result = requests.get(link, headers=headers)    
    rows = result.json()
    index = 0

    for event in rows['recruitments']:  
        index = index + 1
        company_info = event["type"]["label"] + "、"
        company_name = ""
        company_year = ""
        company_dt = ""

        if event["targetYear"] > 0: 
            company_info = company_info + str(event["targetYear"]) + "卒、"
            company_year = str(event["targetYear"]) + "卒"
        
        for gt in event["groupedJobTypes"] :
            company_info = company_info + gt["name"] + "、"
        
        company_info = company_info + "\n"

        company_info = company_info + event["title"] + "\n"

        for company in event["companies"] :
            company_info = company_info + company["name"] + "\n" 
            company_name = company["name"]
        
        schedule_info = ""
        num = 0
        for sh in event["schedules"] :
            num = num + 1
            schedule_info = schedule_info + str(num) + " : "
            schedule_info = schedule_info + sh["name"] + "、" + sh["place"]
            schedule_info = schedule_info + " ~ "
            schedule_info = schedule_info + formatDateTime(sh["entryStart"]) + "、"
            schedule_info = schedule_info + formatDateTime(sh["entryEnd"]) + "、"
            schedule_info = schedule_info + formatDateTime(sh["eventStart"]) + "、"
            schedule_info = schedule_info + formatDateTime(sh["eventEnd"])
            schedule_info = schedule_info + " ~ "
            schedule_info = schedule_info + "\n"

            if num == 1: company_dt = formatDateTime(sh["entryEnd"])

        tree.insert("", tk.END, values=(index, company_info, schedule_info, company_name, company_year, company_dt)) 
        
    tree.tag_configure("bold", font=("Arial", 10, ""))
    for item in tree.get_children():
        tree.item(item, tags=("bold",))

def csvMake (): 
    workbook = xlsxwriter.Workbook("Gaishishukatsu.xlsx")
    worksheet = workbook.add_worksheet()
    num = 1
    
    worksheet.write("A" + str(num), "締切日（降順）")
    worksheet.write("B" + str(num), "サイト名")
    worksheet.write("C" + str(num), "企業名")
    worksheet.write("D" + str(num), "就職年度")
    
    for item in tree.get_children():
        num = num + 1
        # print(tree.item(item)['values'][5])

        worksheet.write("A" + str(num), tree.item(item)['values'][5])
        worksheet.write("B" + str(num), "外資就活ドットコム")
        worksheet.write("C" + str(num), tree.item(item)['values'][3])
        worksheet.write("D" + str(num), tree.item(item)['values'][4])

    workbook.close()

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="データ取得する", command=loadData)
filemenu.add_command(label="Excel作成する", command=csvMake)
filemenu.add_command(label="API連携する", command=loadData)
menubar.add_cascade(label="メニュー", menu=filemenu)
root.config(menu=menubar)

tree.pack(fill=tk.BOTH, expand=True, pady=0)

root.mainloop()
# index = 0
# for x in result:        
#     index = index + 1
#     print(index)
#     print(x.decode('utf8', 'ignore'))


