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
root.title("https://www.onecareer.jp/")

style = ttk.Style()
style.configure('Treeview', rowheight=30)
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

tree.column("#0", 			anchor=tk.CENTER, width=10)
tree.column("#1", 			anchor=tk.CENTER, width=50)
tree.column("#2", 		anchor=tk.CENTER, width=50)
tree.column("#3", 			anchor=tk.CENTER, width=200)
tree.column("#4", 			anchor=tk.CENTER, width=200)
tree.column("#5", 	anchor=tk.CENTER, width=200)
tree.column("#6", 	anchor=tk.CENTER, width=300)
tree.column("#7", 			anchor=tk.CENTER, width=500)
# tree.column("product_category", anchor=tk.CENTER, width=100)
# tree.column("product_mana_no", 	anchor=tk.CENTER, width=100)
# tree.column("product_sale_no", 	anchor=tk.CENTER, width=100)
# tree.column("product_id", 		anchor=tk.CENTER, width=100)
# tree.column("login_id", 		anchor=tk.CENTER, width=100)
# tree.column("bid_amount", 		anchor=tk.CENTER, width=100)
# tree.column("closed_date", 		anchor=tk.CENTER, width=100)
tree.heading("#0", 				anchor=tk.CENTER, text="")
tree.heading("#1", 				anchor=tk.CENTER, text="No")
tree.heading("#2", 				anchor=tk.CENTER, text="本日締切!")
tree.heading("#3", 		anchor=tk.CENTER, text="開始時間")
tree.heading("#4", 			anchor=tk.CENTER, text="完了時間")
tree.heading("#5", 				anchor=tk.CENTER, text="カテゴリー")
tree.heading("#6", 		anchor=tk.CENTER, text="会社名")
tree.heading("#7",		anchor=tk.CENTER, text="タイトル")



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
    
    link = "https://oc2-api.www.onecareer.jp/api/v1/promotion_events?category=entry&graduates="

    result = requests.get(link, headers=headers)    
    rows = result.json()
    print(rows)
    index = 0

    for event in rows['promotion_events']:
        index = index + 1
    
        tree.insert("", tk.END, values=(index, hasLiveSchedule(event['has_live_schedule']), formatDateTime(event['closest_schedule']['start_at'] ), formatDateTime(event['closest_schedule']['time_limit_at']), event['company']['business_category_name'], 
                                        event['company']['name'], event['title']))   

        
    link = "https://oc2-api.www.onecareer.jp/api/v1/events?page=1&per=1000&sort=time_limit_at&category=entry&business_subcategory_ids=&start_at=&end_at=&days=&company_ids=&companies_to_follow=false&graduates=&search_by_current_user_graduate_year=false&prefectures[online]=false&prefectures[prefecture_ids]=&is_onecareer_hosting_event=false&upper_limit_of_company_review=5&lower_limit_of_company_review=0&since_time_limit_at=&until_time_limit_at=&held_months="     

    result = requests.get(link, headers=headers)
    rows = result.json()
    for event in rows['events']:
        index = index + 1
    
        tree.insert("", tk.END, values=(index, hasLiveSchedule(event['has_live_schedule']), formatDateTime(event['closest_schedule']['start_at'] ), formatDateTime(event['closest_schedule']['time_limit_at']), event['company']['business_category_name'], 
                                        event['company']['name'], event['title']))
        
    tree.tag_configure("bold", font=("Arial", 10, ""))
    for item in tree.get_children():
        tree.item(item, tags=("bold",))

def csvMake (): 
    workbook = xlsxwriter.Workbook("OneCarrer.xlsx")
    worksheet = workbook.add_worksheet()
    num = 1
    
    worksheet.write("A" + str(num), "締切日（降順）")
    worksheet.write("B" + str(num), "サイト名")
    worksheet.write("C" + str(num), "企業名")
    worksheet.write("D" + str(num), "タイトル")
    
    for item in tree.get_children():
        num = num + 1
        # print(tree.item(item)['values'][5])

        worksheet.write("A" + str(num), tree.item(item)['values'][2])
        worksheet.write("B" + str(num), "One Carrer")
        worksheet.write("C" + str(num), tree.item(item)['values'][5])
        worksheet.write("D" + str(num), tree.item(item)['values'][6])

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


