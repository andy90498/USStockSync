# 🧱 標準函式庫
import os
import sys
import shutil
import time
import threading
import datetime
from pathlib import Path
#GUI模組（tkinter + ttkbootstrap）
import tkinter as tk
from tkinter import messagebox, simpledialog, scrolledtext, ttk
from ttkbootstrap import Style
import requests
from bs4 import BeautifulSoup
from pandas.io.formats.format import return_docstring

#設定
from config import EXCEL_PATH, CREDENTIALS_PATH, BASE, APPDATA
from sync import (
    smart_write_to_google_sheet,
    sync_group_to_google_sheet,
    enqueue_group_sync,
    start_group_sync_worker
)
from group import (
    save_group_data,
    load_group_data,
    rename_group as rename_group_excel,
    delete_group as delete_group_excel,
    update_single_group_column
)
from data_fetcher import (
    is_file_locked,
    close_excel_instances,
    fetch_overview_metrics,
    fetch_financial_metrics,
    fetch_total_debt,
    write_to_excel
)
######################################################import尾端########################################################

def get_internal_resource(filename):
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / filename
    return Path(__file__).parent / filename

APPDATA.mkdir(parents=True, exist_ok=True)                                                                              #於APPDATA中建立該變數值的資料夾


class StockApp:
    def __init__(self, root):                                                                                           #GUI元件建立及排版
        self.root = root
        self.root.title("📈 US Stock Offline Data Update System")
        self.file_path = EXCEL_PATH
        self.codes = []

        self.style = Style("minty")
        self.mainframe = tk.Frame(root)
        self.mainframe.grid(row=0, column=0, padx=20, pady=20)                                                          #主畫面主容器

        self.title_label = tk.Label(self.mainframe, text="📈 US StockSync", font=("Arial", 20, "bold"))
        self.title_label.grid(row=0, column=0, columnspan=3, pady=10)                                                   #標題 Label

        self.status_label = tk.Label(self.mainframe, text="Stock Code", font=("Arial", 10))
        self.status_label.grid(row=1, column=0, padx=3)                                                                 #股票代碼文字說明 Label


        # 建一個frame準備將listbox和scrollbar放一起
        listbox_frame = tk.Frame(self.mainframe)
        listbox_frame.grid(row=2, column=0, padx=5, pady=5, sticky="n")                                                 #listbox群組
        # 建立側邊的捲動bar
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical")
        # 建立股票代碼的Listbox
        self.stock_listbox = tk.Listbox(listbox_frame, width=30, height=21, yscrollcommand=scrollbar.set)
        self.stock_listbox.grid(row=0, column=0, sticky="ns")
        # 將Scrollbar綁定到Listbox
        scrollbar.config(command=self.stock_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")


        self.entry = tk.Entry(self.mainframe)
        self.entry.grid(row=3, column=0, pady=5)                                                                        #股票代碼輸入框
        self.entry.bind("<Return>", lambda event: self.add_code())
        self.entry.focus_set()

        #將群組名稱初始化為Default
        self.current_group = "Default Group"
        self.groups = {self.current_group: []}

        try:
            self.groups = load_group_data()
            if self.groups:
                #如果True 就指定第一項名稱到變數中
                self.current_group = list(self.groups.keys())[0]
        except:
            #如果找不到，就建一個
            self.groups = {self.current_group: []}
            #最後呼叫更新群組資料的def
        self.refresh_stock_list()

        #建立群組功能的Frame
        group_frame = tk.Frame(self.mainframe)
        group_frame.grid(row=1, column=1, sticky="w")                                                                   # 加上靠左對齊(w) 群組功能區的Frame（放在 Listbox 上方）
        self.group_var = tk.StringVar(value=self.current_group)
        self.group_menu = ttk.Combobox(group_frame, textvariable=self.group_var, state="readonly", width=18)
        self.group_menu['values'] = list(self.groups.keys())
        self.group_menu.grid(row=0, column=0, padx=(0, 5), pady=5, sticky="w")                                          # 靠左放，右邊空隙小一點        # 群組下拉選單 Combobox
        tk.Button(group_frame, text="➕ Add", command=self.add_group).grid(row=0, column=1, padx=(0, 3), pady=5)        # 新增群組按鈕
        tk.Button(group_frame, text="✏️ Rename", command=self.rename_group).grid(row=0, column=2, padx=(0, 3), pady=5)  # 重新命名群組按鈕
        tk.Button(group_frame, text="🗑 Del", command=self.delete_group).grid(row=0, column=3, padx=0, pady=5)           # 刪除群組按鈕

        #建立 股票代碼功能區的Frame
        btn_frame = tk.Frame(self.mainframe)
        btn_frame.grid(row=7, column=0, padx=5)                                                                         # 新增/刪除代碼按鈕的容器 Frame
        tk.Button(btn_frame, text="+ Add", command=self.add_code).grid(row=0, column=0, padx=5)                         # 新增代碼按鈕
        tk.Button(btn_frame, text="🗑 Del", command=self.remove_code).grid(row=0, column=1, padx=5)                      # 刪除代碼按鈕


        tk.Button(self.mainframe, text="🔄 Update Data", command=self.update_data).grid(row=7, column=1, padx=10)       # 更新資料按鈕

        self.status_label = tk.Label(self.mainframe, text="Last updated record", font=("Arial", 10))
        self.status_label.grid(row=0, column=0, padx=10)                                                                # 資料更新時間狀態 Label

        self.log_area = scrolledtext.ScrolledText(self.mainframe, width=90, height=20, font=("Courier New", 10))
        self.log_area.grid(row=2, column=1, columnspan=2,padx=10, pady=5,sticky="n")                                    # Log區域Text widget(console使用)
        self.log_area.config(state="disabled")  # 預設編輯模式為唯讀

        #進度條群組(沒建Frame)
        self.progress = ttk.Progressbar(self.mainframe, length=750, mode="determinate")
        self.progress.grid(row=3, column=1, padx=10, pady=5, sticky="w")                                                # 進度條
        self.progress_label = tk.Label(self.mainframe, text="0%", font=("Arial", 10), bg="SystemButtonFace")
        self.progress_label.place(in_=self.progress, relx=0.5, rely=0.5, anchor="center")                               # 進度條百分比的文字label
        #透過trace_add來追蹤group_var的變化，連動觸發start_group_sync_worker，後方的spreadsheet_id於上傳github後刪除
        self.group_var.trace_add("write", self.on_group_change)
        #spreadsheet_id取得方式 為google sheet網址中的 https://docs.google.com/spreadsheets/d/這個位置/edit?gid=0#gid=0
        start_group_sync_worker(spreadsheet_id="<刪除>")

    def add_code(self):
        code = self.entry.get().strip()

        # ✨ 彩蛋
        easter_triggers = ["vermilion", "作者", "author", "developers"]
        if code.lower() in [kw.lower() for kw in easter_triggers]:
            messagebox.showinfo(
                "🎨 關於作者 Vermilion",
                "你好，這個工具由 Vermilion 精心開發，\n\n誠心希望你會喜歡，感謝你的支持與鼓勵 🧡，\n\n如果有任何問題，歡迎來信 andy90498.vc@gmail.com。"
            )
            self.entry.delete(0, tk.END)  # 清空輸入框
            return

        code = code.upper()
        if code and code not in self.groups[self.current_group]:
            self.groups[self.current_group].append(code)
            self.stock_listbox.insert(tk.END, code)
            self.entry.delete(0, tk.END)

        #呼叫def將新增的股票代碼存入陣列中，並且透過enqueue排程上傳self.group
        update_single_group_column(self.current_group, self.groups[self.current_group])
        enqueue_group_sync(self.groups)


    def remove_code(self):
        selected = self.stock_listbox.curselection()
        if selected:
            idx = selected[0]
            del self.groups[self.current_group][idx]
            self.stock_listbox.delete(idx)
        # 呼叫def將刪除之後剩下的股票代碼存入陣列中，並且透過enqueue排程上傳self.group
        update_single_group_column(self.current_group, self.groups[self.current_group])
        enqueue_group_sync(self.groups)


    def update_data(self):
        #背景執行target(這邊指定_crawl_data是為了避免在爬蟲的過程中console不會更新，提高流暢的感覺
        threading.Thread(target=self._crawl_data, daemon=True).start()

    def _crawl_data(self):
        codes = self.groups[self.current_group]
        total = len(codes)

        if not codes:
            messagebox.showinfo("提示", "請先新增股票代碼")
            return

        today = datetime.datetime.now().strftime("%Y-%m-%d")
        self.log_message(
            f"----------------{today} System Initialization----------------\n"
            "           ✅ Connection successful.    ❌ Connection failed"
        )

        #初始化進度條及值
        self.progress["value"] = 0
        self.progress["maximum"] = total

        #初始化爬蟲的字典、陣列、還有成功連線的數量
        extra_data_dict = {}
        all_data_rows = []  # ← 累積成功資料
        success_count = 0

        #從codes的第一筆開始迴圈到最後一筆
        for index, code in enumerate(codes, 1):
            #先從連線過程會遇到的情況開始列舉
            try:
                #發出request
                response = requests.get(f"https://stockanalysis.com/stocks/{code}/", timeout=10)
                if response.status_code != 200:
                    raise Exception("連線失敗")

                #BeautifulSoup為解析指定網站的原始碼
                soup = BeautifulSoup(response.text, "html.parser")
                #先抓標題(這邊會抓到公司名稱)
                title = soup.title.string if soup.title else "Unknown resolution"
                title = title.replace(f" ({code}) Stock Price & Overview", "").strip()
                self.log_message(f"✅ {code} | {title}")

                # 資料擷取完成後的各變數，各變數來源自後方的def(這邊共送出三次request)
                code, shares_out, pe_ratio, price_target = fetch_overview_metrics(code)
                financials = fetch_financial_metrics(code)
                td_data = fetch_total_debt(code)
                financials.update(td_data)

                # 加入成功清單
                extra_data_dict[code] = {
                    **financials,
                    "公司名稱": title
                }
                all_data_rows.append((code, shares_out, pe_ratio, price_target))
                success_count += 1

            except Exception as e:
                self.log_message(f"{code} ❌ attach fail：{e}")

            # 更新進度條
            self.progress["value"] = index
            self.progress.update_idletasks()
            #這邊index一定要是int，不然會出錯
            self.progress_label.config(text=f"{int((index / total) * 100)}%")

        # 爬完之後，最後再一起寫入成功的資料
        if all_data_rows:
            #寫入本地excel檔
            write_to_excel(
                all_data_rows,
                extra_data_dict=extra_data_dict,
                file_path=str(self.file_path),
                sheet_name="Crawl_data",
            )
            #寫入雲端google sheet(spreadsheet_id於上傳github時刪除)
            smart_write_to_google_sheet(
                all_data_rows,
                extra_data_dict=extra_data_dict,
                spreadsheet_id="<刪除>",
                sheet_name="Crawl_data",
                log_fn=self.log_message
            )

        #整理時間格式，並寫入GUI左上角的label
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status_label.config(text=f"Last updated：{now}")
        self.log_message(f"----------------Update Completed（Finished {success_count} /  {total}）----------------")

    def update_group_menu(self):
        #將字典的值丟到combox的選項裡面
        self.group_menu['values'] = list(self.groups.keys())
        #不管值有無變化，一併更新到current_group
        self.group_var.set(self.current_group)
        self.refresh_stock_list()

    def on_group_change(self, *args):
        #使用者切換群組
        selected = self.group_var.get()
        #確認群組名稱正確
        if selected in self.groups:
            #將選取的群組名稱設定為current，並更新清單
            self.current_group = selected
            self.refresh_stock_list()

    def refresh_stock_list(self):
        #先刪除listbox的所有資料
        self.stock_listbox.delete(0, tk.END)
        #再從current_group中把每一項list加回來，因通常是之前先設好最新的current_group才會呼叫這段來更新
        for code in self.groups[self.current_group]:
            self.stock_listbox.insert(tk.END, code)

    def add_group(self):
        #彈出askstring引導輸入群組名稱
        name = simpledialog.askstring("輸入群組名稱", "請輸入新的群組名稱：")
        #如果沒有輸入或是輸入空格就return掉
        if not name or name.strip() == "":
            return  # 使用者取消或留空就直接返回
        name = name.strip()

        #彩蛋
        easter_triggers = ["vermilion", "作者", "author", "developers"]
        if name.lower() in [kw.lower() for kw in easter_triggers]:
            messagebox.showinfo(
                "🎨 關於作者 Vermilion",
                "你好，這個工具由 Vermilion 精心開發，\n\n誠心希望你會喜歡，感謝你的支持與鼓勵 🧡，\n\n如果有任何問題，歡迎來信 andy90498.vc@gmail.com。"
            )
            return

        if name in self.groups:
            messagebox.showinfo("錯誤", "群組名稱已存在")
            return

        #將剛剛合格的群組name新增到groups字典中
        self.groups[name] = []
        self.current_group = name
        self.update_group_menu()
        self.refresh_stock_list()
        save_group_data(self.groups)
        enqueue_group_sync(self.groups)#排程進enqueue進行背景同步


    def rename_group(self):
        old_name = self.current_group
        new_name = simpledialog.askstring("重新命名群組", "請輸入新的群組名稱：", initialvalue=old_name)
        if not new_name or new_name.strip() == "":
            return  #沒有值或是只有空格return

        new_name = new_name.strip()
        if new_name == old_name:
            return  #名字相同return
        if new_name in self.groups:
            messagebox.showinfo("錯誤", "群組名稱已存在")
            return  #字典中查到同名return

        # 將舊群組資料轉移並刪除舊的
        self.groups[new_name] = self.groups.pop(old_name)
        self.current_group = new_name
        self.update_group_menu()
        self.refresh_stock_list()
        #呼叫轉移本地excel資料的def
        rename_group_excel(old_name, new_name)
        save_group_data(self.groups)
        #因為google sheet都是整頁clear掉才重新輸入，所以不用特地重寫一個rename的def
        enqueue_group_sync(self.groups)  # 推進最新資料


    def delete_group(self):
        if len(self.groups) == 1:
            messagebox.showinfo("提示", "至少需要保留一個群組")
            return
        del self.groups[self.current_group]
        self.current_group = list(self.groups.keys())[0]
        self.update_group_menu()
        self.refresh_stock_list()
        save_group_data(self.groups)
        enqueue_group_sync(self.groups)


    def log_message(self, message):
        #console使用，先建好時間模板的變數
        now = time.strftime("%H:%M:%S")
        full_message = f"[{now}] {message}\n"
        #tk模板預設唯獨，state要改normal才能寫入
        self.log_area.config(state="normal")
        #匯入訊息
        self.log_area.insert(tk.END, full_message)
        #滑動到最下方
        self.log_area.see(tk.END)
        #設定唯獨
        self.log_area.config(state="disabled")

    def run_crawler(self):
        #先辨識current的群組，再透過list把這群組的每筆資料傳入codes這個陣列中
        codes = list(self.groups[self.current_group])
        #建立一個迴圈，內容是將每筆股票代碼丟到爬蟲裡面並回傳結果到results這個陣列中
        results = [fetch_overview_metrics(code) for code in codes]
        #將每一筆結果丟到本地excel中
        write_to_excel(results)


def try_close_excel_or_abort():
    #先透過is_file_locked確認excel目前locked的狀態是不是true (EXCEL_PATH來自config.py)
    if is_file_locked(EXCEL_PATH):
        choice = messagebox.askyesno("Excel 檔案開啟", "偵測到 Excel 正在使用該檔案。\n要關閉後繼續嗎？")
        if not choice:
            #這邊的return是回給呼叫這個def的前一程序，並使用布林值會更方便判斷
            return False
        #選是的話，就呼叫關閉程式的def，目標指向EXCEL_PATH
        closed = close_excel_instances(os.path.basename(EXCEL_PATH))
        return closed
    return True

#主程序
if __name__ == "__main__":
    #自動補檔案，分別是金鑰跟EXCEL檔
    for filename in ["爬蟲更新區.xlsm", "credentials.json"]:
        dest = APPDATA / filename
        if not dest.exists():
            try:
                source = get_internal_resource(filename)
                shutil.copyfile(source, dest)
                print(f"✅ 補上：{dest.name}")
            except Exception as e:
                print(f"❌ 無法補 {filename}：{e}")
    #
    if not EXCEL_PATH.exists():
        shutil.copy(BASE / "爬蟲更新區.xlsm", EXCEL_PATH)
    if not CREDENTIALS_PATH.exists():
        shutil.copy(BASE / "credentials.json", CREDENTIALS_PATH)

    #詢問要不要強制關閉EXCEL
    try_close_excel_or_abort()

    #建立Tkinter容器
    root = tk.Tk()
    #建立GUI的物件
    app = StockApp(root)
    #啟動GUI
    root.mainloop()