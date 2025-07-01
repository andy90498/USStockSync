import gspread
from google.oauth2.service_account import Credentials
import threading
import queue
import time
import re
from datetime import datetime
from tkinter import messagebox
from config import CREDENTIALS_PATH, EXCEL_PATH

_group_queue = queue.Queue()
_latest_groups = None

#數字轉column
def get_column_letter(n):
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)  #26進位制
        result = chr(65 + rem) + result
    return result

#將EXCEL的標題欄位從網頁原始格式轉成自定義
def clean_column_name(name):
    if name == "公司名稱":
        return name

    #處理EPS、FREE CASH FLOW、DEBT + TTM
    #正則表達式比對EPS (TTM) 或 Free Cash Flow (TTM) 或 Total Debt (TTM)
    #.*萬用字元
    match_ttm = re.search(r"(EPS|Free Cash Flow|Total Debt).*\(TTM\)", name)
    if match_ttm:
        #如果比對成功，將字串修改為 EPS(TTM)、Free Cash Flow(TTM)、Total Debt(TTM)
        return f"{match_ttm.group(1)}(TTM)"

    #處理FY年份欄位
    # 正則表達式比對EPS (年分) 或 Free Cash Flow (年分) 或 Total Debt (年分)
    #\d{4} 表西元年
    match_year = re.search(r"(EPS|Free Cash Flow|Total Debt).*?\(FY (\d{4})\)", name)
    if match_year:
        return f"{match_year.group(1)}({match_year.group(2)})"

    return name.strip()

def extract_year(col):
    match = re.search(r"\((TTM|\d{4})\)", col)
    #如果有抓到TTM或是年分
    if match:
        val = match.group(1)
        #如果是TTM則回傳9999、如果是年份則回傳年分、如果都不是則回傳-1
        return 9999 if val == "TTM" else int(val)
    return -1


def sort_headers_with_priority(headers, priority_list):
    headers_set = set(headers)
    ordered = []
    for col in priority_list:
        if col in headers_set:
            ordered.append(col)
    for col in headers:
        if col not in ordered:
            ordered.append(col)
    return ordered

def smart_write_to_google_sheet(
    data_rows,
    extra_data_dict=None,
    spreadsheet_id=None,
    sheet_name="Crawl_data",
    log_fn=None,
):
    def do_update():
        if not spreadsheet_id:
            log_fn(f"❌ 未寫入 Spreadsheet ID")
            return

        credentials = Credentials.from_service_account_file(
            CREDENTIALS_PATH,
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        gc = gspread.authorize(credentials)
        sh = gc.open_by_key(spreadsheet_id)

        try:
            worksheet = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=sheet_name, rows="1000", cols="40")

        rows = worksheet.get_all_values()
        current_headers = rows[0] if rows else []

        static_cols = ["股票代碼", "公司名稱", "SharesOut", "PE Ratio", "Price Target"]

        dynamic_keys = set()
        if extra_data_dict:
            for val in extra_data_dict.values():
                for key in val:
                    if key != "公司名稱":
                        dynamic_keys.add(key)

        cleaned_keys = [clean_column_name(key) for key in dynamic_keys]
        sorted_dynamic_cols = sorted(
            cleaned_keys, key=lambda c: -extract_year(c) if extract_year(c) != -1 else 0
        )

        #標題順序
        desired_order = [
            "股票代碼", "公司名稱", "SharesOut", "PE Ratio", "Price Target",
            "EPS(TTM)", "EPS(2025)", "EPS(2024)", "EPS(2023)", "EPS(2022)", "EPS(2021)", "EPS(2020)",
            "Free Cash Flow(TTM)", "Free Cash Flow(2025)","Free Cash Flow(2024)", "Free Cash Flow(2023)", "Free Cash Flow(2022)","Free Cash Flow(2021)","Free Cash Flow(2020)",
            "Total Debt(TTM)","Total Debt(2025)","Total Debt(2024)", "Total Debt(2023)","Total Debt(2022)","Total Debt(2021)","Total Debt(2020)", "更新時間"
        ]

        # 先建立欄位集合，包含資料出現過的欄位
        raw_all_headers = static_cols + sorted_dynamic_cols + ["更新時間"]

        # 若有desired_order有出現但raw_all_headers沒有則append在後面
        for col in desired_order:
            if col not in raw_all_headers:
                raw_all_headers.append(col)

        #依據上面的順序排列標題
        all_headers = sort_headers_with_priority(raw_all_headers, desired_order)

        if current_headers != all_headers:
            if not current_headers:
                worksheet.append_row(all_headers)
            else:
                worksheet.update("A1", [all_headers])
            current_headers = all_headers

        col_index_map = {col: idx + 1 for idx, col in enumerate(current_headers)}
        existing_rows = worksheet.get_all_values()[1:]

        code_to_row = {
            row[0].strip().upper(): idx + 2
            for idx, row in enumerate(existing_rows)
            if row and row[0].strip()
        }

        for code, shares, pe, pt in data_rows:
            code_key = code.strip().upper()
            extra = extra_data_dict.get(code, {}) if extra_data_dict else {}
            company_name = extra.get("公司名稱", "-")

            row_data = [""] * len(current_headers)
            row_data[col_index_map["股票代碼"] - 1] = code
            row_data[col_index_map["公司名稱"] - 1] = company_name
            row_data[col_index_map["SharesOut"] - 1] = shares
            row_data[col_index_map["PE Ratio"] - 1] = pe
            row_data[col_index_map["Price Target"] - 1] = pt

            for cleaned_key in col_index_map:
                if cleaned_key in static_cols or cleaned_key == "更新時間":
                    continue  # 靜態欄位交由前面處理

                # 嘗試從 extra 資料中找對應的原始 key
                raw_match = next((rk for rk in extra if clean_column_name(rk) == cleaned_key), None)

                if raw_match:
                    val = extra.get(raw_match)
                    row_data[col_index_map[cleaned_key] - 1] = val if val not in [None, "", "NA", "N/A"] else "-"
                else:
                    row_data[col_index_map[cleaned_key] - 1] = "-"

            row_data[col_index_map["更新時間"] - 1] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            try:
                if code_key in code_to_row:
                    row_num = code_to_row[code_key]
                    end_col_letter = get_column_letter(len(current_headers))
                    cell_range = f"A{row_num}:{end_col_letter}{row_num}"
                    worksheet.update(cell_range, [row_data])
                    log_fn(f"📌 Renew{code}|{company_name} to {row_num} row")
                else:
                    worksheet.append_row(row_data)
                    existing_rows = worksheet.get_all_values()[1:]
                    new_row = len(existing_rows) + 1  # 或 +2，看你 header 是否包含在內
                    log_fn(f"➕ add{code}|{company_name} to {new_row} row")



            except gspread.exceptions.APIError as e:
                log_fn(f"⚠️ 寫入 {code} |{company_name}發生錯誤：{e}")

    threading.Thread(target=do_update, daemon=True).start()

def sync_group_to_google_sheet(groups_dict, spreadsheet_id, sheet_name="Group"):
    def do_sync():
        credentials = Credentials.from_service_account_file(
            CREDENTIALS_PATH,
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        gc = gspread.authorize(credentials)

        try:
            sh = gc.open_by_key(spreadsheet_id)
            if sheet_name in [ws.title for ws in sh.worksheets()]:
                worksheet = sh.worksheet(sheet_name)
                worksheet.clear()
            else:
                worksheet = sh.add_worksheet(title=sheet_name, rows="100", cols="30")
        except Exception as e:
            print(f"❌ 同步 Group 分頁失敗：{e}")
            return

        max_len = max((len(codes) for codes in groups_dict.values()), default=0)
        rows = [list(groups_dict.keys())]
        for i in range(max_len):
            row = [groups_dict[g][i] if i < len(groups_dict[g]) else "" for g in groups_dict]
            rows.append(row)

        worksheet.update("A1", rows)

    threading.Thread(target=do_sync, daemon=True).start()

def _group_sync_worker(spreadsheet_id, sheet_name="Group"):
    global _latest_groups
    last_sync_time = 0
    debounce_seconds = 10

    while True:
        try:
            groups = _group_queue.get(timeout=2)
            _latest_groups = groups
            last_sync_time = time.time()
        except queue.Empty:
            if _latest_groups and (time.time() - last_sync_time > debounce_seconds):
                sync_group_to_google_sheet(_latest_groups, spreadsheet_id, sheet_name)
                _latest_groups = None

def start_group_sync_worker(spreadsheet_id, sheet_name="Group"):
    t = threading.Thread(target=_group_sync_worker, args=(spreadsheet_id, sheet_name), daemon=True)
    t.start()

def enqueue_group_sync(groups_dict):
    _group_queue.put(groups_dict)