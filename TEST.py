import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import load_workbook
import json
import threading
import queue
import traceback

CONFIG_FILE = 'config.json'

# 全域 log 佇列（背景執行緒寫入，主執行緒讀取並顯示在 UI）
LOG_QUEUE = queue.Queue()


def log(message):
    """輸出 log 到終端機，並將訊息放入佇列供 UI 顯示"""
    print(message)
    LOG_QUEUE.put(("log", message))


# ==================== Excel 處理函式 ====================

def getCellValue(sheet, search_header, search_value, target_header, header_row):
    """
    根據指定欄位 A 的值找到對應列，並取得該列的另一個欄位的值。

    Args:
        sheet: Excel 工作表物件
        search_header: 欄位 A 的標題（作為搜尋條件的欄位）
        search_value: 要搜尋的值
        target_header: 要取得值的欄位標題
        header_row: 標題所在的列數
    """
    # 取得欄位 A 的欄索引
    search_col = get_column_index(sheet, search_header, header_row)
    # 取得要更新欄位的欄索引
    target_col = get_column_index(sheet, target_header, header_row)

    if search_col is None:
        log(f"getCellValue 錯誤：{sheet.title}找不到search_col標題「{search_header}」")
        return
    if target_col is None:
        log(f"getCellValue 錯誤：{sheet.title}找不到target_col標題「{target_header}」")
        return

    target_value_sum = 0
    target_value = 0
    # 從標題列的下一列開始搜尋
    for row in range(header_row + 1, sheet.max_row + 1):
        # 讀取欄位 A 的值
        cell_value = sheet.cell(row=row, column=search_col).value
        if cell_value == search_value:
            # 找到目標列，取得指定欄位的值
            target_value = sheet.cell(row=row, column=target_col).value
            if(target_value is None or str(target_value).strip() == ""):
                target_value = 0
            log(f"在{sheet.title}中，第 {row} 列，欄位 {target_col} 的值為 {target_value}")
            target_value_sum += float(target_value)
            log(f"找到 {search_value}，對應的 {target_header} 值為：{target_value}\n")

    return target_value, target_value_sum


def get_column_index(sheet, header_name, header_row):
    """
    根據標題名稱取得對應的欄位索引（數字）。

    Args:
        sheet: Excel 工作表物件
        header_name: 要尋找的標題名稱
        header_row: 標題列所在的行號

    Returns:
        int: 欄位索引（1 為 A 欄），若找不到則回傳 None
    """
    for col in range(1, sheet.max_column + 1):
        header = sheet.cell(row=header_row, column=col).value
        if header is None and str(header).strip() == "":
            continue  # 跳過空白欄位
        if str(header).strip() == str(header_name).strip():
            return col
    return None


def find_and_update(workbook, sheet, file_path, search_header, search_value, update_header, new_value, header_row):
    """
    根據指定欄位 A 的值找到對應列，並修改該列的另一個欄位。

    Args:
        workbook: Excel 工作簿物件
        sheet: Excel 工作表物件
        file_path: Excel 檔案路徑
        search_header: 欄位 A 的標題（作為搜尋條件的欄位）
        search_value: 要搜尋的值
        update_header: 要修改的欄位標題
        new_value: 要寫入的新值
        header_row: 標題所在的列數
    """
    # 取得欄位 A 的欄索引
    search_col = get_column_index(sheet, search_header, header_row)
    # 取得要更新欄位的欄索引
    update_col = get_column_index(sheet, update_header, header_row)

    if search_col is None:
        log(f"find_and_update 錯誤：{sheet.title}找不到search_col標題「{search_header}」")
        return
    if update_col is None:
        log(f"find_and_update錯誤：{sheet.title}找不到update_col標題「{update_header}」")
        return

    # 從標題列的下一列開始搜尋
    for row in range(header_row + 1, sheet.max_row + 1):
        # 讀取欄位 A 的值
        cell_value = sheet.cell(row=row, column=search_col).value
        if cell_value == search_value:
            # 找到目標列，直接更新指定欄位（不需判斷該儲存格原本的內容）
            sheet.cell(row=row, column=update_col).value = new_value
            log(f"已更新第 {row} 列，欄位 {update_col} 的值為 {new_value}")

    # 儲存變更
    workbook.save(file_path)
    log("檔案已儲存")


def read_column_by_header(sheet, header_name, header_row):
    """
    以標題名稱讀取 Excel 特定欄位的資料。

    Args:
        sheet: Excel 工作表物件
        header_name: 要讀取的欄位標題
        header_row: 標題所在的列數（預設為第 1 列）

    Returns:
        list: 該欄位的所有值
    """
    # 1. 讀取標題列，找出目標標題的欄位索引
    headers = []
    for cell in sheet[header_row]:
        headers.append(cell.value)

    if header_name not in headers:
        raise ValueError(f"找不到標題「{header_name}」，現有標題：{headers}")

    col_index = headers.index(header_name) + 1  # openpyxl 欄位索引從 1 開始

    # 2. 讀取該欄位的所有資料（跳過標題列）
    result = []
    for row in sheet.iter_rows(min_col=col_index, max_col=col_index, values_only=True):
        value = row[0]
        # 篩除 None 與空白字串
        if value is not None and str(value).strip() != "":
            result.append(value)

    return result


def find_and_update_with_col(workbook, sheet, file_path, search_header, search_value_list, update_header, new_value_list, header_row):
    """
    根據指定欄位 A 的值找到對應列，並修改該列的另一個欄位。

    Args:
        workbook: Excel 工作簿物件
        sheet: Excel 工作表物件
        file_path: Excel 檔案路徑
        search_header: 欄位 A 的標題（作為搜尋條件的欄位）
        search_value: 要搜尋的值
        update_header: 要修改的欄位標題
        new_value: 要寫入的新值
        header_row: 標題所在的列數
    """
    # 取得欄位 A 的欄索引
    search_col = get_column_index(sheet, search_header, header_row)
    # 取得要更新欄位的欄索引
    update_col = get_column_index(sheet, update_header, header_row)

    if search_col is None:
        log(f"find_and_update 錯誤：{sheet.title}找不到search_col標題「{search_header}」")
        return
    if update_col is None:
        log(f"find_and_update錯誤：{sheet.title}找不到update_col標題「{update_header}」")
        return

    # 從標題列的下一列開始搜尋
    for row in range(header_row + 1, sheet.max_row + 1):
        # 讀取欄位 A 的值
        cell_value = sheet.cell(row=row, column=search_col).value
        for search_value, new_value in zip(search_value_list, new_value_list):
            if str(cell_value).strip() == str(search_value).strip():
                # 找到目標列，直接更新指定欄位（不需判斷該儲存格原本的內容）
                sheet.cell(row=row, column=update_col).value = new_value
                log(f"已更新第 {row} 列，欄位 {update_col} 的值為 {new_value}")

    # 儲存變更
    workbook.save(file_path)
    log("檔案已儲存")


# ==================== 執行主邏輯 ====================

def run_erp_process(file_path, test_type):
    """執行 ERP Excel 更新流程（在背景執行緒執行，不阻塞 UI）"""
    try:
        with open(CONFIG_FILE, encoding="utf-8") as config_file:
            config = json.load(config_file)
    except Exception as e:
        log(f"錯誤：無法讀取 {CONFIG_FILE}：{e}")
        LOG_QUEUE.put(("error", f"無法讀取 {CONFIG_FILE}：{e}"))
        return

    mainChangeSheet = config["mainChangeSheet"]["sheetName"]
    mainSelectHeader = config["mainChangeSheet"]["selectHeader"]
    mainHeaderRow = config["mainChangeSheet"]["headerRow"]

    dataResourceList = config["dataResource"]
    selectDataList = config["selectDataList"]
    selectResourceList = config["selectResourceList"]

    log(f"已選取檔案：{file_path}")
    # 使用 openpyxl 讀取 Excel 檔案
    workbook = load_workbook(file_path)
    log(f"工作表名稱：{workbook.sheetnames}")

    sheet_names = workbook.sheetnames
    if(test_type == 2 or test_type == 3):
        dataResourceList = selectResourceList

    filtered_data_resource = [
        item for item in dataResourceList
        if item['sheetName'] in sheet_names
    ]

    # 比對設定檔 dataResourceList 與 Excel 檔案 sheetnames 的工作表名稱
    # 主表 mainChangeSheet 不納入比對，從兩邊集合中排除
    config_sheet_names = {item['sheetName'] for item in dataResourceList} - {mainChangeSheet}
    excel_sheet_names = set(sheet_names) - {mainChangeSheet}
    # 資料清單有、檔案缺少的工作表
    missing_in_file = config_sheet_names - excel_sheet_names
    # 檔案有、資料清單缺少的工作表
    missing_in_config = excel_sheet_names - config_sheet_names

    if missing_in_file or missing_in_config:
        warning_lines = []
        for name in sorted(missing_in_file):
            warning_lines.append(f"工作表「{name}」存在於 指定工作表中，但 Excel 檔案中缺少")
        for name in sorted(missing_in_config):
            warning_lines.append(f"工作表「{name}」存在於 Excel 檔案中，但 指定工作表中 缺少")
        warning_msg = "\n".join(warning_lines)
        # 送出提示（不寫入 log），並等待使用者關閉提示視窗後才繼續執行
        wait_event = threading.Event()
        LOG_QUEUE.put(("warning", (warning_msg, wait_event)))
        wait_event.wait()

    workbook = load_workbook(file_path)
    sheet = workbook[mainChangeSheet]
    if(test_type == 1 or test_type == 3):  # 執行指定欄位
        pluNoList = selectDataList
    else:
        pluNoList = read_column_by_header(sheet, mainSelectHeader, mainHeaderRow)
    pluNoList = [str(x).strip() for x in pluNoList]
    log("讀取到的品號清單：")
    for item in pluNoList:
        log(f"  - {item}")

    for item in filtered_data_resource:
        dataSheetName = item["sheetName"]
        dataSelectHeader = item["selectHeader"]
        dataTargetHeader = item["targetHeader"]
        dataHeaderRow = item["headerRow"]
        additionalProcessing = item.get("additionalProcessing", "")
        mainChangerHeader = item["mainChangerHeader"]
        new_value_list = []
        for pluNo in pluNoList:
            dataSheet = workbook[dataSheetName]
            targetValue, targetValueSum = getCellValue(dataSheet, dataSelectHeader, pluNo, dataTargetHeader, dataHeaderRow)

            if additionalProcessing == "-":
                # 如果需要額外處理，這裡可以加入額外的邏輯
                targetValueSum = -targetValueSum
            if additionalProcessing == "No Sum":
                new_value_list.append(targetValue)
            else:
                new_value_list.append(targetValueSum)
        find_and_update_with_col(workbook, sheet, file_path, mainSelectHeader, pluNoList, mainChangerHeader, new_value_list, mainHeaderRow)

    log("=== ERP Excel 更新完成！ ===")
    LOG_QUEUE.put(("done", None))


# ==================== UI 設定編輯器 ====================

RESOURCE_FIELDS = [
    ('mainChangerHeader', '主表更新欄位名稱'),
    ('sheetName', '資料來源工作表名稱'),
    ('selectHeader', '資料來源工作表條件欄位名稱'),
    ('headerRow', '資料來源工作表標題列'),
    ('targetHeader', '資料來源欄位名稱'),
    ('additionalProcessing', '額外處理'),
    ]


class ConfigEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ERP Excel 設定編輯器")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)
        # 預設以全螢幕（最大化）顯示
        self.root.state('zoomed')

        self.config = self.load_config()

        # 建立可捲動的 Canvas 容器（讓整個 UI 都能捲動）
        self.main_canvas = tk.Canvas(self.root, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(self.root, orient='vertical', command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        self.main_scrollbar.pack(side='right', fill='y')
        self.main_canvas.pack(side='left', fill='both', expand=True)

        # 內容 frame（放在 canvas 內）
        self.content_frame = ttk.Frame(self.main_canvas)
        self.content_window = self.main_canvas.create_window((0, 0), window=self.content_frame, anchor='nw')

        # 內容高度變化時更新捲動範圍；canvas 寬度變化時讓內容跟著調整
        self.content_frame.bind(
            '<Configure>',
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox('all'))
        )
        self.main_canvas.bind(
            '<Configure>',
            lambda e: self.main_canvas.itemconfig(self.content_window, width=e.width)
        )

        self.create_main_section()
        self.create_notebook()
        self.create_action_buttons()
        self.create_log_section()

        # 支援滑鼠滾輪捲動（在所有 UI 建立後遞迴綁定所有子元件，確保分頁內部也能捲動整個 UI）
        self._bind_mousewheel(self.content_frame)

        # 定期檢查 log 佇列，將背景執行緒的 log 顯示在 UI
        self.process_log_queue()

    def _on_mousewheel(self, event):
        """滑鼠滾輪捲動整個 UI"""
        self.main_canvas.yview_scroll(int(-event.delta / 120), 'units')
        return 'break'  # 阻止事件繼續傳遞，避免分頁內部同時捲動

    def _bind_mousewheel(self, widget):
        """遞迴綁定滾輪事件到子元件，讓整個 UI 都能捲動。
        跳過有自己捲動功能的元件（Treeview、Listbox、Text、Scrollbar），
        保留它們原本的捲動與拖動功能。
        """
        # 跳過有自己捲動功能的元件，避免攔截它們的滾輪事件
        # 特別注意：Scrollbar 必須跳過，否則綁定 MouseWheel 會阻止滾動條拖動
        if isinstance(widget, (ttk.Treeview, tk.Listbox, tk.Text, ttk.Scrollbar, tk.Scrollbar)):
            return
        try:
            widget.bind('<MouseWheel>', self._on_mousewheel)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._bind_mousewheel(child)

    def load_config(self):
        try:
            with open(CONFIG_FILE, encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取 {CONFIG_FILE}：\n{e}")
            return {}

    # ---------- 主工作表設定 ----------
    def create_main_section(self):
        main_frame = ttk.LabelFrame(self.content_frame, text="主工作表設定 (mainChangeSheet)", padding=10)
        main_frame.pack(fill='x', padx=10, pady=5)

        main_cfg = self.config.get('mainChangeSheet', {})

        ttk.Label(main_frame, text="工作表名稱 :").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.main_sheet_name_var = tk.StringVar(value=main_cfg.get('sheetName', ''))
        ttk.Entry(main_frame, textvariable=self.main_sheet_name_var, width=40).grid(row=0, column=1, sticky='w', padx=5, pady=3)

        ttk.Label(main_frame, text="主要條件欄位 :").grid(row=0, column=2, sticky='w', padx=5, pady=3)
        self.main_select_header_var = tk.StringVar(value=main_cfg.get('selectHeader', ''))
        ttk.Entry(main_frame, textvariable=self.main_select_header_var, width=30).grid(row=0, column=3, sticky='w', padx=5, pady=3)

        ttk.Label(main_frame, text="標題列 :").grid(row=0, column=4, sticky='w', padx=5, pady=3)
        self.main_header_row_var = tk.StringVar(value=str(main_cfg.get('headerRow', 1)))
        ttk.Entry(main_frame, textvariable=self.main_header_row_var, width=10).grid(row=0, column=5, sticky='w', padx=5, pady=3)

    # ---------- Notebook 分頁 ----------
    def create_notebook(self):
        # 使用 clam 主題並設定分頁樣式，讓分頁更明顯
        style = ttk.Style()
        style.theme_use('clam')
        style.map('TNotebook.Tab',
                  background=[('selected', '#007BFF'), ('!selected', '#E8E8E8')],
                  foreground=[('selected', 'white'), ('!selected', '#333333')])

        self.notebook = ttk.Notebook(self.content_frame)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)

        # dataResource 分頁
        self.data_resource_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.data_resource_tab, text="資料來源工作表設定")
        self.create_data_resource_tab()

        # selectResourceList 分頁
        self.select_resource_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.select_resource_tab, text="指定資料來源工作表")
        self.create_select_resource_tab()

        # selectDataList 分頁
        self.select_data_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.select_data_tab, text="指定主表欄位")
        self.create_select_data_tab()


    # ---------- dataResource ----------
    def create_data_resource_tab(self):
        toolbar = ttk.Frame(self.data_resource_tab)
        toolbar.pack(fill='x', padx=5, pady=5)
        ttk.Button(toolbar, text="新增", command=self.add_data_resource).pack(side='left', padx=3)
        ttk.Button(toolbar, text="編輯", command=self.edit_data_resource).pack(side='left', padx=3)
        ttk.Button(toolbar, text="刪除", command=self.delete_data_resource).pack(side='left', padx=3)
        ttk.Button(toolbar, text="複製到指定資料來源", command=self.copy_to_select_resource).pack(side='left', padx=3)

        columns = ('mainChangerHeader', 'sheetName', 'selectHeader', 'headerRow', 'targetHeader', 'additionalProcessing')
        # 先建立 Treeview，再建立捲軸（避免 command 引用尚未建立的 widget）
        self.data_resource_tree = ttk.Treeview(self.data_resource_tab, columns=columns, show='headings', height=15)

        # command=self.data_resource_tree.yview 讓滾動條能點擊拖動調整內容
        data_scrollbar = ttk.Scrollbar(self.data_resource_tab, orient='vertical', command=self.data_resource_tree.yview)
        data_scrollbar.pack(side='right', fill='y')

        headers = {
            'mainChangerHeader': '主表更新欄位名稱',
            'sheetName': '資料來源工作表名稱',
            'selectHeader': '資料來源工作表條件欄位名稱',
            'headerRow': '資料來源工作表標題列',
            'targetHeader': '資料來源欄位名稱',
            'additionalProcessing': '額外處理',
        }
        widths = {
            'mainChangerHeader': 130,
            'sheetName': 220,
            'selectHeader': 120,
            'headerRow': 60,
            'targetHeader': 120,
            'additionalProcessing': 100,
        }
        for col in columns:
            self.data_resource_tree.heading(col, text=headers[col])
            self.data_resource_tree.column(col, width=widths[col], anchor='center')

        self.data_resource_tree.pack(fill='both', expand=True, padx=5, pady=5)
        self.data_resource_tree.bind('<Double-1>', lambda e: self.edit_data_resource())

        self.data_resource_tree.configure(yscrollcommand=data_scrollbar.set)

        self.refresh_data_resource_tree()

    def refresh_data_resource_tree(self):
        self.data_resource_tree.delete(*self.data_resource_tree.get_children())
        for item in self.config.get('dataResource', []):
            self.data_resource_tree.insert('', 'end', values=(
                item.get('mainChangerHeader', ''),
                item.get('sheetName', ''),
                item.get('selectHeader', ''),
                item.get('headerRow', ''),
                item.get('targetHeader', ''),
                item.get('additionalProcessing', ''),
            ))

    def get_selected_data_resource_index(self):
        selection = self.data_resource_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "請先選取一個項目")
            return None
        return self.data_resource_tree.index(selection[0])

    def add_data_resource(self):
        self.open_resource_dialog(self.data_resource_tab, None, is_data_resource=True)

    def edit_data_resource(self):
        idx = self.get_selected_data_resource_index()
        if idx is not None:
            self.open_resource_dialog(self.data_resource_tab, idx, is_data_resource=True)

    def delete_data_resource(self):
        idx = self.get_selected_data_resource_index()
        if idx is None:
            return
        if messagebox.askyesno("確認", "確定要刪除這個資料來源嗎？"):
            del self.config['dataResource'][idx]
            self.refresh_data_resource_tree()

    def copy_to_select_resource(self):
        """將選取的資料來源複製到指定資料來源 (selectResourceList)"""
        idx = self.get_selected_data_resource_index()
        if idx is None:
            return
        item = self.config['dataResource'][idx]
        # 深層複製，避免修改時互相影響
        import copy
        new_item = copy.deepcopy(item)
        self.config.setdefault('selectResourceList', []).append(new_item)
        self.refresh_select_resource_tree()
        messagebox.showinfo("成功", f"已複製「{item.get('mainChangerHeader', '')}」到指定資料來源")

    # ---------- 資源編輯對話框 ----------
    def open_resource_dialog(self, parent, index, is_data_resource=True):
        dialog = tk.Toplevel(parent)
        dialog.title("編輯資料來源" if is_data_resource else "編輯指定資源")
        dialog.grab_set()
        dialog.resizable(False, False)

        key = 'dataResource' if is_data_resource else 'selectResourceList'
        items = self.config.setdefault(key, [])
        item = items[index] if index is not None else {}

        vars = {}
        row = 0
        for field, label in RESOURCE_FIELDS:
            ttk.Label(dialog, text=f"{label} :").grid(row=row, column=0, sticky='w', padx=10, pady=5)
            var = tk.StringVar(value=str(item.get(field, '')))
            vars[field] = var
            if field == 'additionalProcessing':
                # 額外處理使用下拉選單
                combo = ttk.Combobox(
                    dialog, textvariable=var, width=37, state='readonly',
                    values=['', '-', 'No Sum']
                )
                combo.grid(row=row, column=1, sticky='w', padx=10, pady=5)
            else:
                ttk.Entry(dialog, textvariable=var, width=40).grid(row=row, column=1, sticky='w', padx=10, pady=5)
            row += 1

        def on_save():
            try:
                new_item = {}
                for field, _ in RESOURCE_FIELDS:
                    value = vars[field].get().strip()
                    if field == 'headerRow':
                        new_item[field] = int(value)
                    else:
                        new_item[field] = value
                if index is not None:
                    items[index] = new_item
                else:
                    items.append(new_item)
                if is_data_resource:
                    self.refresh_data_resource_tree()
                else:
                    self.refresh_select_resource_tree()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("錯誤", "標題列 (headerRow) 必須是數字", parent=dialog)

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="儲存", command=on_save).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side='left', padx=5)

    # ---------- selectDataList ----------
    def create_select_data_tab(self):
        frame = ttk.Frame(self.select_data_tab)
        frame.pack(fill='both', expand=True, padx=5, pady=5)

        list_frame = ttk.Frame(frame)
        list_frame.pack(side='left', fill='both', expand=True)

        # 先建立 Listbox，再建立捲軸（避免 command 引用尚未建立的 widget）
        self.select_data_listbox = tk.Listbox(list_frame, height=20)

        # command=self.select_data_listbox.yview 讓滾動條能點擊拖動調整內容
        list_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.select_data_listbox.yview)
        list_scrollbar.pack(side='right', fill='y')
        self.select_data_listbox.pack(side='left', fill='both', expand=True)
        self.select_data_listbox.configure(yscrollcommand=list_scrollbar.set)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(side='right', fill='y', padx=10)

        ttk.Label(btn_frame, text="品號:").pack(pady=5)
        self.select_data_entry = ttk.Entry(btn_frame, width=20)
        self.select_data_entry.pack(pady=5)
        ttk.Button(btn_frame, text="新增", command=self.add_select_data).pack(pady=5, fill='x')
        ttk.Button(btn_frame, text="編輯", command=self.edit_select_data).pack(pady=5, fill='x')
        ttk.Button(btn_frame, text="刪除", command=self.delete_select_data).pack(pady=5, fill='x')
        ttk.Button(btn_frame, text="清空", command=self.clear_select_data).pack(pady=5, fill='x')

        self.select_data_listbox.bind('<Double-1>', lambda e: self.edit_select_data())

        self.refresh_select_data_listbox()

    def refresh_select_data_listbox(self):
        self.select_data_listbox.delete(0, tk.END)
        for item in self.config.get('selectDataList', []):
            self.select_data_listbox.insert(tk.END, item)

    def get_selected_select_data_index(self):
        selection = self.select_data_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "請先選取一個品號")
            return None
        return selection[0]

    def add_select_data(self):
        value = self.select_data_entry.get().strip()
        if not value:
            messagebox.showwarning("提示", "請輸入品號")
            return
        self.config.setdefault('selectDataList', []).append(value)
        self.select_data_entry.delete(0, tk.END)
        self.refresh_select_data_listbox()

    def edit_select_data(self):
        idx = self.get_selected_select_data_index()
        if idx is None:
            return
        current_value = self.config['selectDataList'][idx]

        dialog = tk.Toplevel(self.select_data_tab)
        dialog.title("編輯品號")
        dialog.grab_set()
        dialog.resizable(False, False)

        ttk.Label(dialog, text="品號:").grid(row=0, column=0, sticky='w', padx=10, pady=10)
        var = tk.StringVar(value=current_value)
        entry = ttk.Entry(dialog, textvariable=var, width=30)
        entry.grid(row=0, column=1, sticky='w', padx=10, pady=10)

        def on_save():
            new_value = var.get().strip()
            if not new_value:
                messagebox.showwarning("提示", "品號不能為空白", parent=dialog)
                return
            self.config['selectDataList'][idx] = new_value
            self.refresh_select_data_listbox()
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="儲存", command=on_save).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side='left', padx=5)

    def delete_select_data(self):
        idx = self.get_selected_select_data_index()
        if idx is None:
            return
        del self.config['selectDataList'][idx]
        self.refresh_select_data_listbox()

    def clear_select_data(self):
        if messagebox.askyesno("確認", "確定要清空所有品號嗎？"):
            self.config['selectDataList'] = []
            self.refresh_select_data_listbox()

    # ---------- selectResourceList ----------
    def create_select_resource_tab(self):
        toolbar = ttk.Frame(self.select_resource_tab)
        toolbar.pack(fill='x', padx=5, pady=5)
        ttk.Button(toolbar, text="新增", command=self.add_select_resource).pack(side='left', padx=3)
        ttk.Button(toolbar, text="編輯", command=self.edit_select_resource).pack(side='left', padx=3)
        ttk.Button(toolbar, text="刪除", command=self.delete_select_resource).pack(side='left', padx=3)

        columns = ('mainChangerHeader', 'sheetName', 'selectHeader', 'headerRow', 'targetHeader', 'additionalProcessing')
        # 先建立 Treeview，再建立捲軸（避免 command 引用尚未建立的 widget）
        self.select_resource_tree = ttk.Treeview(self.select_resource_tab, columns=columns, show='headings', height=15)

        # command=self.select_resource_tree.yview 讓滾動條能點擊拖動調整內容
        select_resource_scrollbar = ttk.Scrollbar(self.select_resource_tab, orient='vertical', command=self.select_resource_tree.yview)
        select_resource_scrollbar.pack(side='right', fill='y')

        headers = {
            'mainChangerHeader': '主表更新欄位名稱',
            'sheetName': '資料來源工作表名稱',
            'selectHeader': '資料來源工作表條件欄位名稱',
            'headerRow': '資料來源工作表標題列',
            'targetHeader': '資料來源欄位名稱',
            'additionalProcessing': '額外處理',
        }
        widths = {
            'mainChangerHeader': 130,
            'sheetName': 220,
            'selectHeader': 120,
            'headerRow': 60,
            'targetHeader': 120,
            'additionalProcessing': 100,
        }
        for col in columns:
            self.select_resource_tree.heading(col, text=headers[col])
            self.select_resource_tree.column(col, width=widths[col], anchor='center')

        self.select_resource_tree.pack(fill='both', expand=True, padx=5, pady=5)
        self.select_resource_tree.bind('<Double-1>', lambda e: self.edit_select_resource())

        self.select_resource_tree.configure(yscrollcommand=select_resource_scrollbar.set)

        self.refresh_select_resource_tree()

    def refresh_select_resource_tree(self):
        self.select_resource_tree.delete(*self.select_resource_tree.get_children())
        for item in self.config.get('selectResourceList', []):
            self.select_resource_tree.insert('', 'end', values=(
                item.get('mainChangerHeader', ''),
                item.get('sheetName', ''),
                item.get('selectHeader', ''),
                item.get('headerRow', ''),
                item.get('targetHeader', ''),
                item.get('additionalProcessing', ''),
            ))

    def get_selected_select_resource_index(self):
        selection = self.select_resource_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "請先選取一個項目")
            return None
        return self.select_resource_tree.index(selection[0])

    def add_select_resource(self):
        self.open_resource_dialog(self.select_resource_tab, None, is_data_resource=False)

    def edit_select_resource(self):
        idx = self.get_selected_select_resource_index()
        if idx is not None:
            self.open_resource_dialog(self.select_resource_tab, idx, is_data_resource=False)

    def delete_select_resource(self):
        idx = self.get_selected_select_resource_index()
        if idx is None:
            return
        if messagebox.askyesno("確認", "確定要刪除這個指定資源嗎？"):
            del self.config['selectResourceList'][idx]
            self.refresh_select_resource_tree()

    # ---------- 操作按鈕 ----------
    def create_action_buttons(self):
        btn_frame = tk.Frame(self.content_frame)
        btn_frame.pack(fill='x', padx=10, pady=10)

        # 儲存 / 重新載入按鈕
        tk.Button(
            btn_frame, text="儲存設定", command=self.save_config,
            bg="#007BFF", fg="white", font=("Microsoft JhengHei", 11, "bold"),
            activebackground="#0056b3", activeforeground="white",
            relief="raised", bd=2, padx=12, pady=5, cursor="hand2"
        ).pack(side='left', padx=5)
        tk.Button(
            btn_frame, text="重新載入", command=self.reload_config,
            bg="#6C757D", fg="white", font=("Microsoft JhengHei", 11, "bold"),
            activebackground="#5a6268", activeforeground="white",
            relief="raised", bd=2, padx=12, pady=5, cursor="hand2"
        ).pack(side='left', padx=5)

        # 分隔線
        ttk.Separator(btn_frame, orient='vertical').pack(side='left', fill='y', padx=10)

        # 執行按鈕（不同顏色區分）
        run_buttons = [
            ("執行全部", "#28A745", 0),
            ("執行指定主表欄位", "#FD7E14", 1),
            ("執行指定工作表來源", "#6F42C1", 2),
            ("執行指定欄位與工作表", "#DC3545", 3),
        ]
        for text, color, test_type in run_buttons:
            tk.Button(
                btn_frame, text=text, command=lambda t=test_type: self.run_test(t),
                bg=color, fg="white", font=("Microsoft JhengHei", 11, "bold"),
                activebackground=color, activeforeground="white",
                relief="raised", bd=2, padx=12, pady=5, cursor="hand2"
            ).pack(side='left', padx=5)

    # ---------- Log 顯示區域 ----------
    def create_log_section(self):
        log_frame = ttk.LabelFrame(self.content_frame, text="執行 Log", padding=5)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # 工具列（清空 Log）
        log_toolbar = ttk.Frame(log_frame)
        log_toolbar.pack(fill='x', padx=5, pady=3)
        ttk.Button(log_toolbar, text="清空 Log", command=self.clear_log).pack(side='left')

        # Log 顯示區域（Text + 捲軸）— 加大顯示高度
        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # 先建立 Text，再建立捲軸（避免 command 引用尚未建立的 widget）
        self.log_text = tk.Text(text_frame, height=20, state='disabled', wrap='word',
                                bg="#1E1E1E", fg="#00FF00",
                                font=("Consolas", 10))

        # command=self.log_text.yview 讓滾動條能點擊拖動調整內容
        log_scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.log_text.yview)
        log_scrollbar.pack(side='right', fill='y')
        self.log_text.pack(side='left', fill='both', expand=True)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

    def _is_at_bottom(self):
        """檢查 log 顯示區域是否捲動到底部（用於判斷是否要自動捲動）"""
        try:
            # yview() 回傳 (top, bottom)，bottom 接近 1.0 表示在底部
            return self.log_text.yview()[1] >= 0.99
        except Exception:
            return True

    def append_log(self, message):
        """將 log 訊息加入顯示區域"""
        at_bottom = self._is_at_bottom()
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        # 只有使用者在底部時才自動捲動，避免干擾使用者查看歷史 log
        if at_bottom:
            self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')

    def append_log_batch(self, messages):
        """批次將多條 log 訊息加入顯示區域（減少 UI 更新次數，避免阻塞主執行緒）"""
        if not messages:
            return
        at_bottom = self._is_at_bottom()
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, "\n".join(messages) + "\n")
        # 只有使用者在底部時才自動捲動，避免干擾使用者查看歷史 log
        if at_bottom:
            self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')

    def process_log_queue(self):
        """從佇列讀取 log / 錯誤 / 完成訊息（在主執行緒執行，確保 UI 執行緒安全）

        改進重點：
        1. 每次最多處理 MAX_PER_TICK 條訊息，避免一次處理大量訊息阻塞主執行緒
        2. 批次收集 log 後一次插入，減少 Text widget 的更新次數
        3. messagebox 使用 after(0, ...) 延遲顯示，避免阻塞主執行緒
        """
        MAX_PER_TICK = 100  # 每次最多處理的訊息數量，避免阻塞 UI
        processed = 0
        batch_messages = []  # 批次收集 log 訊息

        try:
            while processed < MAX_PER_TICK:
                msg_type, payload = LOG_QUEUE.get_nowait()
                processed += 1

                if msg_type == 'log':
                    batch_messages.append(payload)
                elif msg_type == 'error':
                    # 先處理已收集的 log，再處理錯誤
                    if batch_messages:
                        self.append_log_batch(batch_messages)
                        batch_messages = []
                    self.append_log(f"錯誤：{payload}")
                    # 使用 after 延遲顯示 messagebox，避免阻塞主執行緒
                    self.root.after(0, lambda p=payload: messagebox.showerror("錯誤", p))
                elif msg_type == 'warning':
                    # 工作表名稱等提示：顯示提示視窗（不寫入 log），關閉後解除背景執行緒等待
                    if batch_messages:
                        self.append_log_batch(batch_messages)
                        batch_messages = []
                    warning_msg, wait_event = payload
                    self.root.after(0, lambda m=warning_msg, ev=wait_event: self.show_warning_dialog(m, ev))
                elif msg_type == 'done':
                    if batch_messages:
                        self.append_log_batch(batch_messages)
                        batch_messages = []
                    self.append_log("處理完成")
                    self.set_running(False)
                    # 使用 after 延遲顯示 messagebox，避免阻塞主執行緒
                    self.root.after(0, lambda: messagebox.showinfo("完成", "ERP Excel 更新完成！"))
        except queue.Empty:
            pass

        # 批次寫入剩餘的 log
        if batch_messages:
            self.append_log_batch(batch_messages)

        # 每 100ms 檢查一次佇列
        self.root.after(100, self.process_log_queue)

    def show_warning_dialog(self, message, wait_event):
        """顯示提示視窗；使用者關閉後解除背景執行緒的等待，使其繼續執行"""
        messagebox.showwarning("提示", message)
        wait_event.set()

    def set_running(self, running):
        """記錄執行狀態（避免重複執行）"""
        self.running = running

    def clear_log(self):
        """清空 Log 顯示區域"""
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')

    def save_config(self):
        try:
            # 主工作表設定
            self.config['mainChangeSheet'] = {
                'sheetName': self.main_sheet_name_var.get().strip(),
                'selectHeader': self.main_select_header_var.get().strip(),
                'headerRow': int(self.main_header_row_var.get().strip()),
            }
            # selectDataList 從 Listbox 收集
            self.config['selectDataList'] = [
                self.select_data_listbox.get(i) for i in range(self.select_data_listbox.size())
            ]

            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "設定已儲存至 config.json")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗：\n{e}")

    def reload_config(self):
        if messagebox.askyesno("確認", "重新載入會放棄未儲存的修改，確定嗎？"):
            self.config = self.load_config()
            main_cfg = self.config.get('mainChangeSheet', {})
            self.main_sheet_name_var.set(main_cfg.get('sheetName', ''))
            self.main_select_header_var.set(main_cfg.get('selectHeader', ''))
            self.main_header_row_var.set(str(main_cfg.get('headerRow', 1)))
            self.refresh_data_resource_tree()
            self.refresh_select_data_listbox()
            self.refresh_select_resource_tree()

    def run_test(self, test_type):
        # 先儲存設定，確保執行時讀取到最新設定
        self.save_config()

        # 在主執行緒選取檔案（背景執行緒不能操作 tkinter）
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title="請選擇 Excel 檔案",
            filetypes=[
                ("Excel 檔案", "*.xlsx *.xlsm *.xltx *.xltm"),
                ("舊版 Excel", "*.xls"),
            ]
        )

        if not file_path:
            self.append_log("未選取任何檔案")
            return

        self.append_log(f"已選取檔案：{file_path}")
        self.append_log("開始處理...")
        self.set_running(True)

        # 在背景執行緒執行處理，避免阻塞 UI
        thread = threading.Thread(target=self._worker, args=(file_path, test_type), daemon=True)
        thread.start()

    def _worker(self, file_path, test_type):
        """背景執行緒：執行 ERP 處理與例外處理"""
        try:
            run_erp_process(file_path, test_type)
        except Exception as e:
            log(f"執行失敗：{e}")
            traceback.print_exc()
            LOG_QUEUE.put(("error", f"執行失敗：{str(e)}"))
            LOG_QUEUE.put(("done", None))


def main():
    root = tk.Tk()
    app = ConfigEditorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()