import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import os
import webbrowser

# 預設執行者文字（當 J 欄為空時填入）
DEFAULT_EXECUTOR = "未指派"

# 處理 I 欄文字，依據特定分隔符（-, by, ok）拆分成兩部分
def split_i_column(val):
    if pd.isna(val):
        return pd.NA, pd.NA
    val = str(val).strip()

    # 搜尋結尾分隔符，分成前段與後段
    match = re.search(r'(.+?)\s*(?:[-]|by|ok)\s*(.{0,20})$', val, re.IGNORECASE)
    if match:
        part1 = match.group(1).strip()
        part2 = match.group(2).strip()
        return part1, part2
    else:
        return val, pd.NA  # 若無法拆分則保留原文並回傳空值

# 主要資料處理邏輯
def process_excel(input_path, output_path):
    try:
        # 讀取 Excel，從第三列開始當作欄位標題
        df = pd.read_excel(input_path, header=2, engine='openpyxl')

        # 拆分 I 欄為「處理記錄」與「執行者」
        df[['處理記錄', '執行者']] = df.iloc[:, 8].apply(lambda x: pd.Series(split_i_column(x)))

        # 若執行者為空則填入預設內容
        df['執行者'] = df['執行者'].fillna(DEFAULT_EXECUTOR)

        # 合併 H 欄（index 7）與新的 I 欄（處理記錄），並處理「。。」為「。」
        combined = df.iloc[:, 7].fillna('').astype(str) + "。" + "\n" + "\n" + df['處理記錄'].fillna('').astype(str) + "。"
        df['處理記錄'] = combined.str.replace('。。', '。', regex=False)

        # 刪除原始 A(0), C(2), D(3), F(5) , h(7) 五欄
        drop_indices = sorted([0, 2, 3, 5, 7], reverse=True)  # 反向刪除避免欄位位移錯誤
        for idx in drop_indices:
            if idx < len(df.columns):
                df.drop(df.columns[idx], axis=1, inplace=True)
        
        df = df[['叫修日期', '地點', '問題類別', '處理記錄', '執行者']]

        # 將結果寫出到 Excel 檔案
        df.to_excel(output_path, index=False, engine='openpyxl')
        return True

    except Exception as e:
        # 若有錯誤則顯示錯誤視窗
        messagebox.showerror("錯誤", f"處理失敗：{str(e)}")
        return False


# 建立 GUI 畫面
def run_gui():
    root = tk.Tk()
    root.title("總處報修資料欄位轉換工具 v1.0")
    root.geometry("500x150")
    root.resizable(False, False)

    # 設定來源與輸出路徑的變數
    input_path = tk.StringVar()
    output_path = tk.StringVar()

    # 選擇來源檔案的函式
    def select_input():
        path = filedialog.askopenfilename(filetypes=[("維修資料Excel檔", "*.xlsx")])
        if path:
            input_path.set(path)

    # 選擇輸出位置的函式
    def select_output():
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("匯出結果Excel檔", "*.xlsx")])
        if path:
            output_path.set(path)

    # 按下「執行處理」時的邏輯
    def execute():
        if not input_path.get() or not output_path.get():
            messagebox.showwarning("注意", "請選擇來源與輸出檔案")
            return
        success = process_excel(input_path.get(), output_path.get())
        if success:
            show_success_popup()

    def show_success_popup():
        popup = tk.Toplevel(root)
        popup.title("合併完成")
        popup.geometry("300x150")
        popup.resizable(False, False)

        tk.Label(popup, text="✅ 合併成功！", font=("Arial", 14)).pack(pady=10)
        tk.Label(popup, text="檔案已成功儲存。").pack()

        # 加入超連結
        def open_link(event):
            webbrowser.open_new("https://github.com/adsa562/DGBAS_ComputerMaintenanceReport-Export")

        link_label = tk.Label(popup, text="🔗Github項目連結", fg="blue", cursor="hand2")
        link_label.pack(pady=5)
        link_label.bind("<Button-1>", open_link)

        # 關閉按鈕
        tk.Button(popup, text="關閉", command=popup.destroy).pack(pady=10)


    # 建立主要區塊
    frame = tk.Frame(root)
    frame.pack(padx=20, pady=20)

    # 輸入檔案行（Label + Entry + Button）
    
    tk.Label(frame, text="來源檔案：").grid(row=0, column=0, sticky="w")
    tk.Entry(frame, textvariable=input_path, width=40).grid(row=0, column=1, padx=5)
    tk.Button(frame, text="選擇", command=select_input).grid(row=0, column=2)

    # 輸出檔案行
    tk.Label(frame, text="匯出至：").grid(row=1, column=0, sticky="w", pady=(10, 0))
    tk.Entry(frame, textvariable=output_path, width=40).grid(row=1, column=1, padx=5, pady=(10, 0))
    tk.Button(frame, text="選擇", command=select_output).grid(row=1, column=2, pady=(10, 0))

    # 執行按鈕
    tk.Button(root, text="轉換！！", command=execute, bg="green", fg="white", width=20).pack(pady=10)

    root.mainloop()

# 主程式進入點
if __name__ == "__main__":
    run_gui()
