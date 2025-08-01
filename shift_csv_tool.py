import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import re
from datetime import datetime

def parse_and_export(raw_data, year):
    date_pattern = re.compile(r'(\d{1,2})/(\d{1,2})\([A-Za-z]+\)\[追加\]')
    time_pattern = re.compile(r'(\d{1,2}:\d{2})-(\d{1,2}:\d{2})')

    events = []
    current_date = None

    for line in raw_data.splitlines():
        line = line.strip()
        if not line:
            continue

        date_match = date_pattern.match(line)
        if date_match:
            month, day = map(int, date_match.groups())
            current_date = datetime(year, month, day)
            continue

        if '[休' in line or not current_date:
            continue

        time_match = time_pattern.search(line)
        if time_match:
            start_time, end_time = time_match.groups()
            start_dt = datetime.combine(current_date.date(), datetime.strptime(start_time, '%H:%M').time())
            end_dt = datetime.combine(current_date.date(), datetime.strptime(end_time, '%H:%M').time())

            events.append({
                'Subject': '大阪某テーマパーク勤務',
                'Start Date': start_dt.strftime('%Y/%m/%d'),
                'Start Time': start_dt.strftime('%H:%M'),
                'End Date': end_dt.strftime('%Y/%m/%d'),
                'End Time': end_dt.strftime('%H:%M'),
                'Description': 'テーマパーク勤務シフト'
            })

    if not events:
        return 0

    filepath = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")],
        title="保存するCSVファイルを指定"
    )
    if not filepath:
        return 0

    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'Description'])
        writer.writeheader()
        writer.writerows(events)

    return len(events)

def create_gui():
    def on_generate():
        try:
            year = int(entry_year.get())
            raw_data = text_input.get("1.0", tk.END)
            count = parse_and_export(raw_data, year)
            if count > 0:
                messagebox.showinfo("完了", f"{count} 件のイベントをCSVに出力しました。")
            else:
                messagebox.showwarning("警告", "有効なイベントが見つかりませんでした。")
        except ValueError:
            messagebox.showerror("エラー", "年は西暦で入力してください（例: 2025）")

    root = tk.Tk()
    root.title("テーマパーク シフトCSV作成ツール")

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    tk.Label(frame, text="年（西暦）:").grid(row=0, column=0, sticky="w")
    entry_year = tk.Entry(frame)
    entry_year.insert(0, "2025")
    entry_year.grid(row=0, column=1, sticky="w")

    tk.Label(frame, text="スケジュール入力（元データ）:").grid(row=1, column=0, columnspan=2, sticky="w")
    text_input = tk.Text(frame, height=25, width=80)
    text_input.grid(row=2, column=0, columnspan=2, pady=5)

    generate_btn = tk.Button(frame, text="CSV出力", command=on_generate)
    generate_btn.grid(row=3, column=0, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
