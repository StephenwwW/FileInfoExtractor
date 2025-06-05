import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
import threading
from moviepy.editor import VideoFileClip
import humanize

class FileInfoExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Info Extractor")
        self.root.geometry("432x150")  # 最佳GUI大小，適合顯示所有元件

        self.folder_path = tk.StringVar()  # 儲存選擇的資料夾路徑
        self.file_info_list = []  # 儲存掃描結果

        # ====== 資料夾選擇區塊 ======
        label = tk.Label(root, text="Select Root Folder:", anchor='center')
        label.pack(fill='x', pady=4)  # 置中顯示
        frame = tk.Frame(root)
        frame.pack(pady=2)
        tk.Entry(frame, textvariable=self.folder_path, width=45).pack(side=tk.LEFT, padx=2)
        tk.Button(frame, text="Browse", command=self.browse_folder, width=8).pack(side=tk.LEFT)

        # ====== 進度條區塊 ======
        self.progress = ttk.Progressbar(root, orient='horizontal', length=350, mode='determinate')
        self.progress.pack(pady=2)
        self.progress_label = tk.Label(root, text="", anchor='center')
        self.progress_label.pack(fill='x')

        # ====== 操作按鈕區塊 ======
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=8)
        btn_width = 16
        btn_height = 1
        tk.Button(btn_frame, text="Scan", width=btn_width, height=btn_height, command=self.scan_files_thread).pack(side=tk.LEFT, padx=12)
        tk.Button(btn_frame, text="Export to Excel", width=btn_width, height=btn_height, command=self.export_excel).pack(side=tk.LEFT, padx=12)

    def browse_folder(self):
        """彈出資料夾選擇視窗，選擇後將路徑寫入輸入框"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)

    def format_duration(self, seconds):
        """將秒數轉為 hh:mm:ss 格式字串"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def format_size(self, size_bytes):
        """將位元組轉為GB或MB字串，保留兩位小數"""
        if size_bytes >= 1024**3:
            return f"{size_bytes/1024**3:.2f} GB"
        else:
            return f"{size_bytes/1024**2:.2f} MB"

    def scan_files_thread(self):
        """啟動新執行緒進行檔案掃描，避免GUI卡死"""
        t = threading.Thread(target=self.scan_files)
        t.daemon = True
        t.start()

    def scan_files(self):
        """掃描所選資料夾下所有影片檔案，取得檔案資訊並更新進度條"""
        root_folder = self.folder_path.get()
        if not root_folder:
            messagebox.showwarning("Warning", "Please select a folder first!")
            return

        self.file_info_list = []
        video_extensions = ('.mp4', '.ts', '.mkv', '.avi', '.mov')  # 支援的影片副檔名

        # 先統計所有要處理的檔案數量
        total_files = 0
        for root, dirs, files in os.walk(root_folder):
            for file in files:
                if file.lower().endswith(video_extensions):
                    total_files += 1
        if total_files == 0:
            self.progress['value'] = 0
            self.progress_label.config(text="無影片檔案")
            return
        self.progress['maximum'] = total_files
        self.progress['value'] = 0
        self.progress_label.config(text="0%")
        self.root.update_idletasks()

        processed = 0
        try:
            for root, dirs, files in os.walk(root_folder):
                for file in files:
                    if file.lower().endswith(video_extensions):
                        file_path = os.path.join(root, file)
                        try:
                            file_size_bytes = os.path.getsize(file_path)  # 取得檔案大小
                            with VideoFileClip(file_path) as clip:
                                duration_seconds = clip.duration  # 取得影片長度
                            # 計算平均每小時檔案大小
                            if duration_seconds >= 3600:
                                avg_size_per_hour = file_size_bytes / (duration_seconds / 3600)
                                if avg_size_per_hour >= 1024**3:
                                    avg_size_str = f"{avg_size_per_hour/1024**3:.2f} GB/hr"
                                else:
                                    avg_size_str = f"{avg_size_per_hour/1024**2:.2f} MB/hr"
                            else:
                                avg_size_str = '無法計算'
                            file_info = {
                                'File Name': file,  # 只顯示檔案名稱
                                'Video Duration': self.format_duration(duration_seconds),
                                'File Size': self.format_size(file_size_bytes),
                                'File Size in Bytes': file_size_bytes,
                                'Average File Size Per Hour': avg_size_str
                            }
                            self.file_info_list.append(file_info)
                        except Exception as e:
                            print(f"Error processing file {file_path}: {str(e)}")
                            continue
                        # 更新進度條
                        processed += 1
                        self.progress['value'] = processed
                        percent = int(processed / total_files * 100)
                        self.progress_label.config(text=f"{percent}%")
                        self.root.update_idletasks()
            self.progress_label.config(text="掃描完成！")
            messagebox.showinfo("Info", f"Scan complete! {len(self.file_info_list)} files found.")
        except Exception as e:
            self.progress_label.config(text="掃描失敗")
            messagebox.showerror("Error", f"An error occurred during scanning: {str(e)}")

    def export_excel(self):
        """將掃描結果匯出為Excel檔，並自動調整欄寬與對齊"""
        if not self.file_info_list:
            messagebox.showwarning("Warning", "No data to export. Please scan first.")
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save as Excel file"
        )
        if not save_path:
            return
        df = pd.DataFrame(self.file_info_list)
        df.to_excel(save_path, index=False)

        # 用 openpyxl 調整Excel格式：A欄只根據標題寬度，其他欄根據標題自動調整，所有儲存格靠左對齊
        wb = load_workbook(save_path)
        ws = wb.active
        a1_length = len(str(ws['A1'].value))
        ws.column_dimensions['A'].width = a1_length + 2
        for col in ws.iter_cols(min_col=2, max_col=ws.max_column):
            col_letter = col[0].column_letter
            max_length = len(str(col[0].value))
            ws.column_dimensions[col_letter].width = max_length + 2
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='left', vertical='center')
        wb.save(save_path)
        messagebox.showinfo("Success", f"Exported to:\n{save_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileInfoExtractorApp(root)
    root.mainloop() 