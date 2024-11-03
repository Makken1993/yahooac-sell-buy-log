import re
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import yahoo_ac_scraper

class YahooAuctionScraperGUI:
    def __init__(self, master):
        self.master = master
        master.title("Yahoo!オークションスクレイパー")
        master.geometry("600x700")  # ウィンドウサイズを少し大きくしました

        self.is_scraping = False
        self.create_widgets()

    def create_widgets(self):
        # スプレッドシートURL入力
        ttk.Label(self.master, text="スプレッドシートURL:").pack(pady=(20, 0))
        self.spreadsheet_url = ttk.Entry(self.master, width=50)
        self.spreadsheet_url.pack(pady=(0, 20))

        # 使用方法の説明
        ttk.Label(self.master, text="使用方法:").pack(pady=(20, 0))
        usage_text = ("1. テンプレートをコピー: [テンプレートURL]\n"
                      "2. コピーしたスプレッドシートを開く\n"
                      "3. A3以降にYahoo!オークションURLを入力\n"
                      "4. スプレッドシートURLを上の欄に貼り付け\n"
                      "5. 開始・終了行を設定し、スクレイピング開始")
        usage_label = ttk.Label(self.master, text=usage_text, justify="left", wraplength=550)
        usage_label.pack(pady=(0, 20))

        # ラジオボタンの追加
        self.history_type = tk.StringVar(value="purchase")
        self.history_frame = ttk.LabelFrame(self.master, text="履歴タイプ")
        self.history_frame.pack(pady=10, padx=10, fill="x")
        ttk.Radiobutton(self.history_frame, text="購入履歴", variable=self.history_type, value="purchase").pack(
            side="left", padx=5)
        ttk.Radiobutton(self.history_frame, text="売却履歴", variable=self.history_type, value="sale").pack(side="left",
                                                                                                            padx=5)

        # 行選択フレーム
        self.row_selection_frame = ttk.LabelFrame(self.master, text="処理する行の選択")
        self.row_selection_frame.pack(pady=10, padx=10, fill="x")

        ttk.Label(self.row_selection_frame, text="開始行:").grid(row=0, column=0, padx=5, pady=5)
        self.start_row = ttk.Entry(self.row_selection_frame, width=10)
        self.start_row.grid(row=0, column=1, padx=5, pady=5)
        self.start_row.insert(0, "3")  # デフォルト値として3を設定（1行目はヘッダー、2行目はタグ）

        ttk.Label(self.row_selection_frame, text="終了行:").grid(row=0, column=2, padx=5, pady=5)
        self.end_row = ttk.Entry(self.row_selection_frame, width=10)
        self.end_row.grid(row=0, column=3, padx=5, pady=5)

        # スクレイピング開始ボタン
        self.start_button = ttk.Button(self.master, text="スクレイピング開始", command=self.start_scraping)
        self.start_button.pack(pady=10)

        # 処理中断ボタン（初期状態は無効）
        self.stop_button = ttk.Button(self.master, text="処理を中断する", command=self.stop_scraping, state='disabled')
        self.stop_button.pack(pady=10)

        # プログレスバー
        self.progress = ttk.Progressbar(self.master, length=400, mode='indeterminate')
        self.progress.pack(pady=20)

        # ステータス表示
        self.status_label = ttk.Label(self.master, text="")
        self.status_label.pack(pady=10)

        # 結果表示エリア
        self.result_text = tk.Text(self.master, height=10, width=60)
        self.result_text.pack(pady=20)

    def start_scraping(self):
        url = self.spreadsheet_url.get().strip()
        start_row = self.start_row.get().strip()
        end_row = self.end_row.get().strip()

        # URLのバリデーション
        if not url:
            self.show_error("エラー", "スプレッドシートURLを入力してください。")
            return
        if not re.match(r'^https://docs\.google\.com/spreadsheets/d/[a-zA-Z0-9-_]+', url):
            self.show_error("エラー", "無効なスプレッドシートURLです。")
            return

        # 行番号のバリデーション
        try:
            start_row = int(self.zen_to_han(start_row))
            if start_row < 3:
                raise ValueError("開始行は3以上である必要があります。")

            if end_row:
                end_row = int(self.zen_to_han(end_row))
                if end_row < start_row:
                    raise ValueError("終了行は開始行以上である必要があります。")
        except ValueError as e:
            self.show_error("エラー", f"無効な行番号です: {str(e)}")
            return

        # スクレイピング処理の開始
        self.is_scraping = True
        self.start_button['state'] = 'disabled'
        self.stop_button['state'] = 'normal'
        self.progress.start()
        self.status_label.config(text="スクレイピングを開始しています...")

        # スクレイピングを別スレッドで実行
        threading.Thread(target=self.run_scraping, args=(url, start_row, end_row, self.history_type.get()), daemon=True).start()

    def stop_scraping(self):
        self.is_scraping = False
        self.status_label.config(text="処理を中断しています...")
        self.stop_button['state'] = 'disabled'

    def show_error(self, title, message):
        messagebox.showerror(title, message)

    def zen_to_han(self, text):
        return text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))

    def run_scraping(self, url, start_row, end_row, history_type):
        try:
            sheet_name = "ヤフオク購入履歴" if history_type == "purchase" else "ヤフオク売却履歴"
            new_count, skipped_count = yahoo_ac_scraper.main(url, start_row, end_row, sheet_name, lambda: self.is_scraping)

            if self.is_scraping:
                self.master.after(0, self.update_result,
                                  f"新たにスクレイピングしたURL数: {new_count}\nスキップしたURL数: {skipped_count}\n\nスクレイピングが完了しました。")
            else:
                self.master.after(0, self.update_result, "スクレイピングが中断されました。")
        except Exception as e:
            self.master.after(0, self.update_result, f"エラーが発生しました: {str(e)}")
        finally:
            self.master.after(0, self.finish_scraping)

    def update_result(self, message):
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, message)

    def finish_scraping(self):
        self.is_scraping = False
        self.progress.stop()
        self.start_button['state'] = 'normal'
        self.stop_button['state'] = 'disabled'
        self.status_label.config(text="処理が完了しました")

if __name__ == "__main__":
    root = tk.Tk()
    gui = YahooAuctionScraperGUI(root)
    root.mainloop()