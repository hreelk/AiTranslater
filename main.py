import webbrowser
import tkinter as tk

def open_webpage():
    url = entry.get()
    webbrowser.open(url)

# 建立主視窗
window = tk.Tk()

# 建立標籤和輸入框
label = tk.Label(window, text="請輸入URL：")
label.pack()
entry = tk.Entry(window)
entry.pack()

# 建立按鈕
button = tk.Button(window, text="打開網頁", command=open_webpage)
button.pack()

# 啟動主迴圈
window.mainloop()