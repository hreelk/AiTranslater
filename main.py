import tkinter as tk
from selenium import webdriver

def open_webpage():
    # 取得URL
    url = entry.get()
    
    # 創建瀏覽器驅動程式
    driver = webdriver.Chrome()
    
    # 打開網頁
    driver.get(url)
    
    # 啟動主迴圈
    window.mainloop()

# 創建視窗
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
