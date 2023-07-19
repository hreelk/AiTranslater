import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.by import By

def open_webpage():
    # 取得URL
    url = entry.get()
    
    # 創建瀏覽器驅動程式
    driver = webdriver.Chrome()
    
    # 打開網頁
    driver.get(url)
    # 關閉視窗
    window.destroy()
    # 使用Selenium尋找特定元素
    element = driver.find_element(By.ID, "L2")  # 使用適當的定位方式和元素ID
    selected_text = element.text
    print(selected_text)  # 印出選取的文字
    
    # 關閉瀏覽器
    driver.quit()

# 創建視窗
window = tk.Tk()

# 建立標籤和輸入框
label = tk.Label(window, text="請輸入URL：")
label.pack()
entry = tk.Entry(window)
entry.pack()

# 建立按鈕
button = tk.Button(window, text="打開網頁並選取文字", command=open_webpage)
button.pack()

# 啟動主迴圈
window.mainloop()
