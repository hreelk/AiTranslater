import os
import subprocess
import tkinter as tk
from docx import Document
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import time

def open_webpage():

    #檔案儲存位置
    folder_path = "C:/Documents/novel-translate"
    os.makedirs(folder_path, exist_ok=True)
    # 取得URL
    url = entry.get()
    
    # 創建瀏覽器驅動程式
    driver = webdriver.Chrome()
    
    # 打開網頁
    driver.get(url)
    
    # 使用Selenium尋找特定元素
    selected_text = ""
    
    # 迴圈遍歷每個元素
    for i in range(1, 10000):
        # 使用元素的 ID 值來定位元素
        element_id = "L" + str(i)
        try:
            element = driver.find_element(By.ID, element_id)
            selected_text += element.text + "\n"
        except NoSuchElementException:
            break
    
    # 創建 Word 文件
    doc = Document()
    doc.add_paragraph(selected_text)
    
    # 取得使用者輸入的保存名稱
    save_name = entry2.get()
    # 構建完整的文件路徑
    file_name = save_name + ".docx"
    file_path = os.path.join(folder_path, file_name)
    # 儲存 Word 文件
    doc.save(file_path)
    # 關閉瀏覽器
    driver.quit()
    # 關閉視窗
    window.destroy()
    #以debug模式打開chrome
    bat_file_path = 'start_chrome_debug.bat'
    subprocess.call(bat_file_path, shell=True)
    # 創建 ChromeOptions
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9222")
    # 創建瀏覽器驅動程式
    driver = webdriver.Chrome(options=options)
    # 移動到網站
    driver.get("https://chat.forefront.ai")
    print(driver.title)
    #等待網站開好
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "text-th-primary-dark")))
    time.sleep(2)
    #再輸入框位置點左鍵
    pyautogui.click(x=2128, y=1972)
    #輸入
    hi = "hi"
    pyautogui.typewrite(hi)



# 創建視窗
window = tk.Tk()

# 建立標籤和輸入框
label = tk.Label(window, text="請輸入URL：")
label.pack()
entry = tk.Entry(window)
entry.pack()
label2 = tk.Label(window, text="請輸入保存名稱：")
label2.pack()
entry2 = tk.Entry(window)
entry2.pack()

# 建立按鈕
button = tk.Button(window, text="打開網頁並選取文字", command=open_webpage)
button.pack()

# 啟動主迴圈
window.mainloop()
