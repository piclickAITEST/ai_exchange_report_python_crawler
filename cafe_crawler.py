import time
import ctypes
import tkinter as tk
from tkinter import messagebox
from os import system

from selenium import webdriver

# URLs
CAFE24_ADMIN_LOGIN_URL = 'https://eclogin.cafe24.com/Shop/?url=Init&login_mode=2&is_multi=F'

# 카페24 유저 정보
CAFE24_MALL_ID       = 'benitomaster'
CAFE24_USER_ID       = 'guest'
CAFE24_USER_PASSWORD = 'qpslxh88!!'

def display_message_box(title, body):
    root = tk.Tk()
    root.geometry("400x300")
    window_width  = root.winfo_reqwidth()
    window_height = root.winfo_reqheight()
    position_right = int(root.winfo_screenwidth()/2 - window_width/2)
    position_down = int(root.winfo_screenheight()/2 - window_height/2)
    root.geometry("+{}+{}".format(position_right, position_down))
    root.title(title)
    label = tk.Label(root, text=body)
    label.pack(side="top", fill="both", expand=True, padx=20, pady=20)
    button = tk.Button(root, text="OK", command=lambda: root.destroy())
    button.pack(side="bottom", fill="none", expand=True)
    root.mainloop()
    
def login_cafe24_admin(driver, mall_id, user_id, user_password):
    driver.find_element_by_id('mall_id').send_keys(mail_id)
    driver.find_element_by_id('userid').send_keys(user_id)
    driver.find_element_by_id('userpasswd').send_keys(user_password)
    
    driver.find_element_by_xpath('//*[@id="tabStaff"]/div/fieldset/p[1]/a').click()
    driver.sleep(3)

# 로그인 페이지 접속
driver1 = webdriver.Chrome()
driver1.get(CAFE24_ADMIN_LOGIN_URL)

display_message_box('필독! 반드시 읽어주세요!', 'reCAPTCHA가 인증을 반드시 완료한 후 OK 버튼을 눌러주세요!')
login_cafe24_admin(driver1, CAFE24_MALL_ID, CAFE24_USER_ID, CAFE24_USER_PASSWORD)

driver1.get()

# 드라이버 종료
driver1.quit()