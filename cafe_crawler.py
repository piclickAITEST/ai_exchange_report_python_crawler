import time
import ctypes
import tkinter as tk
from tkinter import messagebox
from os import system

from selenium import webdriver

# URLs
CAFE24_ADMIN_LOGIN_URL = 'https://eclogin.cafe24.com/Shop/?url=Init&login_mode=2&is_multi=F'
"https://benitomaster.cafe24.com/admin/php/shop1/s_new/order_list_item.php?rows=20&searchSorting=order_desc&isBusanCall=&isChinaCall=&orderCallnum=&cticall=&realclick=T&tabclick=F&MSK[]=order_id&MSV[]=&orderStatusPayment=all&date_type=order_date&btnDate=7&product_search_type=product_name&find_option=product_no&order_product_name=&order_product_code=&order_product_no=&order_product_text=&order_set_product_no=&layer_order_product_code=&layer_order_product_opt_id=&popup_item_code=&popup_product_code=&payed=&payed_sql_version=&bank_info=&memberType=1&group_no=&isMemAuth=&isBlackList=&isFirstOrder=&isPointfyUsedMember=&shipment_type=all&bunch=&shippedAgain=&shipmentMessage=&delivSeperated=&isReservedOrder=&isSubscriptionOrder=&paystandard=choice&product_total_price1=&product_total_price2=&item_count_start=&item_count_end=&orderPathType=A&search_SaleOpenMarket[]=cafe24&search_SaleOpenMarket[]=mobile&search_SaleOpenMarket[]=mobile_d&search_SaleOpenMarket[]=NCHECKOUT&search_SaleOpenMarket[]=gmarket&search_SaleOpenMarket[]=auction&search_SaleOpenMarket[]=sk11st&search_SaleOpenMarket[]=shopn&search_SaleOpenMarket[]=inpark&search_SaleOpenMarket[]=coupang&search_SaleOpenMarket[]=kakao&search_SaleOpenMarket[]=womanstalk&search_SaleOpenMarket[]=tenten&search_SaleOpenMarket[]=wemake&search_SaleOpenMarket[]=melchi&search_SaleOpenMarket[]=halfclub&search_SaleOpenMarket[]=boribori&search_SaleOpenMarket[]=ogage&search_SaleOpenMarket[]=moongori&search_SaleOpenMarket[]=shopeesg&search_SaleOpenMarket[]=shopeeid&search_SaleOpenMarket[]=shopeemy&search_SaleOpenMarket[]=shopeetw&search_SaleOpenMarket[]=shopeeth&search_SaleOpenMarket[]=shopeeph&search_SaleOpenMarket[]=brich&search_SaleOpenMarket[]=zigzag&search_SaleOpenMarket[]=ably&search_SaleOpenMarket[]=timon&search_SaleOpenMarket[]=musinsa&search_SaleOpenMarket[]=wizwid&search_SaleOpenMarket[]=hottracks&search_SaleOpenMarket[]=akmall&search_SaleOpenMarket[]=daisomall&search_SaleOpenMarket[]=lfmall&search_SaleOpenMarket[]=styleshare&search_SaleOpenMarket[]=aland&search_SaleOpenMarket[]=rakutenkr&search_SaleOpenMarket[]=cjmall&search_SaleOpenMarket[]=lotteon&search_SaleOpenMarket[]=himart&search_SaleOpenMarket[]=tofkof&search_SaleOpenMarket[]=MORUGI&search_SaleOpenMarket[]=11st&mkSaleType=M&mkSaleTypeChg=&inflowPathType=A&inflowPathDetail=0000000000000000000000000000000000&paymethodType=A&paymentMethod[]=cash&paymentMethod[]=card&paymentMethod[]=tcash&paymentMethod[]=icash&paymentMethod[]=cell&paymentMethod[]=deferpay&paymentMethod[]=cvs&paymentMethod[]=point&paymentMethod[]=mileage&paymentMethod[]=deposit&paymentMethod[]=etc&pgListType=A&pgList[]=danal&pgList[]=dacom&pgList[]=payco&pgList[]=paynow&pgList[]=smilepay&pgList[]=eximbay&pgList[]=etc&paymentInfo=&discountMethod=&shop_no_order=1&delvReady=&delvCancel=&orderStatusNotPayCancel=N&orderStatusCancel=N&orderSearchCancelStatus=&orderStatusExchange=N&orderSearchExchangeStatus=&orderStatusReturn=N&orderStatusRefund=N&orderSearchRefundStatus=&orderSearchShipStatus=&orderStatus[]=all&orderStatus[]=N10&orderStatus[]=N20&orderStatus[]=N22&orderStatus[]=N21&orderStatus[]=N30&orderStatus[]=N40&RefundType=&RefundSubType=&sc_id=&second_shipping_company_id=&HopeShipCompanyId=all&post_express_flag=&tabStatus=&paymethod_total_count=&search_invoice_print_flag=all&search_is_escrow_shipping_registered=all&search_print_second_invoice=all&incoming=&is_purchased=&order_fail_code=&isBlackOrder=&start_date=" refdate "&year1=" year "&month1=" monthstr "&day1=" daystr "&start_time=00:00:00&end_date=" refdate "&year2=" year "&month2=" monthstr "&day2=" daystr "&end_time=23:59:59&realclick=T"

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