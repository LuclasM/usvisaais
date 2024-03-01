'''
v0.3
重新预约的时候，如果出错，会自动截图，然后等待5分钟后继续。本次修改改进了，截图时会自动检测文件是否已经存在
如果已经存在则会自动加上数字后缀，例如screenshot1.jpg，screenshot2.jpg等等
v0.5
updated mutlple customer and mutiple prefered office
start a new thread for each updateing
if already send cutomer for a selected office, then skip the office for other customers
skip "no available time" test, seems problem with the website
'''



import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta
import time
import os

import json
import threading



login_url = "https://ais.usvisa-info.com/en-ca/niv/users/sign_in"

#output to console and log file
def duooutput(text):
    print(text)
    
    if not os.path.exists('log.txt'):
        open('log.txt', 'w').close()
        
    with open('log.txt', 'a') as f:
        f.write(text+"\n")




class Reschedule:
    
    available_dates = {}

    app_data = {}

    Canada_offices = ["Calgary", "Ottawa", "Toronto", "Vancouver", "Halifax", "Montreal", "Quebec City"]

    def __init__(self):
        with open('data.txt', 'r') as f:
            content = f.read()
            self.app_data = json.loads(content)
            
            # set is_ready to True for each customer
            for customer in self.app_data['CUSTOMERS'].values():
                customer['IS_READY'] = True

    def open_probe(self, index: int):
        driverX = webdriver.Firefox()
        driverX.get(login_url)
        driverX.find_element(By.ID, "user_email").send_keys(self.app_data.get(f'USERNAME_P{index}'))
        driverX.find_element(By.ID, "user_password").send_keys(self.app_data.get(f'PASSWORD_P{index}'))
        driverX.find_element(By.CSS_SELECTOR, ".icheckbox").click()
        driverX.find_element(By.NAME, "commit").click()
        return driverX


    def open_real(self, customer: dict):
        driverR = webdriver.Firefox()

        driverR.maximize_window()
        screen_width = driverR.execute_script("return window.screen.width;")
        screen_height = driverR.execute_script("return window.screen.height;")
        driverR.set_window_position(screen_width // 2, 0)
        driverR.set_window_size(screen_width // 2, screen_height)

        driverR.get(login_url)
        driverR.find_element(By.ID, "user_email").send_keys(customer["USERNAME"])
        driverR.find_element(By.ID, "user_password").send_keys(customer["PASSWORD"])
        driverR.find_element(By.CSS_SELECTOR, ".icheckbox").click()
        driverR.find_element(By.NAME, "commit").click()
        return driverR




    def run(self, minutes: int):

        duooutput("\n======== Start refreshing ========\n")

        duooutput(f"Start Probe Accounts: {datetime.now()}\n")

        driverP1 = self.open_probe(1)
        driverP2 = self.open_probe(2)
        driverP3 = self.open_probe(3)
        driverP4 = self.open_probe(4)
        driverP5 = self.open_probe(5)
        
        #output data
        for customer in self.app_data['CUSTOMERS'].values():
            duooutput(f"Appointment date for {customer['USERNAME']}: {customer['APPOINTMENT_DATE']}")

        duooutput("\nstart loop\n")
        start_time = datetime.now()
        while datetime.now() - start_time < timedelta(minutes=minutes):
            for i in range(5):
                driverP = eval(f"driverP{i+1}")

                self.refreshP(driverP, self.app_data[F"WEBID_P{i+1}"])
                
                #find earlier date
                started_office  = []
                date_now = datetime.now()
                for customer in self.app_data['CUSTOMERS'].values():
                    preferred_offices = customer['PREFERED_OFFICE']
                    for office in preferred_offices:
                        if office in started_office:
                            continue
                        if self.available_dates[office] is not None:
                            current_date= datetime.strptime(customer['APPOINTMENT_DATE'], "%d %B, %Y")
                            if current_date - date_now > timedelta(days=300) and (current_date - self.available_dates[office]) >= timedelta(days=90)\
                                or\
                                current_date - date_now > timedelta(days=100) and current_date - date_now <= timedelta(days=300) and \
                                (current_date - self.available_dates[office]) >= timedelta(days=30)\
                                or\
                                current_date - date_now > timedelta(days=30) and current_date - date_now <= timedelta(days=100) and \
                                (current_date - self.available_dates[office]) >= timedelta(days=7):

                                if customer['IS_READY']:
                                    thread = threading.Thread(target=self.update_appointment, args=(customer, office))
                                    thread.start()
                                    started_office.append(office)

                duooutput("\nsleep for 63 seconds...\n")
                time.sleep(63)            
                                
                    
        #close drivers
        driverP1.quit()
        driverP2.quit()
        driverP3.quit()
        driverP4.quit()
        driverP5.quit()


    def refreshP(self, driverX: webdriver, webID: str): 
        url = f"https://ais.usvisa-info.com/en-ca/niv/schedule/{webID}/payment"
        driverX.get(url)

        duooutput("\n\n--------------------------------------")
        duooutput("Time: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S") )
        duooutput("--------------------------------------")
        
        for office in self.Canada_offices:
            time : datetime = None
            try:
                # 使用XPath查找包含文本"目标文本"的元素
                xpath_expression = f"//*[contains(text(), '{office}')]"

                # 找到包含文本的元素
                element_with_text = driverX.find_element(By.XPATH, xpath_expression)

                # 从包含文本的元素出发，找到相应的元素（例如，找到它的父元素、子元素等）
                # 这里只是一个示例，你可以根据需要进一步定位相应的元素
                related_element = element_with_text.find_element(By.XPATH, "./following-sibling::td")
                time_text = related_element.text
                time = datetime.strptime(time_text, "%d %B, %Y")
                duooutput(f"{office} time: {time}")
            except:
                duooutput(f"{office} time is not available")

            self.available_dates[office] = time


    def update_appointment(self, customer: dict, office: str):

        customer["IS_READY"] = False

        duooutput(f"\t=== found earlier date, start update appointment ===")
        duooutput(f"\t=== customer: {customer['USERNAME']} ===")
        duooutput(f"\t=== current date: {customer['APPOINTMENT_DATE']} ===")
        duooutput(f"\t=== office: {office} ===")
        duooutput(f"\t=== available date: {self.available_dates[office]} ===")
        
        date_today = datetime.today()
        
        #open page
        driverR = self.open_real(customer)
        driverR.get(customer['APPOINTMENT_URL'])

        availeble_date = self.available_dates[office]
                
        #pick city now only Toronto
        try:
            dropdown = driverR.find_element(By.ID, "appointments_consulate_appointment_facility_id")
            select = Select(dropdown)
            select.select_by_visible_text(office)
        except Exception as e:
            duooutput(f"\tCannot select {office}. Return")
            duooutput(str(e))
        
        #click date picker
        try:
            driverR.find_element(By.ID, "appointments_consulate_appointment_date").click()
            #click next month button
            num_months = availeble_date.year * 12 + availeble_date.month - (date_today.year * 12 + date_today.month)
            for i in range(num_months-1):
                driverR.find_element(By.CSS_SELECTOR, ".ui-icon-circle-triangle-e").click()
            #find day and click
            day = availeble_date.day
            day_text = str(day)
            driverR.find_element(By.LINK_TEXT, day_text).click()
        except Exception as e:
            duooutput("\tCannot select date. Return")
            duooutput(str(e))

        #update time, pick the first available time, because normally these is only one time slot available
        try:
            dropdown = driverR.find_element(By.ID, "appointments_consulate_appointment_time")   
            select = Select(dropdown)
            select.select_by_index(1)
        except Exception as e:
            duooutput("\tCannot select time. Return")
            duooutput(str(e))

        #click submit
        try:
            driverR.find_element(By.ID, "appointments_submit").click()
        except Exception as e:
            duooutput("\tCannot click submit. Return")
            duooutput(str(e))

        #click confirm 
        success = False
        try:
            driverR.find_element(By.LINK_TEXT, "Confirm").click()
            success = True
        except Exception as e:
            duooutput("\tCannot click confirm. Return")
            duooutput(str(e))
            success = False

        if success:
            duooutput("\n\t=== update appointment success ===\n")
            #success, update data
            customer['APPOINTMENT_DATE'] = availeble_date.strftime("%d %B, %Y")
            with open('data.txt', 'w') as f:
                json.dump(self.app_data, f)
        else:
            duooutput("\n\t=== update appointment failed ===\n")

        i = 0
        while os.path.isfile(f"screenshot{i}.png"):
            i += 1
        driverR.get_screenshot_as_file(f"screenshot{i}.png")

        driverR.quit()
        duooutput("\n\tsleep for 5 minutes\n")
        time.sleep(300)
        customer["IS_READY"] = True
        return
            






    


if __name__ == "__main__":
    duooutput("\n======== Start Program ========")
    duooutput("This program will run until it finds an earlier date than the current date\n")

    myupdater = Reschedule()

    while True:
        myupdater.run(60)
        duooutput(f"\nsleep for 1/2 hour, will be back on {datetime.now()+timedelta(minutes=30)}\n")
        time.sleep(1800)


