import sys
import pyautogui as pag
import numpy
from typing import Any
from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5 import QtCore, QtWidgets, uic
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import winsound as sd
import pyperclip
import time
import re



form_class = uic.loadUiType('VaccineMacro.ui')[0]


class MyWindow(QtWidgets.QMainWindow, form_class):
    is_data_disabled = False  # 정보 입력창 비활성화 여부
    time = None  # 스레드로부터 받아오는 현재 시간
      
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.user_password_input.setEchoMode(QtWidgets.QLineEdit.Password)
        self.time_edit.setTime(QtCore.QTime.currentTime())
        self.disableElements(self.time_edit, self.label_auto_start)
 
        self.apply_btn.clicked.connect(self.initData)  # 확인 버튼 클릭
        self.login_btn.clicked.connect(self.startLogin)  # 로그인 버튼 클릭
        self.start_ticketing_btn.clicked.connect(self.startTicketing)  # 시작 버튼 클릭
        self.use_auto_start.stateChanged.connect(self.useAutoStart)  # 자동 시작 여부 변경  
        
    def initData(self):
        if not self.is_data_disabled:  # 정보 입력창 활성화 시
            self.user_id = self.user_id_input.text()  # 아이디
            self.user_password = self.user_password_input.text()  # 비밀번호
            self.use_random_seat = self.use_seat_select.isChecked()  # 알림음
                        
            if self.use_auto_start.isChecked():  # 자동 시작 시 시간 24시간 방식 -> 12시간 방식 변경
                self.time = self.time_edit.time()
                self.time = str(self.time.toPyTime())[:6] + '00'
                if int(self.time[0:2]) > 12:
                    self.new_time_h = int(self.time[0:2]) - 12
                    self.time = str(self.new_time_h) + self.time[2:]
                    if self.new_time_h < 10:
                        self.time = '0' + self.time
                elif self.time[0:2] == '00':
                    self.time = '12' + self.time[2:]
            self.disableElements(self.user_id_input, self.user_password_input, self.use_seat_select,
                                 self.use_auto_start, self.time_edit)  # 요소 비활성화
            self.is_data_disabled = True
            self.apply_btn.setText('수정')
            #QtWidgets.QMessageBox.information(self, '완료', '정보 입력이 완료되었습니다.')
        else:
            self.is_data_disabled = False
            self.apply_btn.setText('확인')
            self.enableElements(self.user_id_input, self.user_password_input, self.use_seat_select,
                                 self.use_auto_start, self.time_edit)
    
    # 요소들 입력받아 한꺼번에 활성화
    def enableElements(self, *elements):
        for element in elements:
            element.setEnabled(True)
 
    # 요소들 입력받아 한꺼번에 비활성화
    def disableElements(self, *elements):
        for element in elements:
            element.setEnabled(False)
 
    # 자동시작 시
    def useAutoStart(self):
        if self.use_auto_start.isChecked():
            self.enableElements(self.time_edit, self.label_auto_start)
        else:
            self.time = None
            self.disableElements(self.time_edit, self.label_auto_start)
 
    # 로그인
    def startLogin(self):
        self.time_th = TimeThread()
        time.sleep(0.5)
        if self.apply_btn.text() == '확인':  # 확인 버튼을 누르지 않았을 때
            QtWidgets.QMessageBox.critical(self, '오류', '확인 버튼을 누르세요.')
        else:
            # Dictionary 형식으로 정보 전달
            ticketing_data_to_send = {
                'user_id': self.user_id,  # 아이디
                'user_pw': self.user_password,  # 비밀번호
                'time': self.time,  # 자동시작 시간
                'use_random_seat': self.use_random_seat,  # 알림음 설정
                'time_signal': self.time_th.time_signal  # 시간 스레드
            }
            self.apply_btn.setEnabled(False)  # 확인(수정) 버튼 비활성화
            self.login_btn.setEnabled(False)  # 로그인 버튼 비활성화
            # 시간/티켓팅 스레드 시작
            self.time_th.start()
            self.time_th.time_signal.connect(self.changeTime)
            self.ticketing_th = TicketingThread(ticketing_data=ticketing_data_to_send)
            self.ticketing_th.start()
 
    # 티켓팅 시작
    def startTicketing(self):
        self.ticketing_th.start_ticketing = True
 
    # 타이머 시작
    def startTimer(self):
        self.time_th = TimeThread()
        self.time_th.start()
        self.time_th.time_signal.connect(self.changeTime)
 
    @QtCore.pyqtSlot(str)
    def changeTime(self, time):
        if time == 'error':  # 웹드라이버(시계) 종료시 알림
            QtWidgets.QMessageBox.critical(self, '오류', '타이머를 끄지 마십시오.')
        else:
            self.now_time.setText(time)
 
class TimeThread(QtCore.QThread):
    time_signal = QtCore.pyqtSignal(str)

    def run(self):
        while True:
            #time.sleep(2)
            self.driver = webdriver.Chrome('chromedriver.exe')
            #네이버 시계 접속
            self.driver.get('https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&query=%EB%84%A4%EC%9D%B4%EB%B2%84+%EC%8B%9C%EA%B3%84&oquery=%EC%8B%9C%EA%B3%84&tqi=UEMiLdprvTossFQzFhCssssssKG-165498')
            while True:
                try:
                    text = self.driver.find_element_by_css_selector('#_cs_domestic_clock > div._timeLayer.time_bx > div > div').text  # 네이버 시계 text
                    text = text.replace('\n', '').replace(' ', '')[0:8]
                except:
                    self.time_signal.emit('error')  # 웹드라이버(시계) 종료 시 error emit
                    break
                self.time_signal.emit(''.join(text))

class TicketingThread(QtCore.QThread):
    def __init__(self, ticketing_data, parent=None):
        QtCore.QThread.__init__(self, parent)
        self.user_id = ticketing_data['user_id']
        self.user_password = ticketing_data['user_pw']
        self.set_time = ticketing_data['time']
        self.use_random_seat = ticketing_data['use_random_seat']
        self.time_signal = ticketing_data['time_signal']
                
        time.sleep(1)
        self.driver = webdriver.Chrome('chromedriver.exe')
        self.driver.get('https://nid.naver.com/nidlogin.login?mode=form&url=https%3A%2F%2Fwww.naver.com')

        self.time_signal.connect(self.checkTime)  # 시간 스레드와 연결(자동시작)
        self.is_logined = False  # 로그인 여부 False
        self.start_ticketing = False  # 시작 여부 False
        self.done = False
        self.VaccineA=[]
        
    
 
    # 로그인
    def login(self):
        self.driver.switch_to.default_content()
        self.IDpath="//*[@id='id']"
        self.PWpath="//*[@id='pw']"
        self.LoginBtnpath="//*[@id='log.login']"
        pyperclip.copy(self.user_id)
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH,self.IDpath))).send_keys(Keys.CONTROL,'v')
        
        pyperclip.copy(self.user_password)
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH,self.PWpath))).send_keys(Keys.CONTROL,'v')
        
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH,self.LoginBtnpath))).click()

        try:
            WebDriverWait(self.driver, 2).until(EC.presence_of_element_located((By.XPATH,self.PWpath))).send_keys(self.user_password)
            WebDriverWait(self.driver, 2).until(EC.presence_of_element_located((By.XPATH,"//*[@id='captcha']"))).click()
            while True:
                LoginBtnpath="//*[@id='log.login']"
                LLL=self.driver.find_element_by_xpath(LoginBtnpath)
                
                if len(LLL)==0:
                    break

        except:
            pass
    


    def beepsound(self):
        fr=2000
        du=1000
        sd.Beep(fr,du)

    def FindVaccine(self):
        self.failed_to_get_ticket = False
        self.driver.get('https://m.place.naver.com/rest/vaccine?vaccineFilter=used')
        time.sleep(3)
        while True:
            self.driver.refresh()
            self.driver.switch_to.default_content()
            """
            try:
                Myposition0=WebDriverWait(self.driver, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"_3_X4H _1dUVm"))) #현재 위치 On
                Myposition0.click()
            except:
                pass
            
                try:
                    Myposition1=WebDriverWait(self.driver, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"_3_X4H _1dUVm vo0oY"))) #현재 위치 Off
                    Myposition1.click()
                    Myposition0=WebDriverWait(self.driver, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"_3_X4H _1dUVm"))) #현재 위치 On
                    Myposition0.click()
                    WebDriverWait(self.driver, 1).until(EC.presence_of_element_located((By.CLASS_NAME,"_2v3t_"))).click() #현재 지도에서 검색
                except:
                    pass
            """
            try:
                Refresh=WebDriverWait(self.driver, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"_1M6rf"))) #새로 고침 버튼
            except: pass

            Tvaccine=[]
            try:
                Tvaccine=WebDriverWait(self.driver, 1).until(EC.presence_of_all_elements_located((By.CLASS_NAME,"_3sd6u"))) #백신 전체 상황
            except:
                continue
            

            j=0
            for i, T in enumerate(Tvaccine):
                temp=T.get_attribute("innerHTML") #백신 각각 내용
                #print(temp)
                if "span" in temp: #확인 요소
                    if temp[0]=='0':
                        pass
                    else:
                        print(T, temp, temp[0])
                        self.VaccineA=T
                        j=i+1
                        break
            
            if j>0: break

        
    def Choice(self):    
        self.driver.switch_to.default_content()
                
        try:
            self.VaccineA.click()
        except: 
            pass
        
        while True:
            SelectVac=[]
            self.driver.switch_to.default_content()##원인
            
            try:
                #SelectVac=WebDriverWait(self.driver, 1).until(EC.presence_of_element_located((By.CLASS_NAME,"_2RpG_ ur_-G")))
                SelectVac=WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH,"//*[@id='app-root']/div/div/div[3]/div/div/ul/li/div[1]/div/div[2]/a")))
                SelectVac.click()
                print(SelectVac,type(SelectVac))
            except:
                pass
                #SelectVac=pag.locateCenterOnScreen('vaccine.png')
                #pag.click(SelectVac)

            try:
                if len(SelectVac)==0:                
                    pass
                else:
                    break
            except:
                
                break


    def Identity(self):    
       
        self.driver.switch_to.default_content()
        try:
            WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME,"btn_certify"))).click()
        except:
            
            pass

        while True:
            self.driver.switch_to.default_content()
            Identpath="//*[@id='container']/div/div[1]/ul/li[1]/span"
            try:
                WebDriverWait(self.driver, 500).until(EC.presence_of_element_located((By.XPATH,Identpath)))
                break
            except:
                
                continue
        
        if self.use_random_seat:
            self.beepsound()
        
        self.driver.switch_to.default_content()
        try:
            Fail=''
            Fail=WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.CLASS_NAME, "info_box_inner"))).get_attribute('textContent')
        except:
            print(14,Fail)
            pass
        Jan=''
        p=re.compile('잔여')
        Jan=p.findall(Fail)
        try:    
            if len(Jan) > 1:
                print(Fail)
                self.done = True

            else:
                print(Fail)
                self.start_ticketing = True
            
                

                
        except:
            
            self.start_ticketing = True
        
        








    @QtCore.pyqtSlot(str)
    def checkTime(self, data):
        self.time = data
    
    def run(self):
        # 로그인되지 않았다면 로그인
        if not self.is_logined:
            self.login()
            self.is_logined = True
        while True:
            if str(self.set_time) == str(self.time):  # 입력한 시간하고 일치한다면
                self.start_ticketing = True
                self.done = False
            if self.start_ticketing:
                self.FindVaccine()
                self.Choice()
                self.Identity()
                if self.done == True:
                    break



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()