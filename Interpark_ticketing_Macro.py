import sys
import random
from typing import Any
from PyQt5 import QtCore, QtGui, QtWidgets, uic
from selenium import webdriver
import pyautogui as pag
import pytesseract
import cv2
import numpy as np
import time
from PIL import ImageGrab, ImageTk
from selenium.webdriver.support import ui
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import re
import smtplib
import winsound as sd
from win32api import GetSystemMetrics




form_class = uic.loadUiType('TicketingMacro_1.ui')[0] 
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
 
class MyWindow(QtWidgets.QMainWindow, form_class):
    is_data_disabled = False  # 정보 입력창 비활성화 여부
    time = None  # 스레드로부터 받아오는 현재 시간

    
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.user_password_input.setEchoMode(QtWidgets.QLineEdit.Password)
        self.seats_number_spinbox.setValue(1)
        self.date_edit.setDate(QtCore.QDate.currentDate())
        self.time_edit.setTime(QtCore.QTime.currentTime())
        self.disableElements(self.time_edit, self.label_auto_start)
 
        self.apply_btn.clicked.connect(self.initData)  # 확인 버튼 클릭
        self.login_btn.clicked.connect(self.startLogin)  # 로그인 버튼 클릭
        self.start_ticketing_btn.clicked.connect(self.startTicketing)  # 시작 버튼 클릭
        self.seats_number_spinbox.valueChanged.connect(self.checkSeatsNumber)  # 좌석 개수 변경
        self.use_auto_start.stateChanged.connect(self.useAutoStart)  # 자동 시작 여부 변경  
        


    def mouseMoveEvent(self, event):
        m_x = event.globalX()
        m_y = event.globalY()

        m_t = "({0}, {1})".format(m_x, m_y)
        self.label_mouset.setText(m_t)

    def initData(self):
        if not self.is_data_disabled:  # 정보 입력창 활성화 시
            self.user_id = self.user_id_input.text()  # 아이디
            self.user_password = self.user_password_input.text()  # 비밀번호
            self.product_code = self.product_code_input.text()  # 공연 코드
            self.seats_number = self.seats_number_spinbox.value()  # 좌석 개수
            self.use_random_seat = self.use_seat_select.isChecked()  # 중간(랜덤) 선택 여부
            self.is_canceled_ticketing = self.canceled_ticket_mode.isChecked()  # 취켓팅 여부
            self.left_x1=self.lineEdit_x1.text() #보안문자 캡쳐 x1
            self.top_y1=self.lineEdit_y1.text() #보안문자 캡쳐 y1
            self.right_x2=self.lineEdit_x2.text() #보안문자 캡쳐 x2
            self.bot_y2=self.lineEdit_y2.text() #보안문자 캡쳐 y2
            self.new_x=self.lineEdit_newx.text() #보안문자 캡쳐 새로고침 x
            self.new_y=self.lineEdit_newy.text() #보안문자 캡쳐 새로고침 y
            self.textinput_x=self.lineEdit_textx.text() #보안문자 캡쳐 텍스트창 x
            self.textinput_y=self.lineEdit_texty.text() #보안문자 캡쳐 텍스트창 y
            self.capchat=self.checkBox_capcha.isChecked() #보안문자 자동입력 관련

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
            self.date = self.date_edit.date()  # 날짜
            self.date = self.date.toPyDate()
            self.date = str(self.date).replace('-', '')
            self.disableElements(self.user_id_input, self.user_password_input, self.product_code_input, self.date_edit,
                                 self.seats_number_spinbox, self.use_seat_select, self.use_auto_start, self.time_edit, self.canceled_ticket_mode,
                                 self.lineEdit_x1, self.lineEdit_y1, self.lineEdit_x2, self.lineEdit_y2, self.lineEdit_newx, self.lineEdit_newy, self.lineEdit_textx, self.lineEdit_texty, self.checkBox_capcha)  # 요소 비활성화
            self.is_data_disabled = True
            self.apply_btn.setText('수정')
            #QtWidgets.QMessageBox.information(self, '완료', '정보 입력이 완료되었습니다.')
        else:
            self.is_data_disabled = False
            self.apply_btn.setText('확인')
            self.enableElements(self.user_id_input, self.user_password_input, self.product_code_input, self.date_edit,
                                self.seats_number_spinbox, self.use_seat_select, self.use_auto_start, self.time_edit, self.canceled_ticket_mode,
                                self.lineEdit_x1, self.lineEdit_y1, self.lineEdit_x2, self.lineEdit_y2, self.lineEdit_newx, self.lineEdit_newy, self.lineEdit_textx, self.lineEdit_texty, self.checkBox_capcha)  # 요소 활성화
 
    # 좌석 개수 변경
    def checkSeatsNumber(self):
        if self.seats_number_spinbox.value() > 4:
            self.seats_number_spinbox.setValue(4)
            QtWidgets.QMessageBox.critical(self, '오류', '최대 좌석 개수는 4개입니다.')
        if self.seats_number_spinbox.value() < 1:
            self.seats_number_spinbox.setValue(1)
            QtWidgets.QMessageBox.critical(self, '오류', '최소 좌석 개수는 1개입니다.')
 
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
        time.sleep(1)
        if self.apply_btn.text() == '확인':  # 확인 버튼을 누르지 않았을 때
            QtWidgets.QMessageBox.critical(self, '오류', '좌석 정보를 입력하세요.')
        else:
            # Dictionary 형식으로 정보 전달
            ticketing_data_to_send = {
                'user_id': self.user_id,  # 아이디
                'user_pw': self.user_password,  # 비밀번호
                'product_code': self.product_code,  # 공연 번호
                'date': self.date,  # 공연 날짜
                'time': self.time,  # 자동시작 시간
                'seats_number': self.seats_number,  # 좌석 개수
                'use_random_seat': self.use_random_seat,  # 랜덤 선택 여부
                'time_signal': self.time_th.time_signal,  # 시간 스레드
                'is_canceled': self.is_canceled_ticketing,  # 취켓팅 여부
                'left': self.left_x1,
                'top': self.top_y1,
                'right': self.right_x2,
                'bot':self.bot_y2,
                'new_x': self.new_x,
                'new_y': self.new_y,
                'textinput_x': self.textinput_x,
                'textinput_y': self.textinput_y,
                'capcha': self.capchat
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
            #time.sleep(1)
            #self.driver.get('https://ticket.interpark.com/Gate/TPLogin.asp')
        
 
 
class TicketingThread(QtCore.QThread):
    def __init__(self, ticketing_data, parent=None):
        QtCore.QThread.__init__(self, parent)
        self.width=GetSystemMetrics(0)
        self.height=GetSystemMetrics(1)
        self.user_id = ticketing_data['user_id']
        self.user_password = ticketing_data['user_pw']
        self.product_code = ticketing_data['product_code']
        self.date = ticketing_data['date']
        self.left_x1 = ticketing_data['left']
        self.top_y1 = ticketing_data['top']
        self.right_x2 = ticketing_data['right']
        self.bot_y2 = ticketing_data['bot']
        self.new_x = ticketing_data['new_x']
        self.new_y = ticketing_data['new_y']
        self.textinput_x = ticketing_data['textinput_x']
        self.textinput_y = ticketing_data['textinput_y']
        self.capcha=ticketing_data['capcha']
        
        #self.product_code = '21003717'#21004639
        #self.date = '20210729'
        #self.playseq='060'

        self.product_code = '21004665' #4665 = DK vs HLE   #21004663 = AF vs LSB  #'21000894'#21004639
        self.date = '' #'20210629'
        self.playseq=''

        self.left_x1 = '408'
        self.top_y1 = '488'
        self.right_x2 = '720'
        self.bot_y2 = '644'
        self.new_x = '472'
        self.new_y = '334'
        self.textinput_x = '531'
        self.textinput_y = '695'
        #self.capcha=True
        #self.capcha=False
        self.set_time = ticketing_data['time']
        self.seats_number = ticketing_data['seats_number']
        self.use_random_seat = ticketing_data['use_random_seat']
        self.time_signal = ticketing_data['time_signal']
        self.is_canceled_ticketing = ticketing_data['is_canceled']
        
        self.kernel2 = np.ones((2, 2), np.uint8)
        self.kernel3 = np.ones((3, 3), np.uint8)
        self.i=0
        
        time.sleep(1)
        self.driver = webdriver.Chrome('chromedriver.exe')
        self.driver.get('https://ticket.interpark.com/Gate/TPLogin.asp')
        
        self.time_signal.connect(self.checkTime)  # 시간 스레드와 연결(자동시작)
        self.is_logined = False  # 로그인 여부 False
        self.start_ticketing = False  # 시작 여부 False
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
    # 요소는 반드시 CSS 셀렉터로만(ID 우선)
 
    # 로그인
    def login(self):
        #self.driver.get('https://ticket.interpark.com/Gate/TPLogin.asp')
        iframes = self.driver.find_elements_by_tag_name('iframe')
        self.driver.switch_to.frame(iframes[0])
        self.driver.find_element_by_id('userId').send_keys(self.user_id)
        self.driver.find_element_by_id('userPwd').send_keys(self.user_password)
        self.driver.find_element_by_id('btn_login').click()
        #self.driver.get('https://ticket.interpark.com/')  # 인터파크 메인 페이지로 강제 접속
        
 
    def selectSeat(self):
        self.failed_to_get_ticket = False  # 취켓팅용-기본적으로 성공으로 표시
        # 직링 생성
        self.url = 'http://poticket.interpark.com/Book/BookSession.asp?GroupCode={}&Tiki=N&Point=N&PlayDate={}&PlaySeq={}&BizCode=&BizMemberCode='.format(self.product_code, self.date, self.playseq)
        #self.url = 'http://poticket.interpark.com/Book/BookSession.asp?GroupCode={}&Tiki=N&Point=N&PlayDate={}&PlaySeq=&BizCode=&BizMemberCode='.format(self.product_code, self.date)
        self.driver.get(self.url)
        print(self.date)
        time.sleep(0.1)
        self.Lockscreen()

    def beepsound(self):
        fr=2000
        du=100
        sd.Beep(fr,du)

    def Lockscreen(self): 
         
        try: 
            
            exit_path='//a[@class="closeBtn"]'
            self.driver.find_element_by_xpath(exit_path).click()
            time.sleep(0.2)
            #time.sleep(0.5)
            
            
        except:
            try: 
            
                exit_path='//a[@class="closeBtn"]'
                self.driver.find_element_by_xpath(exit_path).click()
                time.sleep(1)
                #time.sleep(0.5)
                
                
            except:
                checker=pag.locateCenterOnScreen('check.png')
            
                pag.click(checker)
                

                exit_path='//a[@class="closeBtn"]'
                self.driver.find_element_by_xpath(exit_path).click()
                time.sleep(0.3)

                #next_path='//a[@id="LargeNextBtnLink"]'
                #self.driver.find_element_by_xpath(next_path).click()
                Nextbtn=pag.locateCenterOnScreen('nextbtn.png')
                try:
                    pag.click(Nextbtn)
                    time.sleep(0.5)
                
                    checker=pag.locateCenterOnScreen('check.png')
                    pag.click(checker)
                    time.sleep(0.3)
                except:
                    pass
            
            

        ## 알림창 닫기 클릭
        #exit_path='//a[@class="closeBtn"]'
        #self.driver.find_element_by_xpath(exit_path).click()
        #time.sleep(0.3)
        #xpath_btn="//a[@id='SmallNextBtnLink']"
        #wait=ui.WebDriverWait(self.driver,20)
        #Next_btn=wait.until(EC.presence_of_element_located(By.XPATH,xpath_btn)).click()
        #wait=ui.WebDriverWait(self.driver,5)
        #img_path="//div[@id='divRecaptcha']/div[@id='capchaInner']/div[@id='capchaImg']"
        #imgimg=wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#capchaImg')))
        #print(imgimg.get_attribute('src'))
        #time.sleep(5)
        #img_path="//img[@id='imgCaptcha']"
        #imgimg=self.driver.find_element_by_xpath(img_path).get_attribute('src')
        #print(imgimg)

    def cracksecuretext(self):
                      
        ## 보안문자 해독
        try:

            boo=pag.locateCenterOnScreen('boo.png')
            da=pag.locateCenterOnScreen('da.png')
            
            #print(boo,da)
            #boo=(412,433)
            #da=(706,463)
            
            self.left_x1 = boo[0]
            self.top_y1 = boo[1]+int((da[1]-boo[1])*1.83*(self.height/1080))
            self.right_x2 = da[0]
            self.bot_y2 = boo[1]+int((da[1]-boo[1])*6.5*(self.height/1080))
        except:
            time.sleep(0.5)
            boo=pag.locateCenterOnScreen('boo.png')
            da=pag.locateCenterOnScreen('da.png')
            
            self.left_x1 = boo[0]
            self.top_y1 = boo[1]+int((da[1]-boo[1])*1.83*(self.height/1080))
            self.right_x2 = da[0]
            self.bot_y2 = boo[1]+int((da[1]-boo[1])*6.5*(self.height/1080))

        self.new_x = boo[0]+int((da[0]-boo[0])*1.313*(self.width/1920))
        self.new_y = boo[1]-int((da[1]-boo[1])*4.1*(self.height/1080))
        self.textinput_x = boo[0]+int((da[0]-boo[0])/2*(self.width/1920))
        self.textinput_y = boo[1]+int((da[1]-boo[1])*8.7*(self.height/1080))
        self.seatdone=(boo[0]+int((da[0]-boo[0])*1.983),boo[1]+int((da[1]-boo[1])*14.37*(self.height/1080)))
        #이미지 분석
        if self.capcha==True:
            imgcount=1
            while True:
                #time.sleep(0.1)
                #path="./text.jpg" # 캡쳐 저장할 로컬 주소
                img_file=pag.screenshot(region=(int(self.left_x1),int(self.top_y1),int(self.right_x2)-int(self.left_x1) ,int(self.bot_y2)-int(self.top_y1))) # 캡쳐 로컬 주소에 저장
                
                img_frame = np.array(img_file)
                img  = cv2.cvtColor(img_frame, cv2.COLOR_RGB2BGR)
                gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
                thres = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                erod1 = cv2.erode(thres, self.kernel2, iterations = 1)
                erod2 = cv2.erode(thres, self.kernel2, iterations = 2)
                erod3 = cv2.erode(thres, self.kernel3, iterations = 1)

                #gtxt=pytesseract.image_to_string(gray,lang='eng').replace(" ","").replace("\n","")
                #ttxt=pytesseract.image_to_string(thres,lang='eng').replace(" ","").replace("\n","")
                #e1txt=pytesseract.image_to_string(erod1,lang='eng').replace(" ","").replace("\n","")
                #e2txt=pytesseract.image_to_string(erod2,lang='eng').replace(" ","").replace("\n","")
                #e3txt=pytesseract.image_to_string(erod3,lang='eng').replace(" ","").replace("\n","")

                #gtxt=pytesseract.image_to_string(gray,config="--user-patterns yourpath/xxx.patterns").replace(" ","").replace("\n","")
                #ttxt=pytesseract.image_to_string(thres,config="--user-patterns yourpath/xxx.patterns").replace(" ","").replace("\n","")
                #e1txt=pytesseract.image_to_string(erod1,config="--user-patterns yourpath/xxx.patterns").replace(" ","").replace("\n","")
                #e2txt=pytesseract.image_to_string(erod2,config="--user-patterns yourpath/xxx.patterns").replace(" ","").replace("\n","")
                #e3txt=pytesseract.image_to_string(erod3,config="--user-patterns yourpath/xxx.patterns").replace(" ","").replace("\n","")
                
                #gtxt=pytesseract.image_to_string(gray,lang='eng',config="--user-patterns /xxx2.patterns").replace(" ","").replace("\n","").replace("O","Q").replace("I","T").replace("F","T")
                #ttxt=pytesseract.image_to_string(thres,lang='eng',config="--user-patterns /xxx2.patterns").replace(" ","").replace("\n","").replace("O","Q").replace("I","T").replace("F","T")
                #e1txt=pytesseract.image_to_string(erod1,lang='eng',config="--user-patterns /xxx2.patterns").replace(" ","").replace("\n","").replace("O","Q").replace("I","T").replace("F","T")
                #e2txt=pytesseract.image_to_string(erod2,lang='eng',config="--user-patterns /xxx2.patterns").replace(" ","").replace("\n","").replace("O","Q").replace("I","T").replace("F","T")
                #e3txt=pytesseract.image_to_string(erod3,lang='eng',config="--user-patterns /xxx2.patterns").replace(" ","").replace("\n","")
                
                #gtxt=pytesseract.image_to_string(gray,lang='eng',config='--psm 6').replace(" ","").replace("\n","").replace("O","Q").replace("I","T").replace("F","T")
                ttxt=pytesseract.image_to_string(thres,lang='eng',config='--psm 6').replace(" ","").replace("\n","").replace("O","Q").replace("I","T").replace("F","T")
                e1txt=pytesseract.image_to_string(erod1,lang='eng',config='--psm 6').replace(" ","").replace("\n","").replace("O","Q").replace("I","T").replace("F","T")
                e2txt=pytesseract.image_to_string(erod2,lang='eng',config='--psm 6').replace(" ","").replace("\n","").replace("O","Q").replace("I","T").replace("F","T")
                e3txt=pytesseract.image_to_string(erod3,lang='eng',config="--psm 6").replace(" ","").replace("\n","")
                
                print(ttxt,e1txt,e2txt,e3txt)
                ggtx, tttx, e1etx, e2etx, e3etx = '','','','',''
                p=re.compile('[ABCDEKLMNPQRSTUWXZckpswxz]') #f,g,h,I,J,O,V,Y
                #gtxt6="".join(p.findall(gtxt))
                ttxt6="".join(p.findall(ttxt))
                e1txt6="".join(p.findall(e1txt))
                e2txt6="".join(p.findall(e2txt))
                e3txt6="".join(p.findall(e3txt))
                #for g in gtxt6: ggtx=ggtx+g
                #for t in gtxt6: tttx=tttx+t
                #for e1 in gtxt6: e1etx=e1etx+e1
                #for e2 in gtxt6: e2etx=e2etx+e2
                #for e3 in gtxt6: e3etx=e3etx+e3
                """
                if len(ggtx)==6:
                    txt6=ggtx
                elif len(tttx)==6:
                    txt6=tttx
                elif len(e1etx)==6:
                    txt6=e1etx
                elif len(e2etx)==6:
                    txt6=e2etx
                elif len(e3etx)==6:
                    txt6=e3etx
                else:
                    txt6="NOREAD" """

                
                all_t=ttxt6+' '+e1txt6+' '+e2txt6+' '+e3txt6
                #all_t=gtxt6+' '+ttxt6+' '+e1txt6+' '+e2txt6
                #all_t=gtxt6+' '+ttxt6+' '+e1txt6+' '+e2txt6+' '+e3txt6
                #all_t=e2txt6+' '+e1txt6+' '+ttxt6+' '+gtxt6
                #all_t=all_t.replace("\n","")
                print(all_t)

                p=re.compile('[ABCDEKLMNPQRSTUWXZckpswxz]{6}')
                allt6=p.findall(all_t)
                print(allt6)
                txt6=''
                try:
                    for a in allt6:
                        txt6=a
                    if len(txt6)!=6:
                        txt6="NOREAD"

                except:
                    txt6="NOREAD"


                try:
                    pag.click(int(self.textinput_x),int(self.textinput_y))
                    pag.typewrite(txt6)
                    pag.typewrite('\n')
                except:
                    pass

                #screen=ImageGrab.grab()
                check_white=(int(self.new_x),int(self.new_y))
                if ImageGrab.grab().getpixel(check_white) != (255,255,255): # check_white가 흰색이 아니라면
                    self.use_random_seat=True
                    self.seats_number=1

                    #self.driver.refresh()
                    #time.sleep(0.5)
                    #self.Lockscreen()
                    break
                else:
                    cv2.imwrite('aaa{}.jpg'.format(imgcount), gray)
                    imgcount+=1
                    
                    time.sleep(0.3)
                    continue
        else:
            try:
                pag.click(int(self.textinput_x),int(self.textinput_y))
            except:
                pag.click(int(self.textinput_x),int(self.textinput_y))

            check_white=(int(self.new_x),int(self.new_y))
            while True:
                if ImageGrab.grab().getpixel(check_white) != (255,255,255): # check_white가 흰색이 아니라면
                    self.use_random_seat=True
                    self.seats_number=1

                    #self.driver.refresh()
                    #time.sleep(0.5)
                    #self.Lockscreen()
                    break
                     
           

    def selectseat3(self):
        # 요소를 찾을 때까지 무한 루프(찾지 못하면 예외)
        while True:
            try:
                first_iframe = self.driver.find_element_by_id('ifrmSeat')  # 첫번째 아이프레임
            except:
                continue
            else:
                break
        self.driver.switch_to.frame(first_iframe)
        while True:
            try:
                next_iframe = self.driver.find_element_by_id('ifrmSeatDetail')  # 두번째 좌석선택 아이프레임
            except:
                continue
            else:
                break
        self.driver.switch_to.frame(next_iframe)

    def selectSeat2(self):
        # 요소를 찾을 때까지 무한 루프(찾지 못하면 예외)
        while True:
            try:
                self.first_iframe = self.driver.find_element_by_id('ifrmSeat')  # 첫번째 아이프레임
            except:
                continue
            else:
                break
        self.driver.switch_to.frame(self.first_iframe)
        while True:
            try:
                next_iframe = self.driver.find_element_by_id('ifrmSeatDetail')  # 두번째 좌석선택 아이프레임
            except:
                continue
            else:
                break
        self.driver.switch_to.frame(next_iframe)
        self.i=0
        self.loop_time = 0  #취켓팅용
        while True:
            #self.selectseat3()
            try:
                ets = self.driver.find_elements_by_class_name('stySeat')
                #print(type(ets))
                elements=[]
                for eee in ets:
                    #print(type(eee.get_attribute('alt')))
                    #try: 
                    if eee.get_attribute('alt')!="":
                        #print(eee.get_attribute('src'))
                        elements=elements+[eee]
                    #except:
                    #print(8)
                    #pass

                #print(elements)
                #elements = self.driver.find_elements_by_class_name('stySeat')
                
                self.loop_time += 1
                loop=1
                # 취켓팅하면서 루프를 100번 돌면 루프 탈출
                if self.is_canceled_ticketing and (self.loop_time > loop):
                    self.failed_to_get_ticket = True  # 잔여 좌석 인식 실패(개수 0)
                    #print(5)
                    self.i += 1
                    print(self.i)
                    self.driver.refresh()
                    #time.sleep(0.3)
                    
                    self.Lockscreen()

                    if self.i > 1000:
                        
                        self.driver.get('https://ticket.interpark.com/Gate/TPLogout.asp')
                        self.is_logined = False
                        
                        if not self.is_logined:
                            self.login()
                            self.is_logined = True

                        if str(self.set_time) == str(self.time):  # 입력한 시간하고 일치한다면
                            self.start_ticketing = True

                        if self.start_ticketing:
                            self.selectSeat()  # 티켓팅 시작
                            #self.Lockscreen()
                            self.cracksecuretext()
                            self.i=0
                            self.loop_time=0
                            self.selectseat3()   
                            
                            continue
                        

                    
                    time.sleep(0.1)
                    
                    
                    check_white1=(int(self.new_x),int(self.new_y))
                    if ImageGrab.grab().getpixel(check_white1) == (255,255,255): # check_white가 흰색이 아니라면
                        
                        self.cracksecuretext()
                    self.loop_time=0
                    self.selectseat3()            
                    continue

                elif len(elements) == 0:  # 예외는 나지 않았지만 좌석이 인식되지 못한 경우
                    
                    #time.sleep(1)
                    continue
                else:
                    
                    #print(7)
                    break
            
            except:
                self.driver.get('https://ticket.interpark.com/Gate/TPLogout.asp')
                self.is_logined = False
                
                if not self.is_logined:
                    self.login()
                    self.is_logined = True

                if str(self.set_time) == str(self.time):  # 입력한 시간하고 일치한다면
                    self.start_ticketing = True

                if self.start_ticketing:
                    self.selectSeat()  # 티켓팅 시작
                    #self.Lockscreen()
                    #time.sleep(0.3)
                    self.cracksecuretext()
                    self.i=0
                    self.loop_time=0
                    self.selectseat3()
                continue
            

        while len(elements) > 0:
            # 첫번째부터 1개 선택
            print("!!!!!!!!!!!!!!!!!!!Choice!!!!!!!!!!!!!!!!!")
            
            if (not self.use_random_seat) and (self.seats_number == 1):
                self.driver.find_element_by_css_selector('#TmgsTable > tbody > tr > td > img:nth-child(3)').click()  # 바로 첫번째 요소 선택
                try:
                    self.driver.switch_to.default_content()  # 제일 바깥으로 아이프레임 벗어남
                    self.driver.switch_to.frame(self.first_iframe)  # 첫번째 아이프레임으로 다시 변경
                    self.driver.find_element_by_id('NextStepImage').click()  # 다음 버튼 클릭
                    self.driver.find_element_by_class_name('title')  # 임의의 요소 찾아서 크롬 종료 방어
                except:
                    pass

            # 여러개 선택
            elif not self.use_random_seat:
                seat_count = 0
                for element in elements:
                    if seat_count < self.seats_number:
                        element.click()
                        seat_count += 1
                    try:
                        self.driver.switch_to.default_content()  # 제일 바깥으로 아이프레임 벗어남
                        self.driver.switch_to.frame(self.first_iframe)  # 첫번째 아이프레임으로 다시 변경
                        self.driver.find_element_by_id('NextStepImage').click()  # 다음 버튼 클릭
                        self.driver.find_element_by_class_name('title')  # 임의의 요소 찾아서 크롬 종료 방어
                    except:
                        pass

            # 랜덤/여러개 선택
            elif self.use_random_seat:
                start_seat = random.randint(0, len(elements) - self.seats_number)
                if start_seat < self.seats_number + start_seat:
                    for i in range(start_seat, start_seat + self.seats_number):
                        #if elements[i]
                        try:
                            print(1) 
                            elements[i].click()
                            j=i
                            
                        except: 
                            print(2)
                            pass

                try:
                    #print(3)
                    self.driver.switch_to.default_content()  # 제일 바깥으로 아이프레임 벗어남
                    #print(4)
                    self.driver.switch_to.frame(self.first_iframe)  # 첫번째 아이프레임으로 다시 변경
                    
                    #pag.click(self.seatdone)
                    
                    #WebDriverWait(self.driver, 0.5).until(EC.presence_of_element_located((By.ID, "NextStepImage"))).click()
                    
                    self.driver.find_element_by_id('NextStepImage').click()  # 다음 버튼 클릭
                    
                    
                    self.driver.find_element_by_class_name('title')  # 임의의 요소 찾아서 크롬 종료 방어
                    #time.sleep(0.5)
                    self.beepsound()
                    break
                
                except:
                    #seatdone=pag.locateCenterOnScreen('seatdone.png')
                    pag.click(self.seatdone)
                    self.beepsound()
                
                time.sleep(1)    
                checker=pag.locateCenterOnScreen('check.png')
                
                try:
                    
                    if len(checker)!=0 and j>0:
                        time.sleep(0.1)
                        checker=pag.locateCenterOnScreen('check.png')
                
                        pag.click(checker)
                        elements[j].click()
                            
                        elements.pop(j)
                        continue

                    elif len(checker)!=0 and j==0:
                        time.sleep(0.1)
                        checker=pag.locateCenterOnScreen('check.png')
                
                        pag.click(checker)
                        self.driver.refresh()
                        #time.sleep(0.5)
                        self.Lockscreen()
            

                        self.i=0
                        self.loop_time=0
                        self.selectSeat2()
                        pass

                    else:
                        break
                
                    
                except:
                    elements.pop(j)
                        
                    break

            else:
                #pag.click(979,861)
                break



    def pay(self):
        try:
            time.sleep(3)

            arrow=pag.locateCenterOnScreen('arrow.png')
            arrow1=(arrow[0],arrow[1]+int((self.bot_y2-self.top_y1)/3))
            pag.click(arrow)
            pag.moveTo(arrow[0],arrow[1],0.1)
            time.sleep(0.1)
            pag.click(arrow1)
            won=pag.locateCenterOnScreen('won.png')
            won1=(won[0],won[1]+int((self.bot_y2-self.top_y1)/3.5))
            time.sleep(0.2)
            pag.click(won1)
            time.sleep(0.5)
            try:
                rec=pag.locateCenterOnScreen('rec.png')
                pag.click(rec)
                time.sleep(0.1)
            except:
                time.sleep(3)
                pass

            exam=pag.locateCenterOnScreen('exam.png')
            exam1=(exam[0]-30,exam[1])
            pag.click(exam1)
            time.sleep(0.1)
            pag.typewrite("931103")
            pag.click(won1)
            
            time.sleep(1)

            kakao=pag.locateCenterOnScreen('kakao.png')
            ooo=pag.locateCenterOnScreen('o.png')
            kakao1=(ooo[0],kakao[1])
            pag.click(kakao1)
            pag.click(won1)
            
            time.sleep(1)

            rec=pag.locateCenterOnScreen('rec.png')
            pag.click(rec)
            time.sleep(0.1)

            #rec=pag.locateCenterOnScreen('rec.png')
            #pag.click(rec)
            #time.sleep(0.1)
            
            #rec=pag.locateCenterOnScreen('rec.png')
            #pag.click(rec)
            #time.sleep(0.1)
            pag.click(won1)

            screen3=pag.screenshot('pay.png')

            from email.mime.text import MIMEText
            from email import mime
            s=smtplib.SMTP('smtp.gmail.com',587)
            s.starttls()
            s.login('wlfhfdl3@gmail.com','dgkamtpmoznfkgnx')
            msg=MIMEText('내용: 노트북으로 가서 결재 빨리 ㄱㄱ')
            msg['Subject']='제목: 결재요청'
            s.sendmail("wlfhfdl3@gmail.com","uia97624@vitesco.com; qnslqnel93@naver.com; pk0092@naver.com;",msg.as_string())
            s.quit()
            time.sleep(100)
        except:
            from email.mime.text import MIMEText
            from email import mime
            s=smtplib.SMTP('smtp.gmail.com',587)
            s.starttls()
            s.login('wlfhfdl3@gmail.com','dgkamtpmoznfkgnx')
            msg=MIMEText('내용: 에러에러 긴급긴급')
            msg['Subject']='제목: 에러에러'
            s.sendmail("wlfhfdl3@gmail.com","uia97624@vitesco.com; qnslqnel93@naver.com; pk0092@naver.com;",msg.as_string())
            s.quit()
            time.sleep(100)
        #
    #        pass#continue
        
    
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
            if self.start_ticketing:
                self.selectSeat()  # 티켓팅 시작
                self.cracksecuretext()
                self.selectSeat2()
                self.pay()
                if not self.is_canceled_ticketing:  # 취켓팅이 아니고 끝난 경우는 성공으로 처리
                    self.start_ticketing = False
                if not self.failed_to_get_ticket:  # 취켓팅인데 성공한 경우
                    self.start_ticketing = False
                if self.is_canceled_ticketing and self.failed_to_get_ticket:  # 취켓팅인데 실패한 경우
                    pass
 
 
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
    