import os
import sys
import time
import urllib
from urllib import request

import pandas as pd
import pyautogui
import pyperclip
import win32api
import win32con
import win32gui
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import *

qt_path = os.path.dirname(os.path.realpath(__file__)) + '/view.ui'
# form_class = uic.loadUiType(qt_path)[0]
# img_file_name = "arrow.png"

# config
img_path = os.path.dirname(os.path.realpath(__file__)) + '/img/'

conf = 0.90
pyautogui.PAUSE = 0.5

temp = 0

fail_name = []
success_name = []
timesleep = 0


# 화면을 띄우는데 사용되는 Class 선언
class Friend:
    def __init__(self, name=None, message=None, img=None, definition=None):
        self.name = name
        self.message = message
        self.img = img
        self.definition = definition


'''
class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.file_path_button.clicked.connect(self.show_file_open)
        self.send_button.clicked.connect(self.sendButtonFunction)
        self.load_friend_Button.clicked.connect(self.show_friend_file_open)
        self.fail_excel_save.clicked.connect(self.fail_save_excel)
        self.success_excel_save.clicked.connect(self.success_save_excel)
        self.time_minus.clicked.connect(self.click_minus)
        self.time_plus.clicked.connect(self.click_plus)
        url1 = 'http://www.hizonenews.com/img/katalk/ka1.png'
        url2 = 'http://www.hizonenews.com/img/katalk/ka2.png'
        img1 = urllib.request.urlopen(url1).read()
        img2 = urllib.request.urlopen(url2).read()
        pixmap1 = QPixmap()
        pixmap2 = QPixmap()
        pixmap1.loadFromData(img1)
        pixmap2.loadFromData(img2)
        self.ka1.setPixmap(pixmap1)
        self.ka2.setPixmap(pixmap2)

    def click_minus(self):
        now = int(self.time_label.text())
        print(now)
        if(now > 0):
            now -=1
        global timesleep
        timesleep = now
        self.time_label.setText(str(now))

    def click_plus(self):
        now = int(self.time_label.text())
        now +=1
        global timesleep
        timesleep = now
        self.time_label.setText(str(now))

    def fail_save_excel(self):
        fileSave = QFileDialog.getSaveFileName(self, 'Save file', './')
        failData = {'성': '', '이름': ''}
        failData['이름'] = fail_name
        df = pd.DataFrame(failData)
        df.to_excel(fileSave[0] + ".xlsx")

    def success_save_excel(self):
        fileSave = QFileDialog.getSaveFileName(self, 'Save file', './')
        successData = {'성': '', '이름': ''}
        successData['이름'] = success_name
        df = pd.DataFrame(successData)
        df.to_excel(fileSave[0] + ".xlsx")

    def sendButtonFunction(self):
        global fail_name
        global success_name
        fail_name.clear()
        success_name.clear()
        imgName = self.img_file_path_label.toPlainText()
        message = self.message_edittext.toPlainText()
        for i in range(0, temp):
            name = self.friend_table.item(i, 0).text()
            step1 = open_chatroom(name)
            if (step1 == True):
                success_name.append(name)
                definition = []
                for j in range(1, 4):
                    definition.append(self.friend_table.item(i, j).text())
                friend = Friend(name, message, imgName, definition)
                kakao_send_init(friend)
            else:
                fail_name.append(name)

        self.fail_name_table.setRowCount(len(fail_name))
        self.fail_name_table.setColumnCount(1)

        self.success_name_table.setRowCount(len(success_name))
        self.success_name_table.setColumnCount(1)

        for i in range(0, len(fail_name)):
            self.fail_name_table.setItem(i, 0, QTableWidgetItem(fail_name[i]))

        for i in range(0, len(success_name)):
            self.success_name_table.setItem(i, 0, QTableWidgetItem(success_name[i]))

    # btn_2가 눌리면 작동할 함수
    def show_file_open(self):
        fname = QFileDialog.getOpenFileName(None, None, None, "Image Files (*.png *.jpg *.bmp *.avi *.mp4)")
        dir, file = os.path.split(fname[0])
        print(dir)
        print(file)
        fileName = os.path.basename(fname[0])
        pixmap = QPixmap(fname[0]).scaled(251, 251, transformMode=QtCore.Qt.SmoothTransformation)
        self.img_priview.setPixmap(pixmap)
        self.img_file_path_label.setText(fname[0])

    # 친구목록 불러오기
    def show_friend_file_open(self):
        # "Excel Fiels (*.xls *.xlsx)"
        fname = QFileDialog.getOpenFileName(None, None, None, "Excel Fiels (*.xlsx *.xls)")
        if (fname[0] == ''):
            print('none')
        else:
            csv = pd.read_excel(fname[0])
            sung = csv['성'].values.tolist()
            name = csv['이름'].values.tolist()
            tt = []
            for i in range(0, len(sung)):
                if (str(type(sung[i])) == "<class 'float'>"):
                    sung[i] = ""
                if (str(type(name[i])) == "<class 'float'>"):
                    name[i] = ""
                tt.append(sung[i] + name[i])
            global temp
            temp = len(tt)
            self.friend_table.setRowCount(len(tt))
            self.friend_table.setColumnCount(4)

            for i in range(0, len(tt)):
                for j in range(0, 4):
                    item = QTableWidgetItem()
                    self.friend_table.setItem(i, j, item)
            for i in range(0, len(tt)):
                self.friend_table.setItem(i, 0, QTableWidgetItem(tt[i]))


'''


def click_img(imagePath):
    location = pyautogui.locateCenterOnScreen(imagePath, confidence=conf)
    x, y = location
    pyautogui.click(x, y)


def click_img_plus_x(imagePath, pixel):
    location = pyautogui.locateCenterOnScreen(imagePath, confidence=conf)
    x, y = location
    pyautogui.click(x + pixel, y)


def doubleClickImg(imagePath):
    location = pyautogui.locateCenterOnScreen(imagePath, confidence=conf)
    x, y = location
    pyautogui.click(x, y, clicks=2)


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


# # 채팅방 열기
def open_chatroom(chatroom_name):
    # # 친구목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx(hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx(hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx(hwndkakao_edit2_1, None, "Edit", None)

    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)  # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(1)
    try:
        isOpenChatRoom = location = pyautogui.locateCenterOnScreen(resource_path('light_gray_clip1.png'),
                                                                   confidence=conf)
        if (isOpenChatRoom == None):
            return False
        return True
    except TypeError:
        return False


# # 채팅방 전송 준비
def kakao_send_init(friend):
    # # 핸들 _ 채팅방
    chatroom_name = friend.name
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RichEdit20W", None)
    # hwndListControl = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None)
    text = friend.message
    definition = friend.definition
    kakao_dif_sendtext(definition, text)

    if (friend.img != ''):
        kakao_sendimg(friend.img)

    time.sleep(timesleep)
    pyautogui.keyDown('esc')

    # win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)


# 채팅방에 test입력
def kakao_sendtext(text):
    pyperclip.copy(text)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.typewrite('\n', interval=0.1)


def kakao_dif_sendtext(dif, text):
    arr = text.split('\n')
    answer = ''
    for i in range(0, len(arr)):
        arr[i] = arr[i].replace('{정의1}', dif[0])
        arr[i] = arr[i].replace('{정의2}', dif[1])
        arr[i] = arr[i].replace('{정의3}', dif[2])

        if (i != len(arr) - 1):
            arr[i] = arr[i] + '\n'
        answer += arr[i]

    pyperclip.copy(answer)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.typewrite('\n', interval=0.1)


# 채팅방에 사진입력
def kakao_sendimg(img_file_name):
    click_img(resource_path('light_gray_clip1.png'))
    dir, file = os.path.split(img_file_name)
    dir = dir.replace('/', '\\')
    pyperclip.copy(dir)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.typewrite('\\', interval=0.1)
    pyperclip.copy(file)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.typewrite('\n', interval=0.1)


class Ui_dialog(object):
    def setupUi(self, dialog):
        dialog.setObjectName("카카오톡 1대1 다중 메시지 매크로")
        dialog.resize(1120, 768)
        self.send_button = QtWidgets.QPushButton(dialog)
        self.send_button.setGeometry(QtCore.QRect(870, 340, 231, 41))
        self.send_button.setStyleSheet("background-color : rgb(241, 233, 0);")
        self.send_button.setObjectName("send_button")
        self.file_path_button = QtWidgets.QPushButton(dialog)
        self.file_path_button.setGeometry(QtCore.QRect(600, 340, 251, 41))
        self.file_path_button.setObjectName("file_path_button")
        self.label_2 = QtWidgets.QLabel(dialog)
        self.label_2.setGeometry(QtCore.QRect(610, 30, 71, 16))
        self.label_2.setObjectName("label_2")
        self.img_file_path_label = QtWidgets.QTextEdit(dialog)
        self.img_file_path_label.setGeometry(QtCore.QRect(600, 310, 251, 21))
        self.img_file_path_label.setObjectName("img_file_path_label")
        self.message_edittext = QtWidgets.QTextEdit(dialog)
        self.message_edittext.setGeometry(QtCore.QRect(870, 50, 231, 221))
        self.message_edittext.setObjectName("message_edittext")
        self.label_4 = QtWidgets.QLabel(dialog)
        self.label_4.setGeometry(QtCore.QRect(40, 70, 81, 16))
        self.label_4.setObjectName("label_4")
        self.label_10 = QtWidgets.QLabel(dialog)
        self.label_10.setGeometry(QtCore.QRect(160, 0, 961, 861))
        self.label_10.setStyleSheet("background-color : rgb(212, 212, 212);")
        self.label_10.setText("")
        self.label_10.setObjectName("label_10")
        self.label_11 = QtWidgets.QLabel(dialog)
        self.label_11.setGeometry(QtCore.QRect(870, 30, 111, 20))
        font = QtGui.QFont()
        font.setFamily("나눔스퀘어라운드 Light")
        self.label_11.setFont(font)
        self.label_11.setObjectName("label_11")
        self.load_friend_Button = QtWidgets.QPushButton(dialog)
        self.load_friend_Button.setGeometry(QtCore.QRect(10, 10, 141, 41))
        self.load_friend_Button.setObjectName("load_friend_Button")
        self.label_5 = QtWidgets.QLabel(dialog)
        self.label_5.setGeometry(QtCore.QRect(0, 0, 161, 851))
        self.label_5.setStyleSheet("background-color: rgb(235, 235, 235);")
        self.label_5.setText("")
        self.label_5.setObjectName("label_5")
        self.friend_table = QtWidgets.QTableWidget(dialog)
        self.friend_table.setGeometry(QtCore.QRect(170, 30, 421, 351))
        self.friend_table.setObjectName("friend_table")
        self.friend_table.setColumnCount(4)
        self.friend_table.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.friend_table.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.friend_table.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.friend_table.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.friend_table.setHorizontalHeaderItem(3, item)
        self.fail_name_table = QtWidgets.QTableWidget(dialog)
        self.fail_name_table.setGeometry(QtCore.QRect(30, 90, 101, 271))
        self.fail_name_table.setObjectName("fail_name_table")
        self.fail_name_table.setColumnCount(1)
        self.fail_name_table.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.fail_name_table.setHorizontalHeaderItem(0, item)
        self.ka1 = QtWidgets.QLabel(dialog)
        self.ka1.setGeometry(QtCore.QRect(180, 400, 350, 350))
        self.ka1.setStyleSheet("")
        self.ka1.setText("")
        self.ka1.setObjectName("ka1")
        self.ka2 = QtWidgets.QLabel(dialog)
        self.ka2.setGeometry(QtCore.QRect(640, 400, 350, 350))
        self.ka2.setStyleSheet("")
        self.ka2.setText("")
        self.ka2.setObjectName("ka2")
        self.img_priview = QtWidgets.QLabel(dialog)
        self.img_priview.setGeometry(QtCore.QRect(600, 50, 251, 251))
        self.img_priview.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.img_priview.setStyleSheet("background-color : rgb(255, 255, 255);")
        self.img_priview.setAlignment(QtCore.Qt.AlignCenter)
        self.img_priview.setObjectName("img_priview")
        self.success_name_table = QtWidgets.QTableWidget(dialog)
        self.success_name_table.setGeometry(QtCore.QRect(30, 430, 101, 291))
        self.success_name_table.setObjectName("success_name_table")
        self.success_name_table.setColumnCount(0)
        self.success_name_table.setRowCount(0)
        self.fail_excel_save = QtWidgets.QPushButton(dialog)
        self.fail_excel_save.setGeometry(QtCore.QRect(30, 370, 101, 23))
        self.fail_excel_save.setObjectName("fail_excel_save")
        self.label = QtWidgets.QLabel(dialog)
        self.label.setGeometry(QtCore.QRect(40, 410, 81, 16))
        self.label.setObjectName("label")
        self.success_excel_save = QtWidgets.QPushButton(dialog)
        self.success_excel_save.setGeometry(QtCore.QRect(30, 730, 101, 23))
        self.success_excel_save.setObjectName("success_excel_save")
        self.time_label = QtWidgets.QLabel(dialog)
        self.time_label.setGeometry(QtCore.QRect(1040, 310, 41, 21))
        self.time_label.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.time_label.setAlignment(QtCore.Qt.AlignCenter)
        self.time_label.setObjectName("time_label")
        self.time_plus = QtWidgets.QPushButton(dialog)
        self.time_plus.setGeometry(QtCore.QRect(1080, 310, 21, 23))
        self.time_plus.setObjectName("time_plus")
        self.time_minus = QtWidgets.QPushButton(dialog)
        self.time_minus.setGeometry(QtCore.QRect(1020, 310, 21, 23))
        self.time_minus.setObjectName("time_minus")
        self.label_6 = QtWidgets.QLabel(dialog)
        self.label_6.setGeometry(QtCore.QRect(870, 280, 231, 21))
        self.label_6.setObjectName("label_6")
        self.label_5.raise_()
        self.label_10.raise_()
        self.send_button.raise_()
        self.file_path_button.raise_()
        self.label_2.raise_()
        self.img_file_path_label.raise_()
        self.message_edittext.raise_()
        self.label_4.raise_()
        self.label_11.raise_()
        self.load_friend_Button.raise_()
        self.friend_table.raise_()
        self.fail_name_table.raise_()
        self.ka1.raise_()
        self.ka2.raise_()
        self.img_priview.raise_()
        self.success_name_table.raise_()
        self.fail_excel_save.raise_()
        self.label.raise_()
        self.success_excel_save.raise_()
        self.time_label.raise_()
        self.time_plus.raise_()
        self.time_minus.raise_()
        self.label_6.raise_()

        self.retranslateUi(dialog)
        self.file_path_button.clicked.connect(self.show_file_open)
        self.send_button.clicked.connect(self.sendButtonFunction)
        self.load_friend_Button.clicked.connect(self.show_friend_file_open)
        self.fail_excel_save.clicked.connect(self.fail_save_excel)
        self.success_excel_save.clicked.connect(self.success_save_excel)
        self.time_minus.clicked.connect(self.click_minus)
        self.time_plus.clicked.connect(self.click_plus)
        url1 = 'http://www.hizonenews.com/img/katalk/ka1.png'
        url2 = 'http://www.hizonenews.com/img/katalk/ka2.png'
        img1 = urllib.request.urlopen(url1).read()
        img2 = urllib.request.urlopen(url2).read()
        pixmap1 = QPixmap()
        pixmap2 = QPixmap()
        pixmap1.loadFromData(img1)
        pixmap2.loadFromData(img2)
        self.ka1.setPixmap(pixmap1)
        self.ka2.setPixmap(pixmap2)
        QtCore.QMetaObject.connectSlotsByName(dialog)

    def retranslateUi(self, dialog):
        _translate = QtCore.QCoreApplication.translate
        dialog.setWindowTitle(_translate("dialog", "카카오톡 1대1 다중 메시지 매크로"))
        self.send_button.setText(_translate("dialog", "전송"))
        self.file_path_button.setText(_translate("dialog", "찾기"))
        self.label_2.setText(_translate("dialog", "이미지 첨부"))
        self.label_4.setText(_translate("dialog", "전송 실패 목록"))
        self.label_11.setText(_translate("dialog", "메시지"))
        self.load_friend_Button.setText(_translate("dialog", "친구목록 불러오기"))
        item = self.friend_table.horizontalHeaderItem(0)
        item.setText(_translate("dialog", "이름"))
        item = self.friend_table.horizontalHeaderItem(1)
        item.setText(_translate("dialog", "정의1"))
        item = self.friend_table.horizontalHeaderItem(2)
        item.setText(_translate("dialog", "정의2"))
        item = self.friend_table.horizontalHeaderItem(3)
        item.setText(_translate("dialog", "정의3"))
        item = self.fail_name_table.horizontalHeaderItem(0)
        item.setText(_translate("dialog", "이름"))
        self.img_priview.setText(_translate("dialog", "사진 미리보기"))
        self.fail_excel_save.setText(_translate("dialog", "엑셀로 저장"))
        self.label.setText(_translate("dialog", "전송 성공 목록"))
        self.success_excel_save.setText(_translate("dialog", "엑셀로 저장"))
        self.time_label.setText(_translate("dialog", "0"))
        self.time_plus.setText(_translate("dialog", "+"))
        self.time_minus.setText(_translate("dialog", "-"))
        self.label_6.setText(_translate("dialog", "전송 속도(동영상 전송시 10 이상 권장)"))

    def click_minus(self):
        now = int(self.time_label.text())
        print(now)
        if (now > 0):
            now -= 1
        global timesleep
        timesleep = now
        self.time_label.setText(str(now))

    def click_plus(self):
        now = int(self.time_label.text())
        now += 1
        global timesleep
        timesleep = now
        self.time_label.setText(str(now))

    def fail_save_excel(self):
        fileSave = QFileDialog.getSaveFileName(None, 'Save file', './')
        if (fileSave[0] == ''):
            print('none')
        else:
            failData = {'성': '', '이름': ''}
            failData['이름'] = fail_name
            df = pd.DataFrame(failData)
            df.to_excel(fileSave[0] + ".xlsx")

    def success_save_excel(self):
        fileSave = QFileDialog.getSaveFileName(None, 'Save file', './')
        if (fileSave[0] == ''):
            print('none')
        else:
            successData = {'성': '', '이름': ''}
            successData['이름'] = success_name
            df = pd.DataFrame(successData)
            df.to_excel(fileSave[0] + ".xlsx")

    def sendButtonFunction(self):
        global fail_name
        global success_name
        fail_name.clear()
        success_name.clear()
        imgName = self.img_file_path_label.toPlainText()
        message = self.message_edittext.toPlainText()
        for i in range(0, temp):
            name = self.friend_table.item(i, 0).text()
            step1 = open_chatroom(name)
            if (step1 == True):
                success_name.append(name)
                definition = []
                for j in range(1, 4):
                    definition.append(self.friend_table.item(i, j).text())
                friend = Friend(name, message, imgName, definition)
                kakao_send_init(friend)
            else:
                fail_name.append(name)

        self.fail_name_table.setRowCount(len(fail_name))
        self.fail_name_table.setColumnCount(1)

        self.success_name_table.setRowCount(len(success_name))
        self.success_name_table.setColumnCount(1)

        for i in range(0, len(fail_name)):
            self.fail_name_table.setItem(i, 0, QTableWidgetItem(fail_name[i]))

        for i in range(0, len(success_name)):
            self.success_name_table.setItem(i, 0, QTableWidgetItem(success_name[i]))

    # btn_2가 눌리면 작동할 함수
    def show_file_open(self):
        fname = QFileDialog.getOpenFileName(None, None, None, "Image Files (*.png *.jpg *.bmp *.avi *.mp4)")
        dir, file = os.path.split(fname[0])
        print(dir)
        print(file)
        fileName = os.path.basename(fname[0])
        pixmap = QPixmap(fname[0]).scaled(251, 251, transformMode=QtCore.Qt.SmoothTransformation)
        self.img_priview.setPixmap(pixmap)
        self.img_file_path_label.setText(fname[0])

    # 친구목록 불러오기
    def show_friend_file_open(self):
        # "Excel Fiels (*.xls *.xlsx)"
        fname = QFileDialog.getOpenFileName(None, None, None, "Excel Fiels (*.xlsx *.xls)")
        if (fname[0] == ''):
            print('none')
        else:
            csv = pd.read_excel(fname[0])
            sung = csv['성'].values.tolist()
            name = csv['이름'].values.tolist()
            tt = []
            for i in range(0, len(sung)):
                if (str(type(sung[i])) == "<class 'float'>"):
                    sung[i] = ""
                if (str(type(name[i])) == "<class 'float'>"):
                    name[i] = ""
                tt.append(sung[i] + name[i])
            global temp
            temp = len(tt)
            self.friend_table.setRowCount(len(tt))
            self.friend_table.setColumnCount(4)

            for i in range(0, len(tt)):
                for j in range(0, 4):
                    item = QTableWidgetItem()
                    self.friend_table.setItem(i, j, item)
            for i in range(0, len(tt)):
                self.friend_table.setItem(i, 0, QTableWidgetItem(tt[i]))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QDialog()
    ui = Ui_dialog()
    ui.setupUi(dialog)
    dialog.show()
    sys.exit(app.exec_())

    '''
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
'''
