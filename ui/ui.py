from kiwoom.kiwoom import *
from PyQt5.Qtwidgets import *
import sys

class Ui_class():
    def __init__(self):
        print('ui class입니다')

        self.app=QApplication(sys.argv) #36강_  리스트 형태로 담겨져 있는  sys.argv = ['파이썬파일경로','옵션',...]

        self.kiwoom=Kiwoom()

        self.app.exec_() #36강_ 이벤트 루프를 실행시켜주는 명령어 계속 이벤트루프랑 돌도록