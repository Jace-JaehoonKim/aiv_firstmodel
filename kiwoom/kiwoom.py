from PyQt5.QaxContainer import *  # p98

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__() #상속받은 클래스의 init을 실행시키는 명령어
        print('kiwoom클래스입니다')


    def get_ocx_instance