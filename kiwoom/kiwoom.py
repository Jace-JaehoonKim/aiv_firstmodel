from PyQt5.QAxContainer import *  # p98
from PyQt5.QtCore import *  # 핸들값 다루기 위한 이벤트루프 클래스
from config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()  # 상속받은 클래스의 init을 실행시키는 명령어

        print('kiwoom클래스입니다')

        ##############evnetloop모음##############
        self.login_event_loop = None
        ########################################
        ##############변수모음##############
        self.account_num = None
        ########################################

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect()
        self.get_account_info()

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)

    def login_slot(self, errCode):
        print(errCode)
        print(errors(errCode))
        self.login_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # pyqt 모듈
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(';')[0]

        print("나의 보유 계좌번호 %s" % self.account_num)
