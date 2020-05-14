from PyQt5.QAxContainer import *  # p98
from PyQt5.QtCore import *  # 핸들값 다루기 위한 이벤트루프 클래스
from config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()  # 상속받은 클래스의 init을 실행시키는 명령어

        print('kiwoom클래스입니다')

        ##############evnetloop모음##############
        self.login_event_loop = None
        self.detail_account_info_event_loop = None
        self.detail_account_info_event_loop_2 = None
        ########################################

        ##############변수모음###################
        self.account_num = None
        ########################################

        self.get_ocx_instance()

        self.event_slots() #이벤트 걸어놓는 곳

        self.signal_login_commConnect()
        self.get_account_info()  # 계좌번호 가져오는것
        self.detail_account_info() # 예수금 요청
        self.detail_account_mystock() # 계좌평가 잔고내역 요청



    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

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



    def detail_account_info(self):
        print("나의 예수금을 요청하는 부분 ")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000" )
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분","00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        # 내가지은요청이름 / TR코드 / preNext / 화면번호(그룹핑해주는용도)
        self.dynamicCall("CommRqData(String,String, int, String)" , "예수금상세현황요청",  "opw00001", "0", "2000")

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self ,sPrevNext ="0"):
        print("계좌평가잔고내역요청")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000" )
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분","00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        # 내가지은요청이름 / TR코드 / preNext / 화면번호(그룹핑해주는용도)
        self.dynamicCall("CommRqData(String,String, int, String)" , "계좌평가잔고내역요청",  "opw00018", sPrevNext, "2000")

        self.detail_account_info_event_loop_2 = QEventLoop()
        self.detail_account_info_event_loop_2.exec_()


    def trdata_slot(self ,sScrNo,sRQName,sTrCode,sRecordName,sPrevNext ):
        '''
        tr요청을 받는 슬롯이다
        :param sScrNo: 스크린번호
        :param sRQName: 내가 요청했을때 지은 이읆
        :param sTrCode: 요청id tr코드
        :param sRecordName: 사용안함
        :param sPrevNext: 다음페이지가 있는지
        :return:
        '''

        if sRQName == "예수금상세현황요청":
            #GetCommData (Commget은 안씀) TR코드 / 레코드이름 / TR반복부 / TR에서 얻어오려는 출력항목이름
            deposit = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode , sRQName,0,"예수금")
            print("예수금 %s" %int(deposit))
            withdraw_pos = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능액 %s" %int(withdraw_pos))

            self.detail_account_info_event_loop.exit()

        if sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode , sRQName,0,"총매입금액")
            total_buy_money_result = int(total_buy_money)

            print("총매입금액 :%s" %total_buy_money_result)

            total_profit_loss_rate = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode , sRQName,0,"총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)
            print("총수익률(%%) : %s" % total_profit_loss_rate_result)

            self.detail_account_info_event_loop_2.exit()
