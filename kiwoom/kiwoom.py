from PyQt5.QAxContainer import *  # p98
from PyQt5.QtCore import *  # 핸들값 다루기 위한 이벤트루프 클래스
from PyQt5.QtTest import *

import os
import sys

from config.errorCode import *
from config.kiwoomType import *
from config.log_class import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()  # 상속받은 클래스의 init을 실행시키는 명령어
        self.realType = RealType()

        print("kiwoom() class start. ")


        ####### event loop를 실행 ###############
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트루프
        self.detail_account_info_event_loop = QEventLoop()  # 예수금 요청용 이벤트루프
        self.calculator_event_loop = QEventLoop()
        ########################################

        ###########계좌관련 dictionary###########
        self.account_stock_dict = {}            #매수 종목
        self.not_account_stock_dict = {}        #미체결 종목
        ########################################
        ######## 종목 정보 가져오기################
        self.portfolio_stock_dict = {}          #포트폴리오딕셔너리
        self.jango_dict = {}
        ########################################

        ##############변수모음###################
        self.account_num = None                 #계좌번호
        self.account_password = "0000"          #계좌비밀번호
        self.deposit = 0                        #예수금
        self.output_deposit = 0                 #출금가능 금액
        self.use_money = 0                      #실제 투자에 사용할 금액
        self.use_money_percent = 0.9            #예수금증 사용할 금액비율
        self.betting_size = None                #한번 베팅에 사용할 퍼센트
        self.total_buy_money  =0                #총매입금액
        self.total_profit_loss_money = 0        #총평가손익금액
        self.total_profit_loss_rate = 0.0       #총수익률(%)
        ########################################

        ######### 요청 스크린 번호#################
        self.screen_my_info = "2000"            #계좌 관련한 스크린 번호
        self.screen_calculation_stock = "4000"  #계산용 스크린 번호
        self.screen_real_stock = "5000"         #종목별 할당할 스크린 번호
        self.screen_meme_stock = "6000"         #종목별 할당할 주문용스크린 번호
        self.screen_start_stop_real = "1000"    #장 시작/종료 실시간 스크린번호
        ########################################

        ########### 종목 분석 용##################
        self.calcul_data = []
        ########################################




        self.get_ocx_instance()

        self.event_slots()                      #이벤트 걸어놓는 곳
        self.real_event_slot()                  #실시간 이벤트 시그널 / 슬롯 연결
        print(" ")
        print("로그인")
        self.signal_login_commConnect()         #로그인 요청 시그널 포함
        self.signal_get_account_info()          #계좌번호 가져오는것
        print(" ")
        self.detail_account_info()              #예수금 요청
        print(" ")
        self.detail_account_mystock()           #계좌평가 잔고내역 요청
        print(" ")
        QTimer.singleShot(5000, self.not_concluded_account)  # 5초 뒤에 미체결 종목들 가져오기 실행

        QTest.qWait(10000)
        self.read_code()
        print("번호할당 ")
        self.screen_number_setting()
        print("번호할당끝 ")


        QTest.qWait(5000)
        print(" 실시간 장 운영확인")
        # 실시간 시그널 관련 함수
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, '',
                         self.realType.REALTYPE['장시작시간']['장운영구분'], "0")

        # 실시간 종목 틱데이터 함수
        for code in self.portfolio_stock_dict.keys():
            screen_num = self.portfolio_stock_dict[code]['스크린번호']
            fids = self.realType.REALTYPE['주식체결']['체결시간']
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_num, code, fids, "1")
            print("실시간 등록 코드 %s스크린번호 %s  fid번호 %s " %(code,screen_num,fids))



    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")    # 레지스트리에 저장된 api 모듈 불러오기


# 시그널(시작신호:받는행위요청하기) -> 이벤트(함수발생:받는행위하기) -> 슬랏(함수결과:받는곳) 순서로 되어있음
# 코딩은 이벤트 모아서 정리하고, 시그널-슬랏은 기능끼리 모아서 정리했음


#이벤트 정리해두는곳
    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)    #로그인이벤트를 로그인 슬랏에 담음
        self.OnReceiveTrData.connect(self.trdata_slot)  #TR받는이벤트를 TR데이터 슬랏에 담음

    def real_event_slot(self):
        self.OnReceiveRealData.connect(self.realdata_slot)  # 실시간 이벤트 연결
        self.OnReceiveChejanData.connect(self.chejan_slot)  # 종목 주문체결 관련한 이벤트





#로그인Signal
    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")              # 로그인 요청 시그널 pyqt 모듈로 Commconnect함수를 호출한다

        self.login_event_loop.exec_()                  # 로그인 이벤트루프 실행
#계좌정보시그널
    def signal_get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")  # 계좌번호 반환
        self.account_num = account_list.split(';')[0]
        print("나의 보유 계좌번호 %s" % self.account_num)
        user_name = self.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        print("계좌 소유주 %s" % user_name)
        server_gubun = self.dynamicCall("GetLoginInfo(QString)", "GetServerGubun")
        print("서버 구분(모의투자:1) %s" % server_gubun)

#로그인Slot
    def login_slot(self, errCode):
        print(errCode)
        print(errors(errCode))

        self.login_event_loop.exit()



# 조회구분 1도 있던데 뭐지!!!


#예수금확인TR : 인풋입력후 TR요청 1차시그널 발생시키기
    def detail_account_info(self , sPrevNext="0"):
        print("예수금정보 요청")
    #예수금 확인 기본정보 입력
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", self.account_password )
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분","00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
    #예수금 확인 TR요청 1차시그널 : 내가지은요청이름 / TR코드 / preNext / 화면번호(그룹핑해주는용도) -> OnReceiveTRData() 이벤트로감
        self.dynamicCall("CommRqData(QString,QString, int, QString)","예수금상세현황요청","opw00001", sPrevNext ,self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

#계좌평가잔고내역TR : 인풋입력후 TR요청 1차시그널발생시키기
    def detail_account_mystock(self ,sPrevNext ="0"):
        print("계좌평가잔고내역요청")
    #계좌평가잔고 확인 기본정보 입력
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", self.account_password )
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분","00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
    #계좌평가잔고 TR요청 1차시그널 : 내가지은요청이름 / TR코드 / preNext / 화면번호(그룹핑해주는용도) -> OnReceiveTRData() 이벤트로감
        self.dynamicCall("CommRqData(QString,QString, int, QString)","계좌평가잔고내역요청","opw00018",sPrevNext,self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

# 미체결TR : 인풋입력후 TR요청 1차시그널발생시키기
    def not_concluded_account(self, sPrevNext="0"):
        print("미체결잔고내역요청")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1") #미체결
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0") #매수매도전체
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()




#Trdata슬랏 : TR요청 1차시그널을 받아서 / OnReceiveTRData() 이벤트발생하여 / TR 데이터가 담기는 곳

    def trdata_slot(self , sScrNo , sRQName , sTrCode , sRecordName , sPrevNext ): # 반환되는값 5개
        '''
        TR요청의 결과를 받는 슬롯이다
        :param sScrNo: 스크린번호
        :param sRQName: 내가 요청했을때 지은 이름
        :param sTrCode: 요청했던 tr코드
        :param sRecordName: 사용안함
        :param sPrevNext: 다음페이지가 있는지여부
        '''

        '''
         GetCommData() 
         : OnReceiveTRData()이벤트가 호출될 때 조회데이터를 얻어오는 함수(2차시그널)로, 반드시 OnReceiveTRData()이벤트가 호출될때 그 안에서 사용 
        '''

        if sRQName == '예수금상세현황요청':
        # 예수금확인하기2차시그널 GetCommData() :  TR코드(opw00001) / Request이름/ TR반복부 / TR에서 얻어오려는 출력항목이름
            deposit = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode , sRQName , 0 , "예수금" )
            self.deposit = int(deposit)
            print("- 예수금 %s" % self.deposit)
            output_deposit = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode , sRQName, 0 , "출금가능금액")
            self.output_deposit = int(output_deposit)
            print("- 출금가능액 %s" % self.output_deposit)

            #예수금중 배팅에 관한 부분
#            self.use_money = int(self.deposit) * self.use_money_percent
#            self.betting_money =  self.use_money * self.betting_size

            self.detail_account_info_event_loop.exit()








        elif sRQName == '계좌평가잔고내역요청':

        # 계좌평잔확인2차시그널 GetCommData() :  TR코드(opw00018) / Request이름/ TR반복부 / TR에서 얻어오려는 출력항목이름
            total_buy_money = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode , sRQName , 0 ,"총매입금액" )
            self.total_buy_money = int(total_buy_money)
            print("- 총매입금액 :%s" % self.total_buy_money)

            total_profit_loss_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0,"총평가손익금액")
            self.total_profit_loss_money = int(total_profit_loss_money)
            print("- 총평가손익액 :%s" % self.total_profit_loss_money)

            total_profit_loss_rate = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode , sRQName , 0 ,"총수익률(%)" )
            self.total_profit_loss_rate = float(total_profit_loss_rate)
            print("- 총수익률(%%) : %s" % self.total_profit_loss_rate)


            rows = self.dynamicCall("GetRepeatCnt(Qstring,Qstring)", sTrCode , sRQName) #보유종목 카운트하기
            for i in range(rows):
                code = self.dynamicCall("GetCommData(Qstring,Qstring,int,Qstring)" , sTrCode , sRQName , i ,
                                           "종목번호")
                code = code.strip()[1:]

                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                           "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                            "보유수량")  # 보유수량 : 000000000000010
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                             "매입가")  # 매입가 : 000000000054100
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "수익률(%)")  # 수익률 : -000000001.94
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                 "현재가")  # 현재가 : 000000003450
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,
                                                       i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                     "매매가능수량")

                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                if code in self.account_stock_dict:  # dictionary 에 해당 종목이 있나 확인 없으면 넣어주기
                    pass
                else:
                    self.account_stock_dict[code] = {}

                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({'매매가능수량': possible_quantity})

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()







        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                        "종목코드")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                           "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                            "주문번호") #중요한 값
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                "주문상태")  # 접수,확인,체결
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                  "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                               "주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                               "주문구분")  # -매도, +매수, -매도정정, +매수정정
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                "미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                               "체결량")

                code = code.strip()
                code_nm = code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-') #매수매도 기호지우기
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict: # dictionary 에 해당 종목이 있나 확인 없으면 넣어주기
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}


                self.not_account_stock_dict[order_no].update({'종목코드': code})
                self.not_account_stock_dict[order_no].update({'종목명': code_nm})
                self.not_account_stock_dict[order_no].update({'주문번호': order_no})
                self.not_account_stock_dict[order_no].update({'주문상태': order_status})
                self.not_account_stock_dict[order_no].update({'주문수량': order_quantity})
                self.not_account_stock_dict[order_no].update({'주문가격': order_price})
                self.not_account_stock_dict[order_no].update({'주문구분': order_gubun})
                self.not_account_stock_dict[order_no].update({'미체결수량': not_quantity})
                self.not_account_stock_dict[order_no].update({'체결량': ok_quantity})

                print("미체결종목 : %s" % self.not_account_stock_dict[order_no] )

            self.detail_account_info_event_loop.exit()











    def calculator_fnc(self):
        '''
        종목 분석관련 함수 모음
        :return:
        '''
        code_list = self.get_code_list_by_market("10")


        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock)  # 스크린 연결 끊기
            self.day_kiwoom_db(code=code)




    def read_code(self):
        print("포트폴리오오읽는중 ")
        if os.path.exists("files/condition_stock.txt"):
            f = open("files/condition_stock.txt", "r", encoding="utf8")  # "r"을 인자로 던져주면 파일 내용을 읽어 오겠다는 뜻이다.

            lines = f.readlines()  # 파일에 있는 내용들이 모두 읽어와 진다.
            for line in lines:  # 줄바꿈된 내용들이 한줄 씩 읽어와진다.
                if line != "":
                    ls = line.split("\t")

                    stock_code = ls[0]
                    stock_name = ls[1]
                    stock_size = float(ls[2].split("\n")[0])

                    self.portfolio_stock_dict.update({stock_code: {"종목명": stock_name, "포트폴리오사이즈": stock_size}})
                    # { "205405":{"종목명" : "삼성" , "최대포트폴리오사이즈" : 0.0785} ,"205405":{"종목명" : "LG" , "최대포트폴리오사이즈" : 0.0785}
            f.close()
        else:
           print("포트폴리오읽기끝남")


    def screen_number_setting(self):

        screen_overwrite = []
        # 계좌평가잔고내역에 있는 종목들
        for code in self.account_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)
        # 미체결에 있는 종목들
        for order_number in self.not_account_stock_dict.keys():
            code = self.not_account_stock_dict[order_number]['종목코드']
            if code not in screen_overwrite:
                screen_overwrite.append(code)
        # 포트폴리로에 담겨있는 종목들
        for code in self.portfolio_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)

        # 스크린번호 할당
        cnt = 0
        for code in screen_overwrite:

            temp_screen = int(self.screen_real_stock)
            meme_screen = int(self.screen_meme_stock)

            if (cnt % 50) == 0:
                temp_screen += 1
                self.screen_real_stock = str(temp_screen)      #스크린번호하나당 종목코드 50개씩만 할당하겠다

            if (cnt % 50) == 0:
                meme_screen += 1
                self.screen_meme_stock = str(meme_screen)      #스크린번호하나당 종목코드 50개씩만 할당하겠다

            if code in self.portfolio_stock_dict.keys():       #종목코드 있을경우
                self.portfolio_stock_dict[code].update({"스크린번호": str(self.screen_real_stock)})
                self.portfolio_stock_dict[code].update({"주문용스크린번호": str(self.screen_meme_stock)})

            elif code not in self.portfolio_stock_dict.keys(): #종목코드 없을 경우
                self.portfolio_stock_dict.update(
                    {code: {"스크린번호": str(self.screen_real_stock), "주문용스크린번호": str(self.screen_meme_stock)}})

            cnt += 1





















    # 실시간 데이터 얻어오기 종목코드 / 리얼타입한글로나오는부분 / 데이터전문
    def realdata_slot(self, sCode, sRealType, sRealData, stock_dict=None):

        if sRealType == "장시작시간":
            fid = self.realType.REALTYPE[sRealType]['장운영구분']  # (0:장시작전, 2:장종료전(20분), 3:장시작, 4,8:장종료(30분), 9:장마감)
            value = self.dynamicCall("GetCommRealData(QString, int)", sCode, fid) #종목코드와 Fid번호
            print("장 운영확인 시작")
            if value == '0':
                print("장 시작 전")
            elif value == '3':
                print("장 시작")
            elif value == "2":
                print("장 종료, 동시호가로 넘어감")
            elif value == "4":
                print("3시30분 장 종료")


        #틱 체결 데이터
        elif sRealType == "주식체결":
            a = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['체결시간'])  # 출력 HHMMSS
            b = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['현재가'])  # 출력 : +(-)2520
            b = abs(int(b))

            c = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['전일대비'])  # 출력 : +(-)2520
            c = abs(int(c))

            d = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['등락율'])  # 출력 : +(-)12.98
            d = float(d)

            e = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['(최우선)매도호가'])  # 출력 : +(-)2520
            e = abs(int(e))

            f = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['(최우선)매수호가'])  # 출력 : +(-)2515
            f = abs(int(f))

            g = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['거래량'])  # 출력 : +240124  매수일때, -2034 매도일 때
            g = abs(int(g))

            h = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['누적거래량'])  # 출력 : 240124
            h = abs(int(h))

            i = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['고가'])  # 출력 : +(-)2530
            i = abs(int(i))

            j = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['시가'])  # 출력 : +(-)2530
            j = abs(int(j))

            k = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['저가'])  # 출력 : +(-)2530
            k = abs(int(k))


            if sCode not in self.portfolio_stock_dict:
                self.portfolio_stock_dict.update({sCode: {}})
            self.portfolio_stock_dict[sCode].update({"체결시간": a})
            self.portfolio_stock_dict[sCode].update({"현재가": b})
            self.portfolio_stock_dict[sCode].update({"전일대비": c})
            self.portfolio_stock_dict[sCode].update({"등락율": d})
            self.portfolio_stock_dict[sCode].update({"(최우선)매도호가": e})
            self.portfolio_stock_dict[sCode].update({"(최우선)매수호가": f})
            self.portfolio_stock_dict[sCode].update({"거래량": g})
            self.portfolio_stock_dict[sCode].update({"누적거래량": h})
            self.portfolio_stock_dict[sCode].update({"고가": i})
            self.portfolio_stock_dict[sCode].update({"시가": j})
            self.portfolio_stock_dict[sCode].update({"저가": k})

            print(self.portfolio_stock_dict[sCode])