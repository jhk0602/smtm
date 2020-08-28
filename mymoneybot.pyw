# -*- coding: utf-8 -*-

import sys, os
import datetime, time

import pickle
import uuid
import base64
import subprocess
from subprocess import Popen

import PyQt5
from PyQt5 import QtCore, QtGui, uic
from PyQt5 import QAxContainer
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import (QApplication, QLabel, QLineEdit, QMainWindow, QDialog, QMessageBox, QProgressBar)
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *

import numpy as np
from numpy import NaN, Inf, arange, isscalar, asarray, array

import pandas as pd
import pandas.io.sql as pdsql
from pandas import DataFrame, Series
from pandas.lib import Timestamp

import talib as ta

import mysql.connector

import logging
import logging.handlers


로봇거래계좌번호 = None

주문딜레이 = 0.25
초당횟수제한 = 5

##
## 키움증권 제약사항 - 3.7초에 한번 읽으면 지금까지는 괜찮음
주문지연 = 3700


로봇스크린번호시작 = 9000
로봇스크린번호종료 = 9999


MySQL_POOL_SIZE = 2

데이타베이스_설정값 = {
    'host': '127.0.0.1',
    'user': 'admin',
    'password': 'ahnlab',
    'database': 'moneybot',
    'raise_on_warnings': True,
}

class NumpyMySQLConverter(mysql.connector.conversion.MySQLConverter):
    """ A mysql.connector Converter that handles Numpy types """

    def _float32_to_mysql(self, value):
        return float(value)

    def _float64_to_mysql(self, value):
        return float(value)

    def _int32_to_mysql(self, value):
        return int(value)

    def _int64_to_mysql(self, value):
        return int(value)

    def _timestamp_to_mysql(self, value):
        return value.to_datetime()

def mysqlconn():
    conn = mysql.connector.connect(pool_name="stockpool", pool_size=MySQL_POOL_SIZE, **데이타베이스_설정값)
    conn.set_converter_class(NumpyMySQLConverter)
    return conn

class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, data=None, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._data = data
        if data is None:
            self._data = DataFrame()

    def rowCount(self, parent=None):
        # return len(self._data.values)
        return len(self._data.index)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                # return QtCore.QVariant(str(self._data.values[index.row()][index.column()]))
                return str(self._data.values[index.row()][index.column()])
        # return QtCore.QVariant()
        return None

    def headerData(self, column, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self._data.columns[column]
        return int(column + 1)

    def update(self, data):
        self._data = data
        self.reset()

    def reset(self):
        self.beginResetModel()
        # unnecessary call to actually clear data, but recommended by design guidance from Qt docs
        # left blank in preliminary testing
        self.endResetModel()

    def flags(self, index):
        return QtCore.Qt.ItemIsEnabled

##
## 포트폴리오에 사용되는 주식정보 클래스
class CPortStock(object):
    def __init__(self, 매수일, 종목코드, 종목명, 매수가, 매도가1차, 매도가2차, 손절가, 수량, 매수단위수=1, STATUS=''):
        self.매수일 = 매수일
        self.종목코드 = 종목코드
        self.종목명 = 종목명
        self.매수가 = 매수가
        self.매도가1차 = 매도가1차
        self.매도가2차 = 매도가2차
        self.손절가 = 손절가
        self.수량 = 수량
        self.매수단위수 = 매수단위수
        self.STATUS = STATUS

        self.이전매수일 = 매수일
        self.이전매수가 = 0
        self.이전수량 = 0
        self.이전매수단위수 = 0

    def 평균단가(self):
        if self.이전매수단위수 > 0:
            return ((self.매수가 * self.수량) + (self.이전매수가 * self.이전수량)) // (self.수량 + self.이전수량)
        else:
            return self.매수가

##
## CTrade 거래로봇용 베이스클래스
class CTrade(object):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        """
        :param sName: 로봇이름
        :param UUID: 로봇구분용 id
        :param kiwoom: 키움OpenAPI
        :param parent: 나를 부른 부모 - 보통은 메인윈도우
        """
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None # 거래용계좌번호
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False # 실행상태

        self.portfolio = dict() # 포트폴리오 관리 {'종목코드':종목정보}
        self.현재가 = dict() # 각 종목의 현재가

    def GetStatus(self):
        """
        :return: 포트폴리오의 상태
        """
        result = []
        for p, v in self.portfolio.items():
            result.append('%s(%s)[P%s/V%s/D%s]' % (v.종목명.strip(), v.종목코드, v.매수가, v.수량, v.매수일))

        return [self.__class__.__name__, self.sName, self.UUID, self.sScreenNo, self.running, len(self.portfolio), ','.join(result)]

    def GenScreenNO(self):
        """
        :return: 키움증권에서 요구하는 스크린번호를 생성
        """
        self.SmallScreenNumber += 1
        if self.SmallScreenNumber > 9999:
            self.SmallScreenNumber = 0

        return self.sScreenNo * 10000 + self.SmallScreenNumber

    def GetLoginInfo(self, tag):
        """
        :param tag:
        :return: 로그인정보 호출
        """
        return self.kiwoom.dynamicCall('GetLoginInfo("%s")' % tag)

    def KiwoomConnect(self):
        """
        :return: 키움증권OpenAPI의 CallBack에 대응하는 처리함수를 연결
        """
        self.kiwoom.OnEventConnect[int].connect(self.OnEventConnect)
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].connect(self.OnReceiveChejanData)
        self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].connect(self.OnReceiveRealData)
        # logger.info("%s : connected" % self.sName)

    def KiwoomDisConnect(self):
        """
        :return: Callback 연결해제
        """
        try:
            self.kiwoom.OnEventConnect[int].disconnect(self.OnEventConnect)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].disconnect(self.OnReceiveTrCondition)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveChejanData[str, int, str].disconnect(self.OnReceiveChejanData)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveConditionVer[int, str].disconnect(self.OnReceiveConditionVer)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveRealCondition[str, str, str, str].disconnect(self.OnReceiveRealCondition)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveRealData[str, str, str].disconnect(self.OnReceiveRealData)
        except Exception:
            pass
        # logger.info("%s : disconnected" % self.sName)

    def KiwoomAccount(self):
        """
        :return: 계좌정보를 읽어옴
        """
        ACCOUNT_CNT = self.GetLoginInfo('ACCOUNT_CNT')
        ACC_NO = self.GetLoginInfo('ACCNO')

        self.account = ACC_NO.split(';')[0:-1]
        # logger.debug("보유 계좌수: %s 계좌번호: %s [%s]" % (ACCOUNT_CNT, self.account[0], ACC_NO))

    def KiwoomSendOrder(self, sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo):
        """
        OpenAPI 메뉴얼 참조
        :param sRQName:
        :param sScreenNo:
        :param sAccNo:
        :param nOrderType:
        :param sCode:
        :param nQty:
        :param nPrice:
        :param sHogaGb:
        :param sOrgOrderNo:
        :return:
        """
        order = self.kiwoom.dynamicCall('SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                                        [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo])
        return order

        # -거래구분값 확인(2자리)
        #
        # 00 : 지정가
        # 03 : 시장가
        # 05 : 조건부지정가
        # 06 : 최유리지정가
        # 07 : 최우선지정가
        # 10 : 지정가IOC
        # 13 : 시장가IOC
        # 16 : 최유리IOC
        # 20 : 지정가FOK
        # 23 : 시장가FOK
        # 26 : 최유리FOK
        # 61 : 장전 시간외단일가매매
        # 81 : 장후 시간외종가
        # 62 : 시간외단일가매매
        #
        # -매매구분값 (1 자리)
        # 1 : 신규매수
        # 2 : 신규매도
        # 3 : 매수취소
        # 4 : 매도취소
        # 5 : 매수정정
        # 6 : 매도정정

    def KiwoomSetRealReg(self, sScreenNo, sCode, sRealType='0'):
        """
        OpenAPI 메뉴얼 참조
        :param sScreenNo:
        :param sCode:
        :param sRealType:
        :return:
        """
        ret = self.kiwoom.dynamicCall('SetRealReg(QString, QString, QString, QString)', sScreenNo, sCode, '9001;10', sRealType)
        return ret

    def KiwoomSetRealRemove(self, sScreenNo, sCode):
        """
        OpenAPI 메뉴얼 참조
        :param sScreenNo:
        :param sCode:
        :return:
        """
        ret = self.kiwoom.dynamicCall('SetRealRemove(QString, QString)', sScreenNo, sCode)
        return ret

    def OnEventConnect(self, nErrCode):
        """
        OpenAPI 메뉴얼 참조
        :param nErrCode:
        :return:
        """
        logger.info('OnEventConnect', nErrCode)

    def OnReceiveMsg(self, sScrNo, sRQName, sTRCode, sMsg):
        """
        OpenAPI 메뉴얼 참조
        :param sScrNo:
        :param sRQName:
        :param sTRCode:
        :param sMsg:
        :return:
        """
        logger.info('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTRCode, sMsg))

    def OnReceiveTrCondition(self, sScrNo, strCodeList, strConditionName, nIndex, nNext):
        """
        OpenAPI 메뉴얼 참조
        :param sScrNo:
        :param strCodeList:
        :param strConditionName:
        :param nIndex:
        :param nNext:
        :return:
        """
        logger.info('OnReceiveTrCondition [%s] [%s] [%s] [%s] [%s]' % (sScrNo, strCodeList, strConditionName, nIndex, nNext))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        """
        OpenAPI 메뉴얼 참조
        :param sScrNo:
        :param sRQName:
        :param sTRCode:
        :param sRecordName:
        :param sPreNext:
        :param nDataLength:
        :param sErrorCode:
        :param sMessage:
        :param sSPlmMsg:
        :return:
        """
        # logger.info('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))

        if self.sScreenNo != int(sScrNo[:4]):
            return

        # logger.info('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))

        if 'B_' in sRQName or 'S_' in sRQName:
            주문번호 = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "주문번호")
            # logger.debug("화면번호: %s sRQName : %s 주문번호: %s" % (sScrNo, sRQName, 주문번호))

            self.주문등록(sRQName, 주문번호)

    def OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        """
        OpenAPI 메뉴얼 참조
        :param sGubun:
        :param nItemCnt:
        :param sFidList:
        :return:
        """
        # logger.info('OnReceiveChejanData [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))

        # 주문체결시 순서
        # 1 구분:0 GetChejanData(913) = '접수'
        # 2 구분:0 GetChejanData(913) = '체결'
        # 3 구분:1 잔고정보

        if sGubun == "0":
            화면번호 = self.kiwoom.dynamicCall('GetChejanData(QString)', 920)
            if self.sScreenNo != int(화면번호[:4]):
                return

            param = dict()

            param['sGubun'] = sGubun
            param['계좌번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9201)
            param['주문번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9203)
            param['종목코드'] = self.종목코드변환(self.kiwoom.dynamicCall('GetChejanData(QString)', 9001))

            param['주문업무분류'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 912)

            # 접수 / 체결 확인
            # 주문상태(10:원주문, 11:정정주문, 12:취소주문, 20:주문확인, 21:정정확인, 22:취소확인, 90-92:주문거부)
            param['주문상태'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 913) # 접수 or 체결 확인

            param['종목명'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 302).strip()
            param['주문수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 900)
            param['주문가격'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 901)
            param['미체결수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 902)
            param['체결누계금액'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 903)
            param['원주문번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 904)
            param['주문구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 905)
            param['매매구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 906)
            param['매도수구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 907)
            param['체결시간'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 908)
            param['체결번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 909)
            param['체결가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 910)
            param['체결량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 911)

            param['현재가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 10)
            param['매도호가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 27)
            param['매수호가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 28)

            param['단위체결가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 914).strip()
            param['단위체결량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 915)
            param['화면번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 920)

            param['당일매매수수료'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 938)
            param['당일매매세금'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 939)

            param['체결수량'] = int(param['주문수량']) - int(param['미체결수량'])

            # logger.debug('계좌번호:{계좌번호} 체결시간:{체결시간} 주문번호:{주문번호} 체결번호:{체결번호} 종목코드:{종목코드} 종목명:{종목명} 체결량:{체결량} 체결가:{체결가} 단위체결가:{단위체결가} 주문수량:{주문수량} 체결수량:{체결수량} 미체결수량:{미체결수량}'.format(**param))

            if param["주문상태"] == "접수":
                self.접수처리(param)
            if param["주문상태"] == "체결":
                self.체결처리(param)

        if sGubun == "1":
            # logger.debug('OnReceiveChejanData: 잔고통보 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
            param = dict()

            param['sGubun'] = sGubun
            param['계좌번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9201)
            param['종목코드'] = self.종목코드변환(self.kiwoom.dynamicCall('GetChejanData(QString)', 9001))

            param['신용구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 917)
            param['대출일'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 916)

            param['종목명'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 302).strip()
            param['현재가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 10)

            param['보유수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 930)
            param['매입단가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 931)
            param['총매입가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 932)
            param['주문가능수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 933)
            param['당일순매수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 945)
            param['매도매수구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 946)
            param['당일총매도손익'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 950)
            param['예수금'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 951)

            param['매도호가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 27)
            param['매수호가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 28)

            param['기준가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 307)
            param['손익율'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 8019)
            param['신용금액'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 957)
            param['신용이자'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 958)
            param['만기일'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 918)
            param['당일실현손익_유가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 990)
            param['당일실현손익률_유가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 991)
            param['당일실현손익_신용'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 992)
            param['당일실현손익률_신용'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 993)
            param['담보대출수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 959)

            # logger.debug('계좌번호:{계좌번호} 종목명:{종목명} 보유수량:{보유수량} 매입단가:{매입단가} 당일순매수량:{당일순매수량}'.format(**param))

            self.잔고처리(param)

        if sGubun == "3":
            # logger.debug('OnReceiveChejanData: 특이신호 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
            pass

        # sFid별 주요데이터는 다음과 같습니다.
        # "9201" : "계좌번호"
        # "9203" : "주문번호"
        # "9001" : "종목코드"
        # "913" : "주문상태"
        # "302" : "종목명"
        # "900" : "주문수량"
        # "901" : "주문가격"
        # "902" : "미체결수량"
        # "903" : "체결누계금액"
        # "904" : "원주문번호"
        # "905" : "주문구분"
        # "906" : "매매구분"
        # "907" : "매도수구분"
        # "908" : "주문/체결시간"
        # "909" : "체결번호"
        # "910" : "체결가"
        # "911" : "체결량"
        # "10" : "현재가"
        # "27" : "(최우선)매도호가"
        # "28" : "(최우선)매수호가"
        # "914" : "단위체결가"
        # "915" : "단위체결량"
        # "919" : "거부사유"
        # "920" : "화면번호"
        # "917" : "신용구분"
        # "916" : "대출일"
        # "930" : "보유수량"
        # "931" : "매입단가"
        # "932" : "총매입가"
        # "933" : "주문가능수량"
        # "945" : "당일순매수수량"
        # "946" : "매도/매수구분"
        # "950" : "당일총매도손일"
        # "951" : "예수금"
        # "307" : "기준가"
        # "8019" : "손익율"
        # "957" : "신용금액"
        # "958" : "신용이자"
        # "918" : "만기일"
        # "990" : "당일실현손익(유가)"
        # "991" : "당일실현손익률(유가)"
        # "992" : "당일실현손익(신용)"
        # "993" : "당일실현손익률(신용)"
        # "397" : "파생상품거래단위"
        # "305" : "상한가"
        # "306" : "하한가"

    def OnReceiveConditionVer(self, lRet, sMsg):
        """
        OpenAPI 메뉴얼 참조
        :param lRet:
        :param sMsg:
        :return:
        """
        logger.info('OnReceiveConditionVer : [이벤트] 조건식 저장', lRet, sMsg)

    def OnReceiveRealCondition(self, sTrCode, strType, strConditionName, strConditionIndex):
        """
        OpenAPI 메뉴얼 참조
        :param sTrCode:
        :param strType:
        :param strConditionName:
        :param strConditionIndex:
        :return:
        """
        logger.info('OnReceiveRealCondition [%s] [%s] [%s] [%s]' % (sTrCode, strType, strConditionName, strConditionIndex))

    def OnReceiveRealData(self, sRealKey, sRealType, sRealData):
        """
        OpenAPI 메뉴얼 참조
        :param sRealKey:
        :param sRealType:
        :param sRealData:
        :return:
        """
        # logger.info('OnReceiveRealData [%s] [%s] [%s]' % (sRealKey, sRealType, sRealData))

        _now = datetime.datetime.now()
        if _now.strftime('%H:%M:%S') < '09:00:00':
            return

        if sRealKey not in self.실시간종목리스트:
            return

        if sRealType == "주식시세" or sRealType == "주식체결":
            param = dict()

            param['종목코드'] = self.종목코드변환(sRealKey)
            param['체결시간'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 20).strip()
            param['현재가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 10).strip()
            param['전일대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 11).strip()
            param['등락률'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 12).strip()
            param['매도호가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 27).strip()
            param['매수호가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 28).strip()
            param['누적거래량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 13).strip()
            param['시가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 16).strip()
            param['고가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 17).strip()
            param['저가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 18).strip()
            param['거래회전율'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 31).strip()
            param['시가총액'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 311).strip()

            self.실시간데이타처리(param)

    def 종목코드변환(self, code):
        return code.replace('A', '')

    def 정량매수(self, sRQName, 종목코드, 매수가, 수량):
        # sRQName = '정량매수%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO()
        sAccNo = self.sAccount
        nOrderType = 1  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 수량
        nPrice = 매수가
        sHogaGb = self.매수방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03','07','06']:
            nPrice = 0
        sOrgOrderNo = 0

        ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)

        return ret

    def 정액매수(self, sRQName, 종목코드, 매수가, 매수금액):
        # sRQName = '정액매수%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO()
        sAccNo = self.sAccount
        nOrderType = 1  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 매수금액 // 매수가
        nPrice = 매수가
        sHogaGb = self.매수방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03','07','06']:
            nPrice = 0
        sOrgOrderNo = 0

        # logger.debug('주문 - %s %s %s %s %s %s %s %s %s', sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)
        ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)

        return ret

    def 정량매도(self, sRQName, 종목코드, 매도가, 수량):
        # sRQName = '정량매도%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO()
        sAccNo = self.sAccount
        nOrderType = 2  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 수량
        nPrice = 매도가
        sHogaGb = self.매도방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03','07','06']:
            nPrice = 0
        sOrgOrderNo = 0

        ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)

        return ret

    def 정액매도(self, sRQName, 종목코드, 매도가, 매도금액):
        # sRQName = '정액매도%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO()
        sAccNo = self.sAccount
        nOrderType = 2  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 매도금액 // 매도가
        nPrice = 매도가
        sHogaGb = self.매도방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03','07','06']:
            nPrice = 0
        sOrgOrderNo = 0

        ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)

        return ret

    def 주문등록(self, sRQName, 주문번호):
        self.주문번호_주문_매핑[주문번호] = sRQName

    def 초기조건(self):
        pass


Ui_계좌정보조회, QtBaseClass_계좌정보조회 = uic.loadUiType("계좌정보조회.ui")

class 화면_계좌정보(QDialog, Ui_계좌정보조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_계좌정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['종목번호', '종목명', '현재가', '보유수량', '매입가', '매입금액', '평가금액', '수익률(%)', '평가손익', '매매가능수량']
        self.보이는컬럼 = ['종목번호', '종목명', '현재가', '보유수량', '매입가', '매입금액', '평가금액', '주당손익', '평가손익', '매매가능수량']

        self.result = []

        self.KiwoomAccount()

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def KiwoomAccount(self):
        ACCOUNT_CNT = self.kiwoom.dynamicCall('GetLoginInfo("ACCOUNT_CNT")')
        ACC_NO = self.kiwoom.dynamicCall('GetLoginInfo("ACCNO")')

        self.account = ACC_NO.split(';')[0:-1]

        self.comboBox.clear()
        self.comboBox.addItems(self.account)

        logger.debug("보유 계좌수: %s 계좌번호: %s [%s]" % (ACCOUNT_CNT, self.account[0], ACC_NO))

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))

        if sRQName == "계좌평가잔고내역요청":
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-'+S[1:].lstrip('0')
                    row.append( S )
                self.result.append(row)
                # logger.debug("%s" % row)
            if sPreNext == '2':
                self.Request(_repeat=2)
            else:
                self.model.update(DataFrame(data=self.result, columns=self.보이는컬럼))
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        계좌번호 = self.comboBox.currentText().strip()
        logger.debug("계좌번호 %s" % 계좌번호)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", 계좌번호)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "비밀번호입력매체구분", '00')
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "조회구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "계좌평가잔고내역요청", "opw00018", _repeat, '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)

    def robot_account(self):
        global 로봇거래계좌번호

        로봇거래계좌번호 = self.comboBox.currentText().strip()

        conn = mysqlconn()

        cursor = conn.cursor()
        robot_account = pickle.dumps(로봇거래계좌번호, protocol=pickle.HIGHEST_PROTOCOL, fix_imports=True)
        _robot_account = base64.encodebytes(robot_account)
        cursor.execute("REPLACE into mymoneybot_setting(keyword, value) values (%s, %s)", ['robotaccount', _robot_account])
        conn.commit()
        conn.close()


Ui_일별가격정보백업, QtBaseClass_일별가격정보백업 = uic.loadUiType("일별가격정보백업.ui")

class 화면_일별가격정보백업(QDialog, Ui_일별가격정보백업):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_일별가격정보백업, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('가격 정보 백업')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.columns = ['일자', '현재가', '거래량', '시가', '고가', '저가','거래대금']

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

        self.종목코드테이블 = self.parent.StockCodeBuild().copy()
        self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['종목코드'] + " : " + self.종목코드테이블['시장구분'] + ' - ' + self.종목코드테이블['종목명']
        self.종목코드테이블 = self.종목코드테이블.sort_values(['종목코드', '종목명'], ascending=[True, True])
        self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "주식일봉차트조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-'+S[1:].lstrip('0')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2' and self.radioButton_all.isChecked() == True:
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['일자'] = df['일자'].apply(lambda x: x[0:4] + '-'  + x[4:6] + '-' +x[6:])
                df['종목코드'] = self.종목코드[0]
                df = df[['종목코드','일자','현재가','시가','고가','저가','거래량','거래대금']]
                values = list(df.values)

                try:
                    df.ix[df.현재가 == '', ['현재가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.시가 == '', ['시가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.고가 == '', ['고가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.저가 == '', ['저가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래량 == '', ['거래량']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래대금 == '', ['거래대금']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래대금 == '-', ['거래대금']] = 0
                except Exception as e:
                    pass

                conn = mysqlconn()

                cursor = conn.cursor()
                cursor.executemany("replace into 일별주가(종목코드,일자,종가,시가,고가,저가,거래량,거래대금) values( %s, %s, %s, %s, %s, %s, %s, %s )", df.values.tolist())

                conn.commit()
                conn.close()

                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.result = []

                    self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
                    S = '%s %s' % (self.종목코드[0], self.종목코드[1])
                    self.label_codename.setText(S)

                    QTimer.singleShot(주문지연, lambda : self.Request(_repeat=0))
                else:
                    QMessageBox.about(self, "백업완료","백업을 완료하였습니다..")

    def Request(self, _repeat=0):
        logger.info('%s %s' % (self.종목코드[0], self.종목코드[1]))
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "기준일자", self.기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식일봉차트조회", "OPT10081", _repeat, '{:04d}'.format(self.sScreenNo))

    def Backup_One(self):
        idx = self.comboBox.currentIndex()

        self.백업한종목수 = 1
        self.백업할종목코드 = []
        self.종목코드 = self.종목코드테이블[idx:idx + 1][['종목코드','종목명']].values[0]
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')
        self.result = []
        self.Request(_repeat=0)

    def Backup_All(self):
        idx = self.comboBox.currentIndex()
        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[idx:][['종목코드','종목명']].values)
        self.종목코드 = self.백업할종목코드.pop(0)
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')

        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
        S = '%s %s' % (self.종목코드[0], self.종목코드[1])
        self.label_codename.setText(S)

        self.result = []
        self.Request(_repeat=0)


Ui_일별업종정보백업, QtBaseClass_일별업종정보백업 = uic.loadUiType("일별업종정보백업.ui")

class 화면_일별업종정보백업(QDialog, Ui_일별업종정보백업):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_일별업종정보백업, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('업종 정보 백업')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.columns = ['현재가', '거래량', '일자', '시가', '고가', '저가','거래대금', '대업종구분', '소업종구분', '종목정보', '종목정보', '수정주가이벤트', '전일종가']

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

        self.업종코드테이블 = self.parent.SectorCodeBuild().copy()
        self.업종코드테이블['컬럼'] = ">> " + self.업종코드테이블['업종코드'] + " : " + self.업종코드테이블['시장구분'] + ' - ' + self.업종코드테이블['업종명']
        self.comboBox.addItems(self.업종코드테이블['컬럼'].values)

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "업종일봉조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-'+S[1:].lstrip('0')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2' and self.radioButton_all.isChecked() == True:
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['일자'] = df['일자'].apply(lambda x: x[0:4] + '-'  + x[4:6] + '-' +x[6:])
                df['업종코드'] = self.업종코드[0]
                df = df[['업종코드','일자','현재가','시가','고가','저가','거래량','거래대금']]
                values = list(df.values)
                # print(values)

                try:
                    df.ix[df.종가 == '', ['종가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.시가 == '', ['시가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.고가 == '', ['고가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.저가 == '', ['저가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래량 == '', ['거래량']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래대금 == '', ['거래대금']] = 0
                except Exception as e:
                    pass

                conn = mysqlconn()

                cursor = conn.cursor()
                cursor.executemany("replace into 일별업종지수(업종코드,일자,종가,시가,고가,저가,거래량,거래대금) values( %s, %s, %s, %s, %s, %s, %s, %s )", df.values.tolist())

                conn.commit()
                conn.close()

                self.백업한종목수 += 1
                if len(self.백업할업종코드) > 0:
                    self.업종코드 = self.백업할업종코드.pop(0)
                    self.result = []

                    self.progressBar.setValue(int(self.백업한종목수 / (len(self.업종코드테이블.index) - self.comboBox.currentIndex()) * 100))
                    S = '%s %s' % (self.업종코드[0], self.업종코드[1])
                    self.label_codename.setText(S)

                    QTimer.singleShot(주문지연, lambda : self.Request(_repeat=0))
                else:
                    QMessageBox.about(self, "백업완료","백업을 완료하였습니다..")

    def Request(self, _repeat=0):
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "업종코드", self.업종코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "기준일자", self.기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "업종일봉조회", "OPT20006", _repeat, '{:04d}'.format(self.sScreenNo))

    def Backup_One(self):
        idx = self.comboBox.currentIndex()

        self.백업한종목수 = 1
        self.백업할업종코드 = []
        self.업종코드 = self.업종코드테이블[idx:idx + 1][['업종코드','업종명']].values[0]
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')
        self.result = []
        self.Request(_repeat=0)

    def Backup_All(self):
        idx = self.comboBox.currentIndex()
        self.백업한종목수 = 1
        self.백업할업종코드 = list(self.업종코드테이블[idx:][['업종코드','업종명']].values)
        self.업종코드 = self.백업할업종코드.pop(0)
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')

        self.progressBar.setValue(int(self.백업한종목수 / (len(self.업종코드테이블.index) - self.comboBox.currentIndex()) * 100))
        S = '%s %s' % (self.업종코드[0], self.업종코드[1])
        self.label_codename.setText(S)

        self.result = []
        self.Request(_repeat=0)


Ui_분별가격정보백업, QtBaseClass_분별가격정보백업 = uic.loadUiType("분별가격정보백업.ui")

class 화면_분별가격정보백업(QDialog, Ui_분별가격정보백업):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_분별가격정보백업, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('가격 정보 백업')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.columns = ['체결시간', '현재가', '시가', '고가', '저가', '거래량']

        self.result = []

        self.종목코드테이블 = self.parent.StockCodeBuild().copy()
        self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['종목코드'] + " : "  + self.종목코드테이블['시장구분'] + ' - ' +  self.종목코드테이블['종목명']
        self.종목코드테이블 = self.종목코드테이블.sort_values(['종목코드', '종목명'], ascending=[True, True])
        self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "주식분봉차트조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and (S[0] == '-' or S[0] == '+') :
                        S = S[1:].lstrip('0')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2' and self.radioButton_all.isChecked() == True:
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['체결시간'] = df['체결시간'].apply(lambda x: x[0:4]+'-'+x[4:6]+'-'+x[6:8]+' '+x[8:10]+':'+x[10:12]+':'+x[12:])
                df['종목코드'] = self.종목코드[0]
                df['틱범위'] = self.틱범위
                df = df[['종목코드','틱범위','체결시간','현재가','시가','고가','저가','거래량']]
                values = list(df.values)

                try:
                    df.ix[df.현재가 == '', ['현재가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.시가 == '', ['시가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.고가 == '', ['고가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.저가 == '', ['저가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래량 == '', ['거래량']] = 0
                except Exception as e:
                    pass

                conn = mysqlconn()

                cursor = conn.cursor()
                cursor.executemany("replace into 분별주가(종목코드,틱범위,체결시간,종가,시가,고가,저가,거래량) values( %s, %s, %s, %s, %s, %s, %s, %s )", df.values.tolist())

                conn.commit()
                conn.close()
                    
                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.result = []

                    self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
                    S = '%s %s' % (self.종목코드[0], self.종목코드[1])
                    self.label_codename.setText(S)

                    QTimer.singleShot(주문지연, lambda : self.Request(_repeat=0))
                else:
                    QMessageBox.about(self, "백업완료","백업을 완료하였습니다..")

    def Request(self, _repeat=0):
        logger.info('%s %s' % (self.종목코드[0], self.종목코드[1]))
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "틱범위", self.틱범위)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식분봉차트조회", "OPT10080", _repeat, '{:04d}'.format(self.sScreenNo))

    def Backup_One(self):
        idx = self.comboBox.currentIndex()

        self.백업한종목수 = 1
        self.백업할종목코드 = []
        self.종목코드 = self.종목코드테이블[idx:idx + 1][['종목코드','종목명']].values[0]
        self.틱범위 = self.comboBox_min.currentText()[0:3].strip()
        if self.틱범위[0] == '0':
            self.틱범위 = self.틱범위[1:]
        self.result = []
        self.Request(_repeat=0)

    def Backup_All(self):
        idx = self.comboBox.currentIndex()
        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[idx:][['종목코드','종목명']].values)
        self.종목코드 = self.백업할종목코드.pop(0)
        self.틱범위 = self.comboBox_min.currentText()[0:3].strip()
        if self.틱범위[0] == '0':
            self.틱범위 = self.틱범위[1:]
        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
        S = '%s %s' % (self.종목코드[0], self.종목코드[1])
        self.label_codename.setText(S)

        self.result = []
        self.Request(_repeat=0)


Ui_종목별투자자정보백업, QtBaseClass_종목별투자자정보백업 = uic.loadUiType("종목별투자자정보백업.ui")

class 화면_종목별투자자정보백업(QDialog, Ui_종목별투자자정보백업):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_종목별투자자정보백업, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('종목별 투자자 정보 백업')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.columns = ['일자', '현재가', '전일대비', '누적거래대금', '개인투자자', '외국인투자자','기관계','금융투자','보험','투신','기타금융','은행','연기금등','국가','내외국인','사모펀드','기타법인']

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

        self.종목코드테이블 = self.parent.StockCodeBuild().copy()
        self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['종목코드'] + " : "  + self.종목코드테이블['시장구분'] + ' - ' +  self.종목코드테이블['종목명']
        self.종목코드테이블 = self.종목코드테이블.sort_values(['종목코드', '종목명'], ascending=[True, True])
        self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "종목별투자자조회":
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0').replace('--','-')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2' and self.radioButton_all.isChecked() == True:
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                if len(self.result) > 0:
                    df = DataFrame(data=self.result, columns=self.columns)
                    df['일자'] = df['일자'].apply(lambda x: x[0:4] + '-'  + x[4:6] + '-' +x[6:])
                    df['현재가'] = np.abs(pd.to_numeric(df['현재가'], errors='coerce'))
                    df['종목코드'] = self.종목코드[0]
                    df = df[['종목코드']+self.columns]
                    # values = list(df.values)

                    try:
                        df.ix[df.현재가 == '', ['현재가']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.전일대비 == '', ['전일대비']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.누적거래대금 == '', ['누적거래대금']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.개인투자자 == '', ['개인투자자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.외국인투자자 == '', ['외국인투자자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기관계 == '', ['기관계']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.금융투자 == '', ['금융투자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.금융투자 == '', ['금융투자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.보험 == '', ['보험']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.투신 == '', ['투신']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기타금융 == '', ['기타금융']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.은행 == '', ['은행']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.연기금등 == '', ['연기금등']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.국가 == '', ['국가']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.내외국인 == '', ['내외국인']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.사모펀드 == '', ['사모펀드']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기타법인 == '', ['기타법인']] = 0
                    except Exception as e:
                        pass

                    df.dropna(inplace=True)

                    conn = mysqlconn()

                    cursor = conn.cursor()
                    cursor.executemany("replace into 종목별투자자(종목코드,일자,종가,전일대비,누적거래대금,개인투자자,외국인투자자,기관계,금융투자,보험,투신,기타금융,은행,연기금등,국가,내외국인,사모펀드,기타법인) values(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", df.values.tolist())
                    conn.commit()
                    conn.close()

                else:
                    logger.info("%s 데이타없음", self.종목코드)

                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.result = []

                    self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
                    S = '%s %s' % (self.종목코드[0], self.종목코드[1])
                    self.label_codename.setText(S)

                    QTimer.singleShot(주문지연, lambda : self.Request(_repeat=0))
                else:
                    QMessageBox.about(self, "백업완료","백업을 완료하였습니다..")

    def Request(self, _repeat=0):
        logger.info('%s %s' % (self.종목코드[0], self.종목코드[1]))
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "일자", self.기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "금액수량구분", 2) #1:금액, 2:수량
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "매매구분", 0) #0:순매수, 1:매수, 2:매도
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "단위구분", 1) #1000:천주, 1:단주
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "종목별투자자조회", "OPT10060", _repeat, '{:04d}'.format(self.sScreenNo))

    def Backup_One(self):
        idx = self.comboBox.currentIndex()

        self.백업한종목수 = 1
        self.백업할종목코드 = []
        self.종목코드 = self.종목코드테이블[idx:idx + 1][['종목코드','종목명']].values[0]
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')
        self.result = []
        self.Request(_repeat=0)

    def Backup_All(self):
        idx = self.comboBox.currentIndex()
        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[idx:][['종목코드','종목명']].values)
        self.종목코드 = self.백업할종목코드.pop(0)
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')

        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
        S = '%s %s' % (self.종목코드[0], self.종목코드[1])
        self.label_codename.setText(S)

        self.result = []
        self.Request(_repeat=0)


Ui_업종정보, QtBaseClass_업종정보 = uic.loadUiType("업종정보조회.ui")

class 화면_업종정보(QDialog, Ui_업종정보):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_업종정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('업종정보 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['종목코드', '종목명', '현재가', '대비기호', '전일대비', '등락률','거래량','비중','거래대금','상한','상승','보합','하락','하한','상장종목수']

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "업종정보조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-'+S[1:].lstrip('0')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['업종코드'] = self.업종코드
                df.to_csv("업종정보.csv")
                self.model.update(df[['업종코드']+self.columns])
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        self.업종코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-','')

        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "업종코드", self.업종코드)
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "업종정보조회", "OPT20003", _repeat, '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)


Ui_업종별주가조회, QtBaseClass_업종별주가조회 = uic.loadUiType("업종별주가조회.ui")

class 화면_업종별주가(QDialog, Ui_업종별주가조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_업종별주가, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('업종별 주가 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['현재가', '거래량', '일자', '시가', '고가', '저가', '거래대금', '대업종구분','소업종구분','종목정보','수정주가이벤트','전일종가']

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "업종일봉조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-'+S[1:].lstrip('0')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['업종코드'] = self.업종코드
                self.model.update(df[['업종코드']+self.columns])
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        self.업종코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-','')

        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "업종코드", self.업종코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "기준일자", 기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "업종일봉조회", "OPT20006", _repeat, '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)


Ui_일자별주가조회, QtBaseClass_일자별주가조회 = uic.loadUiType("일자별주가조회.ui")

class 화면_일별주가(QDialog, Ui_일자별주가조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_일별주가, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('일자별 주가 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['일자', '현재가', '거래량', '시가', '고가', '저가','거래대금']

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "주식일봉차트조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-'+S[1:].lstrip('0')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['종목코드'] = self.종목코드
                self.model.update(df[['종목코드']+self.columns])
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        self.종목코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-','')

        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "기준일자", 기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식일봉차트조회", "OPT10081", _repeat, '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)


class 화면_종목별투자자(QDialog, Ui_일자별주가조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_종목별투자자, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('종목별 투자자 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['일자', '현재가', '전일대비', '누적거래대금', '개인투자자', '외국인투자자','기관계','금융투자','보험','투신','기타금융','은행','연기금등','국가','내외국인','사모펀드','기타법인']

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "종목별투자자조회":
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['종목코드'] = self.lineEdit_code.text().strip()
                dfnew = df[['종목코드']+self.columns]
                self.model.update(dfnew)
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        종목코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-','')

        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "일자", 기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", 종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "금액수량구분", 2) #1:금액, 2:수량
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "매매구분", 0) #0:순매수, 1:매수, 2:매도
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "단위구분", 1) #1000:천주, 1:단주
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "종목별투자자조회", "OPT10060", _repeat, '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)


Ui_분별주가조회, QtBaseClass_분별주가조회 = uic.loadUiType("분별주가조회.ui")

class 화면_분별주가(QDialog, Ui_분별주가조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_분별주가, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('분별 주가 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['체결시간', '현재가', '시가', '고가', '저가', '거래량']

        self.result = []

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "주식분봉차트조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and (S[0] == '-' or S[0] == '+') :
                        S = S[1:].lstrip('0')
                    row.append( S )
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda : self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['종목코드'] = self.종목코드
                self.model.update(df[['종목코드']+self.columns])
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        self.종목코드 = self.lineEdit_code.text().strip()
        틱범위 = self.comboBox_min.currentText()[0:2].strip()
        if 틱범위[0] == '0':
            틱범위 = 틱범위[1:]
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "틱범위", 틱범위)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식분봉차트조회", "OPT10080", _repeat, '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)


class RealDataTableModel(QAbstractTableModel):
    def __init__(self, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self.realdata = {}
        self.headers = ['종목코드', '현재가' , '전일대비', '등락률' , '매도호가', '매수호가', '누적거래량', '시가' , '고가' , '저가' , '거래회전율', '시가총액']

    def rowCount(self, index=QModelIndex()):
        return len(self.realdata)

    def columnCount(self, index=QModelIndex()):
        return len(self.headers)

    def data(self, index, role=Qt.DisplayRole):
        if (not index.isValid() or not (0 <= index.row() < len(self.realdata))):
            return None

        if role == Qt.DisplayRole:
            rows = []
            for k in self.realdata.keys():
                rows.append(k)
            one_row = rows[index.row()]
            selected_row = self.realdata[one_row]

            return selected_row[index.column()]

        return None

    def headerData(self, column, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.headers[column]
        return int(column + 1)

    def flags(self, index):
        return QtCore.Qt.ItemIsEnabled

    def reset(self):
        self.beginResetModel()
        self.endResetModel()

Ui_실시간정보, QtBaseClass_실시간정보 = uic.loadUiType("실시간정보.ui")

class 화면_실시간정보(QDialog, Ui_실시간정보):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_실시간정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = RealDataTableModel()
        self.tableView.setModel(self.model)

    def KiwoomConnect(self):
        self.kiwoom.OnEventConnect[int].connect(self.OnEventConnect)
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].connect(self.OnReceiveChejanData)
        self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].connect(self.OnReceiveRealData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnEventConnect[int].disconnect(self.OnEventConnect)
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].disconnect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].disconnect(self.OnReceiveChejanData)
        self.kiwoom.OnReceiveConditionVer[int, str].disconnect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].disconnect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].disconnect(self.OnReceiveRealData)

    def KiwoomAccount(self):
        ACCOUNT_CNT = self.kiwoom.dynamicCall('GetLoginInfo("ACCOUNT_CNT")')
        ACC_NO = self.kiwoom.dynamicCall('GetLoginInfo("ACCNO")')

        self.account = ACC_NO.split(';')[0:-1]
        self.plainTextEdit.insertPlainText("보유 계좌수: %s 계좌번호: %s [%s]" % (ACCOUNT_CNT, self.account[0], ACC_NO))
        self.plainTextEdit.insertPlainText("\r\n")

    def KiwoomSendOrder(self, sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo):
        order = self.kiwoom.dynamicCall('SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                                        [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo])
        return order

    def KiwoomSetRealReg(self, sScreenNo, sCode, sRealType='0'):
        ret = self.kiwoom.dynamicCall('SetRealReg(QString, QString, QString, QString)', sScreenNo, sCode, '9001;10', sRealType)
        return ret

    def KiwoomSetRealRemove(self, sScreenNo, sCode):
        ret = self.kiwoom.dynamicCall('SetRealRemove(QString, QString)', sScreenNo, sCode)
        return ret

    def OnEventConnect(self, nErrCode):
        self.plainTextEdit.insertPlainText('OnEventConnect %s' % nErrCode)
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        self.plainTextEdit.insertPlainText('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveTrCondition(self, sScrNo, strCodeList, strConditionName, nIndex, nNext):
        self.plainTextEdit.insertPlainText('OnReceiveTrCondition [%s] [%s] [%s] [%s] [%s]' % (sScrNo, strCodeList, strConditionName, nIndex, nNext))
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[화면번호] : %s" % sScrNo)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[종목리스트] : %s" % strCodeList)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[조건명] : %s" % strConditionName)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[조건명 인덱스 ] : %s" % nIndex)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[연속조회] : %s" % nNext)
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        self.plainTextEdit.insertPlainText('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        self.plainTextEdit.insertPlainText('OnReceiveChejanData [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
        self.plainTextEdit.insertPlainText("\r\n")

        if sGubun == "0":
            param = dict()

            param['계좌번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9201)
            param['주문번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9203)
            param['종목코드'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9001)
            param['종목명'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 302)
            param['주문수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 900)
            param['주문가격'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 901)
            param['원주문번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 904)
            param['체결량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 911)
            param['미체결수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 902)
            param['매도수구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 907)
            param['단위체결가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 914)
            param['화면번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 920)

            self.plainTextEdit.insertPlainText(str(param))
            self.plainTextEdit.insertPlainText("\r\n")

        if sGubun == "1":
            self.plainTextEdit.insertPlainText('OnReceiveChejanData: 잔고통보 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
            self.plainTextEdit.insertPlainText("\r\n")
        if sGubun == "3":
            self.plainTextEdit.insertPlainText('OnReceiveChejanData: 특이신호 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
            self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveConditionVer(self, lRet, sMsg):
        self.plainTextEdit.insertPlainText('OnReceiveConditionVer : [이벤트] 조건식 저장 %s %s' % (lRet, sMsg))

    def OnReceiveRealCondition(self, sTrCode, strType, strConditionName, strConditionIndex):
        self.plainTextEdit.insertPlainText('OnReceiveRealCondition [%s] [%s] [%s] [%s]' % (sTrCode, strType, strConditionName, strConditionIndex))
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("========= 조건조회 실시간 편입/이탈 ==========")
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[종목코드] : %s" % sTrCode)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[실시간타입] : %s" % strType)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[조건명] : %s" % strConditionName)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[조건명 인덱스] : %s" % strConditionIndex)
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveRealData(self, sRealKey, sRealType, sRealData):
        self.plainTextEdit.insertPlainText("[%s] [%s] %s\n" % (sRealKey, sRealType, sRealData))

        if sRealType == "주식시세" or sRealType == "주식체결":
            param = dict()

            param['종목코드'] = sRealKey.strip()
            param['현재가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 10).strip()))
            param['전일대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 11).strip()
            param['등락률'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 12).strip()
            param['매도호가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 27).strip()))
            param['매수호가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 28).strip()))
            param['누적거래량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 13).strip()
            param['시가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 16).strip()))
            param['고가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 17).strip()))
            param['저가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 18).strip()))
            param['거래회전율'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 31).strip()
            param['시가총액'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 311).strip()

            self.model.realdata[sRealKey] = [param['종목코드'], param['현재가'], param['전일대비'], param['등락률'], param['매도호가'], param['매수호가'], param['누적거래량'], param['시가'], param['고가'], param['저가'], param['거래회전율'], param['시가총액']]
            self.model.reset()

            for i in range(len(self.model.realdata[sRealKey])):
                self.tableView.resizeColumnToContents(i)



##
## TickLogger
Ui_TickLogger, QtBaseClass_TickLogger = uic.loadUiType("TickLogger.ui")

class 화면_TickLogger(QDialog, Ui_TickLogger):
    def __init__(self, parent):
        super(화면_TickLogger, self).__init__(parent)
        self.setupUi(self)

class CTickLogger(CTrade):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.portfolio = dict()
        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.buffer = []

        self.d = datetime.date.today()

    def Setting(self, sScreenNo, 종목유니버스):
        self.sScreenNo = sScreenNo
        self.종목유니버스 = 종목유니버스

        self.실시간종목리스트 = 종목유니버스

    def 실시간데이타처리(self, param):

        if self.running == True:
            _체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
            if len(self.buffer) < 100:
                체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
                종목코드 = param['종목코드']
                현재가 = abs(int(float(param['현재가'])))
                전일대비 = int(float(param['전일대비']))
                등락률 = float(param['등락률'])
                매도호가 = abs(int(float(param['매도호가'])))
                매수호가 = abs(int(float(param['매수호가'])))
                누적거래량 = abs(int(float(param['누적거래량'])))
                시가 = abs(int(float(param['시가'])))
                고가 = abs(int(float(param['고가'])))
                저가 = abs(int(float(param['저가'])))
                거래회전율 = abs(float(param['거래회전율']))
                시가총액 = abs(int(float(param['시가총액'])))

                lst = [체결시간, 종목코드, 현재가, 전일대비, 등락률, 매도호가, 매수호가, 누적거래량, 시가, 고가, 저가, 거래회전율, 시가총액]
                self.buffer.append(lst)
                self.parent.statusbar.showMessage(
                    "[%s]%s %s %s %s" % (_체결시간, 종목코드, self.parent.CODE_POOL[종목코드][1], 현재가, 전일대비))
            else:
                df = DataFrame(data=self.buffer, columns=['체결시간', '종목코드', '현재가', '전일대비', '등락률', '매도호가', '매수호가', '누적거래량', '시가', '고가', '저가', '거래회전율', '시가총액'])
                df.to_csv('TickLogger.csv', mode='a', header=False)
                self.buffer = []
                self.parent.statusbar.showMessage("CTickLogger 기록함")

    def 접수처리(self, param):
        pass

    def 체결처리(self, param):
        pass

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        self.running = flag

        ret = 0
        if flag == True:
            self.KiwoomConnect()
            ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.종목유니버스)+';')

        else:
            ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')
            self.KiwoomDisConnect()

            df = DataFrame(data=self.buffer, columns=['체결시간', '종목코드', '현재가', '전일대비', '등락률', '매도호가', '매수호가', '누적거래량', '시가', '고가', '저가', '거래회전율', '시가총액'])
            df.to_csv('TickLogger.csv', mode='a', header=False)
            self.buffer = []
            self.parent.statusbar.showMessage("CTickLogger 기록함")

##
## TickMonitor
class CTickMonitor(CTrade):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.portfolio = dict()
        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.buffer = []

        self.d = datetime.date.today()

        self.모니터링종목 = dict()
        self.누적거래량 = dict()

    def Setting(self, sScreenNo, 종목유니버스):
        self.sScreenNo = sScreenNo
        self.종목유니버스 = 종목유니버스

        self.실시간종목리스트 = 종목유니버스

    def 실시간데이타처리(self, param):

        if self.running == True:

            체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
            종목코드 = param['종목코드']
            현재가 = abs(int(float(param['현재가'])))
            #체결량 = abs(int(float(param['체결량'])))
            전일대비 = int(float(param['전일대비']))
            등락률 = float(param['등락률'])
            매도호가 = abs(int(float(param['매도호가'])))
            매수호가 = abs(int(float(param['매수호가'])))
            누적거래량 = abs(int(float(param['누적거래량'])))
            시가 = abs(int(float(param['시가'])))
            고가 = abs(int(float(param['고가'])))
            저가 = abs(int(float(param['저가'])))
            거래회전율 = abs(float(param['거래회전율']))
            시가총액 = abs(int(float(param['시가총액'])))

            체결량 = 0
            if self.누적거래량.get(종목코드) == None:
                self.누적거래량[종목코드] = 누적거래량
            else:
                체결량 = 누적거래량 - self.누적거래량[종목코드]
                if 체결량 < 2:
                    if self.모니터링종목.get(종목코드) == None:
                        self.모니터링종목[종목코드] = 1
                    else:
                        self.모니터링종목[종목코드] += 1

                    temp = []
                    for k, v in self.모니터링종목.items():
                        if v >= 10:
                            temp.append(k)
                    if len(temp) > 0:
                        logger.info("%s %s" % (체결시간, temp))
                else:
                    pass

                self.누적거래량[종목코드] = 누적거래량

                self.parent.statusbar.showMessage(
                    "[%s]%s %s %s %s" % (체결시간, 종목코드, self.parent.CODE_POOL[종목코드][1], 현재가, 전일대비))

    def 접수처리(self, param):
        pass

    def 체결처리(self, param):
        pass

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        self.running = flag

        ret = 0
        if flag == True:
            ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.종목유니버스)+';')
            self.KiwoomConnect()
        else:
            ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')
            self.KiwoomDisConnect()

##
## TickTradeRSI
Ui_TickTradeRSI, QtBaseClass_TickTradeRSI = uic.loadUiType("TickTradeRSI.ui")
class 화면_TickTradeRSI(QDialog, Ui_TickTradeRSI):
    def __init__(self, parent):
        super(화면_TickTradeRSI, self).__init__(parent)
        self.setupUi(self)

class CTickTradeRSI(CTrade):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.주문결과 = dict()
        self.주문번호_주문_매핑 = dict()
        self.주문실행중_Lock = dict()

        self.portfolio = dict()

        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.d = datetime.date.today()

    def Setting(self, sScreenNo, 포트폴리오수=10, 단위투자금=300*10000, 시총상한=4000, 시총하한=500, 매수방법='00', 매도방법='00'):
        self.sScreenNo = sScreenNo
        self.실시간종목리스트 = []
        self.단위투자금 = 단위투자금
        self.매수방법 = 매수방법
        self.매도방법 = 매도방법
        self.포트폴리오수 = 포트폴리오수
        self.시총상한 = 시총상한
        self.시총하한 = 시총하한

    def get_price(self, code, 시작일자=None, 종료일자=None):

        if 시작일자 == None and 종료일자 == None:
            query = """
            SELECT 일자, 종가, 시가, 고가, 저가, 거래량
            FROM 일별주가
            WHERE 종목코드='%s'
            ORDER BY 일자 ASC
            """ % (code)
        if 시작일자 != None and 종료일자 == None:
            query = """
            SELECT 일자, 종가, 시가, 고가, 저가, 거래량
            FROM 일별주가
            WHERE 종목코드='%s' AND A.일자 >= '%s'
            ORDER BY 일자 ASC
            """ % (code, 시작일자)
        if 시작일자 == None and 종료일자 != None:
            query = """
            SELECT 일자, 종가, 시가, 고가, 저가, 거래량
            FROM 일별주가
            WHERE 종목코드='%s' AND 일자 <= '%s'
            ORDER BY 일자 ASC
            """ % (code, 종료일자)
        if 시작일자 != None and 종료일자 != None:
            query = """
            SELECT 일자, 종가, 시가, 고가, 저가, 거래량
            FROM 일별주가
            WHERE 종목코드='%s' AND 일자 BETWEEN '%s' AND '%s'
            ORDER BY 일자 ASC
            """ % (code, 시작일자, 종료일자)

        conn = mysqlconn()
        df = pdsql.read_sql_query(query, con=conn)
        conn.close()

        df.fillna(0, inplace=True)
        df.set_index('일자', inplace=True)

        df['RSI'] = ta.RSI(np.array(df['종가'].astype(float)))

        df['macdhigh'], df['macdsignal'], df['macdhist'] = ta.MACD(np.array(df['고가'].astype(float)), fastperiod=12, slowperiod=26, signalperiod=9)
        df['macdlow'], df['macdsignal'], df['macdhist'] = ta.MACD(np.array(df['저가'].astype(float)), fastperiod=12, slowperiod=26, signalperiod=9)
        df['macdclose'], df['macdsignal'], df['macdhist'] = ta.MACD(np.array(df['종가'].astype(float)), fastperiod=12, slowperiod=26, signalperiod=9)

        try:
            df['slowk'], df['slowd'] = ta.STOCH(np.array(df['macdhigh'].astype(float)), np.array(df['macdlow'].astype(float)),np.array(df['macdclose'].astype(float)), 
                                            fastk_period=15, slowk_period=15,slowd_period=5)
        except Exception as e:
            logger.info("데이타부족 %s" % code)
            return None

        df.dropna(inplace=True)

        return df

    def 초기조건(self):
        self.parent.statusbar.showMessage("[%s] 초기조건준비" % (self.sName))

        query = """
                SELECT
                    시장구분, 종목코드, 종목명, 주식수, 감리구분, 상장일, 전일종가,
                    CAST(((주식수 * 전일종가) / 100000000) AS UNSIGNED) AS 시가총액,
                    종목상태
                FROM
                    종목코드
                WHERE
                    ((시장구분 IN ('KOSPI' , 'KOSDAQ'))
                        AND (SUBSTR(종목코드, -1) = '0')
                        AND ((주식수 * 전일종가) between %s * (10000 * 10000) and %s * (10000 * 10000))
                        AND (NOT ((종목명 LIKE '%%스팩')))
                        AND (NOT ((종목명 LIKE '%%SPAC')))
                        AND (NOT ((종목상태 LIKE '%%관리종목%%')))
                        AND (NOT ((종목상태 LIKE '%%거래정지%%')))
                        AND (NOT ((감리구분 LIKE '%%투자경고%%')))
                        AND (NOT ((감리구분 LIKE '%%투자주의%%')))
                        AND (NOT ((감리구분 LIKE '%%환기종목%%')))
                        AND (NOT ((종목명 LIKE '%%ETN%%')))
                        AND (NOT ((종목명 LIKE '%%0호')))
                        AND (NOT ((종목명 LIKE '%%1호')))
                        AND (NOT ((종목명 LIKE '%%2호')))
                        AND (NOT ((종목명 LIKE '%%3호')))
                        AND (NOT ((종목명 LIKE '%%4호')))
                        AND (NOT ((종목명 LIKE '%%5호')))
                        AND (NOT ((종목명 LIKE '%%6호')))
                        AND (NOT ((종목명 LIKE '%%7호')))
                        AND (NOT ((종목명 LIKE '%%8호')))
                        AND (NOT ((종목명 LIKE '%%9호')))
                        AND (NOT ((종목명 LIKE '%%0')))
                        AND (NOT ((종목명 LIKE '%%1')))
                        AND (NOT ((종목명 LIKE '%%2')))
                        AND (NOT ((종목명 LIKE '%%3')))
                        AND (NOT ((종목명 LIKE '%%4')))
                        AND (NOT ((종목명 LIKE '%%5')))
                        AND (NOT ((종목명 LIKE '%%6')))
                        AND (NOT ((종목명 LIKE '%%7')))
                        AND (NOT ((종목명 LIKE '%%8')))
                        AND (NOT ((종목명 LIKE '%%9'))))
                ORDER BY 종목코드 ASC
            """ % (self.시총하한, self.시총상한)

        conn = mysqlconn()
        CODES = pdsql.read_sql_query(query, con=conn)
        conn.close()

        NOW = datetime.datetime.now()
        시작일자 = (NOW + datetime.timedelta(days=-366)).strftime('%Y-%m-%d')
        종료일자 = NOW.strftime('%Y-%m-%d')

        pool = dict()
        for market, code, name, 주식수, 시가총액 in CODES[['시장구분','종목코드','종목명','주식수','시가총액']].values.tolist():
            df = self.get_price(code, 시작일자, 종료일자)
            if df is not None and len(df) > 0:
                pool[code] = df

        self.금일매도 = []
        self.매도할종목 = []
        self.매수할종목 = []
        for code, df in pool.items():

            try:
                종가D1, RSID1, macdD1, slowkD1, slowdD1 = df[['종가', 'RSI','macdclose','slowk','slowd']].values[-2]
                종가D0, RSID0, macdD0, slowkD0, slowdD0 = df[['종가', 'RSI','macdclose','slowk','slowd']].values[-1]

                # print(code, 종가D0, RSID0)

                stock = self.portfolio.get(code)
                if stock != None:
                    # print(stock)
                    if RSID1 > 70.0 and RSID0 < 70.0:
                        self.매도할종목.append(code)

                if stock == None:
                    # print(code)
                    if RSID1 < 30.0 and RSID0 > 30.0:
                        self.매수할종목.append(code)
            except Exception as e:
                logger.info("데이타부족 %s" % code)
                print(df)

        pool = None

    def 실시간데이타처리(self, param):
        if self.running == True:

            체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
            종목코드 = param['종목코드']
            현재가 = abs(int(float(param['현재가'])))
            전일대비 = int(float(param['전일대비']))
            등락률 = float(param['등락률'])
            매도호가 = abs(int(float(param['매도호가'])))
            매수호가 = abs(int(float(param['매수호가'])))
            누적거래량 = abs(int(float(param['누적거래량'])))
            시가 = abs(int(float(param['시가'])))
            고가 = abs(int(float(param['고가'])))
            저가 = abs(int(float(param['저가'])))
            거래회전율 = abs(float(param['거래회전율']))
            시가총액 = abs(int(float(param['시가총액'])))

            종목명 = self.parent.CODE_POOL[종목코드][1]

            self.parent.statusbar.showMessage("[%s] %s %s %s %s" % (체결시간, 종목코드, 종목명, 현재가, 전일대비))

            if 종목코드 in self.매도할종목:
                if self.portfolio.get(종목코드) is not None and self.주문실행중_Lock.get('S_%s' % 종목코드) is None:
                    (result, order) = self.정량매도(sRQName='S_%s' % 종목코드, 종목코드=종목코드, 매도가=현재가, 수량=self.portfolio[종목코드].수량)
                    if result == True:
                        self.주문실행중_Lock['S_%s' % 종목코드] = True
                        logger.debug('정량매도 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % ('S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량) )
                    else:
                        logger.debug('정량매도실패 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % ('S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량) )

            if 종목코드 in self.매수할종목 and 종목코드 not in self.금일매도:
                if len(self.portfolio) < self.포트폴리오수 and self.portfolio.get(종목코드) is None and self.주문실행중_Lock.get('B_%s' % 종목코드) is None:
                    if 현재가 < (현재가-전일대비):
                        (result, order) = self.정액매수(sRQName='B_%s' % 종목코드, 종목코드=종목코드, 매수가=현재가, 매수금액=self.단위투자금)
                        if result == True:
                            self.portfolio[종목코드] = CPortStock(종목코드=종목코드, 종목명=종목명, 매수가=현재가, 매도가1차=0, 매도가2차=0, 손절가=0, 수량=0, 매수일=datetime.datetime.now())
                            self.주문실행중_Lock['B_%s' % 종목코드] = True
                            logger.debug('정액매수 : sRQName=%s, 종목코드=%s, 매수가=%s, 단위투자금=%s' % ('B_%s' % 종목코드, 종목코드, 현재가, self.단위투자금) )
                        else:
                            logger.debug('정액매수실패 : sRQName=%s, 종목코드=%s, 매수가=%s, 단위투자금=%s' % ('B_%s' % 종목코드, 종목코드, 현재가, self.단위투자금) )

    def 접수처리(self, param):
        pass

    def 체결처리(self, param):
        종목코드 = param['종목코드']
        주문번호 = param['주문번호']
        self.주문결과[주문번호] = param

        if param['매도수구분'] == '2':  # 매수
            주문수량 = int(param['주문수량'])
            미체결수량 = int(param['미체결수량'])
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매수가 = int(주문[2:])
                단위체결가 = int(0 if (param['단위체결가'] is None or param['단위체결가'] == '') else param['단위체결가'])

                # logger.debug('매수-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))

                P = self.portfolio.get(종목코드)
                if P is not None:
                    P.종목명 = param['종목명']
                    P.매수가 = 단위체결가
                    P.수량 = 주문수량 - 미체결수량
                else:
                    logger.debug('ERROR 포트에 종목이 없음 !!!!')

                if 미체결수량 == 0:
                    try:
                        self.주문실행중_Lock.pop(주문)
                        # logger.info('POP성공 %s ' % 주문)
                    except Exception as e:
                        # logger.info('POP에러 %s ' % 주문)
                        pass

        if param['매도수구분'] == '1':  # 매도
            주문수량 = int(param['주문수량'])
            미체결수량 = int(param['미체결수량'])
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매수가 = int(주문[2:])

                if 미체결수량 == 0:
                    try:
                        self.portfolio.pop(종목코드)
                        # logger.info('포트폴리오POP성공 %s ' % 종목코드)
                        self.금일매도.append(종목코드)
                    except Exception as e:
                        # logger.info('포트폴리오POP에러 %s ' % 종목코드)
                        pass

                    try:
                        self.주문실행중_Lock.pop(주문)
                        # logger.info('POP성공 %s ' % 주문)
                    except Exception as e:
                        # logger.info('POP에러 %s ' % 주문)
                        pass
                else:
                    # logger.debug('매도-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))
                    P = self.portfolio.get(종목코드)
                    if P is not None:
                        P.종목명 = param['종목명']
                        P.수량 = 미체결수량

        # 메인 화면에 반영
        self.parent.RobotView()

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        self.running = flag

        ret = 0
        if flag == True:
            self.sAccount = sAccount
            if self.sAccount is None:
                self.KiwoomAccount()
                self.sAccount = self.account[0]

            self.주문결과 = dict()
            self.주문번호_주문_매핑 = dict()
            self.주문실행중_Lock = dict()

            self.초기조건()

            self.실시간종목리스트 = self.매도할종목 + self.매수할종목 + list(self.portfolio.keys())

            logger.debug("오늘 거래 종목 : %s %s" % (self.sName, ';'.join(self.실시간종목리스트) + ';'))
            self.KiwoomConnect()
            if len(self.실시간종목리스트) > 0:
                ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.실시간종목리스트) + ';')
                logger.debug("실시간데이타요청 등록결과 %s" % ret)
        else:
            ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')
            self.KiwoomDisConnect()

##
## TickFuturesLogger
Ui_TickFuturesLogger, QtBaseClass_TickFuturesLogger = uic.loadUiType("TickFuturesLogger.ui")

class 화면_TickFuturesLogger(QDialog, Ui_TickFuturesLogger):
    def __init__(self, parent):
        super(화면_TickFuturesLogger, self).__init__(parent)
        self.setupUi(self)

class CTickFuturesLogger(CTrade):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.portfolio = dict()
        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.buffer = []

        self.d = datetime.date.today()


    def Setting(self, sScreenNo, 종목유니버스):
        self.sScreenNo = sScreenNo
        self.종목유니버스 = 종목유니버스

        self.실시간종목리스트 = 종목유니버스

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "선옵현재가정보요청":
            param = dict()

            param['현재가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0,"현재가").strip()
            param['대비기호'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "대비기호").strip()
            param['전일대비'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "전일대비").strip()
            param['등락률'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "등락률").strip()
            param['거래량'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "거래량").strip()
            param['거래량대비'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "거래량대비").strip()
            param['기준가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "기준가").strip()
            param['이론가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "이론가").strip()
            param['이론베이시스'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "이론베이시스").strip()
            param['괴리도'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "괴리도").strip()
            param['괴리율'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "괴리율").strip()
            param['시장베이시스'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "시장베이시스").strip()
            param['누적거래대금'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "누적거래대금").strip()
            param['상한가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "상한가").strip()
            param['하한가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "하한가").strip()
            param['CB상한가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "CB상한가").strip()
            param['CB하한가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "CB하한가").strip()
            param['대용가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "대용가").strip()
            param['최종거래일'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "최종거래일").strip()
            param['잔존일수'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "잔존일수").strip()
            param['영업일기준잔존일'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "영업일기준잔존일").strip()
            param['상장중최고가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "상장중최고가").strip()
            param['상장중최고가대비율'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "상장중최고가대비율").strip()
            param['상장중최고가일'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "상장중최고가일").strip()
            param['종목명'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "종목명").strip()
            param['호가시간'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "호가시간").strip()


            param['매도수익율5'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수익율5").strip()
            param['매도건수5'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도건수5").strip()
            param['매도수량5'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수량5").strip()
            param['매도호가5'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도호가5").strip()
            param['매수호가5'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수호가5").strip()
            param['매수수량5'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수량5").strip()
            param['매수건수5'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수건수5").strip()
            param['매수수익율5'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수익율5").strip()

            param['매도수익율4'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수익율4").strip()
            param['매도건수4'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도건수4").strip()
            param['매도수량4'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수량4").strip()
            param['매도호가4'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도호가4").strip()
            param['매수호가4'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수호가4").strip()
            param['매수수량4'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수량4").strip()
            param['매수건수4'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수건수4").strip()
            param['매수수익율4'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수익율4").strip()

            param['매도수익율3'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수익율3").strip()
            param['매도건수3'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도건수3").strip()
            param['매도수량3'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수량3").strip()
            param['매도호가3'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도호가3").strip()
            param['매수호가3'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수호가3").strip()
            param['매수수량3'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수량3").strip()
            param['매수건수3'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수건수3").strip()
            param['매수수익율3'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수익율3").strip()

            param['매도수익율2'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수익율2").strip()
            param['매도건수2'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도건수2").strip()
            param['매도수량2'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수량2").strip()
            param['매도호가2'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도호가2").strip()
            param['매수호가2'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수호가2").strip()
            param['매수수량2'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수량2").strip()
            param['매수건수2'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수건수2").strip()
            param['매수수익율2'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수익율2").strip()

            param['매도수익율1'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수익율1").strip()
            param['매도건수1'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도건수1").strip()
            param['매도수량1'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도수량1").strip()
            param['매도호가1'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도호가1").strip()
            param['매수호가1'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수호가1").strip()
            param['매수수량1'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수량1").strip()
            param['매수건수1'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수건수1").strip()
            param['매수수익율1'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수수익율1").strip()

            param['시가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "시가").strip()
            param['고가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "고가").strip()
            param['저가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "저가").strip()
            param['2차저항'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "2차저항").strip()
            param['1차저항'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "1차저항").strip()
            param['피봇'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "피봇").strip()
            param['1차저지'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "1차저지").strip()
            param['2차저지'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "2차저지").strip()
            param['미결제약정'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "미결제약정").strip()
            param['미결제약정전일대비'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "미결제약정전일대비").strip()
            param['매도호가총건수'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도호가총건수").strip()
            param['매도호가총잔량'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도호가총잔량").strip()
            param['순매수잔량'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "순매수잔량").strip()
            param['매수호가총잔량'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매주호가총잔량").strip()
            param['매수호가총건수'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수호가총건수").strip()
            param['매도호가총잔량직전대비'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매도호가총잔량직전대비").strip()
            param['매수호가총잔량직전대비'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "매수호가총잔량직전대비").strip()
            param['예상체결가'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "예상체결가").strip()
            param['예상체결가전일종가대비기호'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "예상체결가전일종가대비기호").strip()
            param['예상체결가전일종가대비'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "예상체결가전일종가대비").strip()
            param['예상체결가전일종가대비등락율'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "예상체결가전일종가대비등락율").strip()
            param['이자율'] = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "이자율").strip()

    def OnReceiveRealData(self, sRealKey, sRealType, sRealData):
        """
        OpenAPI 메뉴얼 참조
        :param sRealKey:
        :param sRealType:
        :param sRealData:
        :return:
        """
        logger.info('OnReceiveRealData [%s] [%s] [%s]' % (sRealKey, sRealType, sRealData))

        if sRealType == "선물이론가":
            param = dict()
            param['미결제약정전일대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 181).strip()
            param['이론가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 182).strip()
            param['시장베이시스'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 183).strip()
            param['이론베이시스'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 184).strip()
            param['괴리도'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 185).strip()
            param['괴리율'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 186).strip()
            param['미결제약정'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 195).strip()
            param['시초미결제약정수량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 246).strip()
            param['최고미결제약정수량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 247).strip()
            param['최저미결제약정수량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 248).strip()

            self.실시간데이타처리(param)

        if sRealType == "선물호가잔량":
            param = dict()
            param['누적거래량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 13).strip()
            param['호가시간'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 21).strip()
            param['예상체결가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 23).strip()
            param['최우선매도호가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 27).strip()
            param['최우선매수호가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 28).strip()
            param['매도호가1'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 41).strip()
            param['매도호가2'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 42).strip()
            param['매도호가3'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 43).strip()
            param['매도호가4'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 44).strip()
            param['매도호가5'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 45).strip()
            param['매수호가1'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 51).strip()
            param['매수호가2'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 52).strip()
            param['매수호가3'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 53).strip()
            param['매수호가4'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 54).strip()
            param['매수호가5'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 55).strip()
            param['매도호가수량1'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 61).strip()
            param['매도호가수량2'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 62).strip()
            param['매도호가수량3'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 63).strip()
            param['매도호가수량4'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 64).strip()
            param['매도호가수량5'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 65).strip()
            param['매수호가수량1'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 71).strip()
            param['매수호가수량2'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 72).strip()
            param['매수호가수량3'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 73).strip()
            param['매수호가수량4'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 74).strip()
            param['매수호가수량5'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 75).strip()
            param['매도호가직전대비1'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 81).strip()
            param['매도호가직전대비2'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 82).strip()
            param['매도호가직전대비3'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 83).strip()
            param['매도호가직전대비4'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 84).strip()
            param['매도호가직전대비5'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 85).strip()
            param['매수호가직전대비1'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 91).strip()
            param['매수호가직전대비2'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 92).strip()
            param['매수호가직전대비3'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 93).strip()
            param['매수호가직전대비4'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 94).strip()
            param['매수호가직전대비5'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 95).strip()
            param['매도호가건수1'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 101).strip()
            param['매도호가건수2'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 102).strip()
            param['매도호가건수3'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 103).strip()
            param['매도호가건수4'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 104).strip()
            param['매도호가건수5'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 105).strip()
            param['매수호가건수1'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 111).strip()
            param['매수호가건수2'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 112).strip()
            param['매수호가건수3'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 113).strip()
            param['매수호가건수4'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 114).strip()
            param['매수호가건수5'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 115).strip()
            param['매도호가총잔량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 121).strip()
            param['매도호가총잔량직전대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 122).strip()
            param['매도호가총건수'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 123).strip()
            param['매수호가총잔량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 125).strip()
            param['매수호가총잔량직전대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 126).strip()
            param['매도호가총건수'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 127).strip()
            param['순매수잔량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 128).strip()
            param['예상체결가전일종가대비기호'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 238).strip()
            param['예상체결가전일종가대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 200).strip()
            param['예상체결가전일종가대비등락률'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 201).strip()
            param['예상체결가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 291).strip()
            param['예상체결가전일대비기호'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 293).strip()
            param['예상체결가전일대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 294).strip()
            param['예상체결가전일대비등락률'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 295).strip()

            # param['현재가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 10).strip()
            # param['전일대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 11).strip()
            # param['등락률'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 12).strip()
            # param['거래량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 15).strip()
            # param['시가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 16).strip()
            # param['고가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 17).strip()
            # param['저가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 18).strip()
            # param['체결시간'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 20).strip()
            # param['전일대비기호'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 25).strip()
            # param['전일거래량대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 26).strip()
            # param['전일거래량대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 30).strip()
            # param['호가순잔량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 137).strip()
            # param['미결제약정전일대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 181).strip()
            # param['이론가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 182).strip()
            # param['시장베이시스'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 183).strip()
            # param['이론베이시스'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 184).strip()
            # param['괴리도'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 185).strip()
            # param['괴리율'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 186).strip()
            # param['미결제약정'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 195).strip()
            # param['미결제증감'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 196).strip()
            # param['KOSPI200'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 197).strip()
            # param['시초미결제약정수량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 246).strip()
            # param['최고미결제약정수량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 247).strip()
            # param['최저미결제약정수량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 248).strip()


    def Request(self, _repeat=0):
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목유니버스[0])
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "선옵현재가정보요청", "OPT50001", _repeat, '{:04d}'.format(self.sScreenNo))

    def 실시간데이타처리(self, param):

        if self.running == True:
            if len(self.buffer) < 10:
                현재시간 = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                이론가 = param['이론가']

                lst = [현재시간, 이론가]
                logger.info(lst)
                self.buffer.append(lst)
                self.parent.statusbar.showMessage("TickFuturesLogger : [%s]%s" % (현재시간, 이론가))
            else:
                df = DataFrame(data=self.buffer,columns=['현재시간', '이론가'])
                df.to_csv('TickFuturesLogger.csv', mode='a', header=False)
                self.buffer = []
                self.parent.statusbar.showMessage("CTickFuturesLogger 기록함")

    def 접수처리(self, param):
        pass

    def 체결처리(self, param):
        pass

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        self.running = flag

        ret = 0
        if flag == True:
            self.KiwoomConnect()
            self.Request()

        else:
            self.KiwoomDisConnect()

            df = DataFrame(data=self.buffer, columns=['현재시간', '이론가'])
            df.to_csv('TickFuturesLogger.csv', mode='a', header=False)
            self.buffer = []
            self.parent.statusbar.showMessage("TickFuturesLogger 기록함")


##################################################################################
# 메인
##################################################################################

Ui_MainWindow, QtBaseClass_MainWindow = uic.loadUiType("mymoneybot.ui")

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.setWindowTitle("mymoneybot")

        self.시작시각 = datetime.datetime.now()

        self.KiwoomAPI()
        self.KiwoomConnect()
        self.ScreenNumber = 5000

        self.robots = []

        self.dialog = dict()
        # self.dialog['리얼데이타'] = None
        # self.dialog['계좌정보조회'] = None

        self.model = PandasModel()
        self.tableView_robot.setModel(self.model)
        self.tableView_robot.setSelectionBehavior(QTableView.SelectRows)
        self.tableView_robot.setSelectionMode(QTableView.SingleSelection)

        self.tableView_robot.pressed.connect(self.RobotCurrentIndex)
        # self.connect(self.tableView_robot.selectionModel(), SIGNAL("currentRowChanged(QModelIndex,QModelIndex)"), self.RobotCurrentIndex)
        self.tableView_robot_current_index = None

        self.portfolio_model = PandasModel()
        self.tableView_portfolio.setModel(self.portfolio_model)
        self.tableView_portfolio.setSelectionBehavior(QTableView.SelectRows)
        self.tableView_portfolio.setSelectionMode(QTableView.SingleSelection)
        self.portfolio_model.update((DataFrame(columns=['종목코드', '종목명', '라벨', '매수가', '수량', '매수일'])))


        self.robot_columns = ['Robot타입', 'Robot명', 'RobotID', '스크린번호', '실행상태', '포트수', '포트폴리오']

        #TODO: 주문제한 설정
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.limit_per_second)
        # QtCore.QObject.connect(self.timer, QtCore.SIGNAL("timeout()"), self.limit_per_second)
        self.timer.start(1000)

        self.주문제한 = 0
        self.조회제한 = 0
        self.금일백업작업중 = False
        self.종목선정작업중 = False

        self._login = False

        self.CODE_POOL = self.get_code_pool()

    def get_code_pool(self):
        query = """
            select 시장구분, 종목코드, 종목명, 주식수, 전일종가*주식수 as 시가총액
            from 종목코드
            order by 시장구분, 종목코드
        """
        conn = mysqlconn()
        df = pdsql.read_sql_query(query, con=conn)
        conn.close()

        pool = dict()
        for idx, row in df.iterrows():
            시장구분, 종목코드, 종목명, 주식수, 시가총액 = row
            pool[종목코드] = [시장구분, 종목명, 주식수, 시가총액]
        return pool

    def OnQApplicationStarted(self):
        global 로봇거래계좌번호

        conn = mysqlconn()
        cursor = conn.cursor()

        cursor.execute("select value from mymoneybot_setting where keyword='robotaccount'")
        for row in cursor.fetchall():
            _temp = base64.decodestring(bytes(row[0], 'ascii'))
            로봇거래계좌번호 = pickle.loads(_temp)
            # print(로봇거래계좌번호)

        cursor.execute('select uuid, strategy, name, robot from mymoneybot_robots')

        self.robots = []
        for row in cursor.fetchall():
            uuid, strategy, name, robot_encoded = row

            robot = base64.decodebytes(bytes(robot_encoded, 'ascii'))
            r = pickle.loads(robot)

            r.kiwoom = self.kiwoom
            r.parent = self
            r.d = datetime.date.today()

            r.running = False
            # logger.debug(r.sName, r.UUID, len(r.portfolio))
            self.robots.append(r)

        conn.close()
        self.RobotView()

    def OnClockTick(self):
        current = datetime.datetime.now()
        # print(current.strftime('%H:%M:%S'))

        if '15:36:00' < current.strftime('%H:%M:%S') and current.strftime('%H:%M:%S') < '15:36:59' and self.금일백업작업중 == False and self._login == True:
        # 수능일
        # if '17:00:00' < current.strftime('%H:%M:%S') and current.strftime('%H:%M:%S') < '17:00:59' and self.금일백업작업중 == False and self._login == True:
            self.금일백업작업중 = True
            self.Backup(작업=None)

        # if self.시작시각.strftime('%H:%M:%S') > '08:00:00' and self.시작시각.strftime('%H:%M:%S') < '15:30:00' and current.strftime('%H:%M:%S') > '01:00:00':
        #     if len(self.robots) > 0:
        #         self.RobotSave()

        #     for k in self.dialog:
        #         self.dialog[k].KiwoomDisConnect()
        #         try:
        #             self.dialog[k].close()
        #         except Exception as e:
        #             pass

        #     self.close()

        if current.second == 0: # 매 0초
            if current.minute % 10 == 0: # 매 10 분
                # print(current.minute, current.second)
                for r in self.robots:
                    if r.running == True: # 로봇이 실행중이면
                        # print(r.sName, r.running)
                        pass

    def limit_per_second(self):
        self.주문제한 = 0
        self.조회제한 = 0
        # logger.info("초당제한 클리어")

    def robot_selected(self, QModelIndex):
        # print(self.model._data[QModelIndex.row()])
        Robot타입 = self.model._data[QModelIndex.row():QModelIndex.row()+1]['Robot타입'].values[0]

        uuid = self.model._data[QModelIndex.row():QModelIndex.row()+1]['RobotID'].values[0]
        portfolio = None
        for r in self.robots:
            if r.UUID == uuid:
                portfolio = r.portfolio
                model = PandasModel()
                result = []
                for p, v in portfolio.items():
                    result.append((v.종목코드, v.종목명.strip(), p, v.매수가, v.수량, v.매수일))
                self.portfolio_model.update((DataFrame(data=result, columns=['종목코드','종목명','라벨','매수가','수량','매수일'])))

                break

    def robot_double_clicked(self, QModelIndex):
        self.RobotEdit(QModelIndex)
        self.RobotView()

    def RobotCurrentIndex(self, index):
        self.tableView_robot_current_index = index

    def MENU_Action(self, qaction):
        logger.debug("Action Slot %s %s " % (qaction.objectName(), qaction.text()))
        _action = qaction.objectName()
        if _action == "actionExit":
            if len(self.robots) > 0:
                self.RobotSave()

            for k in self.dialog:
                self.dialog[k].KiwoomDisConnect()
                try:
                    self.dialog[k].close()
                except Exception as e:
                    pass

            self.close()
        elif _action == "actionLogin":
            self.KiwoomLogin()
        elif _action == "actionLogout":
            self.KiwoomLogout()
        elif _action == "actionPriceBackupDay":
            #self.F_dailyprice_backup()
            if self.dialog.get('일별가격정보백업') is not None:
                try:
                    self.dialog['일별가격정보백업'].show()
                except Exception as e:
                    self.dialog['일별가격정보백업'] = 화면_일별가격정보백업(sScreenNo=9990, kiwoom=self.kiwoom, parent=self)
                    self.dialog['일별가격정보백업'].KiwoomConnect()
                    self.dialog['일별가격정보백업'].show()
            else:
                self.dialog['일별가격정보백업'] = 화면_일별가격정보백업(sScreenNo=9990, kiwoom=self.kiwoom, parent=self)
                self.dialog['일별가격정보백업'].KiwoomConnect()
                self.dialog['일별가격정보백업'].show()
        elif _action == "actionPriceBackupMin":
            #self.F_minprice_backup()
            if self.dialog.get('분별가격정보백업') is not None:
                try:
                    self.dialog['분별가격정보백업'].show()
                except Exception as e:
                    self.dialog['분별가격정보백업'] = 화면_분별가격정보백업(sScreenNo=9991, kiwoom=self.kiwoom, parent=self)
                    self.dialog['분별가격정보백업'].KiwoomConnect()
                    self.dialog['분별가격정보백업'].show()
            else:
                self.dialog['분별가격정보백업'] = 화면_분별가격정보백업(sScreenNo=9991, kiwoom=self.kiwoom, parent=self)
                self.dialog['분별가격정보백업'].KiwoomConnect()
                self.dialog['분별가격정보백업'].show()
        elif _action == "actionSectorBackupDay":
            #self.F_dailysector_backup()
            if self.dialog.get('일별업종정보백업') is not None:
                try:
                    self.dialog['일별업종정보백업'].show()
                except Exception as e:
                    self.dialog['일별업종정보백업'] = 화면_일별업종정보백업(sScreenNo=9993, kiwoom=self.kiwoom, parent=self)
                    self.dialog['일별업종정보백업'].KiwoomConnect()
                    self.dialog['일별업종정보백업'].show()
            else:
                self.dialog['일별업종정보백업'] = 화면_일별업종정보백업(sScreenNo=9993, kiwoom=self.kiwoom, parent=self)
                self.dialog['일별업종정보백업'].KiwoomConnect()
                self.dialog['일별업종정보백업'].show()
        elif _action == "actionInvestorBackup":
            #self.F_investor_backup()
            if self.dialog.get('종목별투자자정보백업') is not None:
                try:
                    self.dialog['종목별투자자정보백업'].show()
                except Exception as e:
                    self.dialog['종목별투자자정보백업'] = 화면_종목별투자자정보백업(sScreenNo=9992, kiwoom=self.kiwoom, parent=self)
                    self.dialog['종목별투자자정보백업'].KiwoomConnect()
                    self.dialog['종목별투자자정보백업'].show()
            else:
                self.dialog['종목별투자자정보백업'] = 화면_종목별투자자정보백업(sScreenNo=9992, kiwoom=self.kiwoom, parent=self)
                self.dialog['종목별투자자정보백업'].KiwoomConnect()
                self.dialog['종목별투자자정보백업'].show()
        elif _action == "actionDailyPrice":
            #self.F_dailyprice()
            if self.dialog.get('일자별주가') is not None:
                try:
                    self.dialog['일자별주가'].show()
                except Exception as e:
                    self.dialog['일자별주가'] = 화면_일별주가(sScreenNo=9902, kiwoom=self.kiwoom, parent=self)
                    self.dialog['일자별주가'].KiwoomConnect()
                    self.dialog['일자별주가'].show()
            else:
                self.dialog['일자별주가'] = 화면_일별주가(sScreenNo=9902, kiwoom=self.kiwoom, parent=self)
                self.dialog['일자별주가'].KiwoomConnect()
                self.dialog['일자별주가'].show()
        elif _action == "actionMinuitePrice":
            #self.F_minprice()
            if self.dialog.get('분별주가') is not None:
                try:
                    self.dialog['분별주가'].show()
                except Exception as e:
                    self.dialog['분별주가'] = 화면_분별주가(sScreenNo=9903, kiwoom=self.kiwoom, parent=self)
                    self.dialog['분별주가'].KiwoomConnect()
                    self.dialog['분별주가'].show()
            else:
                self.dialog['분별주가'] = 화면_분별주가(sScreenNo=9903, kiwoom=self.kiwoom, parent=self)
                self.dialog['분별주가'].KiwoomConnect()
                self.dialog['분별주가'].show()
        elif _action == "actionInvestors":
            #self.F_investor()
            if self.dialog.get('종목별투자자') is not None:
                try:
                    self.dialog['종목별투자자'].show()
                except Exception as e:
                    self.dialog['종목별투자자'] = 화면_종목별투자자(sScreenNo=9904, kiwoom=self.kiwoom, parent=self)
                    self.dialog['종목별투자자'].KiwoomConnect()
                    self.dialog['종목별투자자'].show()
            else:
                self.dialog['종목별투자자'] = 화면_종목별투자자(sScreenNo=9904, kiwoom=self.kiwoom, parent=self)
                self.dialog['종목별투자자'].KiwoomConnect()
                self.dialog['종목별투자자'].show()
        elif _action == "actionRealDataDialog":
            _code = '122630;114800'
            if self.dialog.get('리얼데이타') is not None:
                try:
                    self.dialog['리얼데이타'].show()
                except Exception as e:
                    self.dialog['리얼데이타'] = 화면_실시간정보(sScreenNo=9901, kiwoom=self.kiwoom, parent=self)
                    self.dialog['리얼데이타'].KiwoomConnect()
                    _screenno = self.dialog['리얼데이타'].sScreenNo
                    self.dialog['리얼데이타'].KiwoomSetRealRemove(_screenno, _code)
                    self.dialog['리얼데이타'].KiwoomSetRealReg(_screenno, _code, sRealType='0')
                    self.dialog['리얼데이타'].show()
            else:
                self.dialog['리얼데이타'] = 화면_실시간정보(sScreenNo=9901, kiwoom=self.kiwoom, parent=self)
                self.dialog['리얼데이타'].KiwoomConnect()
                _screenno = self.dialog['리얼데이타'].sScreenNo
                self.dialog['리얼데이타'].KiwoomSetRealRemove(_screenno, _code)
                self.dialog['리얼데이타'].KiwoomSetRealReg(_screenno, _code, sRealType='0')
                self.dialog['리얼데이타'].show()
        elif _action == "actionAccountDialog":
            if self.dialog.get('계좌정보조회') is not None:
                try:
                    self.dialog['계좌정보조회'].show()
                except Exception as e:
                    self.dialog['계좌정보조회'] = 화면_계좌정보(sScreenNo=7000, kiwoom=self.kiwoom, parent=self)
                    self.dialog['계좌정보조회'].KiwoomConnect()
                    self.dialog['계좌정보조회'].show()
            else:
                self.dialog['계좌정보조회'] = 화면_계좌정보(sScreenNo=7000, kiwoom=self.kiwoom, parent=self)
                self.dialog['계좌정보조회'].KiwoomConnect()
                self.dialog['계좌정보조회'].show()
        elif _action == "actionSectorView":
            #self.F_sectorview()
            if self.dialog.get('업종정보조회') is not None:
                try:
                    self.dialog['업종정보조회'].show()
                except Exception as e:
                    self.dialog['업종정보조회'] = 화면_업종정보(sScreenNo=9900, kiwoom=self.kiwoom, parent=self)
                    self.dialog['업종정보조회'].KiwoomConnect()
                    self.dialog['업종정보조회'].show()
            else:
                self.dialog['업종정보조회'] = 화면_업종정보(sScreenNo=9900, kiwoom=self.kiwoom, parent=self)
                self.dialog['업종정보조회'].KiwoomConnect()
                self.dialog['업종정보조회'].show()
        elif _action == "actionSectorPriceView":
            #self.F_sectorpriceview()
            if self.dialog.get('업종별주가조회') is not None:
                try:
                    self.dialog['업종별주가조회'].show()
                except Exception as e:
                    self.dialog['업종별주가조회'] = 화면_업종별주가(sScreenNo=9900, kiwoom=self.kiwoom, parent=self)
                    self.dialog['업종별주가조회'].KiwoomConnect()
                    self.dialog['업종별주가조회'].show()
            else:
                self.dialog['업종별주가조회'] = 화면_업종별주가(sScreenNo=9900, kiwoom=self.kiwoom, parent=self)
                self.dialog['업종별주가조회'].KiwoomConnect()
                self.dialog['업종별주가조회'].show()
        elif _action == "actionTickLogger":
            self.RobotAdd_TickLogger()
            self.RobotView()
        elif _action == "actionTickMonitor":
            self.RobotAdd_TickMonitor()
            self.RobotView()
        elif _action == "actionTickTradeRSI":
            self.RobotAdd_TickTradeRSI()
            self.RobotView()
        elif _action == "actionTickFuturesLogger":
            self.RobotAdd_TickFuturesLogger()
            self.RobotView()
        elif _action == "actionRobotLoad":
            self.RobotLoad()
            self.RobotView()
        elif _action == "actionRobotSave":
            self.RobotSave()
        elif _action == "actionRobotOneRun":
            self.RobotOneRun()
            self.RobotView()
        elif _action == "actionRobotOneStop":
            self.RobotOneStop()
            self.RobotView()
        elif _action == "actionRobotRun":
            self.RobotRun()
            self.RobotView()
        elif _action == "actionRobotStop":
            self.RobotStop()
            self.RobotView()
        elif _action == "actionRobotRemove":
            self.RobotRemove()
            self.RobotView()
        elif _action == "actionRobotClear":
            self.RobotClear()
            self.RobotView()
        elif _action == "actionRobotView":
            self.RobotView()
            for r in self.robots:
                logger.debug('%s %s %s %s' % (r.sName, r.UUID, len(r.portfolio), r.GetStatus()))
        elif _action == "actionCodeBuild":
            self.종목코드 = self.StockCodeBuild(to_db=True)
            self.업종코드 = self.SectorCodeBuild(to_db=True)
            QMessageBox.about(self, "종목코드 생성"," %s 항목의 종목코드를 생성하였습니다." % (len(self.종목코드.index)))
        elif _action == "actionBackup2":
            QTimer.singleShot(주문지연, lambda: self.Backup(작업=None))
        elif _action == "actionOpenAPI_document":
            self.kiwoom_doc()
        elif _action == "actionTEST":
            futurecodelist = self.kiwoom.dynamicCall('GetFutureList')
            codes = futurecodelist.split(';')
            print(futurecodelist)


    #-------------------------------------------
    # 키움증권 OpenAPI
    def KiwoomAPI(self):
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

    def KiwoomConnect(self):
        self.kiwoom.OnEventConnect[int].connect(self.OnEventConnect)
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].connect(self.OnReceiveChejanData)
        self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].connect(self.OnReceiveRealData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnEventConnect[int].disconnect(self.OnEventConnect)
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].disconnect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].disconnect(self.OnReceiveChejanData)
        self.kiwoom.OnReceiveConditionVer[int, str].disconnect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].disconnect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].disconnect(self.OnReceiveRealData)

    def KiwoomLogin(self):
        self.kiwoom.dynamicCall("CommConnect()")
        self._login = True

    def KiwoomLogout(self):
        if self.kiwoom is not None:
            self.kiwoom.dynamicCall("CommTerminate()")

        self.statusbar.showMessage("해제됨...")

    def KiwoomAccount(self):
        ACCOUNT_CNT = self.kiwoom.dynamicCall('GetLoginInfo("ACCOUNT_CNT")')
        ACC_NO = self.kiwoom.dynamicCall('GetLoginInfo("ACCNO")')

        self.account = ACC_NO.split(';')[0:-1]

        return (ACCOUNT_CNT, ACC_NO)

    def KiwoomSendOrder(self, sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo):
        if self.주문제한 < 초당횟수제한:
            Order = self.kiwoom.dynamicCall('SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                                        [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo])
            self.주문제한 += 1
            return (True, Order)
        else:
            return (False, 0)

        # -거래구분값 확인(2자리)
        #
        # 00 : 지정가
        # 03 : 시장가
        # 05 : 조건부지정가
        # 06 : 최유리지정가
        # 07 : 최우선지정가
        # 10 : 지정가IOC
        # 13 : 시장가IOC
        # 16 : 최유리IOC
        # 20 : 지정가FOK
        # 23 : 시장가FOK
        # 26 : 최유리FOK
        # 61 : 장전 시간외단일가매매
        # 81 : 장후 시간외종가
        # 62 : 시간외단일가매매
        #
        # -매매구분값 (1 자리)
        # 1 : 신규매수
        # 2 : 신규매도
        # 3 : 매수취소
        # 4 : 매도취소
        # 5 : 매수정정
        # 6 : 매도정정

    def KiwoomSetRealReg(self, sScreenNo, sCode, sRealType='0'):
        ret = self.kiwoom.dynamicCall('SetRealReg(QString, QString, QString, QString)', sScreenNo, sCode, '9001;10', sRealType)
        return ret

    def KiwoomSetRealRemove(self, sScreenNo, sCode):
        ret = self.kiwoom.dynamicCall('SetRealRemove(QString, QString)', sScreenNo, sCode)
        return ret

    def KiwoomScreenNumber(self):
        self.screen_number += 1
        if self.screen_number > 8999:
            self.screen_number = 5000
        return self.screen_number

    def OnEventConnect(self, nErrCode):
        # logger.debug('main:OnEventConnect', nErrCode)

        if nErrCode == 0:
            self.kiwoom.dynamicCall("KOA_Functions(QString, QString)", ["ShowAccountWindow", ""])
        else:
            self.statusbar.showMessage("연결실패... %s" % nErrCode)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        # logger.debug('main:OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))
        pass

    def OnReceiveTrCondition(self, sScrNo, strCodeList, strConditionName, nIndex, nNext):
        logger.debug('main:OnReceiveTrCondition [%s] [%s] [%s] [%s] [%s]' % (sScrNo, strCodeList, strConditionName, nIndex, nNext))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('main:OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.ScreenNumber != int(sScrNo):
            return

        if sRQName == "주식일봉차트조회":
            self.주식일봉컬럼 = ['일자', '현재가', '거래량', '시가', '고가', '저가', '거래대금']

            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.주식일봉컬럼:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-'+S[1:].lstrip('0')
                    row.append( S )
                self.종목일봉.append(row)
            if sPreNext == '2' and False: # 과거 모든데이타 백업시 True로 변경할것
                QTimer.singleShot(주문지연, lambda : self.ReguestPriceDaily(_repeat=2))
            else:
                df = DataFrame(data=self.종목일봉, columns=self.주식일봉컬럼)
                df['일자'] = df['일자'].apply(lambda x: x[0:4] + '-'  + x[4:6] + '-' +x[6:])
                df['종목코드'] = self.종목코드[0]
                df = df[['종목코드','일자','현재가','시가','고가','저가','거래량','거래대금']]
                values = list(df.values)

                try:
                    df.ix[df.현재가 == '', ['현재가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.시가 == '', ['시가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.고가 == '', ['고가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.저가 == '', ['저가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래량 == '', ['거래량']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래대금 == '', ['거래대금']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래대금 == '-', ['거래대금']] = 0
                except Exception as e:
                    pass

                conn = mysqlconn()

                cursor = conn.cursor()
                cursor.executemany("replace into 일별주가(종목코드,일자,종가,시가,고가,저가,거래량,거래대금) values( %s, %s, %s, %s, %s, %s, %s, %s )", df.values.tolist())

                conn.commit()
                conn.close()

                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.종목일봉 = []

                    QTimer.singleShot(주문지연, lambda : self.ReguestPriceDaily(_repeat=0))
                else:
                    QTimer.singleShot(주문지연, lambda: self.Backup(작업="주식일봉백업"))

        if sRQName == "종목별투자자조회":
            self.종목별투자자컬럼 = ['일자', '현재가', '전일대비', '누적거래대금', '개인투자자', '외국인투자자','기관계','금융투자','보험','투신','기타금융','은행','연기금등','국가','내외국인','사모펀드','기타법인']

            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.종목별투자자컬럼:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0').replace('--', '-')
                    row.append(S)
                self.종목별투자자.append(row)
            if sPreNext == '2' and False:
                QTimer.singleShot(주문지연, lambda: self.RequestInvestorDaily(_repeat=2))
            else:
                if len(self.종목별투자자) > 0:
                    df = DataFrame(data=self.종목별투자자, columns=self.종목별투자자컬럼)
                    # df['일자'] = pd.to_datetime(df['일자'], format='%Y%m%d')
                    df['일자'] = df['일자'].apply(lambda x: x[0:4] + '-' + x[4:6] + '-' + x[6:])
                    # df['현재가'] = np.abs(df['현재가'].convert_objects(convert_numeric=True))
                    df['현재가'] = np.abs(pd.to_numeric(df['현재가'], errors='coerce'))
                    df['종목코드'] = self.종목코드[0]
                    df = df[['종목코드'] + self.종목별투자자컬럼]
                    # values = list(df.values)

                    try:
                        df.ix[df.현재가 == '', ['현재가']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.전일대비 == '', ['전일대비']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.누적거래대금 == '', ['누적거래대금']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.개인투자자 == '', ['개인투자자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.외국인투자자 == '', ['외국인투자자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기관계 == '', ['기관계']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.금융투자 == '', ['금융투자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.금융투자 == '', ['금융투자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.보험 == '', ['보험']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.투신 == '', ['투신']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기타금융 == '', ['기타금융']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.은행 == '', ['은행']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.연기금등 == '', ['연기금등']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.국가 == '', ['국가']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.내외국인 == '', ['내외국인']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.사모펀드 == '', ['사모펀드']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기타법인 == '', ['기타법인']] = 0
                    except Exception as e:
                        pass

                    df.dropna(inplace=True)

                    conn = mysqlconn()

                    cursor = conn.cursor()
                    cursor.executemany(
                        "replace into 종목별투자자(종목코드,일자,종가,전일대비,누적거래대금,개인투자자,외국인투자자,기관계,금융투자,보험,투신,기타금융,은행,연기금등,국가,내외국인,사모펀드,기타법인) values(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                        df.values.tolist())
                    conn.commit()
                    conn.close()

                else:
                    logger.info("%s 데이타없음",self.종목코드)

                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.종목별투자자 = []

                    QTimer.singleShot(주문지연, lambda: self.RequestInvestorDaily(_repeat=0))
                else:
                    QTimer.singleShot(주문지연, lambda: self.Backup(작업="종목별투자자백업"))

        if sRQName == "주식분봉차트조회":
            self.주식분봉컬럼 = ['체결시간', '현재가', '시가', '고가', '저가', '거래량']

            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.주식분봉컬럼:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and (S[0] == '-' or S[0] == '+'):
                        S = S[1:].lstrip('0')
                    row.append(S)
                self.종목분봉.append(row)
            if sPreNext == '2' and False:
                QTimer.singleShot(주문지연, lambda: self.ReguestPriceMin(_repeat=2))
            else:
                df = DataFrame(data=self.종목분봉, columns=self.주식분봉컬럼)
                df['체결시간'] = df['체결시간'].apply(
                    lambda x: x[0:4] + '-' + x[4:6] + '-' + x[6:8] + ' ' + x[8:10] + ':' + x[10:12] + ':' + x[12:])
                df['종목코드'] = self.종목코드[0]
                df['틱범위'] = self.틱범위
                df = df[['종목코드', '틱범위', '체결시간', '현재가', '시가', '고가', '저가', '거래량']]
                values = list(df.values)

                try:
                    df.ix[df.현재가 == '', ['현재가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.시가 == '', ['시가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.고가 == '', ['고가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.저가 == '', ['저가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래량 == '', ['거래량']] = 0
                except Exception as e:
                    pass

                conn = mysqlconn()

                cursor = conn.cursor()
                cursor.executemany(
                    "replace into 분별주가(종목코드,틱범위,체결시간,종가,시가,고가,저가,거래량) values( %s, %s, %s, %s, %s, %s, %s, %s )",
                    df.values.tolist())

                conn.commit()
                conn.close()

                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.종목분봉 = []

                    QTimer.singleShot(주문지연, lambda: self.ReguestPriceMin(_repeat=0))
                else:
                    QTimer.singleShot(주문지연, lambda: self.Backup(작업="주식분봉백업"))

    def OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        # logger.debug('main:OnReceiveChejanData [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
        pass

        # sFid별 주요데이터는 다음과 같습니다.
        # "9201" : "계좌번호"
        # "9203" : "주문번호"
        # "9001" : "종목코드"
        # "913" : "주문상태"
        # "302" : "종목명"
        # "900" : "주문수량"
        # "901" : "주문가격"
        # "902" : "미체결수량"
        # "903" : "체결누계금액"
        # "904" : "원주문번호"
        # "905" : "주문구분"
        # "906" : "매매구분"
        # "907" : "매도수구분"
        # "908" : "주문/체결시간"
        # "909" : "체결번호"
        # "910" : "체결가"
        # "911" : "체결량"
        # "10" : "현재가"
        # "27" : "(최우선)매도호가"
        # "28" : "(최우선)매수호가"
        # "914" : "단위체결가"
        # "915" : "단위체결량"
        # "919" : "거부사유"
        # "920" : "화면번호"
        # "917" : "신용구분"
        # "916" : "대출일"
        # "930" : "보유수량"
        # "931" : "매입단가"
        # "932" : "총매입가"
        # "933" : "주문가능수량"
        # "945" : "당일순매수수량"
        # "946" : "매도/매수구분"
        # "950" : "당일총매도손일"
        # "951" : "예수금"
        # "307" : "기준가"
        # "8019" : "손익율"
        # "957" : "신용금액"
        # "958" : "신용이자"
        # "918" : "만기일"
        # "990" : "당일실현손익(유가)"
        # "991" : "당일실현손익률(유가)"
        # "992" : "당일실현손익(신용)"
        # "993" : "당일실현손익률(신용)"
        # "397" : "파생상품거래단위"
        # "305" : "상한가"
        # "306" : "하한가"

    def OnReceiveConditionVer(self, lRet, sMsg):
        logger.debug('main:OnReceiveConditionVer : [이벤트] 조건식 저장', lRet, sMsg)

    def OnReceiveRealCondition(self, sTrCode, strType, strConditionName, strConditionIndex):
        logger.debug('main:OnReceiveRealCondition [%s] [%s] [%s] [%s]' % (sTrCode, strType, strConditionName, strConditionIndex))

    def OnReceiveRealData(self, sRealKey, sRealType, sRealData):
        # logger.debug('main:OnReceiveRealData [%s] [%s] [%s]' % (sRealKey, sRealType, sRealData))
        pass

    # ------------------------------------------------------------
    # robot 함수

    def GetUnAssignedScreenNumber(self):
        스크린번호 = 0
        사용중인스크린번호 = []
        for r in self.robots:
            사용중인스크린번호.append(r.sScreenNo)

        for i in range(로봇스크린번호시작, 로봇스크린번호종료+1):
            if i not in 사용중인스크린번호:
                스크린번호 = i
                break
        return 스크린번호

    def RobotRemove(self):
        RobotUUID = self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row()+1]['RobotID'].values[0]

        robot_found = None
        for r in self.robots:
            if r.UUID == RobotUUID:
                robot_found = r
                break

        if robot_found == None:
            return

        reply = QMessageBox.question(self,
                 "로봇 삭제", "로봇을 삭제할까요?\n%s" % robot_found.GetStatus()[0:4],
                 QMessageBox.Yes|QMessageBox.No|QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass
        elif reply == QMessageBox.No:
            pass
        elif reply == QMessageBox.Yes:
            self.robots.remove(robot_found)

    def RobotClear(self):
        reply = QMessageBox.question(self,
                 "로봇 전체 삭제", "로봇 전체를 삭제할까요?",
                 QMessageBox.Yes|QMessageBox.No|QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass
        elif reply == QMessageBox.No:
            pass
        elif reply == QMessageBox.Yes:
            self.robots = []

    def RobotRun(self):
        for r in self.robots:
            r.초기조건()
            # logger.debug('%s %s %s %s' % (r.sName, r.UUID, len(r.portfolio), r.GetStatus()))
            r.Run(flag=True, sAccount=로봇거래계좌번호)

        self.statusbar.showMessage("RUN !!!")

    def RobotStop(self):
        reply = QMessageBox.question(self,
                 "전체 로봇 실행 중지", "전체 로봇 실행을 중지할까요?",
                 QMessageBox.Yes|QMessageBox.No|QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass
        elif reply == QMessageBox.No:
            pass
        elif reply == QMessageBox.Yes:
            for r in self.robots:
                r.Run(flag=False)

            self.RobotSaveSilently()

        self.statusbar.showMessage("STOP !!!")

    def RobotOneRun(self):
        try:
            RobotUUID = self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row()+1]['RobotID'].values[0]
        except Exception as e:
            RobotUUID = ''

        robot_found = None
        for r in self.robots:
            if r.UUID == RobotUUID:
                robot_found = r
                break

        if robot_found == None:
            return

        robot_found.Run(flag=True, sAccount=로봇거래계좌번호)

    def RobotOneStop(self):
        try:
            RobotUUID = self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row()+1]['RobotID'].values[0]
        except Exception as e:
            RobotUUID = ''

        robot_found = None
        for r in self.robots:
            if r.UUID == RobotUUID:
                robot_found = r
                break

        if robot_found == None:
            return

        reply = QMessageBox.question(self,
                 "로봇 실행 중지", "로봇 실행을 중지할까요?\n%s" % robot_found.GetStatus(),
                 QMessageBox.Yes|QMessageBox.No|QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass
        elif reply == QMessageBox.No:
            pass
        elif reply == QMessageBox.Yes:
            robot_found.Run(flag=False)

    def RobotLoad(self):
        reply = QMessageBox.question(self,
                 "로봇 탑제", "저장된 로봇을 읽어올까요?",
                 QMessageBox.Yes|QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass

        elif reply == QMessageBox.Yes:
            conn = mysqlconn()

            cursor = conn.cursor()

            cursor.execute('select uuid, strategy, name, robot from mymoneybot_robots')

            self.robots = []
            for row in cursor.fetchall():
                uuid, strategy, name, robot_encoded = row

                robot = base64.decodebytes(bytes(robot_encoded, 'ascii'))
                r = pickle.loads(robot)

                r.kiwoom = self.kiwoom
                r.parent = self
                r.d = datetime.date.today()

                r.running = False
                # logger.debug(r.sName, r.UUID, len(r.portfolio))
                self.robots.append(r)

            conn.close()

            self.RobotView()

    def RobotSave(self):
        reply = QMessageBox.question(self,
                 "로봇 저장", "현재 로봇을 저장할까요?",
                 QMessageBox.Yes|QMessageBox.No|QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass
        elif reply == QMessageBox.No:
            pass
        elif reply == QMessageBox.Yes:
            self.RobotSaveSilently()

    def RobotSaveSilently(self):
        conn = mysqlconn()
        cursor = conn.cursor()

        cursor.execute('delete from mymoneybot_robots')
        conn.commit()

        for r in self.robots:
            r.kiwoom = None
            r.parent = None

            uuid = r.UUID
            strategy = r.__class__.__name__
            name = r.sName
            robot = pickle.dumps(r, protocol=pickle.HIGHEST_PROTOCOL, fix_imports=True)
            robot_encoded = base64.encodestring(robot)

            cursor.execute("REPLACE into mymoneybot_robots(uuid, strategy, name, robot) values (%s, %s, %s, %s)",
                           [uuid, strategy, name, robot_encoded])
            conn.commit()

            r.kiwoom = self.kiwoom
            r.parent = self

        conn.close()

    def RobotView(self):
        result = []
        for r in self.robots:
            # logger.debug('%s %s %s %s' % (r.sName, r.UUID, len(r.portfolio), r.GetStatus()))
            result.append(r.GetStatus())

        self.model.update(DataFrame(data=result, columns=self.robot_columns))

        # RobotID 숨김
        self.tableView_robot.setColumnHidden(2, True)

        for i in range(len(self.robot_columns)):
            self.tableView_robot.resizeColumnToContents(i)

        # self.tableView_robot.horizontalHeader().setStretchLastSection(True)

    def RobotEdit(self, QModelIndex):
        # print(self.model._data[QModelIndex.row()])
        Robot타입 = self.model._data[QModelIndex.row():QModelIndex.row()+1]['Robot타입'].values[0]
        RobotUUID = self.model._data[QModelIndex.row():QModelIndex.row()+1]['RobotID'].values[0]
        # print(Robot타입, RobotUUID)

        robot_found = None
        for r in self.robots:
            if r.UUID == RobotUUID:
                robot_found = r
                break

        if robot_found == None:
            return

        if Robot타입 == 'CTickLogger':
            self.RobotEdit_TickLogger(robot_found)
        elif Robot타입 == 'CTickMonitor':
            self.RobotEdit_TickMonitor(robot_found)
        elif Robot타입 == 'CTickTradeRSI':
            self.RobotEdit_TickTradeRSI(robot_found)
        elif Robot타입 == 'CTickFuturesLogger':
            self.RobotEdit_TickFuturesLogger(robot_found)

    def RobotAdd_TickLogger(self):
        스크린번호 = self.GetUnAssignedScreenNumber()
        R = 화면_TickLogger(parent=self)
        R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            self.KiwoomAccount()
            r = CTickLogger(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

            self.robots.append(r)

    def RobotEdit_TickLogger(self, robot):
        R = 화면_TickLogger(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.plainTextEdit_base_price.setPlainText(','.join([str(x) for x in robot.종목유니버스]))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

    def RobotAdd_TickMonitor(self):
        스크린번호 = self.GetUnAssignedScreenNumber()
        R = 화면_TickLogger(parent=self)
        R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            self.KiwoomAccount()
            r = CTickMonitor(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

            self.robots.append(r)

    def RobotEdit_TickMonitor(self, robot):
        R = 화면_TickLogger(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.plainTextEdit_base_price.setPlainText(','.join([str(x) for x in robot.종목유니버스]))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

    def RobotAdd_TickTradeRSI(self):
        스크린번호 = self.GetUnAssignedScreenNumber()
        R = 화면_TickTradeRSI(parent=self)
        R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            단위투자금 = int(R.lineEdit_unit.text()) * 10000
            매수방법 = R.comboBox_buy_sHogaGb.currentText().strip()[0:2]
            매도방법 = R.comboBox_sell_sHogaGb.currentText().strip()[0:2]
            시총상한 = int(R.lineEdit_max.text().strip())
            시총하한 = int(R.lineEdit_min.text().strip())
            포트폴리오수 = int(R.lineEdit_portsize.text().strip())

            r = CTickTradeRSI(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 단위투자금=단위투자금, 시총상한=시총상한, 시총하한=시총하한, 포트폴리오수=포트폴리오수, 매수방법=매수방법, 매도방법=매도방법)

            self.robots.append(r)

    def RobotEdit_TickTradeRSI(self, robot):
        R = 화면_TickTradeRSI(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.lineEdit_unit.setText(str(robot.단위투자금 // 10000))
        R.lineEdit_portsize.setText(str(robot.포트폴리오수))
        R.lineEdit_max.setText(str(robot.시총상한))
        R.lineEdit_min.setText(str(robot.시총하한))
        R.comboBox_buy_sHogaGb.setCurrentIndex(R.comboBox_buy_sHogaGb.findText(robot.매수방법, flags=Qt.MatchContains))
        R.comboBox_sell_sHogaGb.setCurrentIndex(R.comboBox_sell_sHogaGb.findText(robot.매도방법, flags=Qt.MatchContains))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            단위투자금 = int(R.lineEdit_unit.text()) * 10000
            매수방법 = R.comboBox_buy_sHogaGb.currentText().strip()[0:2]
            매도방법 = R.comboBox_sell_sHogaGb.currentText().strip()[0:2]

            시총상한 = int(R.lineEdit_max.text().strip())
            시총하한 = int(R.lineEdit_min.text().strip())
            포트폴리오수 = int(R.lineEdit_portsize.text().strip())

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 단위투자금=단위투자금, 시총상한=시총상한, 시총하한=시총하한, 포트폴리오수=포트폴리오수, 매수방법=매수방법, 매도방법=매도방법)

    def RobotAdd_TickFuturesLogger(self):
        스크린번호 = self.GetUnAssignedScreenNumber()
        R = 화면_TickFuturesLogger(parent=self)
        R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            self.KiwoomAccount()
            r = CTickFuturesLogger(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

            self.robots.append(r)

    def RobotEdit_TickFuturesLogger(self, robot):
        R = 화면_TickFuturesLogger(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.plainTextEdit_base_price.setPlainText(','.join([str(x) for x in robot.종목유니버스]))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

    #-------------------------------------------
    # UI 관련함수
    def SectorCodeBuild(self, to_db=False):
        result = [
            ['KOSPI', '001', '종합(KOSPI)'],
            ['KOSPI', '002', '대형주'],
            ['KOSPI', '003', '중형주'],
            ['KOSPI', '004', '소형주'],
            ['KOSPI', '005', '음식료업'],
            ['KOSPI', '006', '섬유의복'],
            ['KOSPI', '007', '종이목재'],
            ['KOSPI', '008', '화학'],
            ['KOSPI', '009', '의약품'],
            ['KOSPI', '010', '비금속광물'],
            ['KOSPI', '011', '철강금속'],
            ['KOSPI', '012', '기계'],
            ['KOSPI', '013', '전기전자'],
            ['KOSPI', '014', '의료정밀'],
            ['KOSPI', '015', '운수장비'],
            ['KOSPI', '016', '유통업'],
            ['KOSPI', '017', '전기가스업'],
            ['KOSPI', '018', '건설업'],
            ['KOSPI', '019', '운수창고'],
            ['KOSPI', '020', '통신업'],
            ['KOSPI', '021', '금융업'],
            ['KOSPI', '022', '은행'],
            ['KOSPI', '024', '증권'],
            ['KOSPI', '025', '보험'],
            ['KOSPI', '026', '서비스업'],
            ['KOSPI', '027', '제조업'],
            ['KOSPI', '603', '변동성지수'],
            ['KOSPI', '604', '코스피고배당50'],
            ['KOSPI', '605', '코스피배당성장50'],
            ['KOSDAQ', '101', '종합(KOSDAQ)'],
            ['KOSDAQ', '102', '벤처지수'],
            ['KOSDAQ', '103', '기타서비스'],
            ['KOSDAQ', '104', '코스닥IT종합'],
            ['KOSDAQ', '105', '코스닥IT벤처'],
            ['KOSDAQ', '106', '제조'],
            ['KOSDAQ', '107', '건설'],
            ['KOSDAQ', '108', '유통'],
            ['KOSDAQ', '109', '숙박/음식'],
            ['KOSDAQ', '110', '운송'],
            ['KOSDAQ', '111', '금융'],
            ['KOSDAQ', '112', '통신방송서비스'],
            ['KOSDAQ', '113', 'IT S/W & SVC'],
            ['KOSDAQ', '114', 'IT H/W'],
            ['KOSDAQ', '115', '음식료/담배'],
            ['KOSDAQ', '116', '섬유/의류'],
            ['KOSDAQ', '117', '종이/목재'],
            ['KOSDAQ', '118', '출판/매체복제'],
            ['KOSDAQ', '119', '화학'],
            ['KOSDAQ', '120', '제약'],
            ['KOSDAQ', '121', '비금속'],
            ['KOSDAQ', '122', '금속'],
            ['KOSDAQ', '123', '기계/장비'],
            ['KOSDAQ', '124', '일반전기전자'],
            ['KOSDAQ', '125', '의료/정밀 기기'],
            ['KOSDAQ', '126', '운송장비/부품'],
            ['KOSDAQ', '127', '기타 제조'],
            ['KOSDAQ', '128', '통신서비스'],
            ['KOSDAQ', '129', '방송서비스'],
            ['KOSDAQ', '130', '인터넷'],
            ['KOSDAQ', '131', '디지털컨텐츠'],
            ['KOSDAQ', '132', '소프트웨어'],
            ['KOSDAQ', '133', '컴퓨터서비스'],
            ['KOSDAQ', '134', '통신장비'],
            ['KOSDAQ', '135', '정보기기'],
            ['KOSDAQ', '136', '반도체'],
            ['KOSDAQ', '137', 'IT 부품'],
            ['KOSDAQ', '138', 'KOSDAQ 100'],
            ['KOSDAQ', '139', 'KOSDAQ MID 300'],
            ['KOSDAQ', '140', 'KOSDAQ SMALL'],
            ['KOSDAQ', '141', '오락,문화'],
            ['KOSDAQ', '142', '코스닥 우량기업'],
            ['KOSDAQ', '143', '코스닥 벤처기업'],
            ['KOSDAQ', '144', '코스닥 중견기업'],
            ['KOSDAQ', '145', '코스닥 신성장기업'],
            ['KOSDAQ', '150', 'KOSDAQ 150'],
            ['KOSDAQ', '302', 'KOSDAQ스타30'],
            ['KOSDAQ', '303', '프리미어지수'],
            ['KOSPI200', '201', 'KOSPI200'],
            ['KOSPI200', '207', 'F-KOSPI200'],
            ['KOSPI200', '208', 'F-KOSPI200인버스'],
            ['KOSPI200', '209', '레버리지KOSPI200'],
            ['KOSPI200', '211', '건설'],
            ['KOSPI200', '212', '중공업'],
            ['KOSPI200', '213', '철강소재'],
            ['KOSPI200', '214', '에너지화학'],
            ['KOSPI200', '215', '정보기술'],
            ['KOSPI200', '216', '금융'],
            ['KOSPI200', '217', '생활소비재'],
            ['KOSPI200', '218', '경기소비재'],
            ['KOSPI200', '224', '에너지 화학 레버리지'],
            ['KOSPI200', '225', '정보기술 레버리지'],
            ['KOSPI200', '226', '금융 레버리지'],
            ['KOSPI200', '227', '경기소비재 레버리지'],
            ['KOSPI200', '250', 'F-KOSPI200 Plus'],
            ['KOSPI200', '251', 'K200 USD F BuySell'],
            ['KOSPI200', '252', 'USD K200 F BuySell'],
            ['KOSPI200', '253', 'K200 F 매수 콜매도'],
            ['KOSPI200', '254', 'K200 F 매도 풋매도'],
            ['KOSPI200', '255', 'K200 중소형주'],
            ['KOSPI100', '401', 'KOSPI100'],
            ['KOSPI100', '402', 'KOSPI50'],
            ['KRX100', '701', 'KRX100'],
            ['KRX100', '702', 'KRX자동차'],
            ['KRX100', '703', 'KRX반도체'],
            ['KRX100', '704', 'KRX바이오'],
            ['KRX100', '705', 'KRX금융'],
            ['KRX100', '707', 'KRX화학에너지'],
            ['KRX100', '708', 'KRX철강'],
            ['KRX100', '710', 'KRX미디어통신'],
            ['KRX100', '711', 'KRX건설'],
            ['KRX100', '713', 'KRX증권'],
            ['KRX100', '714', 'KRX조선'],
            ['KRX100', '715', 'KRX보험'],
            ['KRX100', '716', 'KRX운송'],
            ['KRX100', '721', '사회책임투자지수'],
            ['KRX100', '722', '환경책임투자지수'],
            ['KRX100', '723', '녹색산업지수'],
            ['KRX100', '724', '지배구조우수기업'],
            ['KRX100', '730', 'KTOP 30'],
            ['KRX100', '731', 'KTOP30레버리지'],
            ['KRX100', '750', 'KRX ESG 리더스150']
        ]
        df_code = DataFrame(data=result, columns=['시장구분','업종코드','업종명'])

        if to_db == True:
            conn = mysqlconn()

            cursor = conn.cursor()
            cursor.executemany("replace into 업종코드(시장구분,업종코드,업종명) values( %s, %s, %s )", df_code.values.tolist())
            conn.commit()
            conn.close()

        return df_code

    def StockCodeBuild(self, to_db=False):
        result = []
        markets = [['0','KOSPI'], ['10','KOSDAQ'], ['8','ETF']]
        for [marketcode, marketname] in markets:
            codelist = self.kiwoom.dynamicCall('GetCodeListByMarket(QString)', [marketcode]) # sMarket – 0:장내, 3:ELW, 4:뮤추얼펀드, 5:신주인수권, 6:리츠, 8:ETF, 9:하이일드펀드, 10:코스닥, 30:제3시장
            codes = codelist.split(';')
            for code in codes:
                if code is not '':
                    종목명 = self.kiwoom.dynamicCall('GetMasterCodeName(QString)', [code])
                    주식수 = self.kiwoom.dynamicCall('GetMasterListedStockCnt(QString)', [code])
                    감리구분 = self.kiwoom.dynamicCall('GetMasterConstruction(QString)', [code]) # 감리구분 – 정상, 투자주의, 투자경고, 투자위험, 투자주의환기종목
                    상장일 = datetime.datetime.strptime(self.kiwoom.dynamicCall('GetMasterListedStockDate(QString)', [code]),'%Y%m%d')
                    전일종가 = int(self.kiwoom.dynamicCall('GetMasterLastPrice(QString)', [code]))
                    종목상태 = self.kiwoom.dynamicCall('GetMasterStockState(QString)', [code]) # 종목상태 – 정상, 증거금100%, 거래정지, 관리종목, 감리종목, 투자유의종목, 담보대출, 액면분할, 신용가능

                    result.append([marketname, code, 종목명, 주식수, 감리구분, 상장일, 전일종가, 종목상태])

        df_code = DataFrame(data=result, columns=['시장구분', '종목코드', '종목명', '주식수', '감리구분', '상장일', '전일종가', '종목상태'])
        # df.set_index('종목코드', inplace=True)

        if to_db == True:
            # 테마코드
            themecodes = []
            ret = self.kiwoom.dynamicCall('GetThemeGroupList(int)', [1]).split(';')
            for item in ret:
                [code, name] = item.split('|')
                themecodes.append([code, name])
            # print(themecodes)
            df_theme = DataFrame(data=themecodes, columns=['테마코드', '테마명'])

            # 테마구성종목
            themestocks = []
            for code, name in themecodes:
                codes = self.kiwoom.dynamicCall('GetThemeGroupCode(QString)', [code]).replace('A','').split(';')
                for c in codes:
                    themestocks.append([code, c])
            # print(themestocks)
            df_themecode = DataFrame(data=themestocks, columns=['테마코드', '구성종목'])
			
            df_code['상장일'] = df_code['상장일'].apply(lambda x: (x.to_datetime()).strftime('%Y-%m-%d %H:%M:%S'))
            # print(df_code.values.tolist())
            conn = mysqlconn()

            cursor = conn.cursor()
            cursor.executemany("replace into 종목코드(시장구분,종목코드,종목명,주식수,감리구분,상장일,전일종가,종목상태 ) values( %s, %s, %s, %s, %s, %s, %s, %s )", df_code.values.tolist())
            conn.commit()

            cursor.executemany("replace into 테마코드(테마코드, 테마명) VALUES (%s, %s)", df_theme.values.tolist())
            conn.commit()

            cursor.executemany("replace into 테마종목(테마코드, 구성종목) VALUES (%s, %s)", df_themecode.values.tolist())
            conn.commit()

            conn.close()

        return df_code


    # 유틸리티 함수
    def kiwoom_doc(self):
        kiwoom = QAxContainer.QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        _doc = kiwoom.generateDocumentation()
        f = open("openapi_doc.html", 'w')
        f.write(_doc)
        f.close()

    def ReguestPriceDaily(self, _repeat=0):
        # logger.info("주식일봉백업: %s" % self.종목코드)
        self.statusbar.showMessage("주식일봉백업: %s" % self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "기준일자", self.기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식일봉차트조회", "OPT10081", _repeat, '{:04d}'.format(self.ScreenNumber))

    def BackupPriceDaily(self):
        # 주식일봉차트조회
        self.종목코드테이블 = self.StockCodeBuild().copy()

        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[['종목코드', '종목명']].values)
        self.종목코드 = self.백업할종목코드.pop(0)
        self.기준일자 = datetime.datetime.now().strftime('%Y%m%d')

        self.종목일봉 = []
        self.ReguestPriceDaily(_repeat=0)

    def ReguestPriceMin(self, _repeat=0):
        # logger.info("주식분봉백업: %s" % self.종목코드)
        self.statusbar.showMessage("주식분봉백업: %s" % self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "틱범위", self.틱범위)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식분봉차트조회", "OPT10080", _repeat, '{:04d}'.format(self.ScreenNumber))

    def BackupPriceMin(self):
        # 주식분봉차트조회
        self.종목코드테이블 = self.StockCodeBuild().copy()

        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[['종목코드', '종목명']].values)
        self.종목코드 = self.백업할종목코드.pop(0)

        self.종목분봉 = []
        self.틱범위 = "01"
        self.ReguestPriceMin(_repeat=0)

    def RequestInvestorDaily(self, _repeat=0):
        # logger.info("종목별투자자백업: %s" % self.종목코드)
        self.statusbar.showMessage("종목별투자자백업: %s" % self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "일자", self.기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "금액수량구분", 2)  # 1:금액, 2:수량
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "매매구분", 0)  # 0:순매수, 1:매수, 2:매도
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "단위구분", 1)  # 1000:천주, 1:단주
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "종목별투자자조회", "OPT10060", _repeat,
                                      '{:04d}'.format(self.ScreenNumber))

    def BackupInvestorDaily(self):
        self.종목코드테이블 = self.StockCodeBuild().copy()

        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[['종목코드', '종목명']].values)
        self.종목코드 = self.백업할종목코드.pop(0)
        self.기준일자 = datetime.datetime.now().strftime('%Y%m%d')

        self.종목별투자자 = []
        self.RequestInvestorDaily(_repeat=0)

    def Backup(self, 작업=None):
        if 작업 == None:
            for r in self.robots:
                r.Run(flag=False)

            self.RobotSaveSilently()

            self.종목코드 = self.StockCodeBuild(to_db=True)
            self.업종코드 = self.SectorCodeBuild(to_db=True)

            self.진행중인작업 = {'주식일봉백업': True, '종목별투자자백업': True}
        else:
            self.진행중인작업[작업]=False
            self.statusbar.showMessage("%s 종료" % 작업)

            # 백업이 완료되면 지표를 생성한다.
            results = False
            for k,v in self.진행중인작업.items():
                results = results or v

            if results == False:
                Popen(['python.exe', 'Scripts Daily/증권사레포트수집.py'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                logger.info("['python.exe', 'Scripts Daily/증권사레포트수집.py']")
                Popen(['python.exe', 'Scripts Daily/재무정보수집.py'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                logger.info("['python.exe', 'Scripts Daily/증권사레포트수집.py']")

        for k, v in self.진행중인작업.items():
            if v == True:
                if k == '주식일봉백업':
                    QTimer.singleShot(주문지연, lambda: self.BackupPriceDaily())
                elif k == '종목별투자자백업':
                    QTimer.singleShot(주문지연, lambda: self.BackupInvestorDaily())
                elif k == '주식분봉백업':
                    QTimer.singleShot(주문지연, lambda: self.BackupPriceMin())
                break



if __name__ == "__main__":

    # 1.로그 인스턴스를 만든다.
    logger = logging.getLogger('mymoneybot')
    # 2.formatter를 만든다.
    formatter = logging.Formatter('[%(levelname)s|%(filename)s:%(lineno)s]%(asctime)s>%(message)s')

    loggerLevel = logging.DEBUG
    filename = "LOG/mymoneybot.log"

    # 스트림과 파일로 로그를 출력하는 핸들러를 각각 만든다.
    filehandler = logging.FileHandler(filename)
    streamhandler = logging.StreamHandler()

    # 각 핸들러에 formatter를 지정한다.
    filehandler.setFormatter(formatter)
    streamhandler.setFormatter(formatter)

    # 로그 인스턴스에 스트림 핸들러와 파일 핸들러를 붙인다.
    logger.addHandler(filehandler)
    logger.addHandler(streamhandler)
    logger.setLevel(loggerLevel)
    logger.debug("=============================================================================")
    logger.info("LOG START")

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)

    window = MainWindow()
    window.show()

    QTimer().singleShot(3, window.OnQApplicationStarted)

    clock = QtCore.QTimer()
    clock.timeout.connect(window.OnClockTick)
    clock.start(1000)

    sys.exit(app.exec_())

