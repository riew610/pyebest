import win32com.client
import pythoncom
import os, sys
import inspect
import sqlite3
import time
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

class XASessionEvents:
    status = False
    
    def OnLogin(self,code,msg):
        print("OnLogin : ", code, msg)
        XASessionEvents.status = True
    
    def OnLogout(self):
        print('OnLogout')

    def OnDisconnect(self):
        print('OnDisconnect')

class XAQueryEvents:
    status = False
    
    def OnReceiveData(self,szTrCode):
        print("OnReceiveData : %s" % szTrCode)
        XAQueryEvents.status = True
    
    def OnReceiveMessage(self,systemError,messageCode,message):
        print("OnReceiveMessage : ", systemError, messageCode, message)
        XAQueryEvents.status = True

class XARealEvents:
    pass

def login(id,pwd,cert='',url='demo.ebestsec.co.kr',svrtype=0,port=200001):
    '''
    return result, error_code, message, account_list, session
    '''
    session = win32com.client.DispatchWithEvents("XA_Session.XASession",XASessionEvents)
    result = session.ConnectServer(url,port)

    if not result:
        nErrCode = session.GetLastError()
        strErrMsg = session.GetErrorMessage(nErrCode)
        return (False,nErrCode,strErrMsg,None,session)
    
    # send a message
    session.Login(id,pwd,cert,svrtype,0)
    # wait the message of Login
    while XASessionEvents.status == False:
        pythoncom.PumpWaitingMessages()

    account_list = []
    num_of_account = session.GetAccountListCount()

    for i in range(num_of_account):
        account_list.append(session.GetAccountList(i))
    
    return (True,'','', account_list, session)

def t8424(gubun1=''):
    '''
    업종전체조회
    '''
    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery",XAQueryEvents)
    
    #pathname = os.path.dirname(sys.argv[0])
    #RESDIR = os.path.abspath(pathname)
    MYNAME = inspect.currentframe().f_code.co_name
    INBLOCK = "%sInBlock" % MYNAME
    OUTBLOCK = "%sOutBlock" % MYNAME
    OUTBLOCK1 = "%sOutBlock1" %MYNAME
    RESFILE = r"C:\eBEST\xingAPI\Res\%s.res"%MYNAME

    query.LoadFromResFile(RESFILE)
    query.SetFieldData(INBLOCK,"gubun1",0,gubun1)
    query.Request(0)
    
    while XAQueryEvents.status == False:
        pythoncom.PumpWaitingMessages()
    XAQueryEvents.status = False
    
    data = []
    block_count = query.GetBlockCount(OUTBLOCK)

    for i in range(block_count):
        hname = query.GetFieldData(OUTBLOCK,"hname",i).strip()
        upcode = query.GetFieldData(OUTBLOCK,"upcode",i).strip()
        lst = [hname,upcode]
        data.append(lst)

    df = pd.DataFrame(data=data,columns=['hname','upcode'])
    return df








