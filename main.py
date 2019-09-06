from tkinter import Frame, Label, Entry
from zeep import Client
import tkinter
from tkinter import ttk
root = tkinter.Tk(  )
from xml.sax.saxutils import unescape
import re
import xlwt
from xlwt import Workbook
import datetime
import csv

form = Frame()

"""Functie om data te pakken en te tonen"""
"""Zeep Client"""

"""https://www.programcreek.com/python/example/106302/zeep.Client"""


def callback():
    Adr = ABX.get()
    Add = ABC.get()
    Ad = ABB.get()
    Ty = CBX.get()
    TyStr = str(Ty)
    Qy = QX.get()
    QyStr = str(Qy)
   
    LB1.configure(text = Adr)
    LB2.configure(text = Add)
    LB3.configure(text = Ad)

    wsdl = 'http://www.verstandigbouwen.be/advertiser/getVBData.asmx?WSDL'

    client = Client(wsdl)

    parameters = {
    'username':'VBDSM17',
    'password':'P37151',
    'product':TyStr,    
    'queryName': QyStr,
    'termsOfUse':'Yes',

    }



    with client.settings(raw_response=True):
        """soap_result = client.service.getAddresses(username='****', password='****', product ='guide', getAdddresses= , termsOfUse ='Yes')"""
        soap_result = client.service.getAddresses(**parameters)
   

    # Print out text from Requests response object returned    
    """print (soap_result.text) Dit en UnEsT is de XML dump"""

    unEsT = unescape(soap_result.text)

    AdrDict = {}
    A = []
    B = []

    mai = re.findall('<Email>(.*)</Email>', unEsT)
    dat = re.findall('<Date>(.*)</Date>', unEsT)

    """wb = Workbook()
    sheet1  = wb.add_sheet('Adres', cell_overwrite_ok=True)"""
    
    for ma in mai :
        for da in dat:
            date_str = da[0:10] 
            format_str = '%Y-%m-%d'
            datetime_obj = datetime.datetime.strptime(date_str, format_str)
            startD = '2019-04-30'
            starttime_obj = datetime.datetime.strptime(startD, format_str)
            if datetime_obj > starttime_obj:
                AdrDict.update({ma: da})
                """A.append(ma)
                B.append(da)"""

    with open('Adressen.csv', 'w') as f:
        for key in AdrDict.keys():
            f.write("%s,%s\n"%(key,AdrDict[key]))
                

                
    """t = 1
    for i in range(1, 20):
        
        sheet1.write(0,t, mai[t])
        t = t + 1
       
        
        
    tt = 1    
    for i in range(1, 20):
        
        sheet1.write(1,t, dat[t])"""
        
        
 
    txt1.insert('1.0', AdrDict)

    """for i in range(1, 1000):
        t = 1
        sheet1.write(0,i, A[t])
        sheet1.write(1,i, B[t])
        t +=1
        
    wb.save('Adressen.xls')"""

    

    

"""GUI"""
    
tkinter.Label(root, text="Adres",borderwidth=1 ).grid(row=1,column=1)
ABX = tkinter.Entry(root,borderwidth=1, width =120 )
ABX.grid(row=1,column=8)
ABX.insert(0, 'http://www.verstandigbouwen.be/advertiser/getVBData.asmx?WSDL')

tkinter.Label(root, text="Paswoord",borderwidth=1 ).grid(row=2,column=1)
ABC = tkinter.Entry(root,borderwidth=1, width =120 )
ABC.grid(row=2,column=8)
ABC.insert(0, 'P37151')

tkinter.Label(root, text="Klant",borderwidth=1 ).grid(row=3,column=1)
ABB = tkinter.Entry(root,borderwidth=1, width =120 )
ABB.grid(row=3,column=8)
ABB.insert(0, 'VBDSM17')

LB1 = tkinter.Label(root, text="Adres",borderwidth=1 )
LB1.grid(row=6,column=8)

LB2 = tkinter.Label(root, text="PasW",borderwidth=1 )
LB2.grid(row=7,column=8)

LB3 = tkinter.Label(root, text="Klant",borderwidth=1 )
LB3.grid(row=8,column=8)

BT = tkinter.Button(root, text="SEND", command=callback).grid(row=4,column=1)

CBX = ttk.Combobox(root, values = ["guide", "booklet" , "pdf"])
CBX.grid(row = 4, column = 2)

QX = ttk.Combobox(root, values = ["Nieuwbouwers"])
QX.grid(row = 5, column = 2)

txt1 = tkinter.Text(root, borderwidth = 1, width = 120)
txt1.grid(row=9, column =8)

