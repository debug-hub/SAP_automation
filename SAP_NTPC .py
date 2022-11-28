from pywinauto.application import Application
import win32gui
import win32con
import sys,os, win32com.client
import time
import subprocess
import subprocess
from openpyxl import *
import re
import datetime
import psycopg2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import time

def send_email(vendor_name,body1):
    msg = MIMEMultipart()
    addr = "pmgbillingtest@tatapower-ddl.com"
    msg["From"] = addr
    msg["To"] = "abhimanyu.yadav@sequelstring.com"
##    msg["Cc"] ="Shubham.sharma@sequelstring.com"
    msg["Cc"] =''
    msg['Subject'] = f"{vendor_name}"

    current_date   = datetime.date.today()

    msg.attach(MIMEText(body1, 'html'))  
    
    s = smtplib.SMTP('smtp.tatapower-ddl.com', 587)
    s.starttls()
    s.login("pmgbillingtest", "Power@1234")
    s.sendmail(msg["From"], msg["To"].split(",") + msg["Cc"].split(","), msg.as_string())
    s.quit()
    print("mail sent")
    return "Success"



conn= psycopg2.connect(database="Tata_power", user='postgres',password='tatapower',host='localhost',port='5432')
cursor=conn.cursor()

transmission_li=["Feroze Gandhi Unchahar TPS 1","Feroze Gandhi Unchahar TPS 2","Feroze Gandhi Unchahar TPS 3","Farakka Super Thermal Power Station","Kahalgaon STPS 2","Kahalgaon STPS 1","Rihand Super Therm Pwr Stn 1","Rihand Super Therm Pwr Stn 2",
                 "National Capital Therm Pwr - Dadri 1","National Capital Therm Pwr - Dadri 2","Dadri Gas Power Station","NTPC TRANSMISSION Charges","Singrauli Super Thermal Power Station","Singrauli Small Hydro"]


query="select distinct name_of_station  from tata_power_data where mappingsheet='NTPC' and status='pending'"
cursor.execute(query)

station_li=cursor.fetchall()

query3="select distinct invoice_date from tata_power_data where mappingsheet='NTPC' and status='pending'"
cursor.execute(query3)

invoice_li=cursor.fetchall()


#########################################

def sap_connection(connection1):
    print("xxSAP connection")
    global session,session1
    error = 'SAP Connection Error'
    
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(path)
    time.sleep(10)
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    print('hh')
    if not type(SapGuiAuto) == win32com.client.CDispatch:
            return error
    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return error
    connection = application.OpenConnection(connection1, True)
    if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return error
    if connection.DisabledByServer == True:
            application = None
            SapGuiAuto = None
            return error
    session = connection.Children(0)
    session1 = session
    if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return error
    if session.Info.IsLowSpeedConnection == True:
            connection = None
            application = None
            SapGuiAuto = None
            return error
    return 'Success'

def status_update(data_li,vendor_name,data,start_time,end_time,request):
    for jj in data_li:
        if type(jj)==str:
            jj=int(jj)-1
        x=data[jj]
        invoice_number=x[1]
        if"'" in invoice_number:
            invoice_number=invoice_number.replace("'","''")
        if request != "total invoice mismatch" or request !="total invoice number mismatch":
            query_u="UPDATE tata_power_data SET status = 'Done' WHERE invoice_number= '%s' and name_of_station='%s'"%(invoice_number,vendor_name)
            cursor.execute(query_u)
            conn.commit()
        if"''" in invoice_number:
            invoice_number=invoice_number.replace("''","'")
        query_in="insert into tata_power_log values(%s,%s,%s,%s,%s,%s)"
        row_item=[invoice_number,vendor_name,start_time,end_time,"Done",request]
        cursor.execute(query_in,row_item)
        conn.commit()
        print(invoice_number,"status changed of station",vendor_name)
        
        
    
def vendor_cc(vendor_name):
    path2=r"C:\TataPOWER\Excel_tata\PMG Vendor master.xlsx"
    wb2=load_workbook(path2,data_only=True)
    ws2=wb2["vendor master"]
    for i in range(2,145):
        station=ws2['E' + str(i)].value

        if station==vendor_name:
            vendor_code=ws2['A' + str(i)].value            
            station_code=ws2['B' + str(i)].value

            return vendor_code,station_code    


def data_rowise(j,invoice_no,station_code,unit_bill,vc,fc_peak,carrying,transmission,others,income_tax,incentive,rras,fc_offpeak,fc_offset,incentive_op,incentive_offset,tcs,remarks):
    
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/cmbWA_ZMMTR_PMG_ITEM-PLANT_VEN_CODE[1,{}]".format(j)).key = station_code
    time.sleep(2)
    #incoice no
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-INVOICE_NUM[2,{}]".format(j)).text = invoice_no
##    time.sleep(2)
    #billfromdate
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/ctxtWA_ZMMTR_PMG_ITEM-BILL_PRD_FRM[3,{}]".format(j)).text = Bill_from
##    time.sleep(2)
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/ctxtWA_ZMMTR_PMG_ITEM-BILL_PRD_FRM[3,{}]".format(j)).setFocus()
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/ctxtWA_ZMMTR_PMG_ITEM-BILL_PRD_FRM[3,{}]".format(j)).caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)
    ##unitbills
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-UNIT_BILLED[8,{}]".format(j)).text = unit_bill
    ##Variable_cost
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0002[9,{}]".format(j)).text = vc
    ##Fixed_cost_peak
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0003[10,{}]".format(j)).text = fc_peak
    ##Carying_cost
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0004[11,{}]".format(j)).text = carrying
    ##Others
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0027[12,{}]".format(j)).text = others
    time.sleep(2)
    ##Income_tax
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0028[13,{}]".format(j)).text =income_tax
    ##Incentive
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0029[14,{}]".format(j)).text =incentive
    ##RRAS charges
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0031[15,{}]".format(j)).text = rras
    ##Fixed-cost(offpeak)
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0033[16,{}]".format(j)).text = fc_offpeak
    ##Fixed_cost(off-set)
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0034[17,{}]".format(j)).text = fc_offset
    ##incentive(off-peak)
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0036[18,{}]".format(j)).text = incentive_op
    ##incentive_offset
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0037[19,{}]".format(j)).text = incentive_offset
    ##tcs
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0038[20,{}]".format(j)).text = tcs
    #remarks
    session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-REMARKS[23,{}]".format(j)).text = remarks
    
##    time.sleep(2)
    session.findById("wnd[0]/usr/txtWA_ZMMTR_PMG_HEADER-TOTAL_EST_AMT").setFocus()
    session.findById("wnd[0]/usr/txtWA_ZMMTR_PMG_HEADER-TOTAL_EST_AMT").caretPosition = 5
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

def data_transmission(j,invoice_no,station_code,unit_bill,vc,fc_peak,carrying,transmission,others,income_tax,incentive,rras,fc_offpeak,fc_offset,incentive_op,incentive_offset,tcs,remarks):
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/cmbWA_ZMMTR_PMG_ITEM-PLANT_VEN_CODE[1,{}]".format(j)).key = station_code
    time.sleep(2)
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-INVOICE_NUM[2,{}]".format(j)).text = invoice_no
##    time.sleep(2)
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/ctxtWA_ZMMTR_PMG_ITEM-BILL_PRD_FRM[3,{}]".format(j)).text = Bill_from
##    time.sleep(2)
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/ctxtWA_ZMMTR_PMG_ITEM-BILL_PRD_FRM[3,{}]".format(j)).setFocus()
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/ctxtWA_ZMMTR_PMG_ITEM-BILL_PRD_FRM[3,{}]".format(j)).caretPosition = 8
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)
    #unitbills
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-UNIT_BILLED[8,{}]".format(j)).text = unit_bill
    #Variablecost
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0002[9,{}]".format(j)).text = vc
    #fixed_cost
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0003[10,{}]".format(j)).text = fc_peak
    #Carryingcost
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0004[11,{}]".format(j)).text = carrying
    #Transmission
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0011[12,{}]".format(j)).text = transmission
    #others
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0027[13,{}]".format(j)).text = others
    time.sleep(2)
    #incometax
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0028[14,{}]".format(j)).text =income_tax
    #incentive
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0029[15,{}]".format(j)).text = incentive
    #rras
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0031[16,{}]".format(j)).text = rras
    #fixedcost_offpeak
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0033[17,{}]".format(j)).text = fc_offpeak
    #fixed_costoffset
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0034[18,{}]".format(j)).text = fc_offset
    #incentive_offpeak
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0036[19,{}]".format(j)).text = incentive_op
    #incentive_offset
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0037[20,{}]".format(j)).text = incentive_offset
    #tcs
    session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-C0038[21,{}]".format(j)).text = tcs
    #remarks
    session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM/txtWA_ZMMTR_PMG_ITEM-REMARKS[24,{}]".format(j)).text = remarks
##    time.sleep(2)
    session.findById("wnd[0]/usr/txtWA_ZMMTR_PMG_HEADER-TOTAL_EST_AMT").setFocus()
    session.findById("wnd[0]/usr/txtWA_ZMMTR_PMG_HEADER-TOTAL_EST_AMT").caretPosition = 0
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

try:
    
    for i in range(5):
            sap_conn = sap_connection(connection1 = 'ERP-QUALITY')
            print(sap_conn)
            if sap_conn == 'Success':
                
                    break
            else:
                    pass
except:
    subject='Connection Error '
    body1='Bot was Unable to connect to SAP'
    send_email(subject,body1)
	    

import time
print('download excel')
time.sleep(2)
#timedelta
#########################################

cred_querry="select * from sap_credt"
cursor.execute(cred_querry)
id_li=cursor.fetchall()
print(id_li)
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = id_li[0][0]
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = id_li[0][1]
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
session.findById("wnd[0]").sendVKey(0)
print("sleeeppp")
time.sleep(2)
try:
    xx=session.findById("/app/con[0]/ses[0]/wnd[0]/sbar").text
    print(xx)
    if "password is incorrect" in xx or "failed attempts" in xx:
        subject='Login Error '
        body1='Bot was Unable to Log In to SAP'
        send_email(subject,body1)
except:
    pass

try:
    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
except:
    pass
counter=0

session.findById("wnd[0]/tbar[0]/okcd").text = "ZPMG01"
time.sleep(2)

session.findById("wnd[0]").sendVKey(0)

transmission_li=["Feroze Gandhi Unchahar TPS 1","Feroze Gandhi Unchahar TPS 2","Feroze Gandhi Unchahar TPS 3","Farakka Super Thermal Power Station","Kahalgaon STPS 2","Kahalgaon STPS 1","Rihand Super Therm Pwr Stn 1","Rihand Super Therm Pwr Stn 2",
                 "National Capital Therm Pwr - Dadri 1","National Capital Therm Pwr - Dadri 2","Dadri Gas Power Station","NTPC TRANSMISSION Charges","Singrauli Super Thermal Power Station","Singrauli Small Hydro"]

print(invoice_li)
print(station_li)
for ii in station_li:
    start_time=datetime.datetime.now()
    vendor_name=ii[0]
    value=vendor_name
    print(vendor_name)
    print(ii)

    for i in invoice_li:
        
        invoice_date=i[0]
        print(invoice_date)
        value2=invoice_date
        query1="select name_of_station,invoice_number,invoice_date,bill_from,totalunitsbilled,vc,fixedcost,interest_,incentive,incometax,others_,transmissioncharges,fixedcostpeak,fixedcostoffpeak,fixedcostoffset,rrascharges,carryingcost,incentiveoffpeak,incentiveoffset,remarks2,Totalinvoiceamount from tata_power_data where mappingsheet='NTPC' and status='pending' and name_of_station='%s' and invoice_date='%s'"%(value,value2)
        cursor.execute(query1)
        data=cursor.fetchall()
        if len(data)==0:
            continue
        

        
        
        month=int(invoice_date[3:5])
        odd_months=[1,3,5,7,8,10,12]
        even_months=[4,6,9,11]

        if month == 2:

                end_date=invoice_date.split('.')
                end_date='28'+'.'+end_date[1]+'.'+end_date[2]

        elif month in odd_months:
                end_date=invoice_date.split('.')
                end_date='31'+'.'+end_date[1]+'.'+end_date[2]

        elif month in even_months:
                end_date=invoice_date.split('.')
                end_date='30'+'.'+end_date[1]+'.'+end_date[2]


        due_date=end_date
        tcs=''
    ##    total_amount=ws['ES'+str(i)].value
        vendor_code,station_code=vendor_cc(vendor_name)
        time.sleep(2)
        session.findById("wnd[0]/usr/btnNEW_BILL_ENTRY").press()
        time.sleep(2)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/cmbTYPE_TRANSACTION").key = "P"
        if vendor_name =="NTPC TRANSMISSION Charges":
            session.findById("wnd[0]/usr/cmbWA_ZMMTR_PMG_HEADER-BILL_TYPE").key = "7"
        else:
            session.findById("wnd[0]/usr/cmbWA_ZMMTR_PMG_HEADER-BILL_TYPE").key = "6"
        session.findById("wnd[0]/usr/cmbWA_ZMMTR_PMG_HEADER-BILL_TYPE").setFocus()
        time.sleep(2)
        #vendor_code
        session.findById("wnd[0]").sendVKey(2)
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-LIFNR").setFocus()
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-LIFNR").caretPosition = 0
        session.findById("wnd[0]").sendVKey(2)
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-LIFNR").text = vendor_code
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-LIFNR").setFocus()
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-LIFNR").caretPosition = 7
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)
        #invoice date
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-INV_DATE").text = invoice_date
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-INV_DATE").caretPosition = 10
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-DUE_DATE").setFocus()
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-DUE_DATE").caretPosition = 0                         
        session.findById("wnd[0]").sendVKey(2)
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-DUE_DATE").text = due_date
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-DUE_DATE").setFocus()
        session.findById("wnd[0]/usr/ctxtWA_ZMMTR_PMG_HEADER-DUE_DATE").caretPosition = 10
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)
        ##################function start############
       
        k=1
        sum_=0.0
        print(data)
        for j in range(0,len(data)):
            x=data[j]
            invoice_no=x[1]
            print(invoice_no)

            Bill_from=x[3]
            if "-" in Bill_from:
                Bill_from=Bill_from.replace("-",".")
            
            
            unit_bill=x[4]
            if unit_bill == None:
                unit_bill=""
            
            vc=x[5]
            if vc == None:
                vc=""
            
            fc=x[6]
            if fc==None:
                fc=""
            
            interest=x[7]
            if interest ==None:
                interest== ''
            
            incentive=x[8]
            if incentive == None:
                incentive=""
            
            income_tax=x[9]
            if income_tax== None:
                income_tax=""
            
            others=x[10]
            if others == None:
                others=""
            print(others)
            
            transmission=x[11]
            if transmission == None:
                transmission=""

            fc_peak=x[12]
            if fc_peak == None:
                fc_peak=""
            if fc!='' and fc_peak!='':
                fc_peak=fc+fc_peak
            elif fc!='' and fc_peak=='':
                fc_peak=fc
            
            fc_offpeak=x[13]
            if fc_offpeak==None:
                fc_offpeak=""
            
            fc_offset=x[14]
            if fc_offset==None:
                fc_offset=""
            
            rras=x[15]
            if rras==None:
                rras=""
            
            carrying=x[16]
            if carrying==None:
                carrying=""
            
            incentive_op=x[17]
            if incentive_op==None:
                incentive=''
            
            incentive_offset=x[18]
            if incentive_offset==None:
                incentive_offset=""

            remarks=x[19]
            total_amount=x[-1]
            
            t = 0
            if j >12:
                
                t = 12
                session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM").verticalScrollbar.position = k
                k=k+1
            else:
                t = j
                    
            if vendor_name not in transmission_li:
                data_rowise(t,invoice_no,station_code,unit_bill,vc,fc_peak,carrying,transmission,others,income_tax,incentive,rras,fc_offpeak,fc_offset,incentive_op,incentive_offset,tcs,remarks)
                
            else:
                
                data_transmission(t,invoice_no,station_code,unit_bill,vc,fc_peak,carrying,transmission,others,income_tax,incentive,rras,fc_offpeak,fc_offset,incentive_op,incentive_offset,tcs,remarks)

                print(others)
            sum_=float(total_amount)+sum_

        for  z in range(k,-1,-1):
            session.findById("wnd[0]/usr/tblZMM_PMG_INVOICETC_ITEM").verticalScrollbar.position = z
            k = k-1

        #totalamount
        session.findById("wnd[0]/usr/txtWA_ZMMTR_PMG_HEADER-NUM_OF_INVOCES").text = (len(data))
        session.findById("wnd[0]/usr/txtWA_ZMMTR_PMG_HEADER-TOTAL_EST_AMT").text = sum_
        time.sleep(2)
        sum_=0
        ################
        data_li=[]
        request=''
        
        try:
            session.findById("wnd[0]/usr/btnSAVE").press()
            time.sleep(2)
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
            document_confirm=session.findByid("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").text
            print(document_confirm)
            try:
                session.findById("wnd[0]/usr/btnSUBMIT").press()
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
                document_confirm=session.findByid("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").text
                print(document_confirm)
                request=document_confirm
                data_li=[x for x in range(0,len(data))]
                counter=1
            except:
                document_confirm=session.findByid("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").text
                print(document_confirm)
                if "Invoice number already exist" in document_confirm:
                    session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    session.findById("wnd[1]/usr/btnBUTTON_1").press()
        except:
            try:
                kk=1
                while kk>0:
                    text1=session.findById("/app/con[0]/ses[0]/wnd[1]/usr/txtMESSTXT1").text
                    if "Invoice number already exist" in text1:
                        data_li.append(text1.split(" ")[-1])
                        text2=session.findById("/app/con[0]/ses[0]/wnd[1]/usr/txtMESSTXT2").text
                        request=text2
                        print(text2.split(':')[1].strip(),"aaaaaaaaaa")
                        counter=1
                    if "Please check total invoice amount" in text1:
                        data_li=[x for x in range(0,len(data))]
                        request="total invoice mismatch"
                        print(data_li,request)
                        counter=2
                    if "Please check number of invoices" in text1:
                        data_li=[x for x in range(0,len(data))]
                        request="total invoice number mismatch"
                        print(data_li,request)
                        counter=2
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    
                    
            except:
                session.findById("wnd[1]/usr/btnBUTTON_2").press()
                session.findById("wnd[0]/tbar[0]/btn[12]").press()
                session.findById("wnd[1]/usr/btnBUTTON_1").press()

##                
        end_time=datetime.datetime.now()
        status_update(data_li,vendor_name,data,start_time,end_time,request)

if counter ==1:
    subject='Data Entered Succcessfully '
    body1='Bot has entered the data of NTPC '
    send_email(subject,body1)
    print("yesss")
if counter ==2:
    subject='Data has mismatched value '
    body1='Bot has detected mismatch the data of NTPC '
    send_email(subject,body1)
    print("yesss")
    
try:
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
except:
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

    











