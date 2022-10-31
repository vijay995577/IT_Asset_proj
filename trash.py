import streamlit as st
from PIL import Image
import pandas as pd
import smtplib
from email.message import EmailMessage
from streamlit_option_menu import option_menu


image1=Image.open("thundersoft.png")
#---------------------------------------------------------------------------------#
file_name='Asset data - Copy.xlsx'
client_list=["Thundersoft","Qualcomm","China"]
client_choice={client_list[0]:"TS India Purchases",client_list[1]:"For QC",client_list[2]:"Chinna"}
Emp_details_choice={"Name":"Employee Name".upper(),"Employee Id":"Employee ID"}
#------------------------------QC-------------------------------------------------#
QC_Details={"invoces NO":"INVOCES NUMBER","DC NO":"DC NUMBER",	"MCN":"MNC","Description of Goods":"Description of Goods",	"QTY":"QUANTITY","Serial#":"SERIAL NUMBER","PTAG #":"PTAG NUMBER","Received Date":"DATE OF RECIVED","Return Date":"DATE OF RETURN",	"Received from":"RECEIVED FROM","Status":"STATUS","Project #":"NAME OF PROJECT","Employee ID":"EMPLOYEE ID","EMPLOYEE NAME":"EMPLOYEE NAME",	"Employee Email":"EMPLOYEE EMAIL"}
QC_list_Details=["EMPLOYEE NAME","EMPLOYEE ID","EMPLOYEE EMAIL","NAME OF PROJECT","INVOCES NUMBER","DC NUMBER",'MNC', 'Description of Goods', "QUANTITY","SERIAL NUMBER", 'PTAG NUMBER', 'DATE OF RECIVED', 'DATE OF RETURN', 'RECEIVED FROM', 'STATUS']
#-------------------------------china--------------------------------#
China_Details={"Device":'DEVICE',	"Serial No":'SERIAL NO',	"P Tag":'P TAG',	"Description":'DESCRIPTION',	"Received Date":'RECEIVED DATE',	"Receviced from":'RECEVICED FROM',	"Employee ID":'EMPLOYEE ID'    ,	"EMPLOYEE NAME":'EMPLOYEE NAME'}
China_list_Details=["EMPLOYEE NAME","EMPLOYEE ID","P Tag","Device",	"Serial No","Description","Received Date",	"Receviced from"]
China_list_Details=list(map(lambda x: x.upper(), China_Details))
#-------------------------TS-----------------------------------------------------#
TS_Details={"Product":"PRODUCT",	"Product ID":'PRODUCT ID',	"Ptag":'PTAG',	"Product Serial Number":'PRODUCT SERIAL NUMBER', 	"Issued on":"ISSUED ON",	"Bought For":"BOUGHT FOR" ,	"Employee ID":"EMPLOYEE ID",	"EMPLOYEE NAME":"EMPLOYEE NAME"}
TS_list_Details=["EMPLOYEE NAME",'EMPLOYEE ID','PRODUCT','PRODUCT ID', 'PTAG', 'PRODUCT SERIAL NUMBER', 'ISSUED ON', 'BOUGHT FOR']
#--------------------------------------To-Gether-------------------------------------------#
client_details={client_list[0]:TS_Details,client_list[1]:QC_Details,client_list[2]:China_Details}
#-------------------------------------------------------------------------------------------#
def get_key(value,client):
    variable=client_details[client]
    key=list(variable.keys())[list(variable.values()).index(value)]
    return key

detailslis=''
def print_details(element,details_dict,client):
    global detailslis
    key=get_key(element,client)
    value=list(details_dict[key].keys())[0]
    print('line39',value)
    value=details_dict[key][value]
    print('line40',value)
    details=(element+" : "+str(value) + '\n \n')
    detailslis +=  details
    st.write(details)

def Read_Excel(client,name,choice):
    data= pd.ExcelFile(file_name)
    data = pd.read_excel(data,client_choice[client],header=0)
    value=Emp_details_choice[choice]
    details=data.loc[data[value]==name]
    #print(details)
    details_dict=details.to_dict()
    print('line52',details_dict)

    if details.empty or name=="nan":
        st.error("Enter Correct Details")
    else:
        if client==client_list[1]: #QUALCOM Details
            for element in QC_list_Details:
                print_details(element,details_dict,client)
        elif client==client_list[0]:
            for element in TS_list_Details:
                print_details(element, details_dict,client)
        else:
            for element in China_list_Details:
                print_details(element, details_dict, client)

def send_mail(name,receiver):
    sender = 'vijay.balakula@thundersoft.com'
    #password = st.text_input('give your password here : npybdfkjywludidr,for TS:u%F6xw7UZfaD' )
    #receiver = 'pavan.d@thundersoft.com'
    msg = EmailMessage()
    msg['Subject'] = 'Assert details '+name
    msg['From'] = sender
    msg['To'] = receiver
    msg.set_content(str(detailslis))
    with smtplib.SMTP_SSL('india-hzsmtp.mail-qiye.com', 465) as smtp:
        smtp.login(sender, "u%F6xw7UZfaD")
        smtp.send_message(msg)
    st.success(f'mail sent successfully to please check your email inbox')


def Get_Employee_Names(client):
    data= pd.ExcelFile(file_name)
    data = pd.read_excel(data,client_choice[client],header=0)
    emp_list=list(data["EMPLOYEE NAME"])
    emp_list.insert(0,"Enter your name")

    return  emp_list


with st.sidebar:
    choose=option_menu("Details",["Emp_details","Asset_details"])
#---------------------EMP-Details-------------------------#
if choose=="Emp_details":
    column1,column2,column3=st.columns([5,5,5])
    name=0
    new_var=0
    with column1:
    	st.image(image1)
    Emp_det=st.selectbox("Get Details Through: ",("Name","Employee Id"))

    #-----------------------Name--------------------------#
    if Emp_det=="Name":
        client=st.selectbox("Select your choice",(client_list))
        name_list=Get_Employee_Names(client)
        name=st.selectbox("Enter your Name: ",name_list)
        if name ==name_list[0]:
            pass
        else:
            button1 = st.button('Submit')
            if st.session_state.get('button') != True:
                st.session_state['button'] = button1 # Saved the state
            if st.session_state['button'] == True:
                Read_Excel(client,name,"Name")
                receiver=st.text_input("Enter receiver mail-id to send details")
                if st.button('Send Mail'):
                    send_mail(name,receiver)
    #----------------------------EMP-ID------------------------#
    else:
        client=st.selectbox("Select your choice",(client_list))
        name=st.text_input("Enter Employee id")
        button1 = st.button('Submit')
        if st.session_state.get('button') != True:
            st.session_state['button'] = button1 # Saved the state
        if st.session_state['button'] == True:
            Read_Excel(client,(name),"Employee Id")
            receiver=st.text_input("Enter receiver mail-id to send details")
            if st.button('Send Mail'):
                send_mail(name,receiver)