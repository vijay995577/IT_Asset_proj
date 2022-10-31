import streamlit as st  #for frontend Work
from PIL import Image   #for reading image
import pandas as pd     #for retrieving data from the XLSX sheets
import smtplib          #for logging into the particular mail server
from email.message import EmailMessage     #for creating msg object
from streamlit_option_menu import option_menu   #to create option menu

image1 = Image.open("thundersoft.png")
# ---------------------------------------------------------------------------------#
file_name = 'Asset data - Copy.xlsx'
client_list = ["Thundersoft", "Qualcomm", "China"]
client_choice = {client_list[0]: "TS India Purchases", client_list[1]: "For QC", client_list[2]: "Chinna"}
Emp_details_choice = {"Name": "Employee Name".upper(), "Employee Id": "Employee ID", "Asset Id": "Asset Id",
                      "PTag": "PTAG #", "Serial Number": "Serial No"}
# ------------------------------QC-------------------------------------------------#
QC_Details = {"invoces NO": "INVOCES NUMBER", "DC NO": "DC NUMBER", "MCN": "MNC",
              "Description of Goods": "Description of Goods", "QTY": "QUANTITY", "Serial#": "SERIAL NUMBER",
              "PTAG #": "PTAG NUMBER", "Received Date": "DATE OF RECIVED", "Return Date": "DATE OF RETURN",
              "Received from": "RECEIVED FROM", "Status": "STATUS", "Project #": "NAME OF PROJECT",
              "Employee ID": "EMPLOYEE ID", "EMPLOYEE NAME": "EMPLOYEE NAME", "Employee Email": "EMPLOYEE EMAIL"}
QC_list_Details = ["EMPLOYEE NAME", "EMPLOYEE ID", "NAME OF PROJECT", "INVOCES NUMBER", "DC NUMBER", 'MNC',
                   'Description of Goods', "QUANTITY", "SERIAL NUMBER", 'PTAG NUMBER', 'DATE OF RECIVED',
                   'DATE OF RETURN', 'RECEIVED FROM', 'STATUS']
# -------------------------------china--------------------------------#
China_Details = {"Device": 'DEVICE', "Serial No": 'SERIAL NO', "P Tag": 'P TAG', "Description": 'DESCRIPTION',
                 "Received Date": 'RECEIVED DATE', "Receviced from": 'RECEVICED FROM', "Employee ID": 'EMPLOYEE ID',
                 "EMPLOYEE NAME": 'EMPLOYEE NAME'}
China_list_Details = ["EMPLOYEE NAME", "EMPLOYEE ID", "P Tag", "Device", "Serial No", "Description", "Received Date",
                      "Receviced from"]
China_list_Details = list(map(lambda x: x.upper(), China_Details))
# -------------------------TS-----------------------------------------------------#
TS_Details = {"Product": "PRODUCT", "Asset Id": 'ASSET ID', "Ptag": 'PTAG',
              "Product Serial Number": 'PRODUCT SERIAL NUMBER', "Issued on": "ISSUED ON", "Bought For": "BOUGHT FOR",
              "Employee ID": "EMPLOYEE ID", "EMPLOYEE NAME": "EMPLOYEE NAME"}
TS_list_Details = ["EMPLOYEE NAME", 'EMPLOYEE ID', 'PRODUCT', 'ASSET ID', 'PTAG', 'PRODUCT SERIAL NUMBER', 'ISSUED ON',
                   'BOUGHT FOR']
# --------------------------------------To-Gether-------------------------------------------#
client_details = {client_list[0]: TS_Details, client_list[1]: QC_Details, client_list[2]: China_Details}
# -------------------------------------------------------------------------------------------#
Asset_details_client = {client_list[0]: "Product", client_list[1]: "MCN", client_list[2]: "Device"}


# -------------------------------------------------------------------------------------------#
# for get a key
def get_key(value, client):
    variable = client_details[client]
    key = list(variable.keys())[list(variable.values()).index(value)]
    return key


detailslis = ''

#for pringting details and stroing details
def print_details(element, details_dict, client):
    global detailslis
    key = get_key(element, client)
    value = list(details_dict[key].keys())[0]
    value = details_dict[key][value]
    details = (element + " : " + str(value) + '\n \n')
    detailslis += details
    # st.markdown(html_str, unsafe_allow_html=True)
    st.write(details)

#for read the details from excel sheet using user chioces

def Read_Excel(client, name, choice, status=True):
    data = pd.ExcelFile(file_name)
    data = pd.read_excel(data, client_choice[client], header=0)
    value = Emp_details_choice[choice]
    details = data.loc[data[value] == name]
    details_dict = details.to_dict()
    if details.empty or name == "nan":
        st.error("Enter Correct Details")
    else:
        if client == client_list[1]:  # QUALCOM Details
            for element in QC_list_Details:
                print_details(element, details_dict, client)
        elif client == client_list[0]:
            for element in TS_list_Details:
                print_details(element, details_dict, client)
        else:
            for element in China_list_Details:
                print_details(element, details_dict, client)

#for sending mails whom we want to send the mail with the details

def send_mail(name, receiver):
    sender = 'vijay.balakula@thundersoft.com'
    # password = st.text_input('give your password here : npybdfkjywludidr,for TS:u%F6xw7UZfaD' )
    # receiver = 'pavan.d@thundersoft.com'
    msg = EmailMessage()
    msg['Subject'] = 'Assert details ' + name
    msg['From'] = sender
    msg['To'] = receiver
    msg.set_content(str(detailslis))
    with smtplib.SMTP_SSL('india-hzsmtp.mail-qiye.com', 465) as smtp:
        smtp.login(sender, "u%F6xw7UZfaD")
        smtp.send_message(msg)
    st.success('mail sent successfully to please check your email inbox')

#for getting a particular column

def Get_Column(client, columnName, status=True):
    data = pd.ExcelFile(file_name)
    data = pd.read_excel(data, client_choice[client], header=0)
    if status == False:
        count_values = dict(data[columnName].value_counts())
        unique_list = list(data[columnName].unique())
        unique_list.insert(0, "Select devices")
        return unique_list, count_values
    else:
        emp_list = list(data[columnName])
        emp_list.insert(0, "Enter " + columnName)
        return emp_list


def get_asset_Details(client, columnName, Device_name):
    data = pd.ExcelFile(file_name)
    data = pd.read_excel(data, client_choice[client], header=0)
    result_data = data.loc[data[columnName] == Device_name]
    return result_data


if __name__ == "__main__":
    with st.sidebar:
        choose = option_menu("Details", ["Emp_details", "Asset_details"])    #side bar options
    # ---------------------EMP-Details-------------------------#
    if choose == "Emp_details":
        column1, column2, column3 = st.columns([5, 5, 5])
        name = 0
        new_var = 0
        with column1:
            st.image(image1)
        Emp_det = st.selectbox("Get Details Through: ", ("Name", "Employee Id"))            # Get Details Through name or EMP ID

        # -----------------------Name--------------------------#
        if Emp_det == "Name":                                                             #if user want details based on namr it will excute
            client = st.selectbox("Select your choice", (client_list))
            name_list = Get_Column(client, 'EMPLOYEE NAME')
            name = st.selectbox("Enter Employee Name: ", name_list)                       # it is for show the dropdown and placeholder
            if name == name_list[0]:
                pass
            else:
                button1 = st.button('Submit')
                if st.session_state.get('button') != True:
                    st.session_state['button'] = button1  # Saved the state
                if st.session_state['button'] == True:
                    Read_Excel(client, name, "Name")
                    receiver_ID = st.text_input("Enter receiver mail-id to send details")
                    if st.button('Send Mail'):
                        send_mail(name, receiver_ID)

        # ----------------------------EMP-ID------------------------#
        else:
            client = st.selectbox("Select your choice", (client_list))
            name = st.text_input("Enter Employee id")
            button1 = st.button('Submit')
            if st.session_state.get('button') != True:
                st.session_state['button'] = button1  # Saved the state
            if st.session_state['button'] == True:
                Read_Excel(client, int(name), "Employee Id")
                receiver_ID = st.text_input("Enter receiver mail-id to send details")
                if st.button('Send Mail'):
                    send_mail(name, receiver_ID)

        # ---------------------Asset_details-------------------------#
    else:
        column1, column2, column3 = st.columns([5, 5, 5])                            # it is for asset part
        with column1:
            st.image(image1)
        client = st.selectbox("Select your choice", (client_list))
        # --------------------TS---------------------------------#
        if client == client_list[0]:
            Asset_det = st.selectbox("Get Details Through:", ("Asset Id", "Device"))      # for getting details throgh asset ID or device
            if Asset_det == "Asset Id":  # Asset Id is Unique value in TS Sheet
                name_list = Get_Column(client_list[0], Asset_det)
                name = st.selectbox("Enter Id : ", name_list)
                if name == name_list[0]:
                    pass
                else:
                    button1 = st.button('Submit')
                    if st.session_state.get('button') != True:
                        st.session_state['button'] = button1  # Saved the state

                    if st.session_state['button'] == True:
                        Read_Excel(client, name, Asset_det)
            else:
                name_list, count_values = Get_Column(client_list[0], Asset_details_client[client_list[0]], False)
                device = st.selectbox("Select Device: ", name_list)
                if st.button("submit"):
                    st.write("Number of " + device + " :" + str(count_values[device]))
                    row_data = get_asset_Details(client, Asset_details_client[client_list[0]], device)
                    row_data = row_data[[Asset_details_client[client_list[0]], "Asset Id"]]
                    st.table(row_data)

        # ------------------------QC-------------------------------------#
        elif client == client_list[1]:
            Asset_det = st.selectbox("Get Details Through:", ("PTag", "Device"))
            if Asset_det == "PTag":  # PTag is Unique value in QC Sheet
                name_list = Get_Column(client_list[1], Emp_details_choice[Asset_det])
                name = st.selectbox("Enter your Name: ", name_list)
                if name == name_list[0]:
                    pass
                else:
                    button1 = st.button('Submit')
                    if st.session_state.get('button') != True:
                        st.session_state['button'] = button1  # Saved the state

                    if st.session_state['button'] == True:
                        Read_Excel(client, name, Asset_det)
            else:
                name_list, count_values = Get_Column(client_list[1], Asset_details_client[client_list[1]], False)
                device = st.selectbox("Enter your choice: ", name_list)
                if st.button("submit"):
                    st.write("Number of " + device + " :" + str(count_values[device]))
                    row_data = get_asset_Details(client, Asset_details_client[client_list[1]], device)
                    row_data = row_data[[Asset_details_client[client_list[1]], Emp_details_choice["PTag"]]]
                    st.table(row_data)
        # -------------------------China--------------------------#
        else:
            Asset_det = st.selectbox("Get Details Through:", ("Serial Number", "Device"))
            if Asset_det == "Serial Number":
                name_list = Get_Column(client_list[2], Emp_details_choice[Asset_det])
                name = st.selectbox("Enter ID: ", name_list)
                if name == name_list[0]:
                    pass
                else:
                    button1 = st.button('Submit')
                    if st.session_state.get('button') != True:
                        st.session_state['button'] = button1  # Saved the state

                    if st.session_state['button'] == True:
                        Read_Excel(client, name, Asset_det)
            else:
                name_list, count_values = Get_Column(client_list[2], Asset_details_client[client_list[2]], False)
                device = st.selectbox("Enter your choice: ", name_list)
                if st.button("submit"):
                    st.write("Number of " + device + " :" + str(count_values[device]))
                    row_data = get_asset_Details(client, Asset_details_client[client_list[2]], device)
                    row_data = row_data[[Asset_details_client[client_list[2]], Emp_details_choice["Serial Number"]]]
                    st.table(row_data)