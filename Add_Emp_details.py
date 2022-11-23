import streamlit as st
import pandas as pd

# import datetime
file_name = "Asset data - Copy.xlsx"
client = ["TS India Purchases", "For QC", "Chinna"]
unique_value = ["Asset Id", "PTAG #", "Serial No"]


class Add_Emp_assert_details:

    def __init__(self):
        self.data = pd.ExcelFile(file_name)
        self.TS = pd.read_excel(self.data, client[0], header=0)
        self.QC = pd.read_excel(self.data, client[1], header=0)
        self.CH = pd.read_excel(self.data, client[2], header=0)

    def write(self, columns1, sl_Number, unique, unique2="Employee ID"):
        details = dict()
        with st.form("my_form"):
            for i in columns1:
                if i in ["SL#", "S.No.", "S.No"]:
                    details[i] = st.text_input("Enter " + i, sl_Number)
                elif i in ["Issued on", "Received Date"]:
                    # x = datetime.datetime.now()
                    details[i] = st.date_input(i)
                else:
                    value = st.text_input("Enter " + i)
                    details[i] = value

            submitted = st.form_submit_button("Submit")
            if submitted:
                if details[unique] == "":
                    st.error("Enter " + unique)
                elif details[unique2] == "":
                    st.error("Enter " + unique2)
                else:
                    details = pd.DataFrame(details, index=[0])
                    st.success("Successfuly Added")
                    return details

    def save_data(self):
        with pd.ExcelWriter(file_name) as writer:
            self.TS.to_excel(writer, sheet_name=client[0], index=False)
            self.QC.to_excel(writer, sheet_name=client[1], index=False)
            self.CH.to_excel(writer, sheet_name=client[2], index=False)

    def T_S(self):
        columns1 = self.TS.columns
        sl_Number = list(list(self.TS[columns1[0]]))[-1]
        Ts_Frame = self.write(columns1, sl_Number + 1, unique_value[0])
        self.TS = pd.concat([self.TS, Ts_Frame])
        dt = self.TS
        self.TS["Issued on"] = pd.to_datetime(dt['Issued on']).dt.date
        self.save_data()

    def Q_C(self):
        columns1 = self.QC.columns
        sl_Number = list(self.QC[columns1[0]])
        sl_Number = sl_Number[-1]
        QC_Frame = self.write(columns1, int(sl_Number) + 1, unique_value[1])
        self.QC = pd.concat([self.QC, QC_Frame])
        dt = self.QC
        self.CH["Received Date"] = pd.to_datetime(dt["Received Date"]).dt.date
        self.save_data()

    def C_H(self):
        columns1 = self.CH.columns
        sl_Number = list(self.CH[columns1[0]])
        sl_Number = sl_Number[-1]
        CH_Frame = self.write(columns1, int(sl_Number) + 1, unique_value[2])
        self.CH = pd.concat([self.CH, CH_Frame])
        self.save_data()
