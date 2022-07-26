import streamlit as st
from page_config import page_setup
import pandas as pd
from customer_returns_streamlit import reconcile
import os
import streamlit_authenticator as stauth
import pickle 
from pathlib import Path 
import yaml
from PIL import Image
import time
import plotly.graph_objects as go
import base64
from login_page import login_status

st.set_page_config(layout="wide",initial_sidebar_state ="collapsed")
page_setup()
state = st.session_state

authentication_status = login_status()


if authentication_status == False:
    space, login, space = st.columns([1,3,1])
    with login:
        st.error("Username/Password is incorrect")

if authentication_status:
    #authenticator.logout('Logout', 'sidebar')
    time.sleep(0.1)
    def landing_page():
        st.markdown('''
        <style>
        .css-9s5bis.edgvbvh3 {
        display: block;
        }
        </style>
        ''', unsafe_allow_html=True)
        #with title:
        # emp,title,emp = st.columns([2,2,2])
        # with title:
        
        if 'submit_cus' not in state:
            state.submit_cus= False
        if 'response_cus' not in state:
            state.response_cus = []
        st.markdown("<h2 style='text-align: center; padding:0'>Customer Returns Reconciliation</h2>", unsafe_allow_html=True)
        #st.write('###')
        payment_report, returns_report, reimbursement_report, inventory_ledger, submit = file_upload_form()
        #print(warehouse_reports)
        if submit:
            state.submit_cus = True
            #print(warehouse_reports)
            #print(submit)
                #print(shipment_instructions_df)
            with st.spinner('Please wait'):
                try:
                    delete_temp()
                except:
                    print()
                if payment_report is not None:
                    payment_report_df = pd.read_csv(payment_report,encoding='latin-1')
                if returns_report is not None:
                    returns_report_df = pd.read_csv(returns_report,encoding='latin-1')
                if reimbursement_report is not None:
                    reimbursement_report_df = pd.read_csv(reimbursement_report,encoding='latin-1')
                if inventory_ledger is not None:
                    inventory_ledger_df = pd.read_csv(inventory_ledger,encoding='latin-1')
                reconcile(payment_report_df, returns_report_df, reimbursement_report_df, inventory_ledger_df)
            #state.response = [payment_report_df, returns_report_df, reimbursement_report, inventory_ledger_df]
            emp, but, empty = st.columns([2.05,1.2,1.5])
            with but:
                st.write("###")
                with open('temp/customer_returns_reco.xlsx', 'rb') as my_file:
                    click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'customer_returns_reco.xlsx', 
                    mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') 
                    #print(click) 
            #st.write(workbook) 

        else:
            if state.submit_cus == True:
                emp, but, empty = st.columns([2.05,1.2,1.5]) 
                with but:
                    st.write("###")
                    with open('temp/customer_returns_reco.xlsx', 'rb') as my_file:
                        click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'customer_returns_reco.xlsx', 
                        mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    def delete_temp():
        os.remove("temp/customer_returns_reco.xlsx")

    def file_upload_form():
        colour = "#89CFF0"
        with st.form(key = 'ticker',clear_on_submit=True):
            text, upload = st.columns([2.5,3]) 
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5>{"&nbsp; Upload Payment Report:"}</h5>', unsafe_allow_html=True)
            with upload:
                payment_report = st.file_uploader("",key = 'pay_rep')

            text, upload = st.columns([2.5,3])
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5>{"&nbsp; Upload Returns Report:"}<h5>', unsafe_allow_html=True)
            with upload:
                returns_report = st.file_uploader("",key = 'ret_rep')

            text, upload = st.columns([2.5,3])
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5> {"&nbsp; Upload Reimbursement Report:"}<h5>', unsafe_allow_html=True)
            with upload:
                reimbursement_report = st.file_uploader("",key = 'reim_rep')

            text, upload = st.columns([2.5,3])
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5> {"&nbsp; Upload Inventory Ledger:"}<h5>', unsafe_allow_html=True)
            with upload:
                inventory_ledger = st.file_uploader("",key = 'inv_led')
            
            a,button,b = st.columns([2,1.2,1.5]) 
            with button:
                st.write('###')
                submit = st.form_submit_button(label = "Start Reconciliation")
                #submit = st.button(label="Start Reconciliation")

        return payment_report, returns_report, reimbursement_report, inventory_ledger, submit
        

        

    landing_page()

