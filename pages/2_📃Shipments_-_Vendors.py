import streamlit as st
from page_config import page_setup
import pandas as pd
from japan_function import reconcile
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
        
        if 'submit_ra' not in state:
            state.submit_ra= False
        if 'response_ra' not in state:
            state.response_ra = []
        st.markdown("<h2 style='text-align: center; padding:0'>Shipments - Vendors</h2>", unsafe_allow_html=True)
        #st.write('###')
        packing_lists, inventory_ledger, submit= file_upload_form()
        #print(warehouse_reports)
        # try:
        if submit:
            state.submit_ra = True
            #print(warehouse_reports)
            #print(submit)
                #print(shipment_instructions_df)
            #with st.spinner('Please wait'):
            try:
                delete_temp()
            except:
                print()

            reconcile(packing_lists, inventory_ledger)
            #state.response = [payment_report_df, returns_report_df, reimbursement_report, inventory_ledger_df]
            emp, but, empty = st.columns([2.05,1.2,1.5])
            with but:
                st.write("###")
                with open('temp/fba_reco_japan.xlsx', 'rb') as my_file:
                    click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'fba_reco_japan.xlsx', 
                    mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    #print(click) 
            #st.write(workbook) 

        else:
            if state.submit_ra == True:
                emp, but, empty = st.columns([2.05,1.2,1.5]) 
                with but:
                    st.write("###")
                    with open('temp/fba_reco_japan.xlsx', 'rb') as my_file:
                        click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'fba_reco_japan.xlsx', 
                        mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # except:
        #    st.error("Run failed, kindly check if the inputs are valid")

    def delete_temp():
        os.remove('temp/fba_reco_japan.xlsx')

    def zip_files():
        zipObj = ZipFile("sample.zip", "w")
        zipObj.write("checkpoint")
        zipObj.write("model_hyperparameters.json")
        # close the Zip File
        zipObj.close()

    def file_upload_form():
        colour = "#89CFF0"
        with st.form(key = 'ticker',clear_on_submit=False):
            text, upload = st.columns([2.5,3]) 
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5>{"&nbsp; Upload Packing Lists:"}</h5>', unsafe_allow_html=True)
            with upload:
                packing_lists = st.file_uploader("",key = 'pack_list', accept_multiple_files=True)

            text, upload = st.columns([2.5,3])
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5>{"&nbsp; Upload Inventory Ledger:"}<h5>', unsafe_allow_html=True)
            with upload:
                inventory_ledger = st.file_uploader("",key = 'inv_ledger')

            # text, upload = st.columns([2.5,3])
            # with text:
            #     st.write("###")
            #     st.write("###")
            #     st.write(f'<h5>{"&nbsp; Upload Previous Reconciliation:"}<h5>', unsafe_allow_html=True)
            # with upload:
            #     prev_recon = st.file_uploader("",key = 'prev_reco')
            
            a,button,b = st.columns([2,1.2,1.5]) 
            with button:
                st.write('###')
                submit = st.form_submit_button(label = "Start Reconciliation")
                #submit = st.button(label="Start Reconciliation")

        return packing_lists, inventory_ledger, submit
        

        

    landing_page()

