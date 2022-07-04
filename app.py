import streamlit as st 
import pandas as pd
from shipment_reco_charts import reconcile
import os
import streamlit_authenticator as stauth
import pickle 
from pathlib import Path 
import yaml
import altair as alt
from PIL import Image

st.set_page_config(layout="wide")

names = ['admin', 'amit']
usernames = ['admin', 'amit']

with open('config.yaml') as file:
    config = yaml.safe_load(file)

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days'],
    config['preauthorized']
)


#authenticator = stauth.Authenticate(names, usernames, hashed_passwords, "ship_recon", "admin")

name, authentication_status, username = authenticator.login('Login', 'main')

print("#######")
print(authentication_status)
if authentication_status == False:
    st.error("Username/Password is incorrect")

if authentication_status == None:
    st.warning("Please enter your username and Password")

if authentication_status:
    def landing_page():
        logo = Image.open('images/reconcify_logo.png')
        hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
        st.markdown(hide_streamlit_style, unsafe_allow_html=True) 
        emp,Logo,emp = st.columns([2,2,2])
        with Logo:
            st.image(logo)
        #with title:
        #st.header('Shipment Reconcilliation')
        st.write('###')
        shipment_instructions, warehouse_reports, inventory_ledger, submit = file_upload_form()
        #print(warehouse_reports)
        if submit:
            print(warehouse_reports)
            #print(submit)
            if shipment_instructions is not None:
                shipment_instructions_df = pd.read_excel(shipment_instructions)
                #print(shipment_instructions_df)
            if inventory_ledger is not None:
                inventory_ledger_df = pd.read_csv(inventory_ledger)
            units_booked, excess_units_received, short_units_received, units_received, matching_sku, mismatching_sku = reconcile(shipment_instructions_df, warehouse_reports, inventory_ledger_df)
            
            bar_data = [['Units Booked',units_booked],['Excess Units', excess_units_received]
            ,['Short Units', short_units_received],['Units Recieved', units_received]]
            #val_df = val_df.set_index
            bar_df = pd.DataFrame(bar_data, columns=['Label', 'Units'])  
            #bar_df = bar_df.set_index('Label')       
            # bar_data = {
            #     'Units Booked':units_booked,
            #     'Excess Units Received': excess_units_received,
            #     'Short Units Received': short_units_received,
            #     'Units Recieved':units_received
            # }

            pie_data = [['Matching SKUs',matching_sku],['Mismatching SKUs', mismatching_sku]]
            pie_df = pd.DataFrame(pie_data, columns=['Label', 'Units']) 
            #pie_df = pie_df.set_index('Label')    
            with st.expander('View KPI Charts'):
                bar,pie = st.columns([1.2,1]) 
                with bar:
                    plot_bar_chart(bar_df,'Label','Units')
                    #st.bar_chart(bar_df)
                with pie:
                    #st.bar_chart(pie_df)
                    plot_pie_chart(pie_df,'Label','Units')
            with open('temp/shipment_reco.xlsx', 'rb') as my_file:
                click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'shipment_reco.xlsx', 
                mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', on_click = delete_temp()) 
                #print(click) 
            #st.write(workbook) 
    def plot_bar_chart(data,X,Y):
        chart = (
            alt.Chart(data).configure_title(fontSize=20)
            .mark_bar()
            .encode(
                x=alt.X(X, type="nominal", title="", axis = alt.Axis(labelAngle=0,labelOverlap=False,labelAlign ='center',labelFontSize=10.5)),
                y=alt.Y(Y, type="quantitative", title=""),
                color = alt.Color(X, legend=None),
                # color=alt.condition(
                # alt.datum[Y] > 0,
                # alt.value("#74c476"),  # The positive color
                # alt.value("#d6616b")  # The negative color
                # ),
                tooltip = [alt.Tooltip(Y, title="",format='.1f')]
                #color=alt.Color("variable", type="nominal", title=""),
                #order=alt.Order("variable", sort="descending"),
            )
        ).interactive()
        
        st.altair_chart(chart, use_container_width=True)

    def plot_pie_chart(data,X,Y):
        pie = alt.Chart(data).mark_arc(innerRadius=50).encode(
        theta=alt.Theta(field=Y, type="quantitative"),
        color=alt.Color(field=X, type="nominal",legend=alt.Legend(orient="bottom",title = "",padding= 20)),
        tooltip = [alt.Tooltip(Y, title="",format='.1f')]
        )
        st.altair_chart(pie, use_container_width=True)

    def delete_temp():
        os.remove("temp/shipment_reco.xlsx")

    def file_upload_form():
        with st.form(key = 'ticker'):
            text, upload = st.columns([2.5,3]) 
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5 style="color:#6F8FAF">{"Upload Shipment Instruction:"}</h5>', unsafe_allow_html=True)
            with upload:
                shipment_instructions = st.file_uploader("",key = 'ship_ins')

            text, upload = st.columns([2.5,3])
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5 style="color:#6F8FAF">{"Upload Warehouse Reports:"}<h5>', unsafe_allow_html=True)
            with upload:
                warehouse_reports = st.file_uploader("",key = 'ware_rep', accept_multiple_files=True)

            text, upload = st.columns([2.5,3])
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5 style="color:#6F8FAF"> {"Upload Inventory Ledger:"}<h5>', unsafe_allow_html=True)
            with upload:
                inventory_ledger = st.file_uploader("",key = 'inv_led')
            
            a,b,button = st.columns([2,2,1.5]) 
            with button:
                st.write('###')
                submit = st.form_submit_button(label = "Start Reconciliation")
                #submit = st.button(label="Start Reconciliation")
        return shipment_instructions, warehouse_reports, inventory_ledger, submit
        

        

    landing_page()