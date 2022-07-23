import streamlit as st
import hydralit_components as hc
from page_config import page_setup
import pandas as pd
from shipment_reco_charts import reconcile
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

page_setup()

authentication_status = login_status()

if authentication_status == False:
    st.error("Username/Password is incorrect")

if authentication_status:
    st.text('hi')