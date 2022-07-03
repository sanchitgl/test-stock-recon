import streamlit as st
import streamlit_authenticator as stauth
import pickle 
from pathlib import Path 
import yaml

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
    st.title('this is the home page')