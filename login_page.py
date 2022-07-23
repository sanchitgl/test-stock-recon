import streamlit_authenticator as stauth
import streamlit as st
import yaml

def login_status():
    with open('config.yaml') as file:
        config = yaml.safe_load(file)

    placeholder = st.empty()
    
    authenticator = stauth.Authenticate(
        config['credentials'],
        config['cookie']['name'],
        config['cookie']['key'],
        config['cookie']['expiry_days'],
        config['preauthorized']
    )


    #authenticator = stauth.Authenticate(names, usernames, hashed_passwords, "ship_recon", "admin")
    with placeholder.container():
        space, login, space = st.columns([1,3,1])
        with login:
            name, authentication_status, username = authenticator.login('Login', 'main')

    placeholder.empty()

    return authentication_status
