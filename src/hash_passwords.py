import streamlit_authenticator as stauth

passwords = ["gey0YLITzctkvldPNIUsKQ==", "manager123"]
hashed_passwords = stauth.Hasher(passwords).generate()
print(hashed_passwords)