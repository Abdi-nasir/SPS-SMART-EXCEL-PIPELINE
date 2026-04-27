import streamlit_authenticator as stauth

passwords = ["admin123", "manager123"]
hashed_passwords = stauth.Hasher(passwords).generate()
print(hashed_passwords)