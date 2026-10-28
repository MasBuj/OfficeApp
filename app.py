import streamlit as st
import requests
import urllib.parse
from google.oauth2 import id_token
from google.auth.transport import requests as google_requests

st.set_page_config(page_title="Portal TND BPKP Papua", page_icon="📜", layout="wide")

# 1. SETUP WHITELIST EMAILS
WHITELIST_EMAILS = [
    "mi.ramadhani123@gmail.com",  
    "afri.rodhiatul@bpkp.go.id"
]

# 2. AMBIL KUNCI (dengan .strip() untuk hapus spasi nyasar)
client_id = st.secrets["google_oauth"]["client_id"].strip()
client_secret = st.secrets["google_oauth"]["client_secret"].strip()
redirect_uri = st.secrets["google_oauth"]["redirect_uri"].strip()

# 3. INISIALISASI SESSION STATE
if "user_email" not in st.session_state:
    st.session_state.user_email = None

# 4. PROSES KODE DARI GOOGLE (STATELESS REST API)
if "code" in st.query_params and st.session_state.user_email is None:
    code = st.query_params["code"]

    # Penangkal Double-Fire Streamlit
    if st.session_state.get("last_code") == code:
        st.stop()
    st.session_state.last_code = code

    # Tukar kode dengan tiket akses secara manual
    token_url = "https://oauth2.googleapis.com/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "code": code,
        "grant_type": "authorization_code",
        "redirect_uri": redirect_uri
    }

    response = requests.post(token_url, data=payload)

    if response.status_code == 200:
        # Sukses!
        tokens = response.json()
        token_id = tokens.get("id_token")

        try:
            # Buka brankas token untuk ambil email
            request = google_requests.Request()
            id_info = id_token.verify_oauth2_token(token_id, request, client_id)
            st.session_state.user_email = id_info.get("email")

            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"Gagal membaca identitas: {e}")
    else:
        # Jika Google nolak, tampilkan alasan aslinya!
        st.error(f"Google menolak akses: {response.text}")
        st.stop()

# 5. HALAMAN LOGIN & VALIDASI WHITELIST
if st.session_state.user_email is None:
    st.title("Portal TND BPKP Papua")
    st.warning("Aplikasi ini dilindungi. Silakan login menggunakan email BPKP yang terdaftar.")

    # Generate URL Login manual
    auth_url = f"https://accounts.google.com/o/oauth2/v2/auth?client_id={client_id}&redirect_uri={urllib.parse.quote(redirect_uri)}&response_type=code&scope=openid%20email&prompt=consent"

    st.link_button("Login with Google", auth_url)

else:
    email = st.session_state.user_email
    if email in WHITELIST_EMAILS:
        col1, col2 = st.columns([0.8, 0.2])
        with col1:
            st.success(f"Login berhasil. Selamat datang, {email}")
        with col2:
            if st.button("Logout"):
                st.session_state.user_email = None
                st.rerun()

        st.divider()
        import tnd_generator
        pages = {
            "Navigasi TND": [
                st.Page("tnd_generator.py", title="Surat Penegasan (SP2)", icon="✍️"),
            ]
        }
        st.navigation(pages).run()
    else:
        st.error(f"AKSES DITOLAK: Email '{email}' tidak terdaftar dalam whitelist sistem.")
        if st.button("Kembali / Ganti Akun"):
            st.session_state.user_email = None
            st.rerun()
